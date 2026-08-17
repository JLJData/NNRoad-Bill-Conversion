# -*- coding: utf-8 -*-
"""
账单转换 HTTP 薄服务（供 Office 调用）

启动:
  pip install -r requirements.txt
  # 建议仅本机：uvicorn convert_api:app --host 127.0.0.1 --port 8765
  # 环境变量 CONVERT_API_KEY：非空则除 /health 外须带请求头 X-Api-Key
  # CONVERT_DISABLE_DOCS=1 或已设 API Key 时关闭 /docs

接口:
  GET  /health
  GET  /engines
  GET  /pdf-profiles
  GET  /mapping/defaults?engineId=&pdfProfileId=  引擎默认映射（含 fixedValueWrites）
  POST /convert  multipart: file, engine_id, region, template(可选), pn_meta(json可选), employee_directory(json数组可选)
  POST /pdf-to-source  multipart: file, profile_id(可选自动识别), pn_meta(json可选), template(可选)
  POST /pdf-to-source-batch  multipart: files[], profile_id, …
  POST /vendor-to-source-batch  multipart: files[](pdf/xlsx), profile_id, …  # 按扩展名自动分流
  GET  /region-template?region=Taiwan  地区默认 PN 母版
  POST /excel-snapshot  multipart: file, sheet(可选默认PN), max_cells(可选默认300)
  POST /hf-snapshot     multipart: file, sheet(可选默认PN), max_cells(可选默认300)  # Node HyperFormula
"""
from __future__ import annotations

import base64
import json
import os
import tempfile
import traceback
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, Query, Request, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from starlette.background import BackgroundTask
from starlette.responses import Response

from convert_mapping import resolve_convert_mapping
from mapping_inspect import inspect_pn_headers, inspect_source_headers
from convert_runner import parse_employee_directory_payload, parse_pn_meta_payload, parse_convert_mapping_payload, run_convert
from excel_com_snapshot import snapshot_workbook
from hf_com_snapshot import snapshot_workbook_hf
from engines import list_engines
from pdf_ingest.registry import list_pdf_profiles
from pdf_ingest.runner import run_pdf_to_source, run_pdf_to_source_batch, run_vendor_to_source_batch
from region_templates import list_regions, get_region_template

CONVERT_API_KEY = os.environ.get("CONVERT_API_KEY", "").strip()
_DISABLE_DOCS = os.environ.get("CONVERT_DISABLE_DOCS", "").strip() in ("1", "true", "True", "yes")
_DOCS_OFF = _DISABLE_DOCS or bool(CONVERT_API_KEY)
_BLOCKED_UPLOAD_SUFFIXES = {".html", ".htm", ".shtml", ".xhtml", ".svg"}

app = FastAPI(
    title="HROne Bill Convert Service",
    version="1.0.0",
    docs_url=None if _DOCS_OFF else "/docs",
    redoc_url=None if _DOCS_OFF else "/redoc",
    openapi_url=None if _DOCS_OFF else "/openapi.json",
)
BASE_DIR = Path(__file__).resolve().parent


def _b64_json_header(payload: object) -> str:
    raw = json.dumps(payload, ensure_ascii=True, separators=(",", ":")).encode("utf-8")
    return base64.b64encode(raw).decode("ascii")


def _assert_safe_upload(upload: UploadFile | None) -> None:
    if upload is None or not upload.filename:
        return
    suffix = Path(upload.filename).suffix.lower()
    if suffix in _BLOCKED_UPLOAD_SUFFIXES:
        raise HTTPException(status_code=400, detail=f"blocked file type: {suffix}")


@app.middleware("http")
async def require_api_key(request: Request, call_next) -> Response:
    path = request.url.path
    if path == "/health" or path.startswith("/health/"):
        return await call_next(request)
    if not CONVERT_API_KEY:
        return await call_next(request)
    key = request.headers.get("X-Api-Key", "").strip()
    if key != CONVERT_API_KEY:
        return JSONResponse(status_code=401, content={"ok": False, "msg": "unauthorized"})
    return await call_next(request)


@app.get("/health")
def health():
    return {"ok": True, "service": "bill-convert", "auth": bool(CONVERT_API_KEY)}


@app.get("/engines")
def engines():
    pdfs = list_pdf_profiles()
    return {
        "engines": [
            {
                "engineId": e.engine_id,
                "label": e.label,
                "module": e.module,
                "description": e.description,
                "pdfProfiles": [
                    {
                        "profileId": p.profile_id,
                        "label": p.label,
                        "region": p.region,
                        "description": p.description,
                    }
                    for p in pdfs
                    if e.engine_id in p.engine_ids
                ],
            }
            for e in list_engines()
        ],
        "regions": list_regions(),
    }


@app.get("/mapping/defaults")
def mapping_defaults(
    engineId: str | None = Query(None),
    engine_id: str | None = Query(None),
    pdfProfileId: str | None = Query(None),
    pdf_profile_id: str | None = Query(None),
):
    eid = (engineId or engine_id or "").strip()
    if not eid:
        raise HTTPException(status_code=400, detail="engineId required")
    pid = (pdfProfileId or pdf_profile_id or "").strip()
    raw: dict = {}
    if pid:
        raw["pdfProfileId"] = pid
    mapping = resolve_convert_mapping(eid, raw)
    return {"engineId": eid, "pdfProfileId": pid or None, "mapping": mapping}


@app.get("/pdf-profiles")
def pdf_profiles():
    return {
        "profiles": [
            {
                "profileId": p.profile_id,
                "label": p.label,
                "region": p.region,
                "module": p.module,
                "description": p.description,
                "engineIds": list(p.engine_ids),
            }
            for p in list_pdf_profiles()
        ]
    }


@app.post("/pdf-to-source")
async def pdf_to_source(
    file: UploadFile = File(...),
    profile_id: str | None = Form(None),
    pn_meta: str | None = Form(None),
    template: UploadFile | None = File(None),
):
    """供应商 PDF → 地区源 Excel（当前：eor_uk → UK-L）。"""
    _assert_safe_upload(file)
    _assert_safe_upload(template)
    suffix = Path(file.filename or "bill.pdf").suffix.lower() or ".pdf"
    if suffix != ".pdf":
        raise HTTPException(status_code=400, detail=f"仅支持 .pdf，当前: {suffix}")

    tmp_dir = Path(tempfile.mkdtemp(prefix="pdf2src_"))
    pdf_path = tmp_dir / f"source{suffix}"
    out_path = tmp_dir / "source_from_pdf.xlsx"
    template_path: Path | None = None
    try:
        content = await file.read()
        if not content:
            raise HTTPException(status_code=400, detail="上传文件为空")
        pdf_path.write_bytes(content)

        if template is not None and template.filename:
            tpl_suffix = Path(template.filename).suffix.lower() or ".xlsx"
            if tpl_suffix not in (".xlsx", ".xlsm"):
                raise HTTPException(status_code=400, detail="母版仅支持 .xlsx/.xlsm")
            tpl_bytes = await template.read()
            if tpl_bytes:
                template_path = tmp_dir / f"template{tpl_suffix}"
                template_path.write_bytes(tpl_bytes)

        try:
            meta = parse_pn_meta_payload(pn_meta)
        except Exception as exc:
            raise HTTPException(status_code=400, detail=f"pn_meta 无效: {exc}") from exc

        result = run_pdf_to_source(
            pdf_path,
            out_path,
            profile_id=(profile_id or "").strip() or None,
            template_path=template_path,
            pn_meta=meta,
            registry_dir=tmp_dir,
            fill_fx=True,
        )
        headers = {
            "X-Pdf-Profile": str(result.get("profile_id") or ""),
            "X-Pdf-Region": str(result.get("region") or ""),
            "X-Pdf-Warnings": str(len(result.get("warnings") or [])),
        }
        for w in result.get("warnings") or []:
            print(f"[pdf-warning] {w}")
        return FileResponse(
            path=str(out_path),
            filename=out_path.name,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers,
            background=BackgroundTask(_cleanup_dir, tmp_dir),
        )
    except HTTPException:
        _cleanup_dir(tmp_dir)
        raise
    except KeyError as exc:
        _cleanup_dir(tmp_dir)
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except ValueError as exc:
        _cleanup_dir(tmp_dir)
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:
        _cleanup_dir(tmp_dir)
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"PDF 转换失败: {exc}") from exc


@app.post("/pdf-to-source-batch")
async def pdf_to_source_batch(
    files: list[UploadFile] = File(...),
    profile_id: str | None = Form(None),
    pn_meta: str | None = Form(None),
    template: UploadFile | None = File(None),
):
    """多份供应商 PDF → 一份地区源 Excel（如 TopSource 一人一票）。"""
    for f in files:
        _assert_safe_upload(f)
    _assert_safe_upload(template)
    if not files:
        raise HTTPException(status_code=400, detail="请至少上传一个 PDF")

    tmp_dir = Path(tempfile.mkdtemp(prefix="pdf2src_batch_"))
    out_path = tmp_dir / "source_from_pdf.xlsx"
    template_path: Path | None = None
    pdf_paths: list[Path] = []
    try:
        for i, f in enumerate(files):
            suffix = Path(f.filename or f"bill_{i}.pdf").suffix.lower() or ".pdf"
            if suffix != ".pdf":
                raise HTTPException(status_code=400, detail=f"仅支持 .pdf，当前: {f.filename}")
            content = await f.read()
            if not content:
                raise HTTPException(status_code=400, detail=f"上传文件为空: {f.filename}")
            pdf_path = tmp_dir / f"source_{i}{suffix}"
            pdf_path.write_bytes(content)
            pdf_paths.append(pdf_path)

        if template is not None and template.filename:
            tpl_suffix = Path(template.filename).suffix.lower() or ".xlsx"
            if tpl_suffix not in (".xlsx", ".xlsm"):
                raise HTTPException(status_code=400, detail="母版仅支持 .xlsx/.xlsm")
            tpl_bytes = await template.read()
            if tpl_bytes:
                template_path = tmp_dir / f"template{tpl_suffix}"
                template_path.write_bytes(tpl_bytes)

        try:
            meta = parse_pn_meta_payload(pn_meta)
        except Exception as exc:
            raise HTTPException(status_code=400, detail=f"pn_meta 无效: {exc}") from exc

        result = run_pdf_to_source_batch(
            pdf_paths,
            out_path,
            profile_id=(profile_id or "").strip() or None,
            template_path=template_path,
            pn_meta=meta,
            registry_dir=tmp_dir,
            fill_fx=True,
        )
        headers = {
            "X-Pdf-Profile": str(result.get("profile_id") or ""),
            "X-Pdf-Region": str(result.get("region") or ""),
            "X-Pdf-Warnings": str(len(result.get("warnings") or [])),
            "X-Pdf-Employees": str(result.get("employee_count") or len(pdf_paths)),
        }
        for w in result.get("warnings") or []:
            print(f"[pdf-warning] {w}")
        return FileResponse(
            path=str(out_path),
            filename=out_path.name,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers,
            background=BackgroundTask(_cleanup_dir, tmp_dir),
        )
    except HTTPException:
        _cleanup_dir(tmp_dir)
        raise
    except KeyError as exc:
        _cleanup_dir(tmp_dir)
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except ValueError as exc:
        _cleanup_dir(tmp_dir)
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:
        _cleanup_dir(tmp_dir)
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"PDF 批量转换失败: {exc}") from exc


@app.post("/vendor-plugins/ingest-file")
async def vendor_plugins_ingest_file(
    file: UploadFile = File(...),
    profile_id: str | None = Form(None),
):
    """
    供应商旁路文件识别（上传时调用）：如 Auxilium Admin Fee PDF → Total VAT 事实。
    不要求与 Payroll Draft 同批。
    """
    from bill_convert.vendor_plugins.registry import get_plugins_for_profile

    _assert_safe_upload(file)
    tmp_dir = Path(tempfile.mkdtemp(prefix="vendor_ingest_"))
    try:
        suffix = Path(file.filename or "file.bin").suffix.lower() or ""
        content = await file.read()
        if not content:
            raise HTTPException(status_code=400, detail="上传文件为空")
        path = tmp_dir / f"ingest{suffix}"
        path.write_bytes(content)
        pid = (profile_id or "").strip() or None
        plugins = get_plugins_for_profile(pid)
        if not plugins:
            return {"ok": True, "matched": False, "facts": {}, "message": "无匹配插件"}
        for plugin in plugins:
            try:
                if not plugin.classify_path(path):
                    continue
                facts = plugin.parse_artifacts([path]) or {}
                facts.pop("_warnings", None)
                # 上传识别：写入 latest 键，供后续转换当 curr
                if "auxilium.admin_fee.total_vat" in facts:
                    facts["auxilium.admin_fee.latest_vat"] = facts["auxilium.admin_fee.total_vat"]
                return {
                    "ok": True,
                    "matched": True,
                    "plugin_id": plugin.plugin_id,
                    "facts": facts,
                    "source_file": file.filename,
                }
            except Exception as exc:
                raise HTTPException(status_code=400, detail=f"{plugin.plugin_id}: {exc}") from exc
        return {"ok": True, "matched": False, "facts": {}, "message": "未识别为旁路文件"}
    except HTTPException:
        raise
    except Exception as exc:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"旁路文件识别失败: {exc}") from exc
    finally:
        _cleanup_dir(tmp_dir)


@app.post("/vendor-to-source-batch")
async def vendor_to_source_batch(
    files: list[UploadFile] = File(...),
    profile_id: str | None = Form(None),
    pn_meta: str | None = Form(None),
    convert_mapping: str | None = Form(None),
    template: UploadFile | None = File(None),
):
    """供应商 PDF/Excel → 一份地区源表（按扩展名自动走 PDF 或 Excel 解析）。"""
    for f in files:
        _assert_safe_upload(f)
    _assert_safe_upload(template)
    if not files:
        raise HTTPException(status_code=400, detail="请至少上传一个源文件")

    tmp_dir = Path(tempfile.mkdtemp(prefix="vendor2src_"))
    out_path = tmp_dir / "source_from_vendor.xlsx"
    template_path: Path | None = None
    source_paths: list[Path] = []
    try:
        for i, f in enumerate(files):
            suffix = Path(f.filename or f"bill_{i}.bin").suffix.lower() or ""
            if suffix not in (".pdf", ".xlsx", ".xlsm", ".xls"):
                raise HTTPException(
                    status_code=400,
                    detail=f"仅支持 .pdf / .xlsx / .xlsm，当前: {f.filename}",
                )
            content = await f.read()
            if not content:
                raise HTTPException(status_code=400, detail=f"上传文件为空: {f.filename}")
            path = tmp_dir / f"source_{i}{suffix}"
            path.write_bytes(content)
            source_paths.append(path)

        if template is not None and template.filename:
            tpl_suffix = Path(template.filename).suffix.lower() or ".xlsx"
            if tpl_suffix not in (".xlsx", ".xlsm"):
                raise HTTPException(status_code=400, detail="母版仅支持 .xlsx/.xlsm")
            tpl_bytes = await template.read()
            if tpl_bytes:
                template_path = tmp_dir / f"template{tpl_suffix}"
                template_path.write_bytes(tpl_bytes)

        try:
            meta = parse_pn_meta_payload(pn_meta)
        except Exception as exc:
            raise HTTPException(status_code=400, detail=f"pn_meta 无效: {exc}") from exc
        try:
            mapping = parse_convert_mapping_payload(convert_mapping)
        except Exception as exc:
            raise HTTPException(status_code=400, detail=f"convert_mapping 无效: {exc}") from exc

        result = run_vendor_to_source_batch(
            source_paths,
            out_path,
            profile_id=(profile_id or "").strip() or None,
            template_path=template_path,
            pn_meta=meta,
            registry_dir=tmp_dir,
            fill_fx=True,
            convert_mapping=mapping,
        )
        headers = {
            "X-Pdf-Profile": str(result.get("profile_id") or ""),
            "X-Pdf-Region": str(result.get("region") or ""),
            "X-Source-Kind": str(result.get("source_kind") or ""),
            "X-Pdf-Warnings": str(len(result.get("warnings") or [])),
            "X-Pdf-Employees": str(result.get("employee_count") or len(source_paths)),
        }
        artifact_facts = result.get("artifact_facts")
        if isinstance(artifact_facts, dict) and artifact_facts:
            headers["X-Vendor-Artifact-Facts"] = _b64_json_header(artifact_facts)
        for w in result.get("warnings") or []:
            print(f"[vendor-warning] {w}")
        return FileResponse(
            path=str(out_path),
            filename=out_path.name,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers,
            background=BackgroundTask(_cleanup_dir, tmp_dir),
        )
    except HTTPException:
        _cleanup_dir(tmp_dir)
        raise
    except KeyError as exc:
        _cleanup_dir(tmp_dir)
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except ValueError as exc:
        _cleanup_dir(tmp_dir)
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:
        _cleanup_dir(tmp_dir)
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"供应商源批量转换失败: {exc}") from exc


@app.get("/region-template")
def region_template(region: str = Query(..., min_length=1)):
    try:
        path = get_region_template(region.strip())
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    if not path.is_file():
        raise HTTPException(status_code=404, detail=f"地区母版不存在: {path}")
    return FileResponse(
        path=str(path),
        filename=path.name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.post("/convert")
async def convert(
    file: UploadFile = File(...),
    engine_id: str = Form(...),
    region: str = Form(...),
    template: UploadFile | None = File(None),
    pn_meta: str | None = Form(None),
    employee_directory: str | None = Form(None),
    convert_mapping: str | None = Form(None),
    output_prefix: str | None = Form(None),
):
    _assert_safe_upload(file)
    _assert_safe_upload(template)
    suffix = Path(file.filename or "source.xlsx").suffix.lower() or ".xlsx"
    if suffix not in (".xlsx", ".xlsm"):
        raise HTTPException(status_code=400, detail=f"仅支持 Excel（.xlsx/.xlsm），当前: {suffix}")

    try:
        meta = parse_pn_meta_payload(pn_meta)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"pn_meta 无效: {exc}") from exc
    try:
        emp_dir = parse_employee_directory_payload(employee_directory)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"employee_directory 无效: {exc}") from exc
    try:
        mapping = parse_convert_mapping_payload(convert_mapping)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"convert_mapping 无效: {exc}") from exc

    tmp_dir = Path(tempfile.mkdtemp(prefix="bill_convert_"))
    source_path = tmp_dir / f"source{suffix}"
    prefix = (output_prefix or "PN_auto").strip() or "PN_auto"
    output_path = tmp_dir / f"{prefix}.xlsx"

    try:
        content = await file.read()
        if not content:
            raise HTTPException(status_code=400, detail="上传文件为空")
        source_path.write_bytes(content)

        template_path = None
        if template is not None and template.filename:
            tpl_suffix = Path(template.filename).suffix.lower() or ".xlsx"
            if tpl_suffix not in (".xlsx", ".xlsm"):
                raise HTTPException(status_code=400, detail=f"母版仅支持 .xlsx/.xlsm，当前: {tpl_suffix}")
            tpl_bytes = await template.read()
            if tpl_bytes:
                template_path = tmp_dir / f"template{tpl_suffix}"
                template_path.write_bytes(tpl_bytes)

        result = run_convert(
            engine_id=engine_id.strip(),
            source_path=source_path,
            output_path=output_path,
            region=region.strip(),
            template_path=template_path,
            pn_meta=meta,
            employee_directory=emp_dir,
            convert_mapping=mapping,
            registry_dir=BASE_DIR,
        )
        if not output_path.is_file():
            raise RuntimeError("转换完成但未生成输出文件")

        headers = {
            "X-Convert-Engine": engine_id,
            "X-Convert-Region": region,
            "X-Convert-Employees": str(result.get("employee_count") or 0),
        }
        warnings = result.get("warnings") or []
        if warnings:
            # 响应头避免非 ASCII；条数提示即可，详情在服务日志
            headers["X-Convert-Warnings"] = str(len(warnings))
            for w in warnings:
                print(f"[convert-warning] {w}")
        # ASCII 诊断：映射样式条数、每人实际套用的 China 示例行
        if result.get("mapping_style_count") is not None:
            headers["X-Convert-Mapping-Styles"] = str(int(result.get("mapping_style_count") or 0))
        formula_rows = str(result.get("formula_main_rows") or "").strip()
        if formula_rows:
            headers["X-Convert-Formula-Rows"] = formula_rows[:200]
        match_hint = str(result.get("formula_match_hint") or "").strip()
        if match_hint:
            headers["X-Convert-Formula-Match"] = match_hint[:64]
        fact_updates = result.get("fact_store_updates")
        if isinstance(fact_updates, dict) and fact_updates:
            headers["X-Convert-Fact-Store"] = _b64_json_header(fact_updates)
        # 结果摘要用自定义头传一小段 JSON（可选）；主体仍是文件
        return FileResponse(
            path=str(output_path),
            filename=output_path.name,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers,
            background=BackgroundTask(_cleanup_dir, tmp_dir),
        )
    except HTTPException:
        _cleanup_dir(tmp_dir)
        raise
    except KeyError as exc:
        _cleanup_dir(tmp_dir)
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:
        _cleanup_dir(tmp_dir)
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"转换失败: {exc}") from exc


def _cleanup_dir(path: Path) -> None:
    try:
        import shutil

        shutil.rmtree(path, ignore_errors=True)
    except Exception:
        pass


@app.post("/excel-snapshot")
async def excel_snapshot(
    file: UploadFile = File(...),
    sheet: str | None = Form("PN"),
    max_cells: int = Form(300),
):
    """
    本机桌面 Excel COM 重算后抽取公式格结果（模板发布业务基准）。
    仅 Windows + 已安装 Excel + pywin32。
    """
    _assert_safe_upload(file)
    suffix = Path(file.filename or "source.xlsx").suffix.lower() or ".xlsx"
    if suffix not in (".xlsx", ".xlsm", ".xls"):
        raise HTTPException(status_code=400, detail=f"仅支持 Excel，当前: {suffix}")

    tmp_dir = Path(tempfile.mkdtemp(prefix="excel_snap_"))
    source_path = tmp_dir / f"source{suffix}"
    try:
        content = await file.read()
        if not content:
            raise HTTPException(status_code=400, detail="上传文件为空")
        source_path.write_bytes(content)
        sheet_filter = None if sheet is None or str(sheet).strip() == "" else str(sheet).strip()
        result = snapshot_workbook(source_path, sheet_filter, max(1, min(int(max_cells or 300), 2000)))
        if not result.get("ok"):
            # 业务失败也返回 200 + ok=false，方便前端展示 message
            return JSONResponse(content=result)
        return JSONResponse(content=result)
    except HTTPException:
        raise
    except Exception as exc:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Excel 快照失败: {exc}") from exc
    finally:
        _cleanup_dir(tmp_dir)


@app.post("/hf-snapshot")
async def hf_snapshot(
    file: UploadFile = File(...),
    sheet: str | None = Form("PN"),
    max_cells: int = Form(300),
):
    """Node HyperFormula 重算快照（模板三引擎 / 随机压测）。"""
    _assert_safe_upload(file)
    suffix = Path(file.filename or "source.xlsx").suffix.lower() or ".xlsx"
    if suffix not in (".xlsx", ".xlsm", ".xls"):
        raise HTTPException(status_code=400, detail=f"仅支持 Excel，当前: {suffix}")

    tmp_dir = Path(tempfile.mkdtemp(prefix="hf_snap_"))
    source_path = tmp_dir / f"source{suffix}"
    try:
        content = await file.read()
        if not content:
            raise HTTPException(status_code=400, detail="上传文件为空")
        source_path.write_bytes(content)
        sheet_filter = None if sheet is None or str(sheet).strip() == "" else str(sheet).strip()
        result = snapshot_workbook_hf(source_path, sheet_filter, max(1, min(int(max_cells or 300), 2000)))
        return JSONResponse(content=result)
    except HTTPException:
        raise
    except Exception as exc:
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"HF 快照失败: {exc}") from exc
    finally:
        _cleanup_dir(tmp_dir)


@app.post("/mapping/inspect-source")
async def mapping_inspect_source(
    file: UploadFile = File(...),
    engine_id: str = Form(...),
    convert_mapping: str | None = Form(None),
):
    """上传样例源账单，按当前映射识别表头（供下拉）。"""
    _assert_safe_upload(file)
    suffix = Path(file.filename or "source.xlsx").suffix.lower() or ".xlsx"
    if suffix not in (".xlsx", ".xlsm"):
        raise HTTPException(status_code=400, detail=f"仅支持 .xlsx/.xlsm，当前: {suffix}")
    try:
        mapping = parse_convert_mapping_payload(convert_mapping)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"convert_mapping 无效: {exc}") from exc

    tmp_dir = Path(tempfile.mkdtemp(prefix="map_insp_"))
    source_path = tmp_dir / f"source{suffix}"
    try:
        content = await file.read()
        if not content:
            raise HTTPException(status_code=400, detail="上传文件为空")
        source_path.write_bytes(content)
        result = inspect_source_headers(
            source_path=source_path,
            engine_id=engine_id.strip(),
            convert_mapping=mapping,
        )
        return JSONResponse(content=result)
    finally:
        _cleanup_dir(tmp_dir)


@app.post("/mapping/inspect-pn")
async def mapping_inspect_pn(
    engine_id: str = Form(...),
    convert_mapping: str | None = Form(None),
    region: str | None = Form(None),
    template: UploadFile | None = File(None),
):
    """解析 PN 母版 targetL 表头行（生效母版或地区默认）。"""
    _assert_safe_upload(template)
    try:
        mapping = parse_convert_mapping_payload(convert_mapping)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"convert_mapping 无效: {exc}") from exc

    tmp_dir = Path(tempfile.mkdtemp(prefix="map_pn_"))
    try:
        template_path: Path | None = None
        if template is not None and template.filename:
            tpl_suffix = Path(template.filename).suffix.lower() or ".xlsx"
            if tpl_suffix not in (".xlsx", ".xlsm"):
                raise HTTPException(status_code=400, detail=f"母版仅支持 .xlsx/.xlsm")
            tpl_bytes = await template.read()
            if tpl_bytes:
                template_path = tmp_dir / f"template{tpl_suffix}"
                template_path.write_bytes(tpl_bytes)
        if template_path is None:
            if not region or not str(region).strip():
                raise HTTPException(status_code=400, detail="未上传母版时必须提供 region")
            template_path = Path(get_region_template(region.strip()))
        result = inspect_pn_headers(
            template_path=template_path,
            engine_id=engine_id.strip(),
            convert_mapping=mapping,
        )
        return JSONResponse(content=result)
    finally:
        _cleanup_dir(tmp_dir)


@app.exception_handler(HTTPException)
async def http_exc_handler(_, exc: HTTPException):
    return JSONResponse(status_code=exc.status_code, content={"ok": False, "msg": exc.detail})


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("convert_api:app", host="0.0.0.0", port=8765, reload=False)
