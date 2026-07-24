# -*- coding: utf-8 -*-
"""
账单转换 HTTP 薄服务（供 Office 调用）

启动:
  pip install -r requirements.txt
  python convert_api.py
  # 或: uvicorn convert_api:app --host 0.0.0.0 --port 8765

接口:
  GET  /health
  GET  /engines
  POST /convert  multipart: file, engine_id, region, pn_meta(json可选), employee_directory(json数组可选)
"""
from __future__ import annotations

import tempfile
import traceback
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from starlette.background import BackgroundTask

from convert_runner import parse_employee_directory_payload, parse_pn_meta_payload, run_convert
from engines import list_engines
from region_templates import list_regions

app = FastAPI(title="HROne Bill Convert Service", version="1.0.0")
BASE_DIR = Path(__file__).resolve().parent


@app.get("/health")
def health():
    return {"ok": True, "service": "bill-convert"}


@app.get("/engines")
def engines():
    return {
        "engines": [
            {
                "engineId": e.engine_id,
                "label": e.label,
                "module": e.module,
                "description": e.description,
            }
            for e in list_engines()
        ],
        "regions": list_regions(),
    }


@app.post("/convert")
async def convert(
    file: UploadFile = File(...),
    engine_id: str = Form(...),
    region: str = Form(...),
    pn_meta: str | None = Form(None),
    employee_directory: str | None = Form(None),
    output_prefix: str | None = Form(None),
):
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

    tmp_dir = Path(tempfile.mkdtemp(prefix="bill_convert_"))
    source_path = tmp_dir / f"source{suffix}"
    prefix = (output_prefix or "PN_auto").strip() or "PN_auto"
    output_path = tmp_dir / f"{prefix}.xlsx"

    try:
        content = await file.read()
        if not content:
            raise HTTPException(status_code=400, detail="上传文件为空")
        source_path.write_bytes(content)

        result = run_convert(
            engine_id=engine_id.strip(),
            source_path=source_path,
            output_path=output_path,
            region=region.strip(),
            pn_meta=meta,
            employee_directory=emp_dir,
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


@app.exception_handler(HTTPException)
async def http_exc_handler(_, exc: HTTPException):
    return JSONResponse(status_code=exc.status_code, content={"ok": False, "msg": exc.detail})


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("convert_api:app", host="0.0.0.0", port=8765, reload=False)
