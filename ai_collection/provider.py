"""Provider abstraction and the OpenAI Responses API adapter."""
from __future__ import annotations

import base64
import copy
import hashlib
import json
import mimetypes
import os
import re
import ssl
import time
import urllib.error
import urllib.request
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Callable

from openpyxl import load_workbook

from bill_validation.contracts import ValidationError, nonempty, require

from .contracts import validate_ai_collection, validate_collection_request


class AIProviderError(RuntimeError):
    """Provider transport/model failure, distinct from business disagreement."""


class AIProvider(ABC):
    @property
    @abstractmethod
    def provider_id(self) -> str:
        raise NotImplementedError

    @abstractmethod
    def collect(self, request: dict) -> dict:
        raise NotImplementedError


class MockAIProvider(AIProvider):
    """Deterministic test provider. It performs no network or file access."""

    def __init__(self, response: dict | Callable[[dict], dict], provider_id: str = "mock"):
        require(nonempty(provider_id), "Mock provider id is required")
        self._response = response
        self._provider_id = provider_id
        self.calls = []

    @property
    def provider_id(self) -> str:
        return self._provider_id

    def collect(self, request: dict) -> dict:
        validate_collection_request(request)
        safe_request = copy.deepcopy(request)
        self.calls.append(safe_request)
        try:
            result = self._response(copy.deepcopy(safe_request)) if callable(self._response) else self._response
            return copy.deepcopy(result)
        except AIProviderError:
            raise
        except Exception as exc:
            raise AIProviderError(str(exc)) from exc


def _strict_object(properties: dict, required: list[str] | None = None) -> dict:
    return {
        "type": "object",
        "properties": properties,
        "required": required if required is not None else list(properties),
        "additionalProperties": False,
    }


def _collection_output_schema(request: dict, model: str) -> dict:
    field_names = [item["field"] for item in request["schema"]["fields"]]
    file_ids = [item["fileId"] for item in request["documents"]]
    source = _strict_object({
        "fileId": {"type": "string", "enum": file_ids},
        "location": {"type": "string"},
        "page": {"type": ["integer", "null"], "minimum": 1},
        "rawText": {"type": "string"},
    })
    field = _strict_object({
        "field": {"type": "string", "enum": field_names},
        "status": {"type": "string", "enum": ["found", "missing", "unreadable", "not_applicable"]},
        "value": {"type": ["string", "null"]},
        "currency": {"type": "string", "enum": [request["currency"]]},
        "confidence": {"type": "number", "minimum": 0, "maximum": 1},
        "source": source,
    })
    record = _strict_object({
        "sourceEntity": _strict_object({
            "employeeId": {"type": ["string", "null"]},
            "name": {"type": "string"},
        }),
        "fields": {"type": "array", "items": field},
    })
    unknown = _strict_object({
        "label": {"type": "string"},
        "value": {"type": ["string", "null"]},
        "currency": {"type": ["string", "null"]},
        "source": source,
    })
    issue = _strict_object({"code": {"type": "string"}, "message": {"type": "string"}})
    return _strict_object({
        "engine": {"type": "string", "enum": ["ai"]},
        "provider": {"type": "string", "enum": ["openai"]},
        "model": {"type": "string", "enum": [model]},
        "runId": {"type": "string", "enum": [request["runId"]]},
        "schemaId": {"type": "string", "enum": [request["schema"]["schemaId"]]},
        "schemaVersion": {"type": "string", "enum": [request["schema"]["schemaVersion"]]},
        "period": {"type": "string", "enum": [request["period"]]},
        "currency": {"type": "string", "enum": [request["currency"]]},
        "automaticPassEnabled": {"type": "boolean", "enum": [False]},
        "records": {"type": "array", "items": record},
        "unknownItems": {"type": "array", "items": unknown},
        "issues": {"type": "array", "items": issue},
        "readiness": {"type": "string", "enum": ["ready_for_association", "review_required"]},
    })


def _template_plan_output_schema(template_manifest: dict, collection: dict, schema: dict) -> dict:
    field_names = [item["field"] for item in schema["fields"]]
    write = _strict_object({
        "targetCell": {"type": "string", "pattern": "^[A-Z]{1,3}[1-9][0-9]*$"},
        "recordIndex": {"type": "integer", "minimum": 0, "maximum": max(0, len(collection["records"]) - 1)},
        "sourceKind": {"type": "string", "enum": ["entity_name", "field"]},
        "field": {"type": ["string", "null"], "enum": field_names + [None]},
    })
    issue = _strict_object({"code": {"type": "string"}, "message": {"type": "string"}})
    return _strict_object({
        "planVersion": {"type": "integer", "enum": [1]},
        "templateSha256": {"type": "string", "enum": [template_manifest["templateSha256"]]},
        "sheetName": {"type": "string", "enum": [template_manifest["sheetName"]]},
        "automaticWriteEnabled": {"type": "boolean", "enum": [False]},
        "writes": {"type": "array", "items": write},
        "issues": {"type": "array", "items": issue},
    })


def _dynamic_template_plan_output_schema(template_manifest: dict, documents: list[dict], model: str) -> dict:
    """Schema for template-driven extraction without a predeclared payroll field list."""
    file_ids = [item["fileId"] for item in documents]
    source = _strict_object({
        "fileId": {"type": "string", "enum": file_ids},
        "location": {"type": "string"},
        "page": {"type": ["integer", "null"], "minimum": 1},
        "rawText": {"type": "string"},
    })
    write = _strict_object({
        "targetCell": {"type": "string", "pattern": "^[A-Z]{1,3}[1-9][0-9]*$"},
        "semanticLabel": {"type": "string"},
        "sourceLabel": {"type": "string"},
        "valueType": {"type": "string", "enum": ["text", "decimal", "date"]},
        "value": {"type": "string"},
        "source": source,
        "confidence": {"type": "number", "minimum": 0, "maximum": 1},
    })
    issue = _strict_object({"code": {"type": "string"}, "message": {"type": "string"}})
    return _strict_object({
        "planVersion": {"type": "integer", "enum": [3]},
        "templateSha256": {"type": "string", "enum": [template_manifest["templateSha256"]]},
        "sheetName": {"type": "string", "enum": [template_manifest["sheetName"]]},
        "model": {"type": "string", "enum": [model]},
        "automaticWriteEnabled": {"type": "boolean", "enum": [False]},
        "writes": {"type": "array", "items": write},
        "issues": {"type": "array", "items": issue},
    })


def _openai_ssl_context() -> ssl.SSLContext:
    """Build a Windows-safe TLS client context (system CA path is often missing)."""
    try:
        import certifi
        context = ssl.create_default_context(cafile=certifi.where())
    except Exception:
        context = ssl.create_default_context()
    # Prefer modern TLS; some VPN/proxy middleboxes reset weaker handshakes mid-flight.
    if hasattr(ssl, "TLSVersion"):
        context.minimum_version = ssl.TLSVersion.TLSv1_2
    return context


def _is_transient_openai_transport_error(exc: BaseException) -> bool:
    text = str(exc).casefold()
    markers = (
        "eof occurred in violation of protocol",
        "connection reset",
        "connection aborted",
        "broken pipe",
        "timed out",
        "temporarily unavailable",
        "remote end closed connection",
        "server disconnected",
        "disconnected without sending",
        "ssl",
        "wrong version number",
    )
    return any(item in text for item in markers)


def _default_http_post(url: str, headers: dict[str, str], payload: dict, timeout_seconds: float) -> dict:
    body = json.dumps(payload).encode("utf-8")
    ssl_context = _openai_ssl_context()
    last_error: BaseException | None = None
    attempts = 3
    print(
        f"[ai] calling OpenAI ({len(body)} bytes, timeout={timeout_seconds}s)",
        flush=True,
    )
    for attempt in range(1, attempts + 1):
        started = time.time()
        try:
            print(f"[ai] attempt {attempt}/{attempts} …", flush=True)
            # httpx handles large TLS uploads more reliably than urllib on Windows.
            try:
                import certifi
                import httpx
                verify = certifi.where()
                with httpx.Client(timeout=timeout_seconds, verify=verify, http2=False) as client:
                    response = client.post(url, headers=headers, content=body)
                if response.status_code >= 400:
                    raise AIProviderError(
                        f"OpenAI HTTP {response.status_code}: {response.text[:1000]}"
                    )
                print(
                    f"[ai] OpenAI OK in {time.time() - started:.1f}s "
                    f"(HTTP {response.status_code})",
                    flush=True,
                )
                return response.json()
            except ImportError:
                request = urllib.request.Request(url, data=body, headers=headers, method="POST")
                with urllib.request.urlopen(
                    request, timeout=timeout_seconds, context=ssl_context,
                ) as response:
                    raw = response.read().decode("utf-8")
                print(f"[ai] OpenAI OK in {time.time() - started:.1f}s (urllib)", flush=True)
                return json.loads(raw)
        except AIProviderError:
            raise
        except urllib.error.HTTPError as exc:
            err_body = exc.read().decode("utf-8", errors="replace")[:1000]
            raise AIProviderError(f"OpenAI HTTP {exc.code}: {err_body}") from exc
        except (urllib.error.URLError, TimeoutError, OSError, json.JSONDecodeError, ssl.SSLError) as exc:
            last_error = exc
            print(f"[ai] attempt {attempt} failed after {time.time() - started:.1f}s: {exc}", flush=True)
            if attempt >= attempts or not _is_transient_openai_transport_error(exc):
                break
            time.sleep(min(2 ** (attempt - 1), 4))
        except Exception as exc:
            # httpx network/timeout errors land here.
            last_error = exc
            print(f"[ai] attempt {attempt} failed after {time.time() - started:.1f}s: {exc}", flush=True)
            if attempt >= attempts or not _is_transient_openai_transport_error(exc):
                break
            time.sleep(min(2 ** (attempt - 1), 4))
    raise AIProviderError("OpenAI request failed: " + str(last_error)) from last_error


class OpenAIResponsesProvider(AIProvider):
    """Independent structured extraction using original documents only.

    ``document_resolver`` must resolve an opaque document descriptor to a local
    original file.  The adapter verifies the SHA-256 before the file can leave
    the conversion service.  It never accepts a converted workbook or code
    result in the request contract.
    """

    def __init__(self, *, document_resolver: Callable[[dict], str | Path], api_key: str | None = None,
                 model: str = "gpt-5.6-luna", base_url: str = "https://api.openai.com/v1",
                 reasoning_effort: str = "medium", timeout_seconds: float = 180,
                 max_total_file_bytes: int = 50 * 1024 * 1024,
                 http_post: Callable[[str, dict[str, str], dict, float], dict] | None = None):
        require(callable(document_resolver), "OpenAI document resolver is required")
        require(nonempty(model), "OpenAI model is required")
        require(reasoning_effort in {"none", "low", "medium", "high", "xhigh", "max"},
                "Unsupported OpenAI reasoning effort")
        require(timeout_seconds > 0 and max_total_file_bytes > 0, "Invalid OpenAI provider limits")
        self._document_resolver = document_resolver
        self._api_key = api_key
        self._model = model
        self._base_url = base_url.rstrip("/")
        self._reasoning_effort = reasoning_effort
        self._timeout_seconds = timeout_seconds
        self._max_total_file_bytes = max_total_file_bytes
        self._http_post = http_post or _default_http_post

    @property
    def provider_id(self) -> str:
        return "openai"

    @property
    def model(self) -> str:
        return self._model

    def _file_parts(self, request: dict) -> list[dict]:
        result = []
        total = 0
        for document in request["documents"]:
            path = Path(self._document_resolver(copy.deepcopy(document))).resolve()
            require(path.is_file(), "Original document is missing: " + document["fileId"])
            raw = path.read_bytes()
            require(hashlib.sha256(raw).hexdigest() == document["sha256"],
                    "Original document hash mismatch: " + document["fileId"])
            total += len(raw)
            require(total <= self._max_total_file_bytes, "Original documents exceed OpenAI input limit")
            media_type = document["mediaType"] or mimetypes.guess_type(path.name)[0]
            require(nonempty(media_type), "Original document media type is required")
            item = {
                "type": "input_file",
                "filename": path.name,
                "file_data": "data:" + media_type + ";base64," + base64.b64encode(raw).decode("ascii"),
            }
            if media_type == "application/pdf":
                item["detail"] = "high"
            result.append(item)
        return result

    def _payload(self, request: dict) -> dict:
        semantic_request = copy.deepcopy(request)
        for document in semantic_request["documents"]:
            document.pop("sourceRef", None)
        prompt = {
            "task": "Extract supplier-reported facts independently from the attached original documents.",
            "rules": [
                "Return JSON that matches the supplied schema.",
                "Emit every requested field once per employee.",
                "Keep missing, unreadable, not applicable, and explicit zero distinct.",
                "Do not calculate payroll, infer absent amounts, or use knowledge of a code conversion result.",
                "Preserve signs and decimal text. Record precise source evidence.",
                "Report unexpected financial items in unknownItems.",
            ],
            "request": semantic_request,
        }
        content = self._file_parts(request)
        content.append({"type": "input_text", "text": json.dumps(prompt, ensure_ascii=False)})
        return {
            "model": self._model,
            "reasoning": {"effort": self._reasoning_effort},
            "store": False,
            "input": [{"role": "user", "content": content}],
            "text": {"format": {
                "type": "json_schema",
                "name": "ai_bill_collection",
                "strict": True,
                "schema": _collection_output_schema(request, self._model),
            }},
        }

    def _call_structured(self, payload: dict) -> dict:
        api_key = self._api_key or os.environ.get("OPENAI_API_KEY")
        if not nonempty(api_key):
            raise AIProviderError("OPENAI_API_KEY is not configured")
        response = self._http_post(
            self._base_url + "/responses",
            {"Authorization": "Bearer " + api_key, "Content-Type": "application/json"},
            payload,
            self._timeout_seconds,
        )
        if response.get("status") != "completed":
            raise AIProviderError("OpenAI response did not complete: " + str(response.get("status")))
        for output in response.get("output") or []:
            if output.get("type") != "message":
                continue
            for content in output.get("content") or []:
                if content.get("type") == "refusal":
                    raise AIProviderError("OpenAI refused the extraction request")
                if content.get("type") == "output_text" and isinstance(content.get("text"), str):
                    try:
                        return json.loads(content["text"])
                    except json.JSONDecodeError as exc:
                        raise AIProviderError("OpenAI returned invalid structured JSON") from exc
        raise AIProviderError("OpenAI response contains no structured output")

    def _trust_xlsx_source_evidence(self, plan: dict, documents: list[dict],
                                    template_manifest: dict | None = None) -> dict:
        """Replace model-reported XLSX labels/evidence with cells from the original file."""
        from .template_fill import (
            filter_inconsistent_employee_source_writes,
            inspect_source_employee_layout,
        )

        trusted = copy.deepcopy(plan)
        by_id = {item["fileId"]: item for item in documents}
        workbooks = {}
        source_layouts: dict[str, dict[str, dict]] = {}
        try:
            for item in trusted.get("writes") or []:
                source = item.get("source") if isinstance(item, dict) else None
                if not isinstance(source, dict):
                    continue
                document = by_id.get(source.get("fileId"))
                if not document:
                    continue
                path = Path(self._document_resolver(copy.deepcopy(document))).resolve()
                if path.suffix.lower() not in {".xlsx", ".xlsm"}:
                    continue
                location = str(source.get("location") or "").strip()
                if "!" not in location:
                    continue
                sheet_name, coordinate = location.rsplit("!", 1)
                sheet_name = sheet_name.strip().strip("'").replace("''", "'")
                coordinate = coordinate.strip().replace("$", "")
                if not re.fullmatch(r"[A-Za-z]{1,3}[1-9][0-9]*", coordinate):
                    continue
                if document["fileId"] not in workbooks:
                    workbooks[document["fileId"]] = (
                        load_workbook(path, data_only=False, read_only=False, keep_links=False),
                        load_workbook(path, data_only=True, read_only=False, keep_links=False),
                    )
                formula_wb, value_wb = workbooks[document["fileId"]]
                if sheet_name not in formula_wb.sheetnames or sheet_name not in value_wb.sheetnames:
                    continue
                formula_ws = formula_wb[sheet_name]
                value_ws = value_wb[sheet_name]
                layouts = source_layouts.setdefault(document["fileId"], {})
                if sheet_name not in layouts:
                    layouts[sheet_name] = inspect_source_employee_layout(formula_ws)
                cell = formula_ws[coordinate]
                actual = value_ws[coordinate].value
                label = ""
                for row in range(cell.row - 1, 0, -1):
                    candidate = formula_ws.cell(row, cell.column)
                    if candidate.data_type == "f" or not isinstance(candidate.value, str):
                        continue
                    label = " ".join(candidate.value.split()).strip()
                    if label:
                        break
                if not label or actual is None:
                    continue
                if hasattr(actual, "isoformat"):
                    actual_text = actual.isoformat()[:10]
                    item["valueType"] = "date"
                    item["value"] = actual_text
                elif isinstance(actual, (int, float)) and not isinstance(actual, bool):
                    from decimal import Decimal
                    number = Decimal(str(actual))
                    actual_text = format(number, "f")
                    if "." in actual_text:
                        actual_text = actual_text.rstrip("0").rstrip(".")
                    item["valueType"] = "decimal"
                    item["value"] = actual_text
                else:
                    actual_text = str(actual).strip()
                    item["valueType"] = "text"
                    item["value"] = actual_text
                item["sourceLabel"] = label
                source["rawText"] = f"{label}: {actual_text}"
            # Also inspect every sheet so omitted employees are still discoverable.
            for document in documents:
                try:
                    path = Path(self._document_resolver(copy.deepcopy(document))).resolve()
                except Exception:
                    continue
                if path.suffix.lower() not in {".xlsx", ".xlsm"}:
                    continue
                if document["fileId"] not in workbooks:
                    workbooks[document["fileId"]] = (
                        load_workbook(path, data_only=False, read_only=False, keep_links=False),
                        load_workbook(path, data_only=True, read_only=False, keep_links=False),
                    )
                formula_wb, _value_wb = workbooks[document["fileId"]]
                layouts = source_layouts.setdefault(document["fileId"], {})
                for sheet_name in formula_wb.sheetnames:
                    if sheet_name not in layouts:
                        layouts[sheet_name] = inspect_source_employee_layout(formula_wb[sheet_name])
            filtered = filter_inconsistent_employee_source_writes(
                trusted,
                source_layouts=source_layouts,
                template_manifest=template_manifest,
            )
            if template_manifest is not None and self._template_has_prefilled_identities(template_manifest):
                # CODE (or another step) already anchored person names on -L.
                # Keep summary-row filters, but do not invent extra identity rows / missing-name issues.
                issues = [
                    item for item in (filtered.get("issues") or [])
                    if not (isinstance(item, dict) and item.get("code") == "MISSING_SOURCE_EMPLOYEES")
                ]
                filtered["issues"] = issues
                return filtered
            if template_manifest is not None:
                from .template_fill import seed_missing_employee_identity_writes
                return seed_missing_employee_identity_writes(
                    filtered, template_manifest, source_layouts,
                )
            return filtered
        finally:
            for formula_wb, value_wb in workbooks.values():
                formula_wb.close()
                value_wb.close()

    @staticmethod
    def _template_has_prefilled_identities(template_manifest: dict) -> bool:
        from .template_fill import _is_person_name_label
        name_columns = {
            int(item["column"])
            for item in template_manifest.get("columnContexts") or []
            if isinstance(item, dict) and isinstance(item.get("column"), int) and (
                _is_person_name_label(item.get("primaryLabel"))
                or _is_person_name_label(item.get("pathLabel"))
            )
        }
        if not name_columns:
            return False
        for item in template_manifest.get("nonemptyCells") or []:
            if not isinstance(item, dict) or item.get("valueKind") == "formula":
                continue
            if item.get("column") in name_columns and item.get("value") not in (None, ""):
                return True
        return False

    def _discover_source_employee_roster(self, documents: list[dict]) -> list[dict]:
        """Program-detect named employees from original XLSX bills before prompting."""
        from .template_fill import inspect_source_employee_layout, list_named_source_employees

        source_layouts: dict[str, dict[str, dict]] = {}
        workbooks = []
        try:
            for document in documents:
                try:
                    path = Path(self._document_resolver(copy.deepcopy(document))).resolve()
                except Exception:
                    continue
                if path.suffix.lower() not in {".xlsx", ".xlsm"}:
                    continue
                wb = load_workbook(path, data_only=False, read_only=False, keep_links=False)
                workbooks.append(wb)
                layouts = source_layouts.setdefault(document["fileId"], {})
                for sheet_name in wb.sheetnames:
                    layouts[sheet_name] = inspect_source_employee_layout(wb[sheet_name])
            return list_named_source_employees(source_layouts)
        finally:
            for wb in workbooks:
                wb.close()

    def collect(self, request: dict) -> dict:
        validate_collection_request(request)
        if not nonempty(self._api_key or os.environ.get("OPENAI_API_KEY")):
            raise AIProviderError("OPENAI_API_KEY is not configured")
        return self._call_structured(self._payload(request))

    def plan_template_fill(self, collection: dict, schema: dict, documents: list[dict],
                           template_manifest: dict) -> dict:
        """Map validated AI facts to blank cells in the inspected last ``-L`` sheet."""
        validate_ai_collection(collection, schema, documents, provider_id=self.provider_id)
        require(isinstance(template_manifest, dict)
                and nonempty(template_manifest.get("templateSha256"))
                and nonempty(template_manifest.get("sheetName")), "Invalid template manifest")
        prompt = {
            "task": "Map the validated AI facts into blank input cells of the current last -L worksheet.",
            "rules": [
                "Use worksheet meaning, labels, row context, units and employee layout; do not rely on old coordinates.",
                "Map every employee name and available requested field only when the target is unambiguous.",
                "Never target formulas, headers, merged placeholders, totals, or another worksheet.",
                "Do not include values in the plan; reference only recordIndex and field.",
                "Put ambiguity or missing target structure in issues and leave the cell unwritten.",
            ],
            "schema": copy.deepcopy(schema),
            "aiCollection": copy.deepcopy(collection),
            "template": copy.deepcopy(template_manifest),
        }
        payload = {
            "model": self._model,
            "reasoning": {"effort": self._reasoning_effort},
            "store": False,
            "input": [{"role": "user", "content": [{
                "type": "input_text", "text": json.dumps(prompt, ensure_ascii=False),
            }]}],
            "text": {"format": {
                "type": "json_schema",
                "name": "ai_template_fill_plan",
                "strict": True,
                "schema": _template_plan_output_schema(template_manifest, collection, schema),
            }},
        }
        plan = self._call_structured(payload)
        from .template_fill import validate_template_fill_plan
        validate_template_fill_plan(plan, template_manifest, collection, schema, documents)
        return plan

    def plan_dynamic_template_fill(self, *, documents: list[dict], template_manifest: dict,
                                   run_id: str, period: str, currency: str,
                                   instructions: list[str],
                                   column_mappings: dict[str, str] | None = None,
                                   code_anchored_employees: list[dict] | None = None) -> dict:
        """Read original bills against the current template itself as the output contract."""
        require(bool(documents), "Original documents are required")
        require(isinstance(template_manifest, dict)
                and nonempty(template_manifest.get("templateSha256"))
                and nonempty(template_manifest.get("sheetName")), "Invalid template manifest")
        request = {"documents": copy.deepcopy(documents)}
        column_mappings = copy.deepcopy(column_mappings or {})
        code_anchored = copy.deepcopy(code_anchored_employees or [])
        detected_employees = self._discover_source_employee_roster(documents)
        # Prefer CODE-anchored roster when present; otherwise fall back to source scan.
        roster_for_prompt = [
            {
                "displayName": item.get("displayName"),
                "names": item.get("names") or [],
                "targetRow": item.get("row"),
                "source": "code_identity",
            }
            for item in code_anchored
            if isinstance(item, dict) and nonempty(item.get("displayName"))
        ] or [
            {
                "displayName": item["displayName"],
                "names": item["names"],
                "sheetName": item["sheetName"],
                "sourceRow": item["sourceRow"],
                "fileId": item["fileId"],
                "source": "original_xlsx",
            }
            for item in detected_employees
        ]
        employee_count = len(roster_for_prompt)
        prompt = {
            "task": "Read the attached original supplier bills and fill the current last -L worksheet semantically.",
            "rules": [
                "Treat the current worksheet layout and labels as the field specification; there is no fixed payroll field list.",
                "Infer employees, rows and target meanings from the current worksheet and the original bills.",
                "Use worksheet meaning, labels, row context and units; do not rely on historical coordinates or field names.",
                "Only propose writes to blank input cells in this worksheet.",
                "Choose the employee row carefully. The service will resolve a uniquely known target column from columnMappings and current template labels instead of trusting the proposed column coordinate.",
                "Never target formulas, headers, merged placeholders, totals or another worksheet.",
                "Formula cells (valueKind=formula in template.nonemptyCells) belong to the template. Do not propose writes to them; Excel will compute them.",
                "Blank and numeric 0 are equivalent. Never write 0 / 0.0 / 0.00 into a cell — leave that cell unwritten (blank).",
                "Copy source facts faithfully. Do not invent payroll calculations. Missing or zero amounts should stay blank, not become literal zero writes.",
                "Every write must include precise evidence from an attached original file.",
                "Identity columns are sacred: BU, CN Name, EN Name, Position, Start Date, End Date and FT/PT must map to the template column with the same meaning. Never shift them one column right/left.",
                "Do not write company/BU values into CN Name, employee Chinese names into EN Name, or English names into Position.",
                "Source and template column orders often differ (extra Monthly/Allowance columns). Always match by field label meaning, never by column letter or left-to-right position.",
                "Never copy the source total/summary row into an employee row.",
                "A source row is an employee only when CN Name or EN Name (or equivalent person-name column) is non-empty. Rows with amounts but blank names are totals/padding — never map them.",
                "When codeAnchoredEmployees is present, person names are ALREADY written on those target rows. Do not rewrite names, do not reorder rows, and do not invent extra people.",
                "Match each codeAnchoredEmployees targetRow to the same person in the original bill by CN/EN name, then fill the remaining blank non-formula pay fields on that exact row.",
                "Bills commonly contain MULTIPLE employees. detectedSourceEmployees/codeAnchoredEmployees is the authoritative roster; fill every listed person.",
                f"The current roster lists {employee_count} employee(s). Filling only one person when the roster has more is incorrect.",
                "One employee maps to exactly one target -L row. Never duplicate the same person across multiple target rows.",
                "All values written to one target row must come from that same source employee row. Do not mix a person's identity with another row's hours/salary/fees.",
                "Prefer contiguous blank data rows under the -L header; fill every available unambiguous non-zero, non-formula field for each employee.",
                "semanticLabel must copy the selected target column's template.columnContexts primaryLabel exactly.",
                "sourceLabel must copy the original source field label exactly; rawText must contain that label and value.",
                "Compare the source label with every target column using parent+child pathLabel when present; identical leaf labels under different parents are different fields.",
                "columnMappings contains reviewed supplier-label to target-label hints for this supplier and customer; prefer a matching hint when its target exists in the current template.",
                "If a columnMappings target is absent or stale, use the current template semantics and report ambiguity instead of inventing a target.",
                "Treat qualifiers as meaningful: employee/employer, normal/adjustment, tier/category and similar suffixes are different fields.",
                "If a source fact or target is ambiguous, report an issue and leave it unwritten.",
                "Do not write section/group headers such as Pay Items into employee value cells.",
                "Never write Service Fee / 服务费 (or any source/template column whose meaning is service fee). Leave those cells blank; they are out of scope for AI filling.",
            ],
            "runId": run_id,
            "period": period,
            "currency": currency,
            "profileInstructions": list(instructions),
            "columnMappings": column_mappings,
            "codeAnchoredEmployees": code_anchored,
            "detectedSourceEmployees": roster_for_prompt,
            "documents": [{key: value for key, value in item.items() if key != "sourceRef"}
                          for item in documents],
            "template": copy.deepcopy(template_manifest),
        }
        content = self._file_parts(request)
        content.append({"type": "input_text", "text": json.dumps(prompt, ensure_ascii=False)})
        payload = {
            "model": self._model,
            "reasoning": {"effort": self._reasoning_effort},
            "store": False,
            "input": [{"role": "user", "content": content}],
            "text": {"format": {
                "type": "json_schema",
                "name": "ai_dynamic_template_fill_plan",
                "strict": True,
                "schema": _dynamic_template_plan_output_schema(template_manifest, documents, self._model),
            }},
        }
        from .template_fill import resolve_dynamic_template_fill_targets, validate_dynamic_template_fill_plan

        def _finalize(raw_plan: dict) -> dict:
            return resolve_dynamic_template_fill_targets(
                self._trust_xlsx_source_evidence(
                    raw_plan, documents, template_manifest=template_manifest,
                ),
                template_manifest, column_mappings=column_mappings,
            )

        def _needs_employee_retry(plan: dict) -> bool:
            codes = {
                item.get("code") for item in (plan.get("issues") or []) if isinstance(item, dict)
            }
            return bool(codes & {"MISSING_SOURCE_EMPLOYEES", "EMPLOYEE_IDENTITY_SEEDED"})

        first_plan = _finalize(self._call_structured(payload))
        try:
            validate_dynamic_template_fill_plan(
                first_plan, template_manifest, documents, column_mappings=column_mappings,
            )
        except ValidationError as exc:
            correction = {
                "task": "Return a complete corrected replacement plan.",
                "validationError": str(exc),
                "rules": [
                    "Re-check every write, not only the first rejected write.",
                    "For every sourceLabel, compare all template.columnContexts pathLabel/primaryLabel values again.",
                    "Preserve source values and evidence; correct targetCell and semanticLabel when necessary.",
                    "Never use nameless summary/total source rows; keep one employee on one target row from one source row.",
                    "Cover EVERY employee in detectedSourceEmployees with their own target row and pay fields.",
                    "If no target is unambiguous, omit that write and add a structured issue.",
                ],
                "detectedSourceEmployees": prompt["detectedSourceEmployees"],
                "rejectedPlan": first_plan,
            }
            retry_payload = copy.deepcopy(payload)
            retry_payload["input"][0]["content"].append({
                "type": "input_text",
                "text": json.dumps(correction, ensure_ascii=False),
            })
            corrected_plan = _finalize(self._call_structured(retry_payload))
            validate_dynamic_template_fill_plan(
                corrected_plan, template_manifest, documents, column_mappings=column_mappings,
            )
            return corrected_plan

        if employee_count > 1 and _needs_employee_retry(first_plan):
            correction = {
                "task": "Return a complete replacement plan that covers every detected employee.",
                "validationError": (
                    "Previous plan missed or only partially covered multiple source employees. "
                    "detectedSourceEmployees is authoritative."
                ),
                "rules": [
                    "Create one target -L row per detectedSourceEmployees entry.",
                    "Fill unambiguous pay fields for each employee from that employee's own source row.",
                    "Do not stop after the first employee.",
                    "Never use nameless summary/total source rows.",
                ],
                "detectedSourceEmployees": prompt["detectedSourceEmployees"],
                "rejectedPlan": first_plan,
            }
            retry_payload = copy.deepcopy(payload)
            retry_payload["input"][0]["content"].append({
                "type": "input_text",
                "text": json.dumps(correction, ensure_ascii=False),
            })
            corrected_plan = _finalize(self._call_structured(retry_payload))
            validate_dynamic_template_fill_plan(
                corrected_plan, template_manifest, documents, column_mappings=column_mappings,
            )
            return corrected_plan
        return first_plan


class ProviderRegistry:
    def __init__(self):
        self._providers = {}

    def register(self, provider: AIProvider) -> None:
        require(isinstance(provider, AIProvider) and nonempty(provider.provider_id), "Invalid AI provider")
        require(provider.provider_id not in self._providers, "Duplicate AI provider: " + provider.provider_id)
        self._providers[provider.provider_id] = provider

    def get(self, provider_id: str) -> AIProvider:
        if provider_id not in self._providers:
            raise ValidationError("Unknown AI provider: " + str(provider_id))
        return self._providers[provider_id]


def run_collection(provider: AIProvider, request: dict, schema: dict) -> dict:
    validate_collection_request(request)
    try:
        result = provider.collect(copy.deepcopy(request))
    except AIProviderError:
        raise
    except Exception as exc:
        raise AIProviderError(str(exc)) from exc
    validate_ai_collection(result, schema, request["documents"], run_id=request["runId"],
                           provider_id=provider.provider_id)
    return result
