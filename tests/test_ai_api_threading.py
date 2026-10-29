"""AI requests must leave the ASGI event loop responsive; no model/network calls."""
import asyncio
import io
import threading
import unittest
from pathlib import Path
from unittest.mock import patch

from fastapi import UploadFile
import convert_api as api


class AiApiThreadingTests(unittest.IsolatedAsyncioTestCase):
    async def test_workbook_generation_runs_off_event_loop(self):
        entered = threading.Event()
        release = threading.Event()
        main_thread = threading.get_ident()

        def generate(**kwargs):
            self.assertNotEqual(threading.get_ident(), main_thread)
            entered.set()
            self.assertTrue(release.wait(3), "Test must release the worker")
            Path(kwargs["output_path"]).write_bytes(b"test workbook")
            return {"runId": "test-run", "model": "test-model", "sheetName": "Test-L", "formalResult": False}

        with patch.object(api, "_AI_VALIDATION_ENABLED", True), patch.object(api, "run_ai_comparison_workbook", generate):
            task = asyncio.create_task(api.ai_validation_run(
                files=[UploadFile(filename="original.pdf", file=io.BytesIO(b"test original"))],
                template=UploadFile(filename="template.xlsx", file=io.BytesIO(b"test template")),
                run_id="test-run", period="2026-10", currency="USD", profile_id="test-profile",
                source_hints_json=None, provider="openai", model="test-model", reasoning_effort="low",
                base_url=None, timeout_seconds=180, code_result=None,
            ))
            try:
                for _ in range(200):
                    if entered.is_set():
                        break
                    await asyncio.sleep(0.005)
                self.assertTrue(entered.is_set())
                self.assertFalse(task.done(), "Generation should still be running while the loop is responsive")
                self.assertTrue(api.health()["ok"])
            finally:
                release.set()
            response = await task
            self.assertEqual(response.headers["X-AI-Formal-Result"], "false")
            self.assertTrue(Path(response.path).is_file())
            await response.background()
            self.assertFalse(Path(response.path).exists())

    async def test_generation_failure_cleans_temporary_files(self):
        touched = []

        def fail(**kwargs):
            touched.append(Path(kwargs["output_path"]).parent)
            raise TimeoutError("provider test timeout")

        with patch.object(api, "_AI_VALIDATION_ENABLED", True), patch.object(api, "run_ai_comparison_workbook", fail), \
                patch.object(api.traceback, "print_exc"):
            with self.assertRaises(api.HTTPException) as raised:
                await api.ai_validation_run(
                    files=[UploadFile(filename="original.pdf", file=io.BytesIO(b"source"))],
                    template=UploadFile(filename="template.xlsx", file=io.BytesIO(b"template")),
                    run_id="failed-run", period="2026-10", currency="USD", profile_id="test",
                    source_hints_json=None, provider="openai", model=None, reasoning_effort=None,
                    base_url=None, timeout_seconds=180, code_result=None,
                )
            self.assertEqual(raised.exception.status_code, 500)
            self.assertFalse(touched[0].exists())


if __name__ == "__main__":
    unittest.main()
