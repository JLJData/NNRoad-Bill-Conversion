"""Unit tests for AI instruction suggestion (no network)."""
from __future__ import annotations

import json
import unittest

from ai_collection.provider import AIProviderError
from ai_collection.suggest_prompt import (
    build_suggest_prompt_payload,
    normalize_suggestion_result,
    suggest_ai_instructions,
)


class SuggestPromptTests(unittest.TestCase):
    def test_build_payload_compacts_diffs_and_uses_user_request(self):
        payload = build_suggest_prompt_payload(
            user_request="CODE is closer here; AI missed Medical Insurance backfill.",
            current_instructions=["Skip service fee"],
            comparison={
                "codeSheet": "Payroll-L",
                "aiSheet": "Payroll-L",
                "differenceCount": 1,
                "matchedEmployeeCount": 2,
                "independentComparison": True,
                "diffs": [{
                    "type": "VALUE_MISMATCH",
                    "employeeName": "Alice",
                    "fieldLabel": "Medical Insurance",
                    "codeCell": "D2",
                    "codeValue": "120",
                    "aiCell": "D2",
                    "aiValue": "40",
                }],
                "unmatchedCode": [],
                "unmatchedAi": [],
            },
            model="gpt-test",
            reasoning_effort="low",
        )
        self.assertEqual(payload["model"], "gpt-test")
        self.assertEqual(payload["reasoning"]["effort"], "low")
        prompt = json.loads(payload["input"][0]["content"][0]["text"])
        self.assertIn("CODE is closer", prompt["userRequest"])
        self.assertEqual(prompt["currentInstructions"], ["Skip service fee"])
        self.assertEqual(prompt["diffs"][0]["fieldLabel"], "Medical Insurance")
        self.assertEqual(prompt["comparisonSummary"]["codeSheet"], "Payroll-L")

    def test_normalize_suggestion_result_dedupes_and_trims(self):
        result = normalize_suggestion_result({
            "explanation": "Medical cover lines should be summed.",
            "suggestedInstructions": [
                "Sum Medical Insurance Cover including prior-month backfill.",
                "Sum Medical Insurance Cover including prior-month backfill.",
                "  ",
            ],
        }, model="gpt-test")
        self.assertEqual(result["model"], "gpt-test")
        self.assertEqual(result["suggestedInstructions"], [
            "Sum Medical Insurance Cover including prior-month backfill.",
        ])

    def test_suggest_ai_instructions_uses_http_post_mock(self):
        calls = []

        def fake_post(url, headers, payload, timeout):
            calls.append({"url": url, "payload": payload, "timeout": timeout})
            return {
                "status": "completed",
                "output": [{
                    "type": "message",
                    "content": [{
                        "type": "output_text",
                        "text": json.dumps({
                            "explanation": "AI undercounted medical cover.",
                            "suggestedInstructions": [
                                "For this supplier: sum current and prior-month Medical Insurance Cover into Medical Insurance."
                            ],
                        }),
                    }],
                }],
            }

        result = suggest_ai_instructions(
            user_request="Please draft prompts assuming CODE is correct this time.",
            current_instructions=[],
            comparison={"diffs": [], "differenceCount": 0},
            model="gpt-test",
            reasoning_effort="low",
            base_url="https://example.test/v1",
            api_key="test-key",
            timeout_seconds=33,
            http_post=fake_post,
        )
        self.assertEqual(len(calls), 1)
        self.assertTrue(calls[0]["url"].endswith("/responses"))
        self.assertEqual(calls[0]["timeout"], 33.0)
        self.assertEqual(result["suggestedInstructions"][0].startswith("For this supplier"), True)
        self.assertIn("undercounted", result["explanation"])

    def test_missing_api_key_raises(self):
        with self.assertRaises(AIProviderError):
            suggest_ai_instructions(
                user_request="hello",
                current_instructions=[],
                comparison={},
                model="gpt-test",
                api_key="",
                http_post=lambda *args, **kwargs: {"status": "completed", "output": []},
            )


if __name__ == "__main__":
    unittest.main()
