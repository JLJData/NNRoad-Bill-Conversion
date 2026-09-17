import copy
import unittest

from ai_collection import (
    AIProviderError,
    MockAIProvider,
    ProviderRegistry,
    build_collection_request,
    run_collection,
    validate_ai_collection,
    validate_collection_request,
)
from bill_validation import ValidationError, compare_collections


class AICollectionCompareTests(unittest.TestCase):
    def setUp(self):
        self.schema = {
            "formatVersion": 0, "schemaId": "pay", "schemaVersion": "1",
            "status": "draft_not_enabled", "automaticPassEnabled": False,
            "decimalEncoding": "string", "blankIsZero": False,
            "fields": [
                {"field": "salary", "type": "decimal", "unit": "currency", "requiredOutputEntry": True,
                 "comparison": {"mode": "compare_reported_fact", "diagnosticTolerance": "0.01"}},
                {"field": "sick", "type": "decimal", "unit": "currency", "requiredOutputEntry": True,
                 "comparison": {"mode": "review_recalculated_value", "diagnosticTolerance": "0"}},
            ],
        }
        self.documents = [{"fileId": "source-1", "sha256": "1" * 64,
                           "mediaType": "application/pdf", "sourceRef": "opaque://source-1"}]
        self.code = {
            "engine": "code", "schemaId": "pay", "schemaVersion": "1",
            "resultSha256": "2" * 64, "templateSha256": "3" * 64,
            "period": "2026-03", "currency": "TWD", "automaticPassEnabled": False,
            "records": [self.code_record("e1", " Alice Smith ", "100.00", "-5"),
                        self.code_record("e2", "Bob", "200", "0")],
            "issues": [], "readiness": "ready_for_comparison",
        }
        self.ai = {
            "engine": "ai", "provider": "mock", "model": "fixture-v1", "runId": "run-1",
            "schemaId": "pay", "schemaVersion": "1", "period": "2026-03", "currency": "TWD",
            "automaticPassEnabled": False,
            "records": [self.ai_record("bob", "200.00", "0"),
                        self.ai_record("alice   smith", "100.009", "-5")],
            "unknownItems": [], "issues": [], "readiness": "ready_for_association",
        }

    def code_record(self, entity, name, salary, sick):
        return {"entityKey": entity, "observedName": name, "row": 1, "fields": [
            self.code_field("salary", salary, "B1"), self.code_field("sick", sick, "C1")]}

    def code_field(self, field, value, cell):
        return {"field": field, "status": "found", "value": value, "currency": "TWD",
                "source": {"fileSha256": "2" * 64, "sheet": "Result", "cell": cell}}

    def ai_record(self, name, salary, sick):
        return {"sourceEntity": {"employeeId": None, "name": name}, "fields": [
            self.ai_field("salary", salary, "salary row"), self.ai_field("sick", sick, "sick row")]}

    def ai_field(self, field, value, location):
        return {"field": field, "status": "found", "value": value, "currency": "TWD", "confidence": 0.9,
                "source": {"fileId": "source-1", "page": 1, "location": location, "rawText": str(value)}}

    def request(self):
        return build_collection_request(run_id="run-1", documents=self.documents, schema=self.schema,
                                        instructions=["Extract source facts."], period="2026-03", currency="TWD")

    def test_request_projection_excludes_comparison_rules_and_code_result(self):
        request = self.request()
        self.assertNotIn("comparison", request["schema"]["fields"][0])
        self.assertNotIn("codeResult", request)
        validate_collection_request(request)
        request["codeResult"] = {"salary": "100"}
        with self.assertRaises(ValidationError):
            validate_collection_request(request)

    def test_mock_provider_runs_without_mutating_request_or_response(self):
        request, response = self.request(), copy.deepcopy(self.ai)
        provider = MockAIProvider(response)
        result = run_collection(provider, request, self.schema)
        result["records"][0]["fields"][0]["value"] = "999"
        self.assertEqual(response["records"][0]["fields"][0]["value"], "200.00")
        self.assertEqual(len(provider.calls), 1)
        self.assertEqual(provider.calls[0], request)

    def test_provider_failure_is_not_a_business_mismatch(self):
        def fail(_):
            raise RuntimeError("timeout")
        with self.assertRaisesRegex(AIProviderError, "timeout"):
            run_collection(MockAIProvider(fail), self.request(), self.schema)

    def test_registry_rejects_duplicates_and_unknown_ids(self):
        registry = ProviderRegistry()
        registry.register(MockAIProvider(self.ai))
        with self.assertRaises(ValidationError):
            registry.register(MockAIProvider(self.ai))
        self.assertEqual(registry.get("mock").provider_id, "mock")
        with self.assertRaises(ValidationError):
            registry.get("missing")

    def test_ai_contract_rejects_wrong_provenance_invalid_value_or_field_gap(self):
        mutations = [
            lambda value: value["records"][0]["fields"][0]["source"].update(fileId="other"),
            lambda value: value["records"][0]["fields"][0].update(value="1,000"),
            lambda value: value["records"][0]["fields"].pop(),
            lambda value: value["records"][0]["fields"][0].update(confidence=2),
        ]
        for mutate in mutations:
            invalid = copy.deepcopy(self.ai)
            mutate(invalid)
            with self.subTest(mutate=mutate), self.assertRaises(ValidationError):
                validate_ai_collection(invalid, self.schema, self.documents)

    def test_row_order_and_name_whitespace_do_not_change_association(self):
        result = compare_collections(self.code, self.ai, self.schema, self.documents)
        self.assertEqual(result["summary"]["associatedEntityCount"], 2)
        self.assertEqual(result["summary"]["matchedFieldCount"], 2)
        self.assertEqual(result["comparisonStatus"], "REVIEW_REQUIRED")
        self.assertEqual({d["kind"] for d in result["differences"]}, {"FIELD_REQUIRES_REVIEW"})
        self.assertFalse(result["automaticPassEnabled"])
        self.assertEqual(result["workflowDecision"], "NOT_EVALUATED")

    def test_amount_difference_exceeding_tolerance_is_mismatch(self):
        self.ai["records"][1]["fields"][0]["value"] = "100.02"
        result = compare_collections(self.code, self.ai, self.schema, self.documents)
        mismatch = next(d for d in result["differences"] if d["kind"] == "VALUE_MISMATCH")
        self.assertEqual(mismatch["difference"], "0.02")
        self.assertEqual(result["comparisonStatus"], "MISMATCH")

    def test_missing_and_extra_employee_are_reported(self):
        self.ai["records"] = [self.ai_record("Carol", "1", "0")]
        result = compare_collections(self.code, self.ai, self.schema, self.documents)
        kinds = [d["kind"] for d in result["differences"]]
        self.assertEqual(kinds.count("MISSING_AI_ENTITY"), 2)
        self.assertEqual(kinds.count("EXTRA_AI_ENTITY"), 1)
        self.assertEqual(result["comparisonStatus"], "MISMATCH")

    def test_zero_ai_records_reports_every_code_employee_missing(self):
        self.ai["records"] = []
        self.ai["readiness"] = "review_required"
        result = compare_collections(self.code, self.ai, self.schema, self.documents)
        self.assertEqual([d["kind"] for d in result["differences"]].count("MISSING_AI_ENTITY"), 2)
        self.assertEqual(result["comparisonStatus"], "MISMATCH")

    def test_ambiguous_name_does_not_pair_arbitrarily(self):
        self.code["records"][1]["observedName"] = "Alice Smith"
        self.ai["records"] = [self.ai_record("Alice Smith", "100", "-5")]
        result = compare_collections(self.code, self.ai, self.schema, self.documents)
        self.assertEqual(result["summary"]["associatedEntityCount"], 0)
        self.assertEqual(result["differences"][0]["kind"], "AMBIGUOUS_ENTITY")
        self.assertEqual(result["comparisonStatus"], "REVIEW_REQUIRED")

    def test_unknown_item_and_unavailable_field_require_review(self):
        self.ai["records"][0]["fields"][0].update(status="missing", value=None)
        self.ai["unknownItems"] = [{"label": "Remote work allowance",
                                    "source": {"fileId": "source-1", "location": "row 8"}}]
        result = compare_collections(self.code, self.ai, self.schema, self.documents)
        self.assertTrue({"FIELD_UNAVAILABLE", "UNKNOWN_AI_ITEM"} <= {d["kind"] for d in result["differences"]})
        self.assertEqual(result["comparisonStatus"], "REVIEW_REQUIRED")

    def test_period_or_currency_mismatch_is_rejected(self):
        for key, value in (("period", "2026-04"), ("currency", "USD")):
            original = self.ai[key]
            self.ai[key] = value
            with self.subTest(key=key), self.assertRaises(ValidationError):
                compare_collections(self.code, self.ai, self.schema, self.documents)
            self.ai[key] = original

    def test_pure_reported_fields_can_match_without_workflow_approval(self):
        self.schema["fields"][1]["comparison"]["mode"] = "compare_reported_fact"
        result = compare_collections(self.code, self.ai, self.schema, self.documents)
        self.assertEqual(result["comparisonStatus"], "MATCH")
        self.assertEqual(result["summary"]["matchedFieldCount"], 4)
        self.assertEqual(result["workflowDecision"], "NOT_EVALUATED")

    def test_mock_end_to_end_request_collection_and_comparison(self):
        extracted = run_collection(MockAIProvider(self.ai), self.request(), self.schema)
        result = compare_collections(self.code, extracted, self.schema, self.documents)
        self.assertEqual(result["summary"]["associatedEntityCount"], 2)
        self.assertEqual(result["comparisonStatus"], "REVIEW_REQUIRED")


if __name__ == "__main__":
    unittest.main()
