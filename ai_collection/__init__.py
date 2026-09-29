"""Provider-neutral AI collection orchestration."""

from .contracts import build_collection_request, validate_ai_collection, validate_collection_request
from .provider import (
    AIProvider,
    AIProviderError,
    MockAIProvider,
    OpenAIResponsesProvider,
    ProviderRegistry,
    run_collection,
)
from .template_fill import (
    extract_last_l_identity_rows,
    filter_inconsistent_employee_source_writes,
    inspect_last_l_sheet,
    inspect_source_employee_layout,
    list_named_source_employees,
    prepare_template_with_code_identities,
    resolve_dynamic_template_fill_targets,
    seed_missing_employee_identity_writes,
    validate_dynamic_template_fill_plan,
    validate_template_fill_plan,
    write_ai_template_copy,
    write_dynamic_ai_template_copy,
)
from .runner import list_ai_validation_profiles, run_ai_comparison_workbook

__all__ = [
    "AIProvider",
    "AIProviderError",
    "MockAIProvider",
    "OpenAIResponsesProvider",
    "ProviderRegistry",
    "build_collection_request",
    "run_collection",
    "inspect_last_l_sheet",
    "inspect_source_employee_layout",
    "list_named_source_employees",
    "extract_last_l_identity_rows",
    "prepare_template_with_code_identities",
    "filter_inconsistent_employee_source_writes",
    "seed_missing_employee_identity_writes",
    "resolve_dynamic_template_fill_targets",
    "validate_template_fill_plan",
    "write_ai_template_copy",
    "validate_dynamic_template_fill_plan",
    "write_dynamic_ai_template_copy",
    "list_ai_validation_profiles",
    "run_ai_comparison_workbook",
    "validate_ai_collection",
    "validate_collection_request",
]
