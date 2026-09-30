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
    copy_code_provenance_cells,
    extract_last_l_identity_rows,
    filter_inconsistent_employee_source_writes,
    inspect_last_l_sheet,
    inspect_source_employee_layout,
    list_named_source_employees,
    normalize_code_provenance_cells,
    prepare_template_with_code_identities,
    resolve_dynamic_template_fill_targets,
    seed_missing_employee_identity_writes,
    strip_code_provenance_writes,
    sync_code_owned_regions_from_code_result,
    validate_dynamic_template_fill_plan,
    validate_template_fill_plan,
    write_ai_template_copy,
    write_dynamic_ai_template_copy,
)
from .runner import list_ai_validation_profiles, run_ai_comparison_workbook
from .suggest_prompt import suggest_ai_instructions

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
    "normalize_code_provenance_cells",
    "strip_code_provenance_writes",
    "copy_code_provenance_cells",
    "sync_code_owned_regions_from_code_result",
    "resolve_dynamic_template_fill_targets",
    "validate_template_fill_plan",
    "write_ai_template_copy",
    "validate_dynamic_template_fill_plan",
    "write_dynamic_ai_template_copy",
    "list_ai_validation_profiles",
    "run_ai_comparison_workbook",
    "suggest_ai_instructions",
    "validate_ai_collection",
    "validate_collection_request",
]
