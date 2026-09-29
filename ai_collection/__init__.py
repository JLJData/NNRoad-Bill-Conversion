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
    filter_inconsistent_employee_source_writes,
    inspect_last_l_sheet,
    inspect_source_employee_layout,
    resolve_dynamic_template_fill_targets,
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
    "filter_inconsistent_employee_source_writes",
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
