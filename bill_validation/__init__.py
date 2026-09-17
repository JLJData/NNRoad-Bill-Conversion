"""Independent, read-only collection contracts and code-result extraction."""

from .contracts import ValidationError, validate_collection
from .reader import read_code_collection


def compare_collections(*args, **kwargs):
    # Keep AI provider contracts independent of package import order.
    from .comparator import compare_collections as implementation
    return implementation(*args, **kwargs)

__all__ = ["ValidationError", "compare_collections", "validate_collection", "read_code_collection"]
