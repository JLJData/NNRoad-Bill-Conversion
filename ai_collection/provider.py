"""Minimal provider abstraction; no external provider is configured yet."""
from __future__ import annotations

import copy
from abc import ABC, abstractmethod
from typing import Callable

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
