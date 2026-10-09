"""Strict, detached JSON evidence. No implicit logging or durable-object import."""

from __future__ import annotations

import hashlib
import json
import math
import re
from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field
from types import MappingProxyType
from typing import cast

MAX_EVIDENCE_BYTES = 16 * 1024 * 1024
_DIGEST = re.compile(r"sha256:[0-9a-f]{64}\Z")


class CanonicalTypeError(TypeError):
    """Deliberate domain validation, distinct from caller-container failures."""


class CanonicalValueError(ValueError):
    """Deliberate domain validation, distinct from caller-container failures."""


class EvidenceTooLarge(ValueError):  # noqa: N818 - named bound signal at the private boundary
    """Canonical evidence cannot fit the transaction's declared byte ceiling."""


def string(value: object, name: str, *, nonempty: bool = False) -> str:
    if type(value) is not str:
        raise CanonicalTypeError(f"{name} must be an exact str")
    text = value
    if (nonempty and not text) or any(0xD800 <= ord(c) <= 0xDFFF for c in text):
        raise CanonicalValueError(f"{name} must contain valid Unicode scalar text")
    return text


def fingerprint(value: object) -> str:
    text = string(value, "fingerprint")
    if _DIGEST.fullmatch(text) is None:
        raise CanonicalValueError("invalid SHA-256 fingerprint")
    return text


def freeze_json(value: object, *, _depth: int = 0, _parents: set[int] | None = None) -> object:
    """Validate without coercion; preserve shared acyclic inputs and reject cycles."""
    if _depth > 64:
        raise CanonicalValueError("JSON nesting exceeds 64")
    if value is None or type(value) is bool:
        return value
    if type(value) is str:
        return string(value, "JSON string")
    if type(value) is int:
        number = value
        if not -(2**63) <= number < 2**63:
            raise CanonicalValueError("JSON integer outside signed 64-bit range")
        return number
    if type(value) is float:
        if not math.isfinite(value):
            raise CanonicalValueError("JSON float must be finite")
        return value
    if not isinstance(value, (Mapping, Sequence)) or isinstance(value, (str, bytes, bytearray)):
        raise CanonicalTypeError("value outside the transaction JSON domain")
    parents = set() if _parents is None else _parents
    identity = id(value)
    if identity in parents:
        raise CanonicalValueError("cyclic JSON container")
    parents.add(identity)
    try:
        if isinstance(value, Mapping):
            result: dict[str, object] = {}
            for key, item in value.items():
                key = string(key, "JSON key")
                if key in result:
                    raise CanonicalValueError("duplicate JSON key")
                result[key] = freeze_json(item, _depth=_depth + 1, _parents=parents)
            return MappingProxyType(result)
        return tuple(freeze_json(item, _depth=_depth + 1, _parents=parents) for item in value)
    finally:
        parents.remove(identity)


def thaw_json(value: object) -> object:
    if isinstance(value, Mapping):
        return {key: thaw_json(item) for key, item in value.items()}
    if isinstance(value, tuple):
        return [thaw_json(item) for item in value]
    return value


def frozen_mapping(value: Mapping[str, object]) -> Mapping[str, object]:
    if not isinstance(value, Mapping):
        raise CanonicalTypeError("a mapping is required")
    return cast("Mapping[str, object]", freeze_json(value))


def canonical_json(value: object) -> bytes:
    validated = freeze_json(value)
    result = json.dumps(
        thaw_json(validated),
        sort_keys=True,
        separators=(",", ":"),
        ensure_ascii=True,
        allow_nan=False,
    ).encode("utf-8")
    if len(result) > MAX_EVIDENCE_BYTES:
        raise EvidenceTooLarge("canonical evidence exceeds 16 MiB")
    return result


def digest(value: object) -> str:
    return "sha256:" + hashlib.sha256(canonical_json(value)).hexdigest()


@dataclass(frozen=True)
class EvidenceRecord:
    evidence: Mapping[str, object]
    audit_metadata: Mapping[str, object] = field(default_factory=dict)
    _digest: str = field(init=False, repr=False)

    def __post_init__(self) -> None:
        evidence = frozen_mapping(self.evidence)
        if evidence.get("schema_version") != "1.0":
            raise ValueError("unsupported evidence schema")
        if {"preview_digest", "receipt_digest", "audit_metadata"} & evidence.keys():
            raise ValueError("reserved derived field in evidence")
        object.__setattr__(self, "evidence", evidence)
        object.__setattr__(self, "audit_metadata", frozen_mapping(self.audit_metadata))
        object.__setattr__(self, "_digest", digest(evidence))

    def _dict(self, digest_key: str) -> dict[str, object]:
        result = cast("dict[str, object]", thaw_json(self.evidence))
        result[digest_key] = self._digest
        result["audit_metadata"] = thaw_json(self.audit_metadata)
        return result
