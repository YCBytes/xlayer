"""Immutable, bounded claims for checking saved workbook facts."""

from __future__ import annotations

import json
import math
import re
from collections.abc import Mapping, Sequence
from dataclasses import dataclass, fields
from typing import cast

from xlayer._canonical import (
    CanonicalTypeError,
    CanonicalValueError,
    canonical_json,
    digest,
    fingerprint,
    frozen_mapping,
    string,
    thaw_json,
)
from xlayer._dependencies import _canonical_address
from xlayer._errors import Refusal

VERSION = "1.0"
MIB = 1048576
CLAIM_ID = re.compile(r"[A-Za-z0-9_.-]{1,64}\Z", re.ASCII)


class ExplanationSizeError(ValueError):
    """An operation would exceed its declared encoded/work bound."""


class ClaimError(ValueError):
    """Deliberate validation failure, distinct from caller-container failures."""

    def __init__(
        self,
        reason: str,
        code: str = "invalid_explanation_request",
        *,
        context: Mapping[str, object] | None = None,
    ) -> None:
        super().__init__(reason)
        self.code = code
        self.context = {} if context is None else dict(context)


def object_dict(value: object) -> dict[str, object]:
    """Internal projection of a validated JSON object."""
    return cast("dict[str, object]", value)


def array(value: object) -> list[object]:
    """Internal projection of a validated JSON array."""
    return cast("list[object]", value)


def bounded_copy(value: object, max_bytes: int) -> object:
    """Count public JSON bytes during capture, before repeated aliases expand.

    Deliberate domain errors are distinct from unexpected caller failures.
    """
    remaining = max_bytes
    active: set[int] = set()

    def charge(size: int) -> None:
        nonlocal remaining
        remaining -= size
        if remaining < 0:
            raise ExplanationSizeError("encoded_bytes")

    def walk(item: object, depth: int) -> object:
        if depth > 64:
            raise CanonicalValueError("JSON nesting exceeds 64")
        if item is None or type(item) is bool:
            charge(4 if item is None or item is True else 5)
            return item
        if type(item) is int:
            if not -(2**63) <= item < 2**63:
                raise CanonicalValueError("integer outside signed 64-bit domain")
            charge(len(str(item)))
            return item
        if type(item) is float:
            if not math.isfinite(item):
                raise CanonicalValueError("non-finite number")
            charge(len(repr(item)))
            return item
        if type(item) is str:
            charge(2)
            if len(item) > remaining:
                raise ExplanationSizeError("encoded_bytes")
            for char in item:
                number = ord(char)
                if 0xD800 <= number <= 0xDFFF:
                    raise CanonicalValueError("invalid Unicode scalar")
                if char in '"\\\b\f\n\r\t':
                    charge(2)
                else:
                    charge(12 if number > 0xFFFF else 6 if number < 32 or number > 126 else 1)
            return item
        if not isinstance(item, (Mapping, Sequence)) or isinstance(item, (str, bytes, bytearray)):
            raise CanonicalTypeError("value outside JSON domain")
        identity = id(item)
        if identity in active:
            raise CanonicalValueError("cyclic JSON")
        charge(2)
        if len(item) > remaining + 1:
            raise ExplanationSizeError("encoded_bytes")
        active.add(identity)
        try:
            if isinstance(item, Mapping):
                result: dict[str, object] = {}
                for index, (key, member) in enumerate(item.items()):
                    if type(key) is not str:
                        raise CanonicalTypeError("JSON keys must be exact strings")
                    if key in result:
                        raise CanonicalValueError("duplicate JSON key")
                    charge(1 + int(index > 0))
                    walk(key, depth + 1)
                    result[key] = walk(member, depth + 1)
                return result
            members: list[object] = []
            for index, member in enumerate(item):
                charge(int(index > 0))
                members.append(walk(member, depth + 1))
            return members
        finally:
            active.remove(identity)

    return walk(value, 0)


def bounded_bytes(value: object, max_bytes: int) -> bytes:
    return canonical_json(bounded_copy(value, max_bytes))


@dataclass(frozen=True)
class ExplanationLimits:
    """Default or lower ceilings for claims, comparisons and complete output."""

    max_claims: int = 128
    max_required_facts: int = 128
    max_cells: int = 128
    max_cell_comparisons: int = 512
    max_request_bytes: int = MIB
    max_evidence_bytes: int = MIB
    max_report_bytes: int = MIB
    max_diagnostic_bytes: int = 65536
    max_diagnostics: int = 16

    def __post_init__(self) -> None:
        for item in fields(self):
            value = getattr(self, item.name)
            minimum = 2 if item.name == "max_diagnostic_bytes" else 1
            if type(value) is not int or not minimum <= value <= cast("int", item.default):
                raise ValueError(f"invalid {item.name}")

    def to_dict(self) -> dict[str, object]:
        return {item.name: getattr(self, item.name) for item in fields(self)}


DEFAULT_EXPLANATION_LIMITS = ExplanationLimits()


def check_explanation_limits(limits: ExplanationLimits) -> None:
    if type(limits) is not ExplanationLimits:
        raise TypeError("limits must be an exact ExplanationLimits")
    limits.__post_init__()


@dataclass(frozen=True)
class FieldSpec:
    kind: str = "text"
    nullable: bool = False
    choices: tuple[object, ...] = ()

    def accepts(self, value: object) -> bool:
        if self.choices:
            return any(type(value) is type(choice) and value == choice for choice in self.choices)
        if value is None:
            return self.nullable or self.kind == "scalar"
        if self.kind == "integer":
            return type(value) is int and value >= 0
        if self.kind == "boolean":
            return type(value) is bool
        if self.kind == "scalar":
            return type(value) in {str, int, float, bool}
        if self.kind == "tabs":
            return isinstance(value, list) and all(
                isinstance(tab, dict)
                and set(tab) == {"name", "state", "kind", "tab_index"}
                and type(tab["name"]) is str
                and bool(tab["name"])
                and STATE.accepts(tab["state"])
                and SHEET_KIND.accepts(tab["kind"])
                and FieldSpec("integer").accepts(tab["tab_index"])
                for tab in value
            )
        return type(value) is str

    def schema(self) -> dict[str, object]:
        if self.choices:
            return {"enum": list(self.choices)}
        if self.kind == "tabs":
            return {
                "type": "array",
                "items": schema_object(
                    {
                        "name": {"type": "string", "minLength": 1},
                        "state": STATE.schema(),
                        "kind": SHEET_KIND.schema(),
                        "tab_index": FieldSpec("integer").schema(),
                    }
                ),
            }
        if self.kind == "scalar":
            return {
                "type": ["string", "number", "boolean", "null"],
                "description": "Finite saved scalar; integers must fit signed 64-bit.",
            }
        kind = "string" if self.kind == "text" else self.kind
        result: dict[str, object] = {"type": [kind, "null"] if self.nullable else kind}
        if kind == "integer":
            result.update(
                minimum=0,
                maximum=2**63 - 1,
                description="Native exact int; bool and integral floats are rejected.",
            )
        return result


STATE = FieldSpec(choices=("visible", "hidden", "veryHidden"))
SHEET_KIND = FieldSpec(choices=("worksheet", "chartsheet", "dialogsheet"))
FIELD_SPECS: dict[str, dict[str, FieldSpec]] = {
    "cell": {
        "presence": FieldSpec(choices=("present", "missing")),
        "stored_type": FieldSpec(choices=("n", "b", "e", "s", "str", "inlineStr", "d", None)),
        "value_kind": FieldSpec(
            choices=("number", "boolean", "string", "error", "iso_date", "blank", None)
        ),
        "raw": FieldSpec(nullable=True),
        "stored_value": FieldSpec("scalar"),
        "value_source": FieldSpec(choices=("stored_literal", "saved_formula_cache", "not_present")),
        "formula_result_status": FieldSpec(choices=("not_verified", None)),
        "has_formula": FieldSpec("boolean"),
        "formula_text": FieldSpec(nullable=True),
        "shared_master_text": FieldSpec(nullable=True),
        "formula_kind": FieldSpec(choices=("normal", "shared", "array", "data_table", None)),
        "format_temporal_kind": FieldSpec(
            choices=("date", "time", "date-time", "non_temporal", "unknown", None)
        ),
        "merge_role": FieldSpec(choices=("root", "child", None)),
    },
    "inventory": {
        "sheets": FieldSpec("tabs"),
        "sheet_count": FieldSpec("integer"),
        "defined_name_count": FieldSpec("integer"),
        "active_tab": FieldSpec("integer", True),
        "date1904": FieldSpec("boolean"),
    },
    "sheet": {"state": STATE, "kind": SHEET_KIND, "tab_index": FieldSpec("integer")},
    "impact": {
        "analysis_status": FieldSpec(choices=("complete", "partial", "unknown", "not_evaluated")),
        **{
            name: FieldSpec("integer", True)
            for name in (
                "known_direct_count",
                "known_transitive_count",
                "known_union_count",
                "exact_total_count",
            )
        },
        "scope": FieldSpec(choices=("worksheet_cell_formulas",)),
        "numerical_effects_evaluated": FieldSpec(choices=(False,)),
    },
}


def schema_object(properties: dict[str, object]) -> dict[str, object]:
    return {
        "type": "object",
        "properties": properties,
        "required": list(properties),
        "additionalProperties": False,
    }


def claim_address(value: object) -> str:
    text = string(value, "address", nonempty=True)
    try:
        return _canonical_address(text)
    except ValueError as exc:
        raise ClaimError("invalid_cell_address") from exc


def cell_target(value: object) -> dict[str, object]:
    if not isinstance(value, dict) or set(value) != {"sheet", "address"}:
        raise ClaimError("invalid_cell_target")
    return {
        "sheet": string(value["sheet"], "sheet", nonempty=True),
        "address": claim_address(value["address"]),
    }


def selector(value: object) -> dict[str, object]:
    if not isinstance(value, dict):
        raise ClaimError("selector_must_be_object")
    family, name = value.get("family"), value.get("field")
    if (
        type(family) is not str
        or family not in FIELD_SPECS
        or type(name) is not str
        or name not in FIELD_SPECS[family]
    ):
        raise ClaimError("unsupported_selector", "unsupported_explanation_claim")
    keys = {"family", "field"}
    if family in {"cell", "sheet"}:
        keys.add("sheet")
    if family == "cell":
        keys.add("address")
    if set(value) != keys:
        raise ClaimError("selector_fields_do_not_match_family")
    result: dict[str, object] = {"family": family, "field": name}
    if "sheet" in keys:
        result["sheet"] = string(value["sheet"], "sheet", nonempty=True)
    if "address" in keys:
        result["address"] = claim_address(value["address"])
    return result


def authorized(item: Mapping[str, object], scope: ExplanationScope) -> bool:
    if item["family"] == "cell":
        return item["address"] in scope.cells.get(cast("str", item["sheet"]), ())
    if item["family"] in {"inventory", "sheet"}:
        return scope.include_inventory
    return scope.impact_root is not None


@dataclass(frozen=True)
class ExplanationScope:
    """Application permission for complete observations and required facts."""

    cells: Mapping[str, Sequence[str]]
    include_inventory: bool = False
    impact_root: tuple[str, str] | None = None
    required_facts: Sequence[Mapping[str, object]] = ()

    def __post_init__(self) -> None:
        candidate: object = self.cells
        if not isinstance(candidate, Mapping) or type(self.include_inventory) is not bool:
            raise TypeError("scope requires a mapping and exact bool")
        try:
            captured = object_dict(
                bounded_copy(
                    {
                        "cells": self.cells,
                        "required": self.required_facts,
                        "root": self.impact_root,
                    },
                    MIB,
                )
            )
            cells: dict[str, object] = {}
            count = 0
            for name, members in object_dict(captured["cells"]).items():
                string(name, "sheet", nonempty=True)
                if not isinstance(members, list):
                    raise CanonicalTypeError("addresses must be a sequence")
                if not members or count + len(members) > 128:
                    raise ValueError("scope cell limit")
                normalized = [_canonical_address(string(v, "address")) for v in members]
                if len(set(normalized)) != len(normalized):
                    raise ValueError("duplicate canonical cells")
                cells[name] = normalized
                count += len(normalized)
            object.__setattr__(self, "cells", frozen_mapping(cells))
            root = captured["root"]
            if root is not None:
                if not isinstance(root, list) or len(root) != 2:
                    raise ValueError("one impact root required")
                normalized_root = cell_target({"sheet": root[0], "address": root[1]})
                object.__setattr__(
                    self, "impact_root", (normalized_root["sheet"], normalized_root["address"])
                )
            if not count and not self.include_inventory and root is None:
                raise ValueError("empty explanation scope")
            requirements = captured["required"]
            if not isinstance(requirements, list):
                raise TypeError("required facts must be a sequence")
            if len(requirements) > 128:
                raise ValueError("required facts limit")
            required = [selector(item) for item in requirements]
            if len({canonical_json(item) for item in required}) != len(required):
                raise ValueError("duplicate required selector")
            if any(not authorized(item, self) for item in required):
                raise ValueError("required selector outside scope")
            object.__setattr__(
                self, "required_facts", tuple(frozen_mapping(item) for item in required)
            )
            bounded_bytes(self.to_dict(), MIB)
        except ExplanationSizeError as exc:
            raise ValueError("scope exceeds encoded limit") from exc

    def to_dict(self) -> dict[str, object]:
        return object_dict(
            thaw_json(
                {
                    "scope_contract_version": VERSION,
                    "cells": self.cells,
                    "include_inventory": self.include_inventory,
                    "impact_root": None
                    if self.impact_root is None
                    else {"sheet": self.impact_root[0], "address": self.impact_root[1]},
                    "required_facts": self.required_facts,
                }
            )
        )

    @property
    def scope_digest(self) -> str:
        return digest(self.to_dict())


RECOVERY = {
    "invalid_explanation_request": "correct_request_schema",
    "unsupported_explanation_claim": "use_supported_saved_fact_claims",
    "explanation_binding_mismatch": "reinspect_and_rebind_request",
    "explanation_scope_exceeded": "request_host_scope_authorization",
    "explanation_evidence_missing": "narrow_claim_or_obtain_supported_evidence",
    "explanation_contradicted": "revise_claim_from_scoped_evidence",
    "explanation_requirements_unmet": "include_required_fact_selectors",
    "explanation_limit_exceeded": "narrow_scope_or_claims",
}


def explanation_refusal(
    code: str,
    *,
    details: Mapping[str, object],
    limits: ExplanationLimits = DEFAULT_EXPLANATION_LIMITS,
    operation: str = "check_explanation",
) -> Refusal:
    safe: dict[str, object] = {}
    for candidate in (details, {"diagnostics_truncated": True}, {}):
        try:
            safe = object_dict(bounded_copy(candidate, limits.max_diagnostic_bytes))
            break
        except ExplanationSizeError:
            continue
    return Refusal(
        code,
        "The explanation could not be checked.",
        operation=operation,
        details=safe,
        recovery_options=(
            {"action": RECOVERY.get(code, "resolve_workbook_read_failure_or_narrow_scope")},
        ),
    )


def _expected(claim: dict[str, object], family: str, name: str) -> None:
    spec = FIELD_SPECS[family][name]
    if not spec.accepts(claim["expected"]):
        raise ClaimError(
            "expected_field_shape", context={"path": ["expected"], "expected_schema": spec.schema()}
        )


def _validate_claim(claim: object, ids: set[str], limits: ExplanationLimits) -> int:
    if not isinstance(claim, dict):
        raise ClaimError("claim_must_be_object")
    ident, kind = claim.get("id"), claim.get("kind")
    if type(ident) is not str or not CLAIM_ID.fullmatch(ident) or ident in ids:
        raise ClaimError("invalid_or_duplicate_claim_id", context={"path": ["id"]})
    ids.add(ident)
    if kind == "fact":
        if set(claim) != {"id", "kind", "selector", "expected"}:
            raise ClaimError("invalid_fact_fields")
        chosen = selector(claim["selector"])
        claim["selector"] = chosen
        _expected(claim, cast("str", chosen["family"]), cast("str", chosen["field"]))
        return int(chosen["family"] == "cell")
    if kind == "all_equal":
        if set(claim) != {"id", "kind", "cells", "field", "expected"}:
            raise ClaimError("invalid_group_fields")
        name = claim["field"]
        if type(name) is not str or name not in FIELD_SPECS["cell"]:
            raise ClaimError("unsupported_group_field", "unsupported_explanation_claim")
        members = claim["cells"]
        if not isinstance(members, list) or len(members) < 2:
            raise ClaimError("group_requires_explicit_members")
        if len(members) > limits.max_cells:
            raise ExplanationSizeError("max_cells")
        normalized = [cell_target(item) for item in members]
        if len({canonical_json(item) for item in normalized}) != len(normalized):
            raise ClaimError("duplicate_group_member")
        claim["cells"] = normalized
        _expected(claim, "cell", name)
        return len(normalized)
    if kind == "known_dependent":
        if set(claim) != {"id", "kind", "target", "relation"}:
            raise ClaimError("invalid_membership_fields")
        if claim["relation"] not in ("direct", "transitive"):
            raise ClaimError("unsupported_membership_relation", "unsupported_explanation_claim")
        claim["target"] = cell_target(claim["target"])
        return 0
    raise ClaimError("unsupported_claim_kind", "unsupported_explanation_claim")


def _validate_request(
    payload: Mapping[str, object], limits: ExplanationLimits
) -> dict[str, object]:
    data = object_dict(bounded_copy(payload, limits.max_request_bytes))
    if set(data) != {
        "schema_version",
        "explanation_contract_version",
        "source_fingerprint",
        "scope_digest",
        "claims",
    }:
        raise ClaimError("invalid_envelope_fields")
    for key in ("schema_version", "explanation_contract_version"):
        if type(data[key]) is not str or data[key] != VERSION:
            raise ClaimError(
                "unsupported_version",
                context={"path": [key], "expected_schema": {"type": "string", "const": VERSION}},
            )
    fingerprint(data["source_fingerprint"])
    fingerprint(data["scope_digest"])
    claims = data["claims"]
    if not isinstance(claims, list) or not claims:
        raise ClaimError("claims_must_be_nonempty_array")
    if len(claims) > limits.max_claims:
        raise ExplanationSizeError("max_claims")
    ids: set[str] = set()
    work = 0
    for index, claim in enumerate(claims):
        try:
            work += _validate_claim(claim, ids, limits)
        except (ClaimError, CanonicalTypeError, CanonicalValueError) as exc:
            context: dict[str, object] = {"claim_index": index, "path": ["claims", index]}
            if (
                isinstance(claim, dict)
                and type(claim.get("id")) is str
                and CLAIM_ID.fullmatch(claim["id"])
            ):
                context["claim_id"] = claim["id"]
            if isinstance(exc, ClaimError):
                context.update(exc.context)
                context["path"] = ["claims", index, *array(exc.context.get("path", []))]
            raise ClaimError(
                str(exc),
                exc.code if isinstance(exc, ClaimError) else "invalid_explanation_request",
                context=context,
            ) from exc
        if work > limits.max_cell_comparisons:
            raise ExplanationSizeError("max_cell_comparisons")
    return data


@dataclass(frozen=True, init=False)
class ExplanationRequest:
    """Validated claims. Construct with from_dict() or from_json()."""

    _data: Mapping[str, object]

    def __init__(self) -> None:
        raise TypeError("use ExplanationRequest.from_dict or from_json")

    @classmethod
    def from_dict(
        cls,
        payload: Mapping[str, object],
        *,
        limits: ExplanationLimits = DEFAULT_EXPLANATION_LIMITS,
    ) -> ExplanationRequest | Refusal:
        if not isinstance(payload, Mapping):
            raise TypeError("payload must be a mapping")
        check_explanation_limits(limits)
        try:
            checked = _validate_request(payload, limits)
            result = object.__new__(cls)
            object.__setattr__(result, "_data", frozen_mapping(checked))
            return result
        except ExplanationSizeError as exc:
            return explanation_refusal(
                "explanation_limit_exceeded",
                details={"reason": str(exc)},
                limits=limits,
                operation="parse_explanation_request",
            )
        except (ClaimError, CanonicalTypeError, CanonicalValueError) as exc:
            return explanation_refusal(
                exc.code if isinstance(exc, ClaimError) else "invalid_explanation_request",
                details={
                    "reason": str(exc),
                    **(exc.context if isinstance(exc, ClaimError) else {}),
                },
                limits=limits,
                operation="parse_explanation_request",
            )

    @classmethod
    def from_json(
        cls, payload: str | bytes, *, limits: ExplanationLimits = DEFAULT_EXPLANATION_LIMITS
    ) -> ExplanationRequest | Refusal:
        if type(payload) not in {str, bytes}:
            raise TypeError("payload must be exact str or bytes")
        check_explanation_limits(limits)
        try:
            if len(payload) > limits.max_request_bytes:
                raise ExplanationSizeError("max_request_bytes")
            text = payload.decode("utf-8") if isinstance(payload, bytes) else payload
            if len(text.encode("utf-8")) > limits.max_request_bytes:
                raise ExplanationSizeError("max_request_bytes")

            def pairs(items: list[tuple[str, object]]) -> dict[str, object]:
                data: dict[str, object] = {}
                for key, value in items:
                    if key in data:
                        raise ClaimError("duplicate_json_key")
                    data[key] = value
                return data

            def constant(_value: str) -> object:
                raise ClaimError("nonfinite_json")

            decoded: object = json.loads(text, object_pairs_hook=pairs, parse_constant=constant)
            if not isinstance(decoded, dict):
                raise ClaimError("request_must_be_object")
            return cls.from_dict(decoded, limits=limits)
        except ExplanationSizeError as exc:
            return explanation_refusal(
                "explanation_limit_exceeded",
                details={"reason": str(exc)},
                limits=limits,
                operation="parse_explanation_request",
            )
        except (ValueError, UnicodeError, RecursionError):
            return explanation_refusal(
                "invalid_explanation_request",
                details={"reason": "invalid_json"},
                limits=limits,
                operation="parse_explanation_request",
            )

    def to_dict(self) -> dict[str, object]:
        return object_dict(thaw_json(self._data))

    def canonical_json(self) -> bytes:
        return canonical_json(self._data)

    @classmethod
    def json_schema(
        cls, *, limits: ExplanationLimits = DEFAULT_EXPLANATION_LIMITS
    ) -> dict[str, object]:
        check_explanation_limits(limits)
        address_schema = {
            "type": "string",
            "pattern": "^[A-Za-z]{1,3}[1-9][0-9]{0,6}$",
            "description": "Explicit A1 address within A1..XFD1048576; no ranges.",
        }
        sheet_schema = {"type": "string", "minLength": 1}
        cell = schema_object({"sheet": sheet_schema, "address": address_schema})
        ident = {"type": "string", "pattern": "^[A-Za-z0-9_.-]{1,64}$"}
        selectors: list[object] = []
        rules: list[object] = []
        for family, members in FIELD_SPECS.items():
            props: dict[str, object] = {
                "family": {"const": family},
                "field": {"enum": list(members)},
            }
            if family in {"cell", "sheet"}:
                props["sheet"] = sheet_schema
            if family == "cell":
                props["address"] = address_schema
            selectors.append(schema_object(props))
            for name, spec in members.items():
                rules.append(
                    {
                        "if": {
                            "properties": {
                                "selector": {
                                    "properties": {
                                        "family": {"const": family},
                                        "field": {"const": name},
                                    },
                                    "required": ["family", "field"],
                                }
                            }
                        },
                        "then": {"properties": {"expected": spec.schema()}},
                    }
                )
        fact_schema = schema_object(
            {
                "id": ident,
                "kind": {"const": "fact"},
                "selector": {"oneOf": selectors},
                "expected": {},
            }
        )
        fact_schema["allOf"] = rules
        group_schema = schema_object(
            {
                "id": ident,
                "kind": {"const": "all_equal"},
                "cells": {
                    "type": "array",
                    "items": cell,
                    "minItems": 2,
                    "maxItems": limits.max_cells,
                },
                "field": {"enum": list(FIELD_SPECS["cell"])},
                "expected": {},
            }
        )
        group_schema["allOf"] = [
            {
                "if": {"properties": {"field": {"const": name}}},
                "then": {"properties": {"expected": spec.schema()}},
            }
            for name, spec in FIELD_SPECS["cell"].items()
        ]
        membership = schema_object(
            {
                "id": ident,
                "kind": {"const": "known_dependent"},
                "target": cell,
                "relation": {"enum": ["direct", "transitive"]},
            }
        )
        identity = {"type": "string", "pattern": "^sha256:[0-9a-f]{64}$"}
        result = schema_object(
            {
                "schema_version": {"type": "string", "const": VERSION},
                "explanation_contract_version": {"type": "string", "const": VERSION},
                "source_fingerprint": identity,
                "scope_digest": identity,
                "claims": {
                    "type": "array",
                    "minItems": 1,
                    "maxItems": limits.max_claims,
                    "items": {"oneOf": [fact_schema, group_schema, membership]},
                },
            }
        )
        result["description"] = (
            "Data-only saved facts. Native validation also enforces unique IDs/canonical cells, "
            "Unicode scalars, signed-64-bit integers, finite floats, nesting <=64, exact native "
            "integer types and aggregate comparison/encoded-byte ceilings."
        )
        return object_dict(bounded_copy(result, MIB))
