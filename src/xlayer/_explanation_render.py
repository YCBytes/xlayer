"""Fixed plain-text templates over captured saved facts; no model narrative."""

from __future__ import annotations

from collections.abc import Mapping

from xlayer._canonical import canonical_json
from xlayer._explanation_contract import ExplanationSizeError, array, object_dict

RENDERER_VERSION = "1.0"
LIMITATIONS = {
    "calculation": "not_evaluated",
    "formula_cache_freshness": "not_verified",
    "business_correctness": "not_evaluated",
    "visual_appearance": "not_evaluated",
    "edit_authority": "not_granted",
    "arbitrary_model_prose": "not_verified",
    "whole_task": "not_certified",
    "current_file_freshness": "not_checked",
}
QUALIFICATIONS = (
    "Calculation, visual appearance, business correctness and edit approval were not established.",
    "This checks declared saved facts, not arbitrary model prose or whole-task completion.",
    "Decoded numbers are convenience values; raw text preserves saved digits.",
    "Saved format classification is not a converted date, domain type or edit permission.",
    "The report describes its loaded snapshot; current on-disk freshness was not checked.",
)


def quoted(value: object) -> str:
    return canonical_json(value).decode("ascii")


def render_report(evidence: Mapping[str, object], max_bytes: int) -> str:
    observations = object_dict(evidence["observations"])
    lines: list[str] = []
    size = 0

    def emit(line: str) -> None:
        nonlocal size
        size += len(line.encode("utf-8")) + 1
        if size > max_bytes:
            raise ExplanationSizeError("rendered_text_bytes")
        lines.append(line)

    emit("Checked saved-workbook facts")
    emit(f"Source: {evidence['source_fingerprint']}")
    emit(f"Scope: {evidence['scope_digest']}")
    inventory = observations["inventory"]
    if inventory is not None:
        data = object_dict(object_dict(inventory)["data"])
        tabs = array(data["sheets"])
        emit(
            f"Inventory: {len(tabs)} tabs; active tab index={quoted(data['active_tab'])}; "
            f"date1904={quoted(data['date1904'])}; defined names={data['defined_name_count']}."
        )
        for value in tabs:
            tab = object_dict(value)
            emit(
                f"Tab {tab['tab_index']}: {quoted(tab['name'])}; "
                f"visibility={quoted(tab['state'])}; kind={quoted(tab['kind'])}."
            )
    for inspection in array(observations["cells"]):
        data = object_dict(object_dict(inspection)["data"])
        for value in array(data["cells"]):
            cell = object_dict(value)
            label = f"{quoted(data['sheet'])}!{quoted(cell['address'])}"
            if cell["presence"] == "missing":
                emit(
                    f"{label}: physically missing; no cell is stored here; "
                    "not a saved zero or a present blank."
                )
            else:
                if cell["formula"] is not None:
                    meaning = "unverified saved formula cache"
                elif cell["value_kind"] == "blank":
                    meaning = "present blank literal"
                elif cell["stored_value"] == "":
                    meaning = "stored literal empty text"
                else:
                    meaning = "stored literal"
                emit(
                    f"{label}: present; {meaning}; raw={quoted(cell['raw'])}; "
                    f"decoded={quoted(cell['stored_value'])}; "
                    f"stored type={quoted(cell['stored_type'])}; "
                    f"value kind={quoted(cell['value_kind'])}."
                )
                if cell["formula"] is not None:
                    formula = object_dict(cell["formula"])
                    emit(
                        f"  Formula kind={quoted(formula['kind'])}; "
                        f"own formula text={quoted(formula['text'])}; "
                        f"result status={quoted(cell['formula_result_status'])}; "
                        "not recalculated or verified."
                    )
                    if cell["value_kind"] == "blank":
                        state = (
                            "absent cache element" if cell["raw"] is None else "empty cache element"
                        )
                        emit(f"  {state}; formula remains a formula, not an editable blank input.")
                    elif cell["stored_value"] == "":
                        emit("  Saved empty-text cache; not an absent cache or numeric zero.")
                    if formula["shared_master"] is not None:
                        master = object_dict(formula["shared_master"])
                        emit(
                            f"  Shared master address={quoted(master['address'])}; "
                            f"master ref={quoted(master['ref'])}; "
                            f"master text={quoted(master['text'])}; "
                            "follower expression not expanded."
                        )
                if cell["cell_format"] is not None:
                    style = object_dict(cell["cell_format"])
                    emit(
                        f"  Saved format id={style['format_id']}; "
                        f"code={quoted(style['format_code'])}; "
                        f"temporal kind={quoted(style['temporal_kind'])}; "
                        f"basis={quoted(style['basis'])}."
                    )
            if cell["merge"] is not None:
                merge = object_dict(cell["merge"])
                emit(
                    f"  Merge ref={quoted(merge['ref'])}; root={quoted(merge['root'])}; "
                    f"merge role={quoted(merge['role'])}; "
                    "no root value is copied to a missing child."
                )
    if observations["impact"] is not None:
        data = object_dict(object_dict(observations["impact"])["data"])
        summary = object_dict(data["impact_summary"])
        root = object_dict(array(summary["roots"])[0])
        emit(
            f"Structural impact for {quoted(root['root'])}: {root['analysis_status']}; "
            f"known direct={quoted(root['known_direct_count'])}; "
            f"known transitive={quoted(root['known_transitive_count'])}; "
            f"known union={quoted(summary['known_union_count'])}; "
            f"exact total={quoted(root['exact_total_count'])}."
        )
        emit(
            f"  scope={quoted(summary['scope'])}; "
            f"numerical_effects_evaluated={quoted(summary['numerical_effects_evaluated'])}."
        )
        emit(
            "  Partial known counts are lower bounds; "
            "null means unknown/not evaluated, never zero. "
            "These are potential structural references, not predictions of numerical effects."
        )
        for kind in ("direct", "transitive"):
            emit(
                f"  {kind} sample truncated={quoted(root[kind + '_sample_truncated'])}; "
                f"published sample={quoted(root[kind + '_dependents_sample'])}."
            )
        emit(
            "  Sample absence is not proof of non-dependence; "
            "reference paths and analysis limits remain in the structured evidence."
        )
    emit(
        f"Accepted claims={len(array(evidence['accepted_claims']))}; "
        f"required facts covered={len(array(evidence['required_coverage']))}."
    )
    for qualification in QUALIFICATIONS:
        emit(qualification)
    return "\n".join(lines) + "\n"
