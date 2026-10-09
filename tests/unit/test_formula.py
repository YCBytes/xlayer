"""Handwritten syntax/reference oracles; no formula evaluation."""

from __future__ import annotations

import pytest

from xlayer._ooxml.formula import (
    Endpoint,
    FormulaProblem,
    ParsedFormula,
    ResolvedReference,
    build_name_index,
    build_reference_context,
    parse_formula,
    resolve_references,
)
from xlayer._ooxml.workbook import DefinedName, SheetEntry, WorkbookRegistry


def parse(text: str, *, chars: int = 8192, depth: int = 64) -> ParsedFormula:
    result = parse_formula(text, max_chars=chars, max_nesting=depth)
    assert isinstance(result, ParsedFormula), result
    return result


def problem(text: str, code: str) -> FormulaProblem:
    result = parse_formula(text, max_chars=8192, max_nesting=64)
    assert isinstance(result, FormulaProblem), result
    assert result.code == code
    return result


def test_range_and_literal_have_exact_source_spans() -> None:
    text = 'SUM($A1:B$2,"C12")'
    result = parse(text)
    assert len(result.references) == 1
    token = result.references[0]
    assert token.source_span == (4, 11)
    assert text[slice(*token.source_span)] == "$A1:B$2"
    assert token.first == Endpoint(1, 1, False, True)
    assert token.last == Endpoint(2, 2, True, False)


@pytest.mark.parametrize(
    "text",
    [
        '"A1"',
        '"a""B2"',
        "1e12",
        "TRUE",
        "FALSE",
        "#DIV/0!",
        "#VALUE!",
        "#N/A",
        "#NAME?",
        "#NULL!",
        "#NUM!",
        "#SPILL!",
        "#CALC!",
        "#GETTING_DATA",
        'IF(TRUE,"A1",FALSE)',
    ],
)
def test_literals_do_not_invent_references(text: str) -> None:
    assert parse(text).references == ()


def test_function_context_precedes_cell_context() -> None:
    problem("LOG10(A1)", "unsupported_function")
    assert parse("LOG10").references[0].kind == "cell"
    assert len(parse("SUM(A1,B1)").references) == 2


@pytest.mark.parametrize(
    "text",
    [
        "+A1",
        "--A1",
        "A1%",
        "A1+B1",
        "A1-B1",
        "A1*B1",
        "A1/B1",
        "A1^B1",
        "A1&B1",
        "A1=B1",
        "A1<B1",
        "A1>B1",
        "A1<=B1",
        "A1>=B1",
        "A1<>B1",
        "SUM()",
        "SUM(,A1)",
        "SUM(A1,)",
        "SUM(A1,,B1)",
        "SUM(A1 , B1)",
        "( A1 )",
        "SUM((A1),B1)",
        "1.2e-3+A1",
    ],
)
def test_supported_operators_and_empty_call_arguments(text: str) -> None:
    parse(text)


def test_qualified_span_and_repeated_occurrence_indices() -> None:
    text = "SUM('O''Brien'!$A1,'O''Brien'!$A1)"
    refs = parse(text).references
    assert [r.source_span for r in refs] == [(4, 18), (19, 33)]
    assert [r.occurrence_index for r in refs] == [0, 1]
    assert [r.qualifier for r in refs] == ["O'Brien", "O'Brien"]
    assert [text[slice(*r.source_span)] for r in refs] == ["'O''Brien'!$A1"] * 2


@pytest.mark.parametrize(
    "text",
    [
        "()",
        "A1+",
        "(A1",
        "A1)",
        "A1:B1:C1",
        '"unterminated',
        "=A1",
        "SUM(A1+)",
        "SUM(+)",
        "SUM((,A1))",
        "1 2",
    ],
)
def test_entire_expression_is_required(text: str) -> None:
    assert isinstance(parse_formula(text, max_chars=8192, max_nesting=64), FormulaProblem)


@pytest.mark.parametrize(
    "text,code",
    [
        ('A1+INDIRECT("B1")', "dynamic_reference_function"),
        ("OFFSET(A1,1,1)", "dynamic_reference_function"),
        ("TABLE(A1)", "unsupported_function"),
        ("_xlfn.SUM(A1)", "unsupported_function"),
        ("LET(x,A1,x)", "unsupported_function"),
        ("LAMBDA(x,x)", "unsupported_function"),
        ("A1 B1", "union_or_intersection_reference"),
        ("(A1,B1)", "union_or_intersection_reference"),
        ("SUM(B2:INDEX(A2:C6,5,2))", "formula_parse_failure"),
        ("INDEX((A1,C1),1)", "union_or_intersection_reference"),
        ("S1:S2!A1", "three_dimensional_reference"),
        ("'S1:S2'!A1", "three_dimensional_reference"),
        ("[1]S!A1", "external_workbook_reference"),
        ("'[1]S'!A1", "external_workbook_reference"),
        ("Table[Column]", "unsupported_structured_reference"),
        ("[@Column]", "unsupported_structured_reference"),
        ("A1#", "unsupported_formula_syntax"),
        ("@A1", "unsupported_formula_syntax"),
        ("{1,2}", "unsupported_formula_syntax"),
        ("#REF!", "broken_reference"),
    ],
)
def test_unsupported_syntax_has_precise_issue_code(text: str, code: str) -> None:
    result = problem(text, code)
    assert result.source_span is not None


def test_ascii_endpoint_and_sheet_lexing() -> None:
    refs = parse("a$1+$XFD1048576+'日本 語'!B2+Σheet.a!C3+A0+XFE1").references
    assert refs[0].first == Endpoint(1, 1, True, False)
    assert refs[1].first == Endpoint(1048576, 16384, False, True)
    assert refs[2].qualifier == "日本 語"
    assert refs[3].qualifier == "Σheet.a"
    assert refs[4].kind == refs[5].kind == "defined_name"
    for text in ["$A$0", "A1:XFE1", "A1:A0", "A1٢", "A" + "1" * 5000]:
        assert isinstance(parse_formula(text, max_chars=8192, max_nesting=64), FormulaProblem)


@pytest.mark.parametrize("qualifier", ["2024", "1e3", "4.Quarters", "42_data"])
def test_numeric_leading_sheet_qualifier_is_not_a_numeric_literal(qualifier: str) -> None:
    token = parse(f"{qualifier}!$A1").references[0]
    assert token.qualifier == qualifier and token.first == Endpoint(1, 1, False, True)
    assert token.source_span == (0, len(qualifier) + 4)
    assert parse("1e3+1.5").references == ()


@pytest.mark.parametrize("text", ["''!A1", "''!ScopedRate"])
def test_empty_quoted_qualifier_cannot_be_a_local_reference(text: str) -> None:
    problem(text, "formula_parse_failure")


def test_formula_character_and_frame_boundaries() -> None:
    assert parse("A1" + " " * 8190).text.endswith(" ")
    result = parse_formula("A1" + " " * 8191, max_chars=8192, max_nesting=64)
    assert isinstance(result, FormulaProblem) and result.code == "formula_limit_exceeded"
    assert parse("(" * 64 + "A1" + ")" * 64).max_nesting == 64
    result = parse_formula("(" * 65 + "A1" + ")" * 65, max_chars=8192, max_nesting=64)
    assert isinstance(result, FormulaProblem) and result.code == "formula_limit_exceeded"
    assert ("observed_nesting", 65) in result.information


def test_long_flat_chains_are_stack_safe() -> None:
    assert parse("-" * 2000 + "A1").max_nesting == 0
    assert len(parse("^".join(["A1"] * 2000)).references) == 2000


def resolve(
    text: str,
    names: tuple[DefinedName, ...] = (),
    *,
    row: int = 0,
    col: int = 0,
) -> tuple[ResolvedReference, ...] | FormulaProblem:
    sheets = tuple(
        SheetEntry(name, i + 1, f"r{i}", "visible", i, kind, f"xl/{i}.xml")
        for i, (name, kind) in enumerate(
            [
                ("S", "worksheet"),
                ("T", "worksheet"),
                ("O'Brien", "worksheet"),
                ("日本 語", "worksheet"),
                ("Chart", "chartsheet"),
            ]
        )
    )
    registry = WorkbookRegistry(sheets, names, 0, False, ())
    context = build_reference_context(registry, names=build_name_index(names))
    return resolve_references(
        parse(text), context, formula_sheet="S", row_offset=row, column_offset=col, max_chars=8192
    )


def test_local_name_shadows_global_and_qualified_scope_is_exact() -> None:
    names = (
        DefinedName("ScopedRate", "S!$A$1", None, False),
        DefinedName("ScopedRate", "S!$A$2", "S", True),
        DefinedName("ScopedRate", "T!$A$3", "T", False),
    )
    result = resolve("scopedrate+t!SCOPEDRATE+'日本 語'!a1", names)
    assert isinstance(result, tuple)
    assert [(r.sheet, r.min_row) for r in result] == [("S", 2), ("T", 3), ("日本 語", 1)]
    assert result[0].definition == names[1]


@pytest.mark.parametrize(
    "name",
    [
        "R",
        "C",
        "R1C1",
        "C1R1",
        "INDEX",
        "ACTIVATE",
        "ECMA.CEILING",
        "FORECAST.ETS.STAT",
        "_xlfn.ACOT",
        "_xlws.N",
        "xlpm.N",
        "xlop.N",
        "A" * 256,
    ],
)
def test_used_name_validation(name: str) -> None:
    result = resolve(name, (DefinedName(name, "S!$A$1", None, False),))
    assert isinstance(result, FormulaProblem) and result.code == "invalid_defined_name"


@pytest.mark.parametrize("name", ["_R1C1", "A0", "XFE1", "R0", "A" * 255])
def test_valid_defined_name_boundary(name: str) -> None:
    result = resolve(name, (DefinedName(name, "S!$B$3", None, True),))
    assert isinstance(result, tuple)
    assert (result[0].min_row, result[0].min_column) == (3, 2)


def test_used_name_ambiguity_and_selected_scope() -> None:
    global_one = DefinedName("ScopedRate", "S!$A$1", None, False)
    local = DefinedName("scopedrate", "S!$A$2", "S", False)
    assert isinstance(resolve("ScopedRate", (global_one, global_one)), FormulaProblem)
    good = resolve("ScopedRate", (global_one, global_one, local))
    assert isinstance(good, tuple) and good[0].min_row == 2


def test_coordinate_name_collision_cannot_redirect_reference() -> None:
    result = resolve("LOG10", (DefinedName("LOG10", "#REF!", None, False),))
    assert isinstance(result, tuple)
    assert result[0].token.kind == "cell" and result[0].definition is None
    assert result[0].min_row == 10
    unknown = resolve("A0")
    assert isinstance(unknown, FormulaProblem) and unknown.code == "unresolved_defined_name"


@pytest.mark.parametrize(
    "definition",
    [
        "$A$1",
        "S!A1",
        "1",
        "Other",
        "S!$A$1+1",
        "S!$A$1,S!$A$2",
        "#REF!",
        "[1]S!$A$1",
        "SUM(S!$A$1)",
    ],
)
def test_name_definitions_are_fixed_references_only(definition: str) -> None:
    result = resolve("SUM(ScopedRate)", (DefinedName("ScopedRate", definition, None, False),))
    assert isinstance(result, FormulaProblem) and result.code == "unsupported_name_definition"
    assert result.source_span == (4, 14)


def test_shift_lexical_endpoints_before_normalization() -> None:
    result = resolve("SUM($B2:A$1)", row=1, col=1)
    assert isinstance(result, tuple)
    assert (result[0].min_row, result[0].min_column, result[0].max_row, result[0].max_column) == (
        1,
        2,
        3,
        2,
    )


def test_shared_bounds_and_fixed_name_behavior() -> None:
    result = resolve(
        "T!A1+$B$2+ScopedRate",
        (DefinedName("ScopedRate", "T!$C$3:$D$4", None, False),),
        row=2,
        col=3,
    )
    assert isinstance(result, tuple)
    assert [(r.sheet, r.min_row, r.min_column) for r in result] == [
        ("T", 3, 4),
        ("S", 2, 2),
        ("T", 3, 3),
    ]
    assert result[0].token.source_span == (0, 4)
    bad = resolve("A1", row=-1)
    assert isinstance(bad, FormulaProblem) and bad.code == "shared_reference_out_of_bounds"


@pytest.mark.parametrize("text", ["Missing!A1", "Chart!A1"])
def test_missing_or_nonworksheet_target_is_unresolved(text: str) -> None:
    result = resolve(text)
    assert isinstance(result, FormulaProblem) and result.code == "unresolved_sheet_reference"


@pytest.mark.parametrize("function", ["LEFT", "STDEV", "COVAR", "MEDIAN", "CHOOSE", "INDEX"])
def test_admitted_calls_track_all_static_arguments_not_selected_values(function: str) -> None:
    text = f'{function.lower()}(A1,,SUM(B2,A1),"C3")'
    refs = parse(text).references
    start = len(function) + 1
    assert [text[slice(*r.source_span)] for r in refs] == ["A1", "B2", "A1"]
    assert [r.source_span for r in refs] == [
        (start, start + 2),
        (start + 8, start + 10),
        (start + 11, start + 13),
    ]
    assert [r.occurrence_index for r in refs] == [0, 1, 2]
    assert parse(f"{function}()").references == ()
    assert isinstance(
        parse_formula(f'{function}(A1,INDIRECT("B1"))', max_chars=8192, max_nesting=64),
        FormulaProblem,
    )


@pytest.mark.parametrize("text", ["CHOOSE(B1,C1,D1)", "INDEX(C1:C9,B1)", "SUM(INDEX(A1:A9,2))"])
def test_reference_returning_functions_admit_argument_references_only(text: str) -> None:
    parse(text)


@pytest.mark.parametrize(
    "text", ["SUM(B2:CHOOSE(1,A1,A2))", "INDEX(A1,1):B9", "INDEX((A:A,C:C),1,1,2)"]
)
def test_computed_endpoints_and_unions_never_invent_rectangles(text: str) -> None:
    assert isinstance(parse_formula(text, max_chars=8192, max_nesting=64), FormulaProblem)


@pytest.mark.parametrize(
    "text", ["TABLE(A1)", "_xlfn.RRI(A1)", "_xlfn.INDEX(A1)", "STDEVP(A1)", "COVARIANCE.P(A1)"]
)
def test_function_amendment_does_not_admit_families_or_future_prefixes(text: str) -> None:
    problem(text, "unsupported_function")


@pytest.mark.parametrize(
    "text,row,col,bounds",
    [
        ("SUM($C:D)", 50, 2, (1, 3, 1048576, 6)),
        ("SUM(9:$4)", -3, 80, (4, 1, 6, 16384)),
        ("SUM(D:$C)", 10, -1, (1, 3, 1048576, 3)),
        ("SUM($XFD:$XFD)", -100, 100, (1, 16384, 1048576, 16384)),
        ("SUM($1048576:$1048576)", 100, -100, (1048576, 1, 1048576, 16384)),
        ("SUM(T!d : b)", 0, 0, (1, 2, 1048576, 4)),
        ("SUM('O''Brien'!$4:$9)", 100, 0, (4, 1, 9, 16384)),
        ("SUM(A:A)", 0, 0, (1, 1, 1048576, 1)),
        ("SUM(1:2)", 0, 0, (1, 1, 2, 16384)),
    ],
)
def test_whole_axes_translate_explicit_locks_before_normalization(
    text: str, row: int, col: int, bounds: tuple[int, int, int, int]
) -> None:
    result = resolve(text, row=row, col=col)
    assert isinstance(result, tuple)
    ref = result[0]
    assert (ref.min_row, ref.min_column, ref.max_row, ref.max_column) == bounds
    assert ref.token.kind == "range"


def test_axis_spans_occurrences_and_numeric_leading_qualifiers_are_exact() -> None:
    text = "SUM(2024!$C:$D,2024!$C:$D)"
    refs = parse(text).references
    assert [r.qualifier for r in refs] == ["2024", "2024"]
    assert [r.source_span for r in refs] == [(4, 14), (15, 25)]
    assert [r.occurrence_index for r in refs] == [0, 1]
    assert parse("1e3+1.5").references == ()
    assert parse("Columns").references[0].kind == "defined_name"


@pytest.mark.parametrize(
    "text",
    [
        "0:1",
        "01:2",
        "1:1048577",
        "XFE:XFE",
        "\u0661:2",
        "A:1",
        "1:A",
        "A1:A",
        "A:A1",
        "$A$$:B",
        "A:",
        ":B",
        "A:B:C",
        "A:A!B",
        "1.5:2",
        "1e3:4",
        "1" * 5000 + ":2",
    ],
)
def test_invalid_axis_shapes_bounds_and_ascii_digits_are_refused(text: str) -> None:
    assert isinstance(parse_formula(text, max_chars=8192, max_nesting=64), FormulaProblem)


@pytest.mark.parametrize(
    "text,row,col", [("A:A", 0, -1), ("XFD:XFD", 0, 1), ("1:2", -1, 0), ("1048576:1048576", 1, 0)]
)
def test_shared_axis_overflow_is_not_clipped(text: str, row: int, col: int) -> None:
    result = resolve(text, row=row, col=col)
    assert isinstance(result, FormulaProblem) and result.code == "shared_reference_out_of_bounds"


@pytest.mark.parametrize(
    "definition,bounds", [("T!$C:$D", (1, 3, 1048576, 4)), ("T!$4:$9", (4, 1, 9, 16384))]
)
def test_fixed_absolute_axis_names_ignore_shared_offsets(
    definition: str, bounds: tuple[int, int, int, int]
) -> None:
    result = resolve(
        "ScopedRange", (DefinedName("ScopedRange", definition, "S", False),), row=50, col=50
    )
    assert isinstance(result, tuple) and result[0].sheet == "T"
    ref = result[0]
    assert (ref.min_row, ref.min_column, ref.max_row, ref.max_column) == bounds


@pytest.mark.parametrize(
    "definition", ["T!C:D", "T!$C:D", "$C:$D", "T!4:9", "T!$4:9", "SUM(T!$C:$D)", "T!$C:$D,T!$F:$F"]
)
def test_relative_unqualified_and_computed_axis_names_stay_unsupported(definition: str) -> None:
    result = resolve("ScopedRange", (DefinedName("ScopedRange", definition, None, False),))
    assert isinstance(result, FormulaProblem) and result.code == "unsupported_name_definition"
