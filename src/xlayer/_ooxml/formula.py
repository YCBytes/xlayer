"""Restricted, whole-expression reference admission; never formula evaluation.

Source spans refer to exact OOXML text. A failed expression contributes no
references, even when an earlier prefix was supported. Parsing uses explicit
frames, not an expression AST or recursive descent.
"""

from __future__ import annotations

import re
from collections.abc import Mapping
from dataclasses import dataclass, replace
from types import MappingProxyType
from typing import Literal

from xlayer._ooxml.workbook import DefinedName, SheetEntry, WorkbookRegistry

ReferenceKind = Literal["cell", "range", "defined_name"]

# Syntax-only reserved identifiers, pinned 2026-10-05 to MS-XLSX 2.2.2:
# https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/
# 3d025add-118d-4413-9856-ab65712ec1b0
# Complete function-list, command-list and unprefixed future-function-list.
# The prefixed future/worksheet productions are rejected by prefix below.
# This is NOT a list of supported functions and never needs runtime lookup.
_RESERVED_NAMES = frozenset(
    {
        "A1.R1C1",
        "ABS",
        "ABSREF",
        "ACCRINT",
        "ACCRINTM",
        "ACOS",
        "ACOSH",
        "ACTIVATE",
        "ACTIVATE.NEXT",
        "ACTIVATE.NOTES",
        "ACTIVATE.PREV",
        "ACTIVE.CELL",
        "ACTIVE.CELL.FONT",
        "ADD.ARROW",
        "ADD.BAR",
        "ADD.CHART.AUTOFORMAT",
        "ADD.COMMAND",
        "ADD.LIST.ITEM",
        "ADD.MENU",
        "ADD.OVERLAY",
        "ADD.PRINT.AREA",
        "ADD.TOOL",
        "ADD.TOOLBAR",
        "ADDIN.MANAGER",
        "ADDRESS",
        "ALERT",
        "ALIGNMENT",
        "AMORDEGRC",
        "AMORLINC",
        "AND",
        "APP.ACTIVATE",
        "APP.ACTIVATE.MICROSOFT",
        "APP.MAXIMIZE",
        "APP.MINIMIZE",
        "APP.MOVE",
        "APP.RESTORE",
        "APP.SIZE",
        "APP.TITLE",
        "APPLY.NAMES",
        "APPLY.STYLE",
        "AREAS",
        "ARGUMENT",
        "ARRANGE.ALL",
        "ASC",
        "ASIN",
        "ASINH",
        "ASSIGN.TO.OBJECT",
        "ASSIGN.TO.TOOL",
        "ATAN",
        "ATAN2",
        "ATANH",
        "ATTACH.TEXT",
        "ATTACH.TOOLBARS",
        "ATTRIBUTES",
        "AUTO.OUTLINE",
        "AUTOCORRECT",
        "AVEDEV",
        "AVERAGE",
        "AVERAGEA",
        "AVERAGEIF",
        "AVERAGEIFS",
        "AXES",
        "BAHTTEXT",
        "BEEP",
        "BESSELI",
        "BESSELJ",
        "BESSELK",
        "BESSELY",
        "BETADIST",
        "BETAINV",
        "BIN2DEC",
        "BIN2HEX",
        "BIN2OCT",
        "BINOMDIST",
        "BORDER",
        "BREAK",
        "BRING.TO.FRONT",
        "CALCULATE.DOCUMENT",
        "CALCULATE.NOW",
        "CALCULATION",
        "CALL",
        "CALLER",
        "CANCEL.COPY",
        "CANCEL.KEY",
        "CEILING",
        "CELL",
        "CELL.PROTECTION",
        "CHANGE.LINK",
        "CHAR",
        "CHART.ADD.DATA",
        "CHART.TREND",
        "CHART.WIZARD",
        "CHECK.COMMAND",
        "CHECKBOX.PROPERTIES",
        "CHIDIST",
        "CHIINV",
        "CHITEST",
        "CHOOSE",
        "CLEAN",
        "CLEAR",
        "CLEAR.OUTLINE",
        "CLEAR.PRINT.AREA",
        "CLEAR.ROUTING.SLIP",
        "CLOSE",
        "CLOSE.ALL",
        "CODE",
        "COLOR.PALETTE",
        "COLUMN",
        "COLUMN.WIDTH",
        "COLUMNS",
        "COMBIN",
        "COMBINATION",
        "COMPLEX",
        "CONCAT",
        "CONCATENATE",
        "CONFIDENCE",
        "CONSOLIDATE",
        "CONSTRAIN.NUMERIC",
        "CONVERT",
        "COPY",
        "COPY.CHART",
        "COPY.PICTURE",
        "COPY.TOOL",
        "CORREL",
        "COS",
        "COSH",
        "COUNT",
        "COUNTA",
        "COUNTBLANK",
        "COUNTIF",
        "COUNTIFS",
        "COUPDAYBS",
        "COUPDAYS",
        "COUPDAYSNC",
        "COUPNCD",
        "COUPNUM",
        "COUPPCD",
        "COVAR",
        "CREATE.NAMES",
        "CREATE.OBJECT",
        "CREATE.PUBLISHER",
        "CRITBINOM",
        "CUBEKPIMEMBER",
        "CUBEMEMBER",
        "CUBEMEMBERPROPERTY",
        "CUBERANKEDMEMBER",
        "CUBESET",
        "CUBESETCOUNT",
        "CUBEVALUE",
        "CUMIPMT",
        "CUMPRINC",
        "CUSTOM.REPEAT",
        "CUSTOM.UNDO",
        "CUSTOMIZE.TOOLBAR",
        "CUT",
        "DATA.DELETE",
        "DATA.FIND",
        "DATA.FIND.NEXT",
        "DATA.FIND.PREV",
        "DATA.FORM",
        "DATA.LABEL",
        "DATA.SERIES",
        "DATE",
        "DATEDIF",
        "DATESTRING",
        "DATEVALUE",
        "DAVERAGE",
        "DAY",
        "DAYS360",
        "DB",
        "DBCS",
        "DCOUNT",
        "DCOUNTA",
        "DDB",
        "DEC2BIN",
        "DEC2HEX",
        "DEC2OCT",
        "DEFINE.NAME",
        "DEFINE.STYLE",
        "DEGREES",
        "DELETE.ARROW",
        "DELETE.BAR",
        "DELETE.CHART.AUTOFORMAT",
        "DELETE.COMMAND",
        "DELETE.FORMAT",
        "DELETE.MENU",
        "DELETE.NAME",
        "DELETE.NOTE",
        "DELETE.OVERLAY",
        "DELETE.STYLE",
        "DELETE.TOOL",
        "DELETE.TOOLBAR",
        "DELTA",
        "DEMOTE",
        "DEREF",
        "DEVSQ",
        "DGET",
        "DIALOG.BOX",
        "DIRECTORY",
        "DISABLE.INPUT",
        "DISC",
        "DISPLAY",
        "DMAX",
        "DMIN",
        "DOCUMENTS",
        "DOLLAR",
        "DOLLARDE",
        "DOLLARFR",
        "DPRODUCT",
        "DSTDEV",
        "DSTDEVP",
        "DSUM",
        "DUPLICATE",
        "DURATION",
        "DVAR",
        "DVARP",
        "ECHO",
        "ECMA.CEILING",
        "EDATE",
        "EDIT.COLOR",
        "EDIT.DELETE",
        "EDIT.OBJECT",
        "EDIT.REPEAT",
        "EDIT.SERIES",
        "EDIT.TOOL",
        "EDITBOX.PROPERTIES",
        "EDITION.OPTIONS",
        "EFFECT",
        "ELSE",
        "ELSE.IF",
        "ENABLE.COMMAND",
        "ENABLE.OBJECT",
        "ENABLE.TIPWIZARD",
        "ENABLE.TOOL",
        "END.IF",
        "ENTER.DATA",
        "EOMONTH",
        "ERF",
        "ERFC",
        "ERROR",
        "ERROR.TYPE",
        "ERRORBAR.X",
        "ERRORBAR.Y",
        "EVALUATE",
        "EVEN",
        "EXACT",
        "EXEC",
        "EXECUTE",
        "EXP",
        "EXPONDIST",
        "EXTEND.POLYGON",
        "EXTRACT",
        "FACT",
        "FACTDOUBLE",
        "FALSE",
        "FCLOSE",
        "FDIST",
        "FILE.CLOSE",
        "FILE.DELETE",
        "FILES",
        "FILL.AUTO",
        "FILL.DOWN",
        "FILL.GROUP",
        "FILL.LEFT",
        "FILL.RIGHT",
        "FILL.UP",
        "FILTER",
        "FILTER.ADVANCED",
        "FILTER.SHOW.ALL",
        "FIND",
        "FIND.FILE",
        "FINDB",
        "FINV",
        "FISHER",
        "FISHERINV",
        "FIXED",
        "FLOOR",
        "FONT",
        "FONT.PROPERTIES",
        "FOPEN",
        "FOR",
        "FOR.CELL",
        "FORECAST",
        "FORECAST.ETS",
        "FORECAST.ETS.CONFINT",
        "FORECAST.ETS.SEASONALITY",
        "FORECAST.ETS.STAT",
        "FORECAST.LINEAR",
        "FORMAT.AUTO",
        "FORMAT.CHART",
        "FORMAT.CHARTTYPE",
        "FORMAT.FONT",
        "FORMAT.LEGEND",
        "FORMAT.MAIN",
        "FORMAT.MOVE",
        "FORMAT.NUMBER",
        "FORMAT.OVERLAY",
        "FORMAT.SHAPE",
        "FORMAT.SIZE",
        "FORMAT.TEXT",
        "FORMULA",
        "FORMULA.ARRAY",
        "FORMULA.CONVERT",
        "FORMULA.FILL",
        "FORMULA.FIND",
        "FORMULA.FIND.NEXT",
        "FORMULA.FIND.PREV",
        "FORMULA.GOTO",
        "FORMULA.REPLACE",
        "FPOS",
        "FREAD",
        "FREADLN",
        "FREEZE.PANES",
        "FREQUENCY",
        "FSIZE",
        "FTEST",
        "FULL",
        "FULL.SCREEN",
        "FUNCTION.WIZARD",
        "FV",
        "FVSCHEDULE",
        "FWRITE",
        "FWRITELN",
        "GALLERY.3D.AREA",
        "GALLERY.3D.BAR",
        "GALLERY.3D.COLUMN",
        "GALLERY.3D.LINE",
        "GALLERY.3D.PIE",
        "GALLERY.3D.SURFACE",
        "GALLERY.AREA",
        "GALLERY.BAR",
        "GALLERY.COLUMN",
        "GALLERY.CUSTOM",
        "GALLERY.DOUGHNUT",
        "GALLERY.LINE",
        "GALLERY.PIE",
        "GALLERY.RADAR",
        "GALLERY.SCATTER",
        "GAMMADIST",
        "GAMMAINV",
        "GAMMALN",
        "GCD",
        "GEOMEAN",
        "GESTEP",
        "GET.BAR",
        "GET.CELL",
        "GET.CHART.ITEM",
        "GET.DEF",
        "GET.DOCUMENT",
        "GET.FIELD",
        "GET.FORMULA",
        "GET.ITEM",
        "GET.LINK.INFO",
        "GET.MOVIE",
        "GET.NAME",
        "GET.NOTE",
        "GET.OBJECT",
        "GET.TOOL",
        "GET.TOOLBAR",
        "GET.VIEW",
        "GET.WINDOW",
        "GET.WORKBOOK",
        "GET.WORKSPACE",
        "GETPIVOTDATA",
        "GOAL.SEEK",
        "GOTO",
        "GRIDLINES",
        "GROUP",
        "GROWTH",
        "HALT",
        "HARMEAN",
        "HELP",
        "HEX2BIN",
        "HEX2DEC",
        "HEX2OCT",
        "HIDE",
        "HIDE.DIALOG",
        "HIDE.OBJECT",
        "HIDEALL.INKANNOTS",
        "HIDEALL.NOTES",
        "HIDECURR.NOTE",
        "HLINE",
        "HLOOKUP",
        "HOUR",
        "HPAGE",
        "HSCROLL",
        "HYPERLINK",
        "HYPGEOMDIST",
        "IF",
        "IFERROR",
        "IFS",
        "IMABS",
        "IMAGINARY",
        "IMARGUMENT",
        "IMCONJUGATE",
        "IMCOS",
        "IMDIV",
        "IMEXP",
        "IMLN",
        "IMLOG10",
        "IMLOG2",
        "IMPOWER",
        "IMPRODUCT",
        "IMREAL",
        "IMSIN",
        "IMSQRT",
        "IMSUB",
        "IMSUM",
        "INDEX",
        "INDIRECT",
        "INFO",
        "INITIATE",
        "INPUT",
        "INSERT",
        "INSERT.MAP.OBJECT",
        "INSERT.OBJECT",
        "INSERT.PICTURE",
        "INSERT.TITLE",
        "INSERTDATATABLE",
        "INT",
        "INTERCEPT",
        "INTRATE",
        "IPMT",
        "IRR",
        "ISBLANK",
        "ISERR",
        "ISERROR",
        "ISEVEN",
        "ISLOGICAL",
        "ISNA",
        "ISNONTEXT",
        "ISNUMBER",
        "ISO.CEILING",
        "ISODD",
        "ISPMT",
        "ISREF",
        "ISTEXT",
        "ISTHAIDIGIT",
        "JUSTIFY",
        "KURT",
        "LABEL.PROPERTIES",
        "LARGE",
        "LAST.ERROR",
        "LAYOUT",
        "LCM",
        "LEFT",
        "LEFTB",
        "LEGEND",
        "LEN",
        "LENB",
        "LINE.PRINT",
        "LINEST",
        "LINK.COMBO",
        "LINK.FORMAT",
        "LINKS",
        "LIST.NAMES",
        "LISTBOX.PROPERTIES",
        "LN",
        "LOG",
        "LOG10",
        "LOGEST",
        "LOGINV",
        "LOGNORMDIST",
        "LOOKUP",
        "LOWER",
        "MACRO.OPTIONS",
        "MAIL.ADD.MAILER",
        "MAIL.DELETE.MAILER",
        "MAIL.EDIT.MAILER",
        "MAIL.FORWARD",
        "MAIL.LOGOFF",
        "MAIL.LOGON",
        "MAIL.NEXT.LETTER",
        "MAIL.REPLY",
        "MAIL.REPLY.ALL",
        "MAIL.SEND.MAILER",
        "MAIN.CHART",
        "MAIN.CHART.TYPE",
        "MATCH",
        "MAX",
        "MAXA",
        "MAXIFS",
        "MDETERM",
        "MDURATION",
        "MEDIAN",
        "MENU.EDITOR",
        "MERGE.STYLES",
        "MESSAGE",
        "MID",
        "MIDB",
        "MIN",
        "MINA",
        "MINIFS",
        "MINUTE",
        "MINVERSE",
        "MIRR",
        "MMULT",
        "MOD",
        "MODE",
        "MONTH",
        "MOVE.BRK",
        "MOVE.TOOL",
        "MOVIE.COMMAND",
        "MROUND",
        "MSOCHECKS",
        "MULTINOMIAL",
        "N",
        "NA",
        "NAMES",
        "NEGBINOMDIST",
        "NETWORKDAYS",
        "NETWORKDAYS.INTL",
        "NEW",
        "NEW.WINDOW",
        "NEWWEBQUERY",
        "NEXT",
        "NOMINAL",
        "NORMAL",
        "NORMDIST",
        "NORMINV",
        "NORMSDIST",
        "NORMSINV",
        "NOT",
        "NOTE",
        "NOW",
        "NPER",
        "NPV",
        "NUMBERSTRING",
        "OBJECT.PROPERTIES",
        "OBJECT.PROTECTION",
        "OCT2BIN",
        "OCT2DEC",
        "OCT2HEX",
        "ODD",
        "ODDFPRICE",
        "ODDFYIELD",
        "ODDLPRICE",
        "ODDLYIELD",
        "OFFSET",
        "ON.DATA",
        "ON.DOUBLECLICK",
        "ON.ENTRY",
        "ON.KEY",
        "ON.RECALC",
        "ON.SHEET",
        "ON.TIME",
        "ON.WINDOW",
        "OPEN",
        "OPEN.DIALOG",
        "OPEN.LINKS",
        "OPEN.MAIL",
        "OPEN.TEXT",
        "OPTIONS.CALCULATION",
        "OPTIONS.CHART",
        "OPTIONS.EDIT",
        "OPTIONS.GENERAL",
        "OPTIONS.LISTS.ADD",
        "OPTIONS.LISTS.DELETE",
        "OPTIONS.LISTS.GET",
        "OPTIONS.ME",
        "OPTIONS.MENONO",
        "OPTIONS.SAVE",
        "OPTIONS.SPELL",
        "OPTIONS.TRANSITION",
        "OPTIONS.VIEW",
        "OR",
        "OUTLINE",
        "OVERLAY",
        "OVERLAY.CHART.TYPE",
        "PAGE.SETUP",
        "PARSE",
        "PASTE",
        "PASTE.LINK",
        "PASTE.PICTURE",
        "PASTE.PICTURE.LINK",
        "PASTE.SPECIAL",
        "PASTE.TOOL",
        "PATTERNS",
        "PAUSE",
        "PEARSON",
        "PERCENTILE",
        "PERCENTRANK",
        "PERMUT",
        "PHONETIC",
        "PI",
        "PICKLIST",
        "PIVOT.ADD.FIELDS",
        "PIVOT.FIELD",
        "PIVOT.FIELD.GROUP",
        "PIVOT.FIELD.PROPERTIES",
        "PIVOT.FIELD.UNGROUP",
        "PIVOT.ITEM",
        "PIVOT.ITEM.PROPERTIES",
        "PIVOT.REFRESH",
        "PIVOT.SHOW.PAGES",
        "PIVOT.TABLE.CHART",
        "PIVOT.TABLE.WIZARD",
        "PMT",
        "POISSON",
        "POKE",
        "POST.DOCUMENT",
        "POWER",
        "PPMT",
        "PRECISION",
        "PREFERRED",
        "PRESS.TOOL",
        "PRICE",
        "PRICEDISC",
        "PRICEMAT",
        "PRINT",
        "PRINT.PREVIEW",
        "PRINTER.SETUP",
        "PROB",
        "PRODUCT",
        "PROMOTE",
        "PROPER",
        "PROTECT.DOCUMENT",
        "PROTECT.REVISIONS",
        "PUSHBUTTON.PROPERTIES",
        "PV",
        "QUARTILE",
        "QUIT",
        "QUOTIENT",
        "RADIANS",
        "RAND",
        "RANDBETWEEN",
        "RANK",
        "RATE",
        "RECEIVED",
        "REFTEXT",
        "REGISTER",
        "REGISTER.ID",
        "RELREF",
        "REMOVE.LIST.ITEM",
        "REMOVE.PAGE.BREAK",
        "RENAME.COMMAND",
        "RENAME.OBJECT",
        "REPLACE",
        "REPLACE.FONT",
        "REPLACEB",
        "REPT",
        "REQUEST",
        "RESET.TOOL",
        "RESET.TOOLBAR",
        "RESTART",
        "RESULT",
        "RESUME",
        "RETURN",
        "RIGHT",
        "RIGHTB",
        "RM.PRINT.AREA",
        "ROMAN",
        "ROUND",
        "ROUNDBAHTDOWN",
        "ROUNDBAHTUP",
        "ROUNDDOWN",
        "ROUNDUP",
        "ROUTE.DOCUMENT",
        "ROUTING.SLIP",
        "ROW",
        "ROW.HEIGHT",
        "ROWS",
        "RSQ",
        "RTD",
        "RUN",
        "SAVE",
        "SAVE.AS",
        "SAVE.COPY.AS",
        "SAVE.DIALOG",
        "SAVE.NEW.OBJECT",
        "SAVE.TOOLBAR",
        "SAVE.WORKBOOK",
        "SAVE.WORKSPACE",
        "SCALE",
        "SCENARIO.ADD",
        "SCENARIO.CELLS",
        "SCENARIO.DELETE",
        "SCENARIO.EDIT",
        "SCENARIO.GET",
        "SCENARIO.MERGE",
        "SCENARIO.SHOW",
        "SCENARIO.SHOW.NEXT",
        "SCENARIO.SUMMARY",
        "SCROLLBAR.PROPERTIES",
        "SEARCH",
        "SEARCHB",
        "SECOND",
        "SELECT",
        "SELECT.ALL",
        "SELECT.CHART",
        "SELECT.END",
        "SELECT.LAST.CELL",
        "SELECT.LIST.ITEM",
        "SELECT.PLOT.AREA",
        "SELECT.SPECIAL",
        "SELECTION",
        "SEND.KEYS",
        "SEND.MAIL",
        "SEND.TO.BACK",
        "SERIES",
        "SERIES.AXES",
        "SERIES.ORDER",
        "SERIES.X",
        "SERIES.Y",
        "SERIESSUM",
        "SET.CONTROL.VALUE",
        "SET.CRITERIA",
        "SET.DATABASE",
        "SET.DIALOG.DEFAULT",
        "SET.DIALOG.FOCUS",
        "SET.EXTRACT",
        "SET.LIST.ITEM",
        "SET.NAME",
        "SET.PAGE.BREAK",
        "SET.PREFERRED",
        "SET.PRINT.AREA",
        "SET.PRINT.TITLES",
        "SET.UPDATE.STATUS",
        "SET.VALUE",
        "SHARE",
        "SHARE.NAME",
        "SHEET.BACKGROUND",
        "SHORT.MENUS",
        "SHOW.ACTIVE.CELL",
        "SHOW.BAR",
        "SHOW.CLIPBOARD",
        "SHOW.DETAIL",
        "SHOW.DIALOG",
        "SHOW.INFO",
        "SHOW.LEVELS",
        "SHOW.TOOLBAR",
        "SIGN",
        "SIN",
        "SINH",
        "SKEW",
        "SLN",
        "SLOPE",
        "SMALL",
        "SORT",
        "SORT.SPECIAL",
        "SOUND.NOTE",
        "SOUND.PLAY",
        "SPELLING",
        "SPELLING.CHECK",
        "SPLIT",
        "SPREADBASE.DATA.FIELD",
        "SQRT",
        "SQRTPI",
        "STANDARD.FONT",
        "STANDARD.WIDTH",
        "STANDARDIZE",
        "STDEV",
        "STDEVA",
        "STDEVP",
        "STDEVPA",
        "STEP",
        "STEYX",
        "STYLE",
        "SUBSCRIBE.TO",
        "SUBSTITUTE",
        "SUBTOTAL",
        "SUBTOTAL.CREATE",
        "SUBTOTAL.REMOVE",
        "SUM",
        "SUMIF",
        "SUMIFS",
        "SUMMARY.INFO",
        "SUMPRODUCT",
        "SUMSQ",
        "SUMX2MY2",
        "SUMX2PY2",
        "SUMXMY2",
        "SWITCH",
        "SYD",
        "T",
        "TAB.ORDER",
        "TABLE",
        "TAN",
        "TANH",
        "TBILLEQ",
        "TBILLPRICE",
        "TBILLYIELD",
        "TDIST",
        "TERMINATE",
        "TEXT",
        "TEXT.BOX",
        "TEXT.TO.COLUMNS",
        "TEXTJOIN",
        "TEXTREF",
        "THAIDAYOFWEEK",
        "THAIDIGIT",
        "THAIMONTHOFYEAR",
        "THAINUMSOUND",
        "THAINUMSTRING",
        "THAISTRINGLENGTH",
        "THAIYEAR",
        "TIME",
        "TIMEVALUE",
        "TINV",
        "TODAY",
        "TRACER.CLEAR",
        "TRACER.DISPLAY",
        "TRACER.ERROR",
        "TRACER.NAVIGATE",
        "TRANSPOSE",
        "TRAVERSE.NOTES",
        "TREND",
        "TRIM",
        "TRIMMEAN",
        "TRUE",
        "TRUNC",
        "TTEST",
        "TYPE",
        "UNDO",
        "UNGROUP",
        "UNGROUP.SHEETS",
        "UNHIDE",
        "UNLOCKED.NEXT",
        "UNLOCKED.PREV",
        "UNPROTECT.REVISIONS",
        "UNREGISTER",
        "UPDATE.LINK",
        "UPPER",
        "USDOLLAR",
        "VALUE",
        "VAR",
        "VARA",
        "VARP",
        "VARPA",
        "VBA.INSERT.FILE",
        "VBA.MAKE.ADDIN",
        "VBA.PROCEDURE.DEFINITION",
        "VBAACTIVATE",
        "VDB",
        "VIEW.3D",
        "VIEW.DEFINE",
        "VIEW.DELETE",
        "VIEW.GET",
        "VIEW.SHOW",
        "VLINE",
        "VLOOKUP",
        "VOLATILE",
        "VPAGE",
        "VSCROLL",
        "WAIT",
        "WEB.PUBLISH",
        "WEEKDAY",
        "WEEKNUM",
        "WEIBULL",
        "WHILE",
        "WINDOW.MAXIMIZE",
        "WINDOW.MINIMIZE",
        "WINDOW.MOVE",
        "WINDOW.RESTORE",
        "WINDOW.SIZE",
        "WINDOW.TITLE",
        "WINDOWS",
        "WORKBOOK.ACTIVATE",
        "WORKBOOK.ADD",
        "WORKBOOK.COPY",
        "WORKBOOK.DELETE",
        "WORKBOOK.HIDE",
        "WORKBOOK.INSERT",
        "WORKBOOK.MOVE",
        "WORKBOOK.NAME",
        "WORKBOOK.NEW",
        "WORKBOOK.NEXT",
        "WORKBOOK.OPTIONS",
        "WORKBOOK.PREV",
        "WORKBOOK.PROTECT",
        "WORKBOOK.SCROLL",
        "WORKBOOK.SELECT",
        "WORKBOOK.TAB.SPLIT",
        "WORKBOOK.UNHIDE",
        "WORKDAY",
        "WORKDAY.INTL",
        "WORKGROUP",
        "WORKGROUP.OPTIONS",
        "WORKSPACE",
        "XIRR",
        "XNPV",
        "YEAR",
        "YEARFRAC",
        "YIELD",
        "YIELDDISC",
        "YIELDMAT",
        "ZOOM",
        "ZTEST",
    }
)

SUPPORTED_FUNCTIONS = frozenset(
    [
        "ABS",
        "AND",
        "AVERAGE",
        "CHOOSE",
        "COUNT",
        "COUNTA",
        "COVAR",
        "DATE",
        "DAY",
        "FALSE",
        "HLOOKUP",
        "HOUR",
        "IF",
        "IFERROR",
        "INDEX",
        "INT",
        "LEFT",
        "MATCH",
        "MAX",
        "MEDIAN",
        "MIN",
        "MINUTE",
        "MOD",
        "MONTH",
        "NA",
        "NOT",
        "OR",
        "POWER",
        "ROUND",
        "ROUNDDOWN",
        "ROUNDUP",
        "SECOND",
        "SQRT",
        "STDEV",
        "SUM",
        "TEXT",
        "TIME",
        "TRUE",
        "VLOOKUP",
        "YEAR",
    ]
)
_SPACE = " \t\r\n"
_NUMBER = re.compile(r"(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?")
_CELL = re.compile(r"(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]{0,6})\Z")
_CELL_LIKE = re.compile(r"\$?[A-Za-z]+\$?[0-9]+\Z")
_AXIS_COLUMN = re.compile(r"(\$?)([A-Za-z]{1,3})\Z")
_AXIS_ROW = re.compile(r"(\$?)([1-9][0-9]{0,6})\Z")
_ERRORS = (
    "#GETTING_DATA",
    "#DIV/0!",
    "#VALUE!",
    "#SPILL!",
    "#NAME?",
    "#NULL!",
    "#CALC!",
    "#REF!",
    "#NUM!",
    "#N/A",
)


@dataclass(frozen=True)
class Endpoint:
    row: int
    column: int
    row_absolute: bool
    column_absolute: bool


@dataclass(frozen=True)
class AxisEndpoint:
    axis: Literal["row", "column"]
    index: int
    absolute: bool


@dataclass(frozen=True)
class ReferenceToken:
    kind: ReferenceKind
    source_span: tuple[int, int]
    occurrence_index: int
    qualifier: str | None
    first: Endpoint | AxisEndpoint | None
    last: Endpoint | AxisEndpoint | None
    name: str | None


@dataclass(frozen=True)
class ParsedFormula:
    text: str
    references: tuple[ReferenceToken, ...]
    max_nesting: int


@dataclass(frozen=True)
class FormulaProblem:
    code: str
    source_span: tuple[int, int] | None
    reason: str
    information: tuple[tuple[str, str | int], ...] = ()


def _problem(code: str, start: int, end: int, reason: str) -> FormulaProblem:
    return FormulaProblem(code, (start, end), reason)


def _endpoint(text: str) -> Endpoint | None:
    match = _CELL.fullmatch(text)
    if match is None:
        return None
    col = 0
    for char in match[2].upper():
        col = col * 26 + ord(char) - ord("A") + 1
    row = int(match[4])
    if row > 1_048_576 or col > 16_384:
        return None
    return Endpoint(row, col, bool(match[3]), bool(match[1]))


def _axis_endpoint(text: str) -> AxisEndpoint | None:
    """Called only at a range site; bare axis-like words remain names/literals."""
    column = _AXIS_COLUMN.fullmatch(text)
    if column is not None:
        index = 0
        for char in column[2].upper():
            index = index * 26 + ord(char) - ord("A") + 1
        return AxisEndpoint("column", index, bool(column[1])) if index <= 16384 else None
    row = _AXIS_ROW.fullmatch(text)
    if row is not None:
        index = int(row[2])  # bounded ASCII length before conversion
        return AxisEndpoint("row", index, bool(row[1])) if index <= 1048576 else None
    return None


def _word_end(text: str, start: int) -> int:
    end = start
    while end < len(text) and (text[end].isalnum() or text[end] in "_.\\$?"):
        end += 1
    return end


def _skip_space(text: str, start: int) -> int:
    while start < len(text) and text[start] in _SPACE:
        start += 1
    return start


def _quoted_end(text: str, start: int, quote: str) -> tuple[int, str] | None:
    position = start + 1
    decoded: list[str] = []
    while position < len(text):
        char = text[position]
        if char == quote:
            if position + 1 < len(text) and text[position + 1] == quote:
                decoded.append(quote)
                position += 2
                continue
            return position + 1, "".join(decoded)
        decoded.append(char)
        position += 1
    return None


def _read_operand(
    text: str, start: int, occurrence: int
) -> tuple[int, ReferenceToken | None, bool] | FormulaProblem:
    """Read an atomic value/reference or the opening of an admitted call."""
    char = text[start]
    if char == '"':
        quoted = _quoted_end(text, start, '"')
        if quoted is None:
            return _problem("formula_parse_failure", start, len(text), "unterminated string")
        return quoted[0], None, False
    if char == "#":
        upper = text[start:].upper()
        for error in _ERRORS:
            if upper.startswith(error):
                if error == "#REF!":
                    return _problem(
                        "broken_reference", start, start + len(error), "broken reference"
                    )
                return start + len(error), None, False
        return _problem("formula_parse_failure", start, start + 1, "unknown error literal")
    if char in "@{":
        return _problem(
            "unsupported_formula_syntax", start, start + 1, "unsupported operator/array"
        )
    if char == "[":
        code = (
            "external_workbook_reference"
            if "!" in text[start:]
            else "unsupported_structured_reference"
        )
        return _problem(code, start, len(text), "bracketed reference")
    word_end = _word_end(text, start)
    has_qualifier = word_end < len(text) and text[word_end] == "!"
    if (
        not has_qualifier
        and char.isascii()
        and (
            char.isdigit()
            or (
                char == "."
                and start + 1 < len(text)
                and text[start + 1].isascii()
                and text[start + 1].isdigit()
            )
        )
    ):
        match = _NUMBER.match(text, start)
        if match is not None:
            end = match.end()
            after_number = _skip_space(text, end)
            if after_number == len(text) or text[after_number] != ":":
                return end, None, False

    qualifier: str | None = None
    position = start
    if char == "'":
        quoted = _quoted_end(text, start, "'")
        if quoted is None:
            return _problem(
                "formula_parse_failure", start, len(text), "unterminated sheet qualifier"
            )
        position, qualifier = quoted
        if not qualifier:
            return _problem("formula_parse_failure", start, position, "empty sheet qualifier")
        if position >= len(text) or text[position] != "!":
            return _problem("formula_parse_failure", start, position, "quoted name requires !")
        position += 1
    else:
        end = _word_end(text, start)
        if end == start:
            return _problem("formula_parse_failure", start, start + 1, "expected operand")
        word = text[start:end]
        after = _skip_space(text, end)
        if after < len(text) and text[after] == "(":
            upper_word = word.upper()
            if upper_word not in SUPPORTED_FUNCTIONS:
                code = (
                    "dynamic_reference_function"
                    if upper_word in {"INDIRECT", "OFFSET"}
                    else "unsupported_function"
                )
                return _problem(code, start, end, "function not admitted")
            return after + 1, None, True
        if end < len(text) and text[end] == "!":
            if not all(c.isalnum() or c in "_." for c in word):
                return _problem("formula_parse_failure", start, end, "invalid unquoted sheet")
            qualifier = word
            position = end + 1
        elif word.upper() in {"TRUE", "FALSE"}:
            return end, None, False

    if qualifier is not None:
        if "[" in qualifier or "]" in qualifier:
            return _problem("external_workbook_reference", start, position, "external qualifier")
        if ":" in qualifier:
            return _problem("three_dimensional_reference", start, position, "sheet range qualifier")
        if text[position : position + 5].upper() == "#REF!":
            return _problem("broken_reference", start, position + 5, "broken qualified reference")
    end = _word_end(text, position)
    if end == position:
        return _problem(
            "formula_parse_failure", start, min(position + 1, len(text)), "missing reference"
        )
    word = text[position:end]
    if end < len(text) and text[end] == "[":
        return _problem(
            "unsupported_structured_reference", start, len(text), "structured reference"
        )
    first = _endpoint(word)
    colon = _skip_space(text, end)
    if colon < len(text) and text[colon] == ":":
        next_start = _skip_space(text, colon + 1)
        next_end = _word_end(text, next_start)
        if next_end < len(text) and text[next_end] == "!":
            return _problem(
                "three_dimensional_reference", start, next_end + 1, "sheet range qualifier"
            )
        second_word = text[next_start:next_end]
        last = _endpoint(second_word)
        if first is None or last is None:
            first_axis, last_axis = _axis_endpoint(word), _axis_endpoint(second_word)
            if first_axis is None or last_axis is None or first_axis.axis != last_axis.axis:
                return _problem("formula_parse_failure", start, next_end, "invalid range endpoints")
            return (
                next_end,
                ReferenceToken(
                    "range", (start, next_end), occurrence, qualifier, first_axis, last_axis, None
                ),
                False,
            )
        return (
            next_end,
            ReferenceToken("range", (start, next_end), occurrence, qualifier, first, last, None),
            False,
        )
    if first is not None:
        return (
            end,
            ReferenceToken("cell", (start, end), occurrence, qualifier, first, first, None),
            False,
        )
    if "$" in word or any(c.isdigit() and not c.isascii() for c in word):
        return _problem("formula_parse_failure", start, end, "invalid coordinate characters")
    if _CELL_LIKE.fullmatch(word) is not None and len(word) > 255:
        return _problem("formula_parse_failure", start, end, "excessive coordinate digits")
    return (
        end,
        ReferenceToken("defined_name", (start, end), occurrence, qualifier, None, None, word),
        False,
    )


def parse_formula(text: str, *, max_chars: int, max_nesting: int) -> ParsedFormula | FormulaProblem:
    """Admit a complete supported expression, atomically and without recursion."""
    if len(text) > max_chars:
        return FormulaProblem(
            "formula_limit_exceeded",
            (0, len(text)),
            "formula too long",
            (("observed_chars", len(text)),),
        )
    refs: list[ReferenceToken] = []
    # True denotes a call frame; False denotes ordinary grouping.
    frames: list[bool] = []
    want_operand = True
    can_empty = False
    last_reference = False
    max_depth = 0
    position = 0

    def failure(code: str, start: int, end: int, reason: str) -> FormulaProblem:
        return FormulaProblem(code, (start, end), reason, (("observed_nesting", max_depth),))

    while position < len(text):
        prior = position
        position = _skip_space(text, position)
        if position == len(text):
            break
        char = text[position]
        if char == ")":
            if not frames or (want_operand and not can_empty):
                return failure(
                    "formula_parse_failure", position, position + 1, "unexpected close/empty group"
                )
            frames.pop()
            want_operand, can_empty, last_reference = False, False, False
            position += 1
            continue
        if char == ",":
            if not frames or not frames[-1]:
                return failure(
                    "union_or_intersection_reference", position, position + 1, "union outside call"
                )
            if want_operand and not can_empty:
                return failure(
                    "formula_parse_failure", position, position + 1, "unfinished argument"
                )
            want_operand, can_empty, last_reference = True, True, False
            position += 1
            continue
        if want_operand:
            if char in "+-":
                can_empty = False
                position += 1
                continue
            if char == "(":
                frames.append(False)
                position += 1
                can_empty = False
            else:
                atom = _read_operand(text, position, len(refs))
                if isinstance(atom, FormulaProblem):
                    return replace(
                        atom, information=(*atom.information, ("observed_nesting", max_depth))
                    )
                position, ref, is_call = atom
                if ref is not None:
                    refs.append(ref)
                if is_call:
                    frames.append(True)
                    can_empty = True
                else:
                    want_operand, can_empty = False, False
                    last_reference = ref is not None
            max_depth = max(max_depth, len(frames))
            if max_depth > max_nesting:
                return FormulaProblem(
                    "formula_limit_exceeded",
                    (prior, position),
                    "formula nesting limit",
                    (("observed_nesting", max_depth),),
                )
            continue
        if char == "%":
            last_reference = False
            position += 1
            continue
        if char in "+-*/^&=<>":
            position += 1
            if char in "<>" and position < len(text) and text[position] in "=>":
                pair = char + text[position]
                if pair in {"<=", ">=", "<>"}:
                    position += 1
            want_operand, can_empty, last_reference = True, False, False
            continue
        if char in "#@{":
            return failure(
                "unsupported_formula_syntax", position, position + 1, "unsupported postfix/operator"
            )
        code = (
            "union_or_intersection_reference"
            if last_reference and position > prior
            else "formula_parse_failure"
        )
        return failure(code, position, position + 1, "unexpected adjacent operand/token")
    if frames or want_operand:
        return failure(
            "formula_parse_failure", max(0, len(text) - 1), len(text), "unfinished expression"
        )
    return ParsedFormula(text, tuple(refs), max_depth)


@dataclass(frozen=True)
class NameIndex:
    entries: Mapping[tuple[str | None, str], tuple[DefinedName, ...]]

    def __post_init__(self) -> None:
        object.__setattr__(
            self,
            "entries",
            MappingProxyType({key: tuple(value) for key, value in self.entries.items()}),
        )


@dataclass(frozen=True)
class ReferenceContext:
    sheets: Mapping[str, SheetEntry]
    names: NameIndex | None

    def __post_init__(self) -> None:
        object.__setattr__(self, "sheets", MappingProxyType(dict(self.sheets)))


@dataclass(frozen=True)
class ResolvedReference:
    token: ReferenceToken
    sheet: str
    min_row: int
    min_column: int
    max_row: int
    max_column: int
    definition: DefinedName | None


def build_name_index(definitions: tuple[DefinedName, ...]) -> NameIndex:
    buckets: dict[tuple[str | None, str], list[DefinedName]] = {}
    for definition in definitions:
        scope = None if definition.scope_sheet is None else definition.scope_sheet.casefold()
        buckets.setdefault((scope, definition.name.casefold()), []).append(definition)
    return NameIndex({key: tuple(value) for key, value in buckets.items()})


def build_reference_context(
    registry: WorkbookRegistry, *, names: NameIndex | None
) -> ReferenceContext:
    return ReferenceContext({sheet.name.casefold(): sheet for sheet in registry.sheets}, names)


_NAME = re.compile(r"[A-Za-z_\\][A-Za-z0-9_.\\]{0,254}\Z")
_R1C1 = re.compile(r"(?:R([0-9]{1,7})(?:C([0-9]{1,5}))?|C([0-9]{1,5})(?:R([0-9]{1,7}))?)\Z", re.I)


def _valid_name(name: str) -> bool:
    if _NAME.fullmatch(name) is None:
        return False
    upper = name.upper()
    if upper in _RESERVED_NAMES or upper in {"R", "C"} or _endpoint(name) is not None:
        return False
    if upper.startswith(("_XLFN.", "_XLWS.", "XLPM.", "XLOP.")):
        return False
    match = _R1C1.fullmatch(name)
    if match is not None:
        row = match[1] or match[4]
        col = match[2] or match[3]
        valid_row = row is None or 1 <= int(row) <= 1_048_576
        valid_col = col is None or 1 <= int(col) <= 16_384
        if valid_row and valid_col:
            return False
    return True


def _resolve_bounds(
    token: ReferenceToken,
    context: ReferenceContext,
    formula_sheet: str,
    row_offset: int,
    column_offset: int,
    definition: DefinedName | None = None,
) -> ResolvedReference | FormulaProblem:
    qualifier = formula_sheet if token.qualifier is None else token.qualifier
    sheet = context.sheets.get(qualifier.casefold())
    if sheet is None or sheet.kind != "worksheet":
        return FormulaProblem(
            "unresolved_sheet_reference", token.source_span, "target is not a registered worksheet"
        )
    if token.first is None or token.last is None:
        return FormulaProblem("formula_parse_failure", token.source_span, "missing endpoints")
    if isinstance(token.first, AxisEndpoint) and isinstance(token.last, AxisEndpoint):
        first, last = token.first, token.last
        if first.axis != last.axis:
            return FormulaProblem("formula_parse_failure", token.source_span, "mixed range axes")
        offset = row_offset if first.axis == "row" else column_offset
        limit = 1048576 if first.axis == "row" else 16384
        indices = tuple(e.index + (0 if e.absolute else offset) for e in (first, last))
        if not all(1 <= index <= limit for index in indices):
            return FormulaProblem(
                "shared_reference_out_of_bounds",
                token.source_span,
                "translated reference leaves worksheet grid",
            )
        low, high = min(indices), max(indices)
        bounds = (low, 1, high, 16384) if first.axis == "row" else (1, low, 1048576, high)
        return ResolvedReference(token, sheet.name, *bounds, definition)
    if not isinstance(token.first, Endpoint) or not isinstance(token.last, Endpoint):
        return FormulaProblem("formula_parse_failure", token.source_span, "mixed endpoint shapes")
    coordinates: list[tuple[int, int]] = []
    for endpoint in (token.first, token.last):
        row = endpoint.row + (0 if endpoint.row_absolute else row_offset)
        col = endpoint.column + (0 if endpoint.column_absolute else column_offset)
        if not (1 <= row <= 1_048_576 and 1 <= col <= 16_384):
            return FormulaProblem(
                "shared_reference_out_of_bounds",
                token.source_span,
                "translated reference leaves worksheet grid",
            )
        coordinates.append((row, col))
    (row1, col1), (row2, col2) = coordinates
    return ResolvedReference(
        token,
        sheet.name,
        min(row1, row2),
        min(col1, col2),
        max(row1, row2),
        max(col1, col2),
        definition,
    )


def _absolute_endpoint(endpoint: Endpoint | AxisEndpoint | None) -> bool:
    if isinstance(endpoint, AxisEndpoint):
        return endpoint.absolute
    return isinstance(endpoint, Endpoint) and endpoint.row_absolute and endpoint.column_absolute


def _resolve_name(
    token: ReferenceToken,
    context: ReferenceContext,
    formula_sheet: str,
    max_chars: int,
) -> ResolvedReference | FormulaProblem:
    name = token.name
    if name is None or not _valid_name(name):
        return FormulaProblem(
            "invalid_defined_name",
            token.source_span,
            "used name outside supported/reserved identifier grammar",
        )
    qualifier = formula_sheet if token.qualifier is None else token.qualifier
    scope = context.sheets.get(qualifier.casefold())
    if scope is None or scope.kind != "worksheet":
        return FormulaProblem(
            "unresolved_sheet_reference", token.source_span, "name scope not a worksheet"
        )
    if context.names is None:
        return FormulaProblem(
            "unresolved_defined_name", token.source_span, "name inventory unavailable"
        )
    entries = context.names.entries
    matches = entries.get((scope.name.casefold(), name.casefold()))
    if matches is None:
        matches = entries.get((None, name.casefold()), ())
    if not matches:
        return FormulaProblem(
            "unresolved_defined_name", token.source_span, "no selected definition"
        )
    if len(matches) != 1:
        return FormulaProblem(
            "ambiguous_defined_name", token.source_span, "duplicate selected definitions"
        )
    definition = matches[0]
    parsed = parse_formula(definition.text, max_chars=max_chars, max_nesting=1)
    if isinstance(parsed, FormulaProblem):
        observed = dict(parsed.information).get("observed_nesting", 0)
        observed_depth = observed if isinstance(observed, int) else 0
        code = (
            "formula_limit_exceeded"
            if parsed.code == "formula_limit_exceeded" and len(definition.text) > max_chars
            else "unsupported_name_definition"
        )
        return FormulaProblem(
            code,
            token.source_span,
            "definition is not a fixed reference",
            (
                ("definition_text", definition.text),
                ("definition_reason", parsed.reason),
                ("observed_definition_chars", len(definition.text)),
                ("observed_definition_nesting", observed_depth),
            ),
        )
    if len(parsed.references) != 1:
        return FormulaProblem(
            "unsupported_name_definition",
            token.source_span,
            "definition not one reference",
            (
                ("observed_definition_chars", len(definition.text)),
                ("observed_definition_nesting", parsed.max_nesting),
            ),
        )
    fixed = parsed.references[0]
    before, after = fixed.source_span
    if (
        fixed.qualifier is None
        or not _absolute_endpoint(fixed.first)
        or not _absolute_endpoint(fixed.last)
        or definition.text[:before].strip(_SPACE)
        or definition.text[after:].strip(_SPACE)
    ):
        return FormulaProblem(
            "unsupported_name_definition",
            token.source_span,
            "definition must be qualified absolute cell/rectangle",
            (
                ("observed_definition_chars", len(definition.text)),
                ("observed_definition_nesting", parsed.max_nesting),
            ),
        )
    result = _resolve_bounds(fixed, context, formula_sheet, 0, 0, definition)
    if isinstance(result, FormulaProblem):
        return FormulaProblem(
            result.code,
            token.source_span,
            result.reason,
            (
                ("definition_text", definition.text),
                ("observed_definition_chars", len(definition.text)),
            ),
        )
    return ResolvedReference(
        token,
        result.sheet,
        result.min_row,
        result.min_column,
        result.max_row,
        result.max_column,
        definition,
    )


def resolve_references(
    parsed: ParsedFormula,
    context: ReferenceContext,
    *,
    formula_sheet: str,
    row_offset: int = 0,
    column_offset: int = 0,
    max_chars: int,
) -> tuple[ResolvedReference, ...] | FormulaProblem:
    """Resolve all sites or none. Fixed definitions never inherit shared offsets."""
    resolved: list[ResolvedReference] = []
    observed_definition_chars = 0
    for token in parsed.references:
        if token.kind == "defined_name":
            result = _resolve_name(token, context, formula_sheet, max_chars)
        else:
            result = _resolve_bounds(token, context, formula_sheet, row_offset, column_offset)
        if isinstance(result, FormulaProblem):
            information = dict(result.information)
            previous = information.get("observed_definition_chars", 0)
            if isinstance(previous, int):
                observed_definition_chars = max(observed_definition_chars, previous)
            information["observed_definition_chars"] = observed_definition_chars
            return replace(result, information=tuple(information.items()))
        if result.definition is not None:
            observed_definition_chars = max(observed_definition_chars, len(result.definition.text))
        resolved.append(result)
    return tuple(resolved)
