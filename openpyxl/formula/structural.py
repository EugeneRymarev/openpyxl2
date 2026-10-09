"""Rewrite A1 references for structural edits, without evaluating formulas.

The axis transform is supplied by the worksheet's atomic edit plan. Copy/paste
translation is different: dollar signs do not prevent structural adjustment.
"""

import re

from openpyxl.formula.tokenizer import Token
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.formula.tokenizer import TokenizerError
from openpyxl.utils.cell import column_index_from_string
from openpyxl.utils.cell import get_column_letter


class FormulaTranslationError(ValueError):
    """A formula cannot be safely adjusted by a structural worksheet edit."""


_CELL = re.compile(r"(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]*)\Z")
_COL = re.compile(r"(\$?)([A-Za-z]{1,3})\Z")
_ROW = re.compile(r"(\$?)([1-9][0-9]*)\Z")


def parse_reference(value):
    if value.endswith("#REF!"):
        qualifier = value[:-5]
        if not qualifier or qualifier.endswith("!"):
            return qualifier[:-1], "!" if qualifier else "", "#REF!", None, None
    prefix, bang, coordinate = value.rpartition("!")
    if not bang:
        coordinate = value
    parts = coordinate.split(":")
    if len(parts) > 2:
        return None
    for pattern in ([_CELL] if len(parts) == 1 else [_CELL, _COL, _ROW]):
        matches = [pattern.fullmatch(part) for part in parts]
        if not all(matches):
            continue
        groups = [m.groups() for m in matches]
        for group in groups:
            if pattern in (_CELL, _COL) and column_index_from_string(group[1]) > 16384:
                return None  # Out-of-grid identifiers can be defined names.
            if (
                pattern in (_CELL, _ROW)
                and int(group[3 if pattern is _CELL else 1]) > 1048576
            ):
                return None
        return prefix, bang, coordinate, pattern, groups
    return None


def rewrite_reference(value, parsed, owner, ws, axis, edit, formula=False):
    prefix, bang, coordinate, pattern, parts = parsed
    if pattern is None:
        return value
    title = prefix if bang else owner
    if bang and title.startswith("'") and title.endswith("'"):
        title = title[1:-1].replace("''", "'")
    if title is None or "[" in title:
        return value
    if ":" in title:
        if formula:
            raise FormulaTranslationError(
                "3D and separately qualified range endpoints are not supported"
            )
        return value
    if title.casefold() != ws.title.casefold():
        return value
    if (pattern is _COL and axis == "row") or (pattern is _ROW and axis == "column"):
        return value
    component = 3 if pattern is _CELL and axis == "row" else 1
    values = [
        int(p[component]) if axis == "row" else column_index_from_string(p[component])
        for p in parts
    ]
    start, end = sorted((values[0], values[-1]))
    # Full-axis references remain full-axis when inserting or deleting within it.
    if formula and len(parts) == 2 and start == 1 and end == edit.limit:
        return value
    try:
        result = edit.interval(start, end)
    except OverflowError:
        if not formula:
            raise
        # Reference overflow is an Excel #REF!, unlike moving an existing cell
        # off the sheet (which the enclosing edit plan rejects).
        first = start + edit.count if start >= edit.index else start
        last = end + edit.count if end >= edit.index else end
        result = (first, min(last, edit.limit)) if first <= edit.limit else None
    if result is None:
        return (prefix + bang if formula and bang else "") + "#REF!"
    if values[0] > values[-1]:
        result = result[::-1]
    if result == (values[0], values[-1]):
        return value
    out = []
    for part, position in zip(parts, result):
        part = list(part)
        replacement = str(position) if axis == "row" else get_column_letter(position)
        if axis == "column" and part[component].islower():
            replacement = replacement.lower()
        part[component] = replacement
        out.append("".join(part))
    return (prefix + bang if bang else "") + ":".join(out)


class _Tokenizer(Tokenizer):
    """Keep source whitespace; recognise @ and spill suffixes without regexing strings."""

    def _parse_whitespace(self):
        value = self.WSPACE_RE.match(self.formula[self.offset :]).group()
        self.items.append(Token(value, Token.WSPACE))
        return len(value)

    def _parse_string(self):
        if self.formula[self.offset] == "'" and self.token == ["@"]:
            self.items.append(Token("@", Token.OP_PRE))
            self.token.clear()
        return super()._parse_string()

    def _parse_error(self):
        if self.token == ["@"]:
            self.items.append(Token("@", Token.OP_PRE))
            self.token.clear()
        if self.token and self.token[-1] != "!":
            following = self.formula[self.offset + 1 : self.offset + 2]
            if following and following not in self.TOKEN_ENDERS + "\n":
                raise TokenizerError("Invalid spill suffix")
            self.token.append("#")
            return 1
        return super()._parse_error()


def rewrite_formula(value, owner, ws, axis, edit):
    """Rewrite only reference operands; keep literals, identifiers and formatting."""
    if not isinstance(value, str):
        raise FormulaTranslationError(
            "Array and data-table formula objects are not supported"
        )
    leading_equals = value.startswith("=")
    try:
        tokenizer = _Tokenizer(value if leading_equals else "=" + value)
        if tokenizer.token_stack:
            raise TokenizerError("Unclosed formula expression")
    except (TokenizerError, IndexError) as exc:
        raise FormulaTranslationError("Cannot tokenize formula: " + str(exc)) from exc
    out = []
    for token in tokenizer.items:
        text = token.value
        if token.type == Token.FUNC and ":" in text:
            raise FormulaTranslationError(
                "Reference operators attached to function calls are not supported: "
                + text
            )
        if token.type == Token.OPERAND and token.subtype == Token.RANGE:
            implicit = text.startswith("@")
            spill = text.endswith("#")
            reference = text[int(implicit) : len(text) - int(spill)]
            parsed = parse_reference(reference)
            if parsed is not None:
                reference = rewrite_reference(
                    reference, parsed, owner, ws, axis, edit, formula=True
                )
                text = ("@" if implicit else "") + reference
                if spill and not reference.endswith("#REF!"):
                    text += "#"
            elif ":" in reference and "[" not in reference:
                raise FormulaTranslationError(
                    "Range expressions with named endpoints are not supported: "
                    + reference
                )
            elif "\t" in reference or "\r" in reference:
                raise FormulaTranslationError(
                    "Tabs and carriage returns between formula operands are not supported"
                )
        out.append(text)
    return ("=" if leading_equals else "") + "".join(out)
