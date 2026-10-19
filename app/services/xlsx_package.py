"""Edit an existing .xlsx or .xlsm package in place.

openpyxl rebuilds the whole workbook on save and drops what it does not model,
such as shapes, form controls, slicers, and VBA. Here only the parts that change
are rewritten: the edited worksheets, workbook.xml, its relationships, and
[Content_Types].xml. Every other part is copied byte for byte.

Strings are written as inline strings, so text that starts with = is never a
formula. Writing over a formula drops calcChain.xml, which Excel rebuilds, and
any write asks Excel to recalculate on open so cached formula results are fresh.
"""

from __future__ import annotations

import io
import posixpath
import re
import zipfile

from lxml import etree
from openpyxl.formula.translate import Translator
from openpyxl.utils.cell import column_index_from_string, get_column_letter, range_boundaries

from app.core.exceptions import OfficeDocumentError

_PARSER = etree.XMLParser(resolve_entities=False, no_network=True)
_CONTENT_TYPES = "[Content_Types].xml"
_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
_DEFAULT_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_WORKSHEET_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
_MACRO_CT = "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
_CELL = re.compile(r"^([A-Z]{1,3})([0-9]+)$")
# Children of <workbook> that may come before <calcPr>, in schema order.
_BEFORE_CALC = (
    "fileVersion",
    "fileSharing",
    "workbookPr",
    "workbookProtection",
    "bookViews",
    "sheets",
    "functionGroups",
    "externalReferences",
    "definedNames",
)


def _unreadable() -> OfficeDocumentError:
    return OfficeDocumentError("INVALID_OFFICE_FILE", "Could not read Excel workbook")


def _parse(data: bytes):
    try:
        return etree.fromstring(data, _PARSER)
    except etree.XMLSyntaxError as exc:
        raise _unreadable() from exc


def _serialize(root) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _rels_path(part: str) -> str:
    folder, name = posixpath.split(part)
    return posixpath.join(folder, "_rels", f"{name}.rels")


def _target(source: str, target: str) -> str:
    if target.startswith("/"):
        return target[1:]
    return posixpath.normpath(posixpath.join(posixpath.dirname(source), target))


def _local(element) -> str:
    return etree.QName(element).localname


def _column(cell) -> int:
    return column_index_from_string(_CELL.match(cell.get("r")).group(1))


# Table column names store control characters as _xHHHH_ escapes.
_NAME_ESCAPES = (("\r", "_x000d_"), ("\n", "_x000a_"), ("\t", "_x0009_"))


def _escape_name(text: str) -> str:
    for char, code in _NAME_ESCAPES:
        text = text.replace(char, code)
    return text


def _unescape_name(text: str) -> str:
    for char, code in _NAME_ESCAPES:
        text = text.replace(code, char)
    return text


def _bracket_end(text: str, start: int) -> int | None:
    """Index of the ] that closes the [ at start. An apostrophe escapes the next character."""
    depth = 0
    index = start
    while index < len(text):
        char = text[index]
        if char == "'":
            index += 2
            continue
        if char == "[":
            depth += 1
        elif char == "]":
            depth -= 1
            if depth == 0:
                return index
        index += 1
    return None


def _has_bracket(text: str) -> bool:
    index = 0
    while index < len(text):
        if text[index] == "'":
            index += 2
            continue
        if text[index] == "[":
            return True
        index += 1
    return False


def _unescape_ref(name: str) -> str:
    chars = []
    index = 0
    while index < len(name):
        if name[index] == "'" and index + 1 < len(name):
            index += 1
        chars.append(name[index])
        index += 1
    return "".join(chars)


def _escape_ref(name: str) -> str:
    return "".join(f"'{char}" if char in "'[]#" else char for char in name)


def _rename_spec(inner: str, old: str, new: str) -> str:
    """Rewrite the part of Table1[...] between the outer brackets."""

    def matches(name: str) -> bool:
        return not name.startswith("#") and _unescape_ref(name).casefold() == old.casefold()

    if not _has_bracket(inner):
        # Table1[[name]] is valid for any column name; Table1[name] is not.
        at = inner.startswith("@")
        name = inner[1:] if at else inner
        return f"{'@' if at else ''}[{_escape_ref(new)}]" if matches(name) else inner
    pieces = []
    index = 0
    while index < len(inner):
        char = inner[index]
        if char == "'":
            pieces.append(inner[index : index + 2])
            index += 2
            continue
        if char == "[":
            end = _bracket_end(inner, index)
            if end is None:
                pieces.append(inner[index:])
                break
            content = inner[index + 1 : end]
            pieces.append(f"[{_escape_ref(new)}]" if matches(content) else inner[index : end + 1])
            index = end + 1
            continue
        pieces.append(char)
        index += 1
    return "".join(pieces)


def rename_column_refs(formula: str, table: str, old: str, new: str) -> str:
    """Point structured references to table column old at new instead.

    Text in double quotes and references to another workbook (after !) are left alone.
    """
    key = table.casefold()
    size = len(table)
    pieces = []
    index = 0
    while index < len(formula):
        char = formula[index]
        if char == '"':
            end = index + 1
            while end < len(formula):
                if formula[end] == '"':
                    if formula[end + 1 : end + 2] == '"':
                        end += 2
                        continue
                    break
                end += 1
            pieces.append(formula[index : end + 1])
            index = end + 1
            continue
        before = formula[index - 1] if index else ""
        if (
            formula[index : index + size].casefold() == key
            and formula[index + size : index + size + 1] == "["
            and not (before and (before.isalnum() or before in "_.!\\"))
        ):
            end = _bracket_end(formula, index + size)
            if end is not None:
                inner = formula[index + size + 1 : end]
                pieces.append(f"{formula[index : index + size]}[{_rename_spec(inner, old, new)}]")
                index = end + 1
                continue
        pieces.append(char)
        index += 1
    return "".join(pieces)


def _header_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, float):
        return repr(value)
    return str(value)


class _Table:
    """A table (ListObject) on a worksheet.

    Excel repairs a workbook whose header cells differ from the table column
    names, so a header write renames the column. Like Excel, an empty header
    becomes ColumnN and a repeated name gets a number, and formulas that name
    the column are rewritten on save. A totals row write replaces that
    column's total function with the written label.
    """

    def __init__(self, root):
        self.root = root
        self.dirty = False
        self.name = root.get("displayName") or root.get("name") or ""
        self.renames: list[tuple[str, str]] = []
        self._ns = etree.QName(root).namespace
        self._min_col, self._min_row, _max_col, self._max_row = range_boundaries(root.get("ref") or "")
        self._header = int(root.get("headerRowCount") or "1") > 0
        self._totals = int(root.get("totalsRowCount") or "0") > 0
        columns = root.find(f"{{{self._ns}}}tableColumns")
        self._columns = [] if columns is None else list(columns.iterchildren(f"{{{self._ns}}}tableColumn"))

    def adjust(self, row: int, column: int, value):
        """The value to store at row and column, after updating the table to match it."""
        index = column - self._min_col
        if not 0 <= index < len(self._columns) or not self._min_row <= row <= self._max_row:
            return value
        target = self._columns[index]
        if self._header and row == self._min_row:
            text = _header_text(value) or f"Column{index + 1}"
            taken = {
                _unescape_name(other.get("name") or "").casefold() for other in self._columns if other is not target
            }
            name, number = text, 1
            while name.casefold() in taken:
                number += 1
                name = f"{text}{number}"
            old = _unescape_name(target.get("name") or "")
            if old != name:
                self.renames.append((old, name))
            target.set("name", _escape_name(name))
            self.dirty = True
            return name
        if self._totals and row == self._max_row:
            target.attrib.pop("totalsRowFunction", None)
            for formula in target.findall(f"{{{self._ns}}}totalsRowFormula"):
                target.remove(formula)
            if isinstance(value, str) and value:
                target.set("totalsRowLabel", value)
            else:
                target.attrib.pop("totalsRowLabel", None)
            self.dirty = True
        return value


class _Sheet:
    def __init__(self, root, tables: list[_Table] | None = None):
        self.root = root
        self.tables = tables or []
        self.ns = etree.QName(root).namespace
        self._c, self._f, self._v, self._is = (self._q(name) for name in ("c", "f", "v", "is"))
        self.data = root.find(self._q("sheetData"))
        if self.data is None:
            raise _unreadable()
        self.formula_removed = False
        self.written = False
        self._number_rows()
        self._rows = {int(row.get("r")): row for row in self.data.iterchildren(self._q("row"))}
        self._last_row = max(self._rows, default=0)
        self._arrays = []
        for formula in self.data.iter(self._q("f")):
            if formula.get("t") in ("array", "dataTable") and formula.get("ref"):
                self._arrays.append(range_boundaries(formula.get("ref")))

    def _q(self, name: str) -> str:
        return f"{{{self.ns}}}{name}"

    def _number_rows(self) -> None:
        # r is optional on rows and cells; a missing one follows the previous.
        row_number = 0
        for row in self.data.iterchildren(self._q("row")):
            row_number = int(row.get("r") or row_number + 1)
            row.set("r", str(row_number))
            column = 0
            for cell in row.iterchildren(self._q("c")):
                ref = cell.get("r")
                if ref:
                    match = _CELL.match(ref)
                    if not match:
                        raise _unreadable()
                    column = column_index_from_string(match.group(1))
                else:
                    column += 1
                    cell.set("r", f"{get_column_letter(column)}{row_number}")

    def next_row(self) -> int:
        last = 0
        for row in self.data.iterchildren(self._q("row")):
            if row.find(self._q("c")) is not None:
                last = max(last, int(row.get("r")))
        return last + 1

    def _row(self, number: int):
        row = self._rows.get(number)
        if row is not None:
            return row
        row = etree.Element(self._q("row"), r=str(number))
        if number > self._last_row:
            self.data.append(row)
            self._last_row = number
        else:
            self._rows[min(n for n in self._rows if n > number)].addprevious(row)
        self._rows[number] = row
        return row

    def _cell(self, row, ref: str, column: int):
        """The cell at ref, and whether it was just created."""
        # Appends write left to right, so a cell past the last one is appended directly.
        last = row[-1] if len(row) else None
        if last is not None and not (last.tag == self._c and _column(last) < column):
            before = None
            for cell in row.iterchildren(self._c):
                if cell.get("r") == ref:
                    return cell, False
                if _column(cell) > column:
                    before = cell
                    break
            if before is not None:
                last = before
            elif last.tag != self._c:
                last = row.find(self._q("extLst"))
            else:
                last = None
        else:
            last = None
        cell = etree.Element(self._c, r=ref)
        if last is None:
            row.append(cell)
        else:
            last.addprevious(cell)
        # Row spans are an optional hint that may no longer cover the new cell.
        row.attrib.pop("spans", None)
        return cell, True

    def write(self, row_number: int, column: int, value) -> None:
        ref = f"{get_column_letter(column)}{row_number}"
        for min_col, min_row, max_col, max_row in self._arrays:
            if min_col <= column <= max_col and min_row <= row_number <= max_row:
                raise OfficeDocumentError("VALIDATION_ERROR", f"cell {ref} is part of an array formula")
        for table in self.tables:
            value = table.adjust(row_number, column, value)
        cell, created = self._cell(self._row(row_number), ref, column)
        if not created:
            formula = cell.find(self._f)
            if formula is not None:
                if formula.get("t") == "shared" and formula.get("ref"):
                    self._unshare(formula, ref)
                self.formula_removed = True
            for child in [child for child in cell if child.tag in (self._f, self._v, self._is)]:
                cell.remove(child)
            cell.attrib.pop("t", None)
        if value is not None:
            self._set_value(cell, value)
        self.written = True

    def _set_value(self, cell, value) -> None:
        if isinstance(value, bool):
            cell.set("t", "b")
            child = etree.Element(self._v)
            child.text = "1" if value else "0"
        elif isinstance(value, (int, float)):
            child = etree.Element(self._v)
            child.text = str(value) if isinstance(value, int) else repr(value)
        else:
            cell.set("t", "inlineStr")
            child = etree.Element(self._is)
            text = etree.SubElement(child, self._q("t"))
            text.set(_XML_SPACE, "preserve")
            try:
                text.text = value
            except ValueError as exc:
                raise OfficeDocumentError("VALIDATION_ERROR", "cell text has characters Excel does not allow") from exc
        # Cell children are f, v, is, extLst in that order; f, v, and is were removed.
        cell.insert(0, child)

    def _unshare(self, master, master_ref: str) -> None:
        """Give each other cell of a shared formula its own copy before the first is overwritten."""
        group = master.get("si")
        source = f"={master.text or ''}"
        for formula in self.data.iter(self._q("f")):
            if formula is master or formula.get("t") != "shared" or formula.get("si") != group:
                continue
            ref = formula.getparent().get("r")
            try:
                translated = Translator(source, origin=master_ref).translate_formula(ref)
            except Exception as exc:
                raise OfficeDocumentError(
                    "VALIDATION_ERROR", f"cell {master_ref} starts a shared formula that cannot be split"
                ) from exc
            formula.text = translated[1:]
            for name in ("t", "si", "ref"):
                formula.attrib.pop(name, None)

    def finish(self) -> None:
        dimension = self.root.find(self._q("dimension"))
        if dimension is None:
            return
        rows, columns = [], []
        for number, row in self._rows.items():
            cells = row.findall(self._q("c"))
            if cells:
                rows.append(number)
                columns.extend((_column(cells[0]), _column(cells[-1])))
        if not rows:
            dimension.set("ref", "A1")
            return
        start = f"{get_column_letter(min(columns))}{min(rows)}"
        end = f"{get_column_letter(max(columns))}{max(rows)}"
        dimension.set("ref", start if start == end else f"{start}:{end}")


class XlsxPackage:
    def __init__(self, raw: bytes):
        try:
            with zipfile.ZipFile(io.BytesIO(raw)) as package:
                self._infos = package.infolist()
                self._data = {info.filename: package.read(info.filename) for info in self._infos}
        except zipfile.BadZipFile as exc:
            raise OfficeDocumentError("INVALID_OFFICE_FILE", "Corrupt Office package") from exc
        self._roots: dict[str, object] = {}
        self._dirty: set[str] = set()
        self._removed: set[str] = set()
        self._added: list[str] = []
        self._sheets: dict[str, _Sheet] = {}
        self._tables: dict[str, _Table] = {}
        self._workbook_path = self._office_document()
        self._workbook_rels_path = _rels_path(self._workbook_path)
        self._workbook = self._root(self._workbook_path)
        self._workbook_rels = self._root(self._workbook_rels_path)
        self._types = self._root(_CONTENT_TYPES)
        self._ns = etree.QName(self._workbook).namespace
        sheets = self._workbook.find(f"{{{self._ns}}}sheets")
        if sheets is None:
            raise _unreadable()
        self._sheet_list = sheets
        self.macro_enabled = self._content_type(self._workbook_path) == _MACRO_CT

    def _root(self, path: str):
        if path not in self._roots:
            if path not in self._data:
                raise _unreadable()
            self._roots[path] = _parse(self._data[path])
        return self._roots[path]

    def _office_document(self) -> str:
        for rel in self._root("_rels/.rels"):
            if (rel.get("Type") or "").endswith("/officeDocument"):
                return _target("", rel.get("Target") or "")
        raise _unreadable()

    def _content_type(self, path: str) -> str | None:
        for item in self._types:
            if _local(item) == "Override" and item.get("PartName") == f"/{path}":
                return item.get("ContentType")
        extension = posixpath.splitext(path)[1][1:].lower()
        for item in self._types:
            if _local(item) == "Default" and (item.get("Extension") or "").lower() == extension:
                return item.get("ContentType")
        return None

    def _relationships(self):
        return [rel for rel in self._workbook_rels if _local(rel) == "Relationship"]

    def _sheet_rel_id(self, sheet) -> str | None:
        for name, value in sheet.attrib.items():
            if etree.QName(name).localname == "id":
                return value
        return None

    def _sheet_part(self, sheet) -> tuple[str, str] | None:
        rel_id = self._sheet_rel_id(sheet)
        for rel in self._relationships():
            if rel.get("Id") == rel_id:
                return _target(self._workbook_path, rel.get("Target") or ""), rel.get("Type") or ""
        return None

    def sheet_count(self) -> int:
        return len(self._sheet_list)

    def find(self, name: str):
        target = name.casefold()
        for sheet in self._sheet_list:
            if (sheet.get("name") or "").casefold() == target:
                return sheet
        return None

    def worksheet(self, name: str) -> _Sheet:
        sheet = self.find(name)
        if sheet is None:
            raise OfficeDocumentError("VALIDATION_ERROR", f"sheet not found: {name}")
        located = self._sheet_part(sheet)
        if located is None or not located[1].endswith("/worksheet"):
            raise OfficeDocumentError("VALIDATION_ERROR", f"sheet is not a worksheet: {name}")
        path = located[0]
        if path not in self._sheets:
            self._sheets[path] = _Sheet(self._root(path), self._sheet_tables(path))
        return self._sheets[path]

    def _sheet_tables(self, path: str) -> list[_Table]:
        rels = _rels_path(path)
        if rels not in self._data:
            return []
        tables = []
        for rel in self._root(rels):
            if _local(rel) != "Relationship" or not (rel.get("Type") or "").endswith("/table"):
                continue
            target = _target(path, rel.get("Target") or "")
            if target not in self._tables:
                self._tables[target] = _Table(self._root(target))
            tables.append(self._tables[target])
        return tables

    def add_worksheet(self, name: str) -> _Sheet:
        folder = posixpath.join(posixpath.dirname(self._workbook_path), "worksheets")
        number = 1
        while posixpath.join(folder, f"sheet{number}.xml") in self._data:
            number += 1
        path = posixpath.join(folder, f"sheet{number}.xml")
        rel_ns, rel_type = _DEFAULT_REL_NS, f"{_DEFAULT_REL_NS}/worksheet"
        for sheet in self._sheet_list:
            for attribute in sheet.attrib:
                qname = etree.QName(attribute)
                if qname.localname == "id" and qname.namespace:
                    rel_ns = qname.namespace
            located = self._sheet_part(sheet)
            if located and located[1].endswith("/worksheet"):
                rel_type = located[1]
        ids = {rel.get("Id") for rel in self._relationships()}
        number_id = 1
        while f"rId{number_id}" in ids:
            number_id += 1
        rel_id = f"rId{number_id}"
        etree.SubElement(
            self._workbook_rels,
            f"{{{_PKG_REL_NS}}}Relationship",
            Id=rel_id,
            Type=rel_type,
            Target=posixpath.relpath(path, posixpath.dirname(self._workbook_path)),
        )
        sheet_ids = [int(sheet.get("sheetId") or 0) for sheet in self._sheet_list]
        entry = etree.SubElement(self._sheet_list, f"{{{self._ns}}}sheet")
        entry.set("name", name)
        entry.set("sheetId", str(max(sheet_ids, default=0) + 1))
        entry.set(f"{{{rel_ns}}}id", rel_id)
        etree.SubElement(self._types, f"{{{_CT_NS}}}Override", PartName=f"/{path}", ContentType=_WORKSHEET_CT)
        root = etree.Element(f"{{{self._ns}}}worksheet", nsmap={None: self._ns, "r": rel_ns})
        etree.SubElement(root, f"{{{self._ns}}}sheetData")
        self._roots[path] = root
        self._data[path] = b""
        self._added.append(path)
        self._dirty.update({self._workbook_path, self._workbook_rels_path, _CONTENT_TYPES})
        sheet = _Sheet(root)
        self._sheets[path] = sheet
        return sheet

    def _drop_calc_chain(self) -> None:
        for rel in self._relationships():
            if (rel.get("Type") or "").endswith("/calcChain"):
                path = _target(self._workbook_path, rel.get("Target") or "")
                self._workbook_rels.remove(rel)
                self._removed.add(path)
                for item in list(self._types):
                    if _local(item) == "Override" and item.get("PartName") == f"/{path}":
                        self._types.remove(item)
                self._dirty.update({self._workbook_rels_path, _CONTENT_TYPES})

    def _recalculate_on_open(self) -> None:
        calc = self._workbook.find(f"{{{self._ns}}}calcPr")
        if calc is None:
            calc = etree.Element(f"{{{self._ns}}}calcPr")
            anchor = None
            for child in self._workbook:
                if isinstance(child.tag, str) and _local(child) in _BEFORE_CALC:
                    anchor = child
            if anchor is None:
                self._workbook.insert(0, calc)
            else:
                anchor.addnext(calc)
        calc.set("fullCalcOnLoad", "1")
        self._dirty.add(self._workbook_path)

    def _rewrite_column_refs(self) -> None:
        """Point formulas at renamed table columns, as Excel does when a header changes."""
        renames = [(table.name, old, new) for table in list(self._tables.values()) for old, new in table.renames]
        if not renames:
            return

        def rewrite(elements) -> bool:
            changed = False
            for element in elements:
                text = element.text
                if not text or "[" not in text:
                    continue
                for table, old, new in renames:
                    text = rename_column_refs(text, table, old, new)
                if text != element.text:
                    element.text = text
                    changed = True
            return changed

        for sheet in self._sheet_list:
            located = self._sheet_part(sheet)
            if located is None or not located[1].endswith("/worksheet"):
                continue
            path = located[0]
            root = self._root(path)
            ns = etree.QName(root).namespace
            tags = [f"{{{ns}}}{name}" for name in ("f", "formula", "formula1", "formula2")]
            if rewrite(element for element in root.iter(*tags)):
                self._dirty.add(path)
            self._sheet_tables(path)
        for path, table in self._tables.items():
            formulas = [f"{{{table._ns}}}{name}" for name in ("calculatedColumnFormula", "totalsRowFormula")]
            if rewrite(table.root.iter(*formulas)):
                table.dirty = True
        if rewrite(self._workbook.iter(f"{{{self._ns}}}definedName")):
            self._dirty.add(self._workbook_path)

    def save(self) -> bytes:
        for path, sheet in self._sheets.items():
            if sheet.written or path in self._added:
                sheet.finish()
                self._dirty.add(path)
        self._rewrite_column_refs()
        self._dirty.update(path for path, table in self._tables.items() if table.dirty)
        if any(sheet.formula_removed for sheet in self._sheets.values()):
            self._drop_calc_chain()
        if any(sheet.written for sheet in self._sheets.values()):
            self._recalculate_on_open()
        buffer = io.BytesIO()
        with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as package:
            for info in self._infos:
                if info.filename in self._removed:
                    continue
                if info.filename in self._dirty:
                    package.writestr(info, _serialize(self._roots[info.filename]), compress_type=zipfile.ZIP_DEFLATED)
                else:
                    package.writestr(info, self._data[info.filename])
            for path in self._added:
                package.writestr(path, _serialize(self._roots[path]))
        return buffer.getvalue()
