from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

"""Write worksheets to xml representations."""

# Python stdlib imports
from io import BytesIO

# compatibility imports

from openpyxl.compat import safe_string, itervalues
from openpyxl.compat.itertools import iteritems, iterkeys

from openpyxl import LXML

# package imports
from openpyxl.utils import COORD_RE
from openpyxl.cell import (
    coordinate_from_string,
    column_index_from_string,
)
from openpyxl.xml.functions import (
    Element,
    SubElement,
    xmlfile,
)
from openpyxl.xml.constants import (
    SHEET_MAIN_NS,
    REL_NS,
)
from openpyxl.formatting import ConditionalFormatting
from openpyxl.worksheet.datavalidation import writer
from openpyxl.worksheet.properties import WorksheetProperties, write_sheetPr


def row_sort(cell):
    """Translate column names for sorting."""
    return column_index_from_string(cell.column)


def write_properties(worksheet):
    wsp = worksheet.sheet_properties
    pr = write_sheetPr(wsp)
    return pr


def write_sheetviews(worksheet):
    views = Element('sheetViews')
    sheetviewAttrs = {'workbookViewId': '0'}
    if not worksheet.show_gridlines:
        sheetviewAttrs['showGridLines'] = '0'
    view = SubElement(views, 'sheetView', sheetviewAttrs)
    selectionAttrs = {
        'activeCell': worksheet.active_cell,
        'sqref': worksheet.selected_cell
    }
    topLeftCell = worksheet.freeze_panes
    if topLeftCell:
        colName, row = coordinate_from_string(topLeftCell)
        column = column_index_from_string(colName)
        pane = 'topRight'
        paneAttrs = {'topLeftCell':topLeftCell, 'state':'frozen'}
        if column > 1:
            paneAttrs['xSplit'] = str(column - 1)
        if row > 1:
            paneAttrs['ySplit'] = str(row - 1)
            pane = 'bottomLeft'
            if column > 1:
                pane = 'bottomRight'
        selectionAttrs['pane'] = pane
        paneAttrs['activePane'] = pane
        view.append(Element('pane', paneAttrs))
        if row > 1 and column > 1:
            SubElement(view, 'selection', pane='topRight')
            SubElement(view, 'selection', pane='bottomLeft')

    SubElement(view, 'selection', selectionAttrs)
    return views


def write_format(worksheet):
    attrs = {'defaultRowHeight': '15', 'baseColWidth': '10'}
    dimensions_outline = [dim.outline_level
                          for dim in itervalues(worksheet.column_dimensions)]
    if dimensions_outline:
        outline_level = max(dimensions_outline)
        if outline_level:
            attrs['outlineLevelCol'] = str(outline_level)
    return Element('sheetFormatPr', attrs)


def write_cols(worksheet):
    """Write worksheet columns to xml.

    <cols> may never be empty -
    spec says must contain at least one child
    """
    cols = []
    for label, dimension in iteritems(worksheet.column_dimensions):
        dimension.style = worksheet._styles.get(label)
        col_def = dict(dimension)
        if col_def == {}:
            continue
        idx = column_index_from_string(label)
        cols.append((idx, col_def))

    if not cols:
        return

    el = Element('cols')

    for idx, col_def in sorted(cols):
        v = "%d" % idx
        cmin = col_def.get('min') or v
        cmax = col_def.get('max') or v
        col_def.update({'min': cmin, 'max': cmax})
        el.append(Element('col', col_def))
    return el


def write_autofilter(worksheet):
    auto_filter = worksheet.auto_filter
    if auto_filter.ref is None:
        return

    el = Element('autoFilter', ref=auto_filter.ref)
    if (auto_filter.filter_columns
        or auto_filter.sort_conditions):
        for col_id, filter_column in sorted(auto_filter.filter_columns.items()):
            fc = SubElement(el, 'filterColumn', colId=str(col_id))
            attrs = {}
            if filter_column.blank:
                attrs = {'blank': '1'}
            flt = SubElement(fc, 'filters', attrs)
            for val in filter_column.vals:
                flt.append(Element('filter', val=val))
        if auto_filter.sort_conditions:
            srt = SubElement(el, 'sortState', ref=auto_filter.ref)
            for sort_condition in auto_filter.sort_conditions:
                sort_attr = {'ref': sort_condition.ref}
                if sort_condition.descending:
                    sort_attr['descending'] = '1'
                srt.append(Element('sortCondtion', sort_attr))
    return el


def write_mergecells(worksheet):
    """Write merged cells to xml."""
    cells = worksheet._merged_cells
    if not cells:
        return

    merge = Element('mergeCells', count='%d' % len(cells))
    for range_string in cells:
        merge.append(Element('mergeCell', ref=range_string))
    return merge


def write_conditional_formatting(worksheet):
    """Write conditional formatting to xml."""
    for range_string, rules in iteritems(worksheet.conditional_formatting.cf_rules):
        if not len(rules):
            # Skip if there are no rules.  This is possible if a dataBar rule was read in and ignored.
            continue
        cf = Element('conditionalFormatting', {'sqref': range_string})
        for rule in rules:
            if rule['type'] == 'dataBar':
                # Ignore - uses extLst tag which is currently unsupported.
                continue
            attr = {'type': rule['type']}
            for rule_attr in ConditionalFormatting.rule_attributes:
                if rule_attr in rule:
                    attr[rule_attr] = str(rule[rule_attr])
            cfr = SubElement(cf, 'cfRule', attr)
            if 'formula' in rule:
                for f in rule['formula']:
                    SubElement(cfr, 'formula').text = f
            if 'colorScale' in rule:
                cs = SubElement(cfr, 'colorScale')
                for cfvo in rule['colorScale']['cfvo']:
                    SubElement(cs, 'cfvo', cfvo)
                for color in rule['colorScale']['color']:
                    SubElement(cs, 'color', dict(color))
            if 'iconSet' in rule:
                iconAttr = {}
                for icon_attr in ConditionalFormatting.icon_attributes:
                    if icon_attr in rule['iconSet']:
                        iconAttr[icon_attr] = rule['iconSet'][icon_attr]
                iconSet = SubElement(cfr, 'iconSet', iconAttr)
                for cfvo in rule['iconSet']['cfvo']:
                    SubElement(iconSet, 'cfvo', cfvo)
        yield cf


def write_datavalidation(worksheet):
    """ Write data validation(s) to xml."""
    # Filter out "empty" data-validation objects (i.e. with 0 cells)
    required_dvs = [x for x in worksheet._data_validations
                    if len(x.cells) or len(x.ranges)]
    if not required_dvs:
        return

    dvs = Element("{%s}dataValidations" % SHEET_MAIN_NS,
                  count=str(len(required_dvs)))
    for dv in required_dvs:
        dvs.append(writer(dv))

    return dvs


def write_header_footer(worksheet):
    header = worksheet.header_footer.getHeader()
    footer = worksheet.header_footer.getFooter()
    if header or footer:
        tag = Element('headerFooter')
        if header:
            SubElement(tag, 'oddHeader').text = header
        if footer:
            SubElement(tag, 'oddFooter').text = footer
        return tag


def write_hyperlinks(worksheet):
    """Write worksheet hyperlinks to xml."""
    tag = Element('hyperlinks')
    for cell in worksheet.get_cell_collection():
        if cell.hyperlink_rel_id is not None:
            attrs = {'display': cell.hyperlink,
                     'ref': cell.coordinate,
                     '{%s}id' % REL_NS: cell.hyperlink_rel_id}
            tag.append(Element('hyperlink', attrs))
    if tag.getchildren():
        return tag


def write_pagebreaks(worksheet):
    breaks = worksheet.page_breaks
    if breaks:
        tag = Element('rowBreaks', {'count': str(len(breaks)),
                                     'manualBreakCount': str(len(breaks))})
        for b in breaks:
            tag.append(Element('brk', id=str(b), man=true, max='16383',
                               min='0'))
        return tag


def write_worksheet(worksheet, shared_strings):
    """Write a worksheet to an xml file."""
    if LXML is True:
        from . lxml_worksheet import write_cell, write_rows

    out = BytesIO()
    NSMAP = {None : SHEET_MAIN_NS}

    with xmlfile(out) as xf:
        with xf.element('worksheet', nsmap=NSMAP):

            props = write_properties(worksheet)
            xf.write(props)

            dim = Element('dimension', {'ref': '%s' % worksheet.calculate_dimension()})
            xf.write(dim)

            xf.write(write_sheetviews(worksheet))
            xf.write(write_format(worksheet))
            cols = write_cols(worksheet)
            if cols is not None:
                xf.write(cols)
            write_rows(xf, worksheet)

            if worksheet.protection.sheet:
                prot = Element('sheetProtection', dict(worksheet.protection))
                xf.write(prot)

            af = write_autofilter(worksheet)
            if af is not None:
                xf.write(af)

            merge = write_mergecells(worksheet)
            if merge is not None:
                xf.write(merge)

            cfs = write_conditional_formatting(worksheet)
            for cf in cfs:
                xf.write(cf)

            dv = write_datavalidation(worksheet)
            if dv is not None:
                xf.write(dv)

            hyper = write_hyperlinks(worksheet)
            if hyper is not None:
                xf.write(hyper)


            options = worksheet.print_options
            if len(dict(options)) > 0:
                new_element = options.write_xml_element()
                xf.write(new_element)
                del new_element

            margins = Element('pageMargins', dict(worksheet.page_margins))
            xf.write(margins)
            del margins

            setup = worksheet.page_setup
            if len(dict(setup)) > 0:
                new_element = setup.write_xml_element()
                xf.write(new_element)
                del new_element

            hf = write_header_footer(worksheet)
            if hf is not None:
                xf.write(hf)

            if worksheet._charts or worksheet._images:
                drawing = Element('drawing', {'{%s}id' % REL_NS: 'rId1'})
                xf.write(drawing)
                del drawing

            # If vba is being preserved then add a legacyDrawing element so
            # that any controls can be drawn.
            if worksheet.vba_controls is not None:
                xml = Element("{%s}legacyDrawing" % SHEET_MAIN_NS,
                              {"{%s}id" % REL_NS : worksheet.vba_controls})
                xf.write(xml)

            pb = write_pagebreaks(worksheet)
            if pb is not None:
                xf.write(pb)

            # add a legacyDrawing so that excel can draw comments
            if worksheet._comment_count > 0:
                comments = Element('legacyDrawing', {'{%s}id' % REL_NS: 'commentsvml'})
                xf.write(comments)

    xml = out.getvalue()
    out.close()
    return xml


def get_rows_to_write(worksheet):
    """Return all rows, and any cells that they contain"""
    # Ensure a blank cell exists if it has a style
    for styleCoord in iterkeys(worksheet._styles):
        if isinstance(styleCoord, str) and COORD_RE.search(styleCoord):
            worksheet.cell(styleCoord)

    # create rows of cells
    cells_by_row = {}
    for cell in itervalues(worksheet._cells):
        cells_by_row.setdefault(cell.row, []).append(cell)

    # make sure rows that only have a height set are returned
    for row_idx in worksheet.row_dimensions:
        if row_idx not in cells_by_row:
            cells_by_row[row_idx] = []

    return cells_by_row

### ElementTree

def write_rows(xf, worksheet):
    """Write worksheet data to xml."""

    cells_by_row = get_rows_to_write(worksheet)

    with xf.element("sheetData"):
        for row_idx in sorted(cells_by_row):
            # row meta data
            row_dimension = worksheet.row_dimensions[row_idx]
            row_dimension.style = worksheet._styles.get(row_idx)
            attrs = {'r': '%d' % row_idx,
                     'spans': '1:%d' % worksheet.max_column}
            attrs.update(dict(row_dimension))

            with xf.element("row", attrs):

                row_cells = cells_by_row[row_idx]
                for cell in sorted(row_cells, key=row_sort):
                    el = write_cell(worksheet, cell)
                    xf.write(el)


def write_cell(worksheet, cell):
    string_table = worksheet.parent.shared_strings
    coordinate = cell.coordinate
    attributes = {'r': coordinate}
    if cell.has_style:
        attributes['s'] = '%d' % cell.style_id

    if cell.data_type != 'f':
        attributes['t'] = cell.data_type

    value = cell.internal_value

    el = Element("c", attributes)
    if value in ('', None):
        return el

    if cell.data_type == 'f':
        shared_formula = worksheet.formula_attributes.get(coordinate, {})
        if shared_formula is not None:
            if (shared_formula.get('t') == 'shared'
                and 'ref' not in shared_formula):
                value = None
        formula = SubElement(el, 'f', shared_formula)
        if value is not None:
            formula.text = value[1:]
            value = None

    if cell.data_type == 's':
        value = string_table.add(value)
    cell_content = SubElement(el, 'v')
    if value is not None:
        cell_content.text = safe_string(value)
    return el
