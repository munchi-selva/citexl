#!/usr/bin/env python3
# -*- coding: utf-8 -*-

################################################################################
# Library of functions for manipulating a book citations spreadsheet.
################################################################################

################################################################################
# Imports
################################################################################
from __future__ import unicode_literals
# Ensures Unicode string compatibility # in Python 2/3
import argparse
import itertools
import re
import json
import sys
from io import BytesIO
# from openpyxl.worksheet.worksheet import Worksheet
from openpyxl import load_workbook
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.workbook.defined_name import DefinedName
from os import path
from enum import IntEnum
#from typing import Dict, List
#import future

import sys
if sys.version_info.major == 2:
    from cjklib import characterlookup

#
# Temporary hack, a copy of .pythonrc config to enable autocomplete when
# running interactively
#
try:
    import readline
except ImportError:
    print("Module readline not available.")
else:
    import rlcompleter
    readline.parse_and_bind("tab: complete")

###############################################################################

###############################################################################
# Constants
###############################################################################

#######################################################
# Default source and destination spreadsheet file names
#######################################################
SOURCE_FILE = '/mnt/d/Books_and_Literature/Notes/Cha/Duke_of_Mount_Deer.xlsx'
DEST_FILE   = '/mnt/d/Books_and_Literature/Notes/Cha/Mount_Deer.xlsx'

#######################
# Worksheet style names
#######################
STYLE_GENERAL   = 'BookRef_General'
STYLE_LINK      = 'BookRef_Link'

########################################
# Column headers for citation worksheets
########################################
COL_HDR_PAGE       = '頁'
COL_HDR_LINE       = '直行'
COL_HDR_CITATION   = '引句'
COL_HDR_CATEGORY   = '範疇'
COL_HDR_PHRASE     = '字詞'
COL_HDR_JYUTPING   = '粵拼'
COL_HDR_DEFN       = '定義'

###############################################################################
# Cell types enum
###############################################################################
CellType = IntEnum('CellType',  'CT_NONE        \
                                 CT_ALL         \
                                 CT_DEFN        \
                                 CT_REFERRING   \
                                 CT_COUNT',

                                 start = -1)
###############################################################################

##############################
# Names of citation worksheets
##############################
CITATION_SHEETS = [
    '一', '二', '三', '四', '五', '六', '七', '八', '九', '十',
    '十一', '十二', '十三', '十四', '十五', '十六', '十七', '十八', '十九', '二十',
    '二十一', '二十二', '二十三', '二十四', '二十五', '二十六', '二十七', '二十八', '二十九', '三十',
    '三一', '三十二', '三十三', '三十四', '三十五', '三十六', '三十七', '三十八', '三十九', '四十',
    '四十一', '四十二', '四十三', '四十四', '四十五', '四十六', '四十七', '四十八', '四十九', '五十'
]

########################################################################
# Component separators for defined name identifiers and reference labels
########################################################################
DEF_NAME_ID_SEP = '_'
REF_LABEL_SEP   = ';'
###############################################################################

#################################################
# Chinese character shape decomposition constants
#################################################
CJK_SHAPE_LTR   = '\u2ff0'         # ⿰    Left to right
CJK_SHAPE_ATB   = u'\u2ff1'         # ⿱    Above to below
CJK_SHAPE_LMR   = u'\u2ff2'         # ⿲    Left to middle and right
CJK_SHAPE_AMB   = u'\u2ff3'         # ⿳    Above to middle and below
CJK_SHAPE_SURR  = u'\u2ff4'         # ⿴    Full surrond
CJK_SHAPE_SA    = u'\u2ff5'         # ⿵    Surround from above
CJK_SHAPE_SB    = u'\u2ff6'         # ⿶    Surround from below
CJK_SHAPE_SL    = u'\u2ff7'         # ⿷    Surround from left
CJK_SHAPE_SUL   = u'\u2ff8'         # ⿸    Surround from upper left
CJK_SHAPE_SUR   = u'\u2ff9'         # ⿹    Surround from upper right
CJK_SHAPE_SLL   = u'\u2ffa'         # ⿺    Surround from lower left
CJK_SHAPE_OL    = u'\u2ffb'         # ⿻    Overlaid

###############################################################################
def header_row(ws):
    """
    Returns the header row of a citation worksheet, which contains names of the
    columns.

    :param ws:  The worksheet
    :returns:   The header row.
    """
    return ws[1]
###############################################################################


###############################################################################
def column_mappings(ws):
    # type (Worksheet) -> dict
    """
    Returns the mappings between a worksheet's column names and letters.
    The mappings are cached for reuse the first time the function is called.

    :param ws:  The worksheet
    :returns:   A dictionary mapping column names to column letters
    """
    if not ws.title in column_mappings.col_dicts:
        col_dict        = dict()
        for header_cell in header_row(ws):
            col_name    = header_cell.value
            col_letter  = header_cell.column_letter
            col_dict[col_name] = col_letter
        column_mappings.col_dicts[ws.title] = col_dict
    return column_mappings.col_dicts[ws.title]

# Faking a static variable that stores the results for column_mappings() via
# the function's __dict__ dictionary attribute.
# See:
#   https://stackoverflow.com/questions/279561/what-is-the-python-equivalent-of-static-variables-inside-a-function
#   https://www.python.org/dev/peps/pep-0232/
column_mappings.col_dicts = dict()
###############################################################################


###############################################################################
def get_col_id(ws, col_name):
    #type (Worksheet, str) -> str
    """
    Returns the letter corresponding to the column with the given name in the
    specified worksheet.

    :param ws:          The worksheet
    :param col_name:i   The column name
    :returns:   The column's letter
    """
    return column_mappings(ws)[col_name]

###############################################################################


###############################################################################
def find_closest_value(ws,
                       col_letter,
                       row):
    # type: (Worksheet, str, int) -> (str, int)
    """
    Finds the first non-empty value that appears in the given column, at or
    above the specified row and TODO: describe the ordinal return value.

    :param  ws:         The worksheet
    :param  col_letter: Column letter
    :param  row:        Row number (1-based)
    :returns: The 
    """

    # Identify non-empty cells at or above the specified row
    non_empty_cells = [c for c in ws[col_letter][:row] if not c.value is None]

    if len(non_empty_cells) != 0:
        return non_empty_cells[-1].value, (row - non_empty_cells[-1].row + 1)

    return None, None

###############################################################################


###############################################################################
def get_def_name_id_and_label(ws, cell):
    # type: (Worksheet, Cell) -> (str, str)
    """
    Returns the defined name and a label for referring to a specified cell that
    is part of book citation.
    The defined name is built from the chapter name (i.e. worksheet title),
    page number, line number of the citation, plus its order amongst all
    citations for that chapter/page/line number.

    :param  ws:  The worksheet
    :param  cell: A cell in the worksheet
    :returns:
    """
    id = label = None

    page_col = get_col_id(ws, COL_HDR_PAGE)
    line_col = get_col_id(ws, COL_HDR_LINE)

    chap_name = ws.title
    page_number, _ = find_closest_value(ws, page_col, cell.row)
    line_number, ref_number = find_closest_value(ws, line_col, cell.row)

    if not page_number is None and not line_number is None:
        # All required reference components are defined... go!
        id = '{}{}{:02d}{}{:02d}{}{:02d}'.format(
                        chap_name, DEF_NAME_ID_SEP,
                        page_number, DEF_NAME_ID_SEP,
                        line_number, DEF_NAME_ID_SEP, ref_number)
        label =  '{}{}{}{}{}'.format(
                            chap_name, REF_LABEL_SEP,
                            page_number, REF_LABEL_SEP, line_number) 

    return id, label
###############################################################################


###############################################################################
def get_refs_for_ws_phrases(wb,
                            ws_name,
                            overwrite,
                            audit_only):
    # type: (Workbook, str, bool, bool) -> None
    """
    Fills in the references for the phrase column in a worksheet containing
    citations of a single chapter.
    This requires building a reference to the first occurrence of each phrase
    in a previous chapter worksheet.

    :param  wb:         The workbook
    :param  ws:         The chapter worksheet name
    :param  overwrite   If true, overwrite any existing content in referring
                        cells if these don't already refer to the referenced
                        cells
    :param  audit_only  If true, show/print the actions for building references
                        without modifying any data
    :returns:
    """
    if ws_name in CITATION_SHEETS and ws_name in wb.sheetnames:
        chapter_index = CITATION_SHEETS.index(ws_name)
        ws = wb.get_sheet_by_name(ws_name)
        #
        # Iterate through all non-empty phrase cells of the chapter worksheet
        #
        for phrase_cell in [c for c in ws[get_col_id(ws, COL_HDR_PHRASE)][1:] if c.value]:
            #
            # Find the first cell to define the phrase
            #
            referenced_cells = find_matches(wb, COL_HDR_PHRASE, [phrase_cell.value], False, CellType.CT_DEFN, 1)
            referenced_cell = referenced_cells[0] if len(referenced_cells) > 0 else None

            if referenced_cell and not referenced_cell == phrase_cell: 
                #
                # Ensure this cell isn't the one providing the definition!
                #
                referenced_ws = referenced_cell.parent
                referenced_chap = referenced_ws.title

                print("{}!{}{} ({}) --> {}!{}{}".format(
                      ws.title,
                      phrase_cell.column_letter, phrase_cell.row,
                      phrase_cell.value, 
                      referenced_chap, referenced_cell.column_letter, referenced_cell.row))

                referring_cell_loc = '{}{}'.format(get_col_id(ws, COL_HDR_DEFN), phrase_cell.row)
                referring_cell = ws[referring_cell_loc]
                build_reference(referenced_ws, referenced_cell,
                                ws, referring_cell, overwrite, audit_only)
###############################################################################


###############################################################################
def build_reference(referenced_ws,
                    referenced_cell,
                    referring_ws,
                    referring_cell,
                    overwrite = False,
                    audit_only = False):
    # type: (Worksheet, Cell, Worksheet, Cell, bool, bool) -> None
    """
    Builds a reference to one cell in another.
    The referencing is achieved by creating a defined name that includes
    the referenced cell, and linking to this name from the referring cell.

    :param  referenced_ws       The referenced worksheet
    :param  referenced_cell     The referenced cell
    :param  referring_ws        The referring worksheet
    :param  referring_cell      The referring cell
    :param  overwrite           If true, overwrite any existing content in the
                                referring cell (if this isn't already a label
                                for the referenced cell)
    :param  audit_only          If true, show/print the actions for building
                                the reference, but do not modify the data
    :returns: Nothing
    """
    #
    # Generate an identifier for the defined name, and the label for referring
    # to the referenced cell (i.e. displayed in the referring cell)
    #
    def_name_id, label = get_def_name_id_and_label(referenced_ws, referenced_cell)
    if not def_name_id is None and not label is None:
        workbook = referenced_ws.parent
        if not def_name_id in workbook.defined_names:
            #
            # Create the defined name
            #
            def_name_destination = '{}!${}${}'.format(referenced_ws.title,
                                                      referenced_cell.column_letter,
                                                      referenced_cell.row)
            def_name = DefinedName(name = def_name_id,
                                   attr_text = def_name_destination)
            print("\tCreate defined name: {}: {}".format(def_name_destination,
                                                       def_name_id))
            if not audit_only:
                workbook.defined_names.append(def_name)

    write_needed = referring_cell.value is None or (overwrite and referring_cell.value != label)
    referring_cell_loc = '{}{}'.format(referring_cell.column_letter,
                                       referring_cell.row)
    if write_needed:
        print('\t{}!{} current value = {}'.format(referring_ws.title,
                                                referring_cell_loc,
                                                referring_cell.value))
        if not audit_only:
            referring_cell.value = label
            referring_cell.hyperlink = Hyperlink(ref = referring_cell_loc,
                                                 location = def_name_id)
            jyutping_cell_loc = '{}{}'.format(get_col_id(referring_ws, COL_HDR_JYUTPING),
                                              referring_cell.row)
            jyutping_cell = referring_ws[jyutping_cell_loc]
            jyutping_cell.value = None
            assign_style(referring_cell)
            assign_style(jyutping_cell)

###############################################################################


###############################################################################
def assign_style(cell):
    # type: (Cell) -> None
    """
    Assigns the appropriate style to a citation worksheet cell
    :param  cell    The cell to be styled
    :returns: Nothing
    """
    cell.style = STYLE_GENERAL if cell.hyperlink is None else STYLE_LINK
###############################################################################


###############################################################################
def style_citation_sheet(ws):
    # type: (Cell) -> None
    """
    Assigns the appropriate style to all cells in a citation worksheet
    :param  ws  The worksheet
    :returns: Nothing
    """
    for row in ws.iter_rows():
        for cell in row:
            assign_style(cell)
###############################################################################

###############################################################################
def style_workbook(wb):
    # type: (Workbook) -> None
    """
    Assigns the appropriate style to all citation worksheets in a workbook
    :param  wb  The workbook
    :returns: Nothing
    """
    for ws in get_citation_sheets(wb):
        style_citation_sheet(ws)
###############################################################################


###############################################################################
def column_values(wb, ws_name, col_letter):
    ws = wb.get_sheet_by_name(ws_name)
    for c in ws[col_letter]:
        if not c.value is None:
            print('{}{} = {}'.format(col_letter, c.row, c.value))
###############################################################################


###############################################################################
def find_matches(wb,
                 col_name,
                 search_terms,
                 do_re_search   = False,
                 cell_type      = CellType.CT_ALL,
                 max_instances  = -1):
    # type: (Workbook, str, List[string], bool, CellType, int) -> List[Cell]
    """
    Finds cells matching certain conditions on a given column.

    :param  wb:             The workbook
    :param  col_name:       The name of the column to be searched
    :param  search_terms:   The search terms to be matched
    :param  do_re_search:   Whether the search terms should be treated as
                            regular expressions
    :param  cell_type:      The type of cells to search for
    :param: max_instances   The maximum number of matched cells to return
    :returns: The matching cells
    """
    matching_cells = list()

    citation_sheets = get_citation_sheets(wb)
    for ws in citation_sheets:
        search_col  = get_col_id(ws, col_name)
        COL_ID_DEFN = get_col_id(ws, COL_HDR_DEFN)

        #
        # Build the list of matches in the worksheet: begin with the cells
        # that provide a value
        #
        ws_matches = [cell for cell in ws[search_col][1:] if cell.value]

        #
        # Filter based on the search terms
        #
        if do_re_search:
            ws_matches  = [cell for cell in ws_matches if any(re.match(term, cell.value) for term in search_terms)]
        else:
            ws_matches  = [cell for cell in ws_matches if cell.value in search_terms]

        #
        # Filter based on cell type
        #
        if cell_type == CellType.CT_DEFN:
           ws_matches = [cell for cell in ws_matches if
                         ws[COL_ID_DEFN][cell.row - 1].style == STYLE_GENERAL]
        elif cell_type == CellType.CT_REFERRING:
           ws_matches = [cell for cell in ws_matches if
                         ws[COL_ID_DEFN][cell.row - 1].style == STYLE_LINK]

        #
        # Add matches to return list, respecting the maximum instances limit
        #
        cells_to_add = len(ws_matches) if max_instances < 0 else max_instances - len(matching_cells)
        matching_cells += ws_matches[:cells_to_add]
        if (max_instances > 0 and len(matching_cells) == max_instances):
            break

    return matching_cells
###############################################################################


###############################################################################
def get_col_display_value(cell,
                          col_name):
    display_value = ''
    display_delim = ''
    if not cell is None:
        ws  = cell.parent
        row = cell.row

        col_value = ws['{}{}'.format(get_col_id(ws, col_name), row)].value
        if col_name == COL_HDR_CITATION:
            display_value = '-' if col_value is None else '"{}"'.format(col_value)
            display_delim = ' '
        elif col_name == COL_HDR_CATEGORY:
            display_value = '<{}>'.format(col_value if col_value else '-')
            #display_value = '<->' if col_value is None else '<{}>'.format(col_value)
            display_delim = ' '
        elif col_name == COL_HDR_PHRASE or col_name == COL_HDR_DEFN:
            display_value = '' if col_value is None else col_value
            display_delim = '\t'
        elif col_name == COL_HDR_JYUTPING:
            display_value = '' if col_value is None else '({})'.format(col_value)
            display_delim = '\t'

    return display_value, display_delim
###############################################################################



###############################################################################
def show_definition(cell,
                    cols_to_show = [COL_HDR_CATEGORY, COL_HDR_PHRASE,
                                    COL_HDR_JYUTPING, COL_HDR_DEFN],
                    show_cell_ref = True):
    # type: (Cell) -> None
    """
    Displays the definition associated with a cell.

    :param  cell: A cell in a citation worksheet
    :returns: Nothing
    """
    if not cell is None:
        ws  = cell.parent
        row = cell.row
        _, cell_label = get_def_name_id_and_label(ws, cell)

        def_line = ''
        if show_cell_ref:
            def_line += '[{}!{}/{}]\t'.format(ws.title, row, cell_label)
        for col_to_show in cols_to_show:
            col_display_value, delim = get_col_display_value(cell, col_to_show)
            def_line += '{}{}'.format(col_display_value, delim)
        print(def_line)
###############################################################################


###############################################################################
def display_matches(wb,
                    search_terms,
                    do_re_search    = False,
                    col_name        = COL_HDR_PHRASE,
                    cell_type       = CellType.CT_DEFN,
                    max_instances   = -1,
                    cols_to_show    = [COL_HDR_CATEGORY, COL_HDR_PHRASE,
                                       COL_HDR_JYUTPING, COL_HDR_DEFN],
                    show_cell_ref   = True):
    # 
    """
    """
    matches = find_matches(wb, col_name, search_terms, do_re_search, cell_type, max_instances)
    for cell in matches:
        show_definition(cell, cols_to_show, show_cell_ref)
###############################################################################


###############################################################################
def show_definedname_cells(wb):
    # type: (Workbook) -> None
    """
    Displays the definition of cells that are associated with a defined name.

    :param  wb: The workbook
    :returns: Nothing
    """
    defined_name_dict = dict()
    cs_names = [cs_name for cs_name in CITATION_SHEETS if cs_name in wb.sheetnames]
    for cs_name in cs_names:
        defined_name_dict[cs_name] = list()

    defined_name_locations = [dn.attr_text for dn in wb.defined_names.definedName]
    for dn_loc in defined_name_locations:
        ws_name, cell_loc = dn_loc.split('!')
        if not cell_loc in defined_name_dict[ws_name]:
            defined_name_dict[ws_name].append(cell_loc)

    for cs_name in cs_names:
        ws = wb.get_sheet_by_name(cs_name)
        for cell_loc in defined_name_dict[cs_name]:
            show_definition(ws[cell_loc])
###############################################################################


###############################################################################
def def_specified(cell):
    # type (Cell) -> bool
    """
    Checks if a given cell is associated with a definition.

    :param  cell: The cell
    :returns Whether the cell matches up with a definition
    """
    if not cell is None:
        def_loc = '{}{}'.format(get_col_id(cell.parent, COL_HDR_DEFN), cell.row)
        return cell.parent[def_loc].style == STYLE_GENERAL
    return False
###############################################################################


###############################################################################
def find_cells_with_no_def(ws,
                           min_num_chars = 1,
                           max_num_chars = 1):
    # type (Worksheet, int, int) -> List
    """
    Finds cells in a citation worksheet that have no definition

    :param ws:              The worksheet
    :param min_num_chars:   The minimum number of characters in the phrase
    :param max_num_chars:   The maximum number of characters in the phrase,
                            if 0 no upper limit is imposed on the phrase length
    :returns The list of cells with no definition
    """
    phrase_cells = list()
    if max_num_chars > 0:
        phrase_cells = [cell for cell in ws[get_col_id(ws, COL_HDR_PHRASE)] if cell.value and len(cell.value) >= min_num_chars and len(cell.value) <= max_num_chars]
    else:
        phrase_cells = [cell for cell in ws[get_col_id(ws, COL_HDR_PHRASE)] if cell.value and len(cell.value) >= min_num_chars]
    phrase_rows = [cell.row for cell in phrase_cells]
    def_cells = [cell.row for cell in ws[get_col_id(ws, COL_HDR_DEFN)] if cell.row in phrase_rows and not cell.value]
    return [cell for cell in phrase_cells if cell.row in def_cells]
###############################################################################



###############################################################################
def find_cells_with_shape_and_value(wb,           # type Workbook
                                    shape,
                                    value,
                                    pos):           # type int
    """
    Finds cells TODO

    :param  wb:     The workbook
    :param  shape:  The ideographic shape
    :param  value:  The value to be found in the shape
    :param
    :returns: list 
    """
    matching_cells = list()

    cjk = characterlookup.CharacterLookup('T')
    sheets = get_citation_sheets(wb)
    for ws in sheets:
        matches = []

        phrase_cells = [cell for cell in ws[get_col_id(ws, COL_HDR_PHRASE)]
                        if def_specified(cell) and len(cjk.getDecompositionEntries(cell.value)) > 0]

        if not shape:
            radical_index = cjk.getKangxiRadicalIndex(value)
            matches = [cell for cell in phrase_cells if cjk.getCharacterKangxiRadicalIndex(cell.value) == radical_index]
        else:
            matches = [cell for cell in phrase_cells if cjk.getDecompositionEntries(cell.value)[0][0] == shape
                and cjk.getDecompositionEntries(cell.value)[0][pos+1][0] == value]

#       matches = [cell for cell in phrase_cells
#                       if def_specified(cell) and
#                          len(cjk.getDecompositionEntries(cell.value)) > 0 and
#                          cjk.getDecompositionEntries(cell.value)[0][0] == shape and
#                          cjk.getDecompositionEntries(cell.value)[0][pos+1][0] == value]


        #decomps = [(cell, cjk.getDecompositionEntries(cell.value)) for cell in
                #phrase_cells if len(cjk.getDecompositionEntries(cell.value)) > 0 and
                #cjk.getDecompositionEntries(cell.value)[0][0] == shape and
                #cjk.getDecompositionEntries(cell.value)[0][pos+1][0] == value
                #]
        #for decomp in decomps:
            #print(decomp)
        matching_cells.extend(matches)

    return matching_cells
###############################################################################

###############################################################################
def get_citation_sheets(wb):
    # type: (Workbook) -> List
    """
    Returns a workbook's citation worksheets

    :param  wb: The workbook
    :returns:   A list of the citation worksheets
    """
    return [wb.get_sheet_by_name(name) for name in CITATION_SHEETS if name in wb.sheetnames]
###############################################################################


###############################################################################
def fill_in_last_sheet(wb):
    # type: (Workbook) -> None
    """
    Fills in a workbook's latest worksheet based on existing definitions, then
    shows the phrases that require definitions.

    :param  wb: The workbook
    :returns:   A list of the citation worksheets
    """
    citation_sheet_names = [c for c in CITATION_SHEETS if c in wb.sheetnames]

    last_sheet_name = citation_sheet_names[-1]
    ws = wb.get_sheet_by_name(last_sheet_name)
    get_refs_for_ws_phrases(wb, last_sheet_name, True, False)

    missing = find_cells_with_no_def(ws, min_num_chars = 1)
    for m in missing:
        print(m.value)

    missing = find_cells_with_no_def(ws, min_num_chars = 2, max_num_chars = 0)
    for m in missing:
        print('{} {}'.format(m.row, m.value))
###############################################################################


###############################################################################
def save_changes(wb,
                 filename = DEST_FILE):
    # type: (Workbook, str) -> None
    """
    Saves the workbook to the chosen file
    :param  wb          The workbook
    :param  filename    The destination filename
    :returns: Nothing
    """
    wb.save(filename)
###############################################################################


###############################################################################
def main():
    """
    TODO

    :returns: None
    """

###############################################################################


if __name__ == "__main__":
    main()
    notes_wb = load_workbook(SOURCE_FILE)

#   fill_in_last_sheet(notes_wb)
#   save_changes(notes_wb)

#   if  sys.version_info.major ==  2:
#       cjk = characterlookup.CharacterLookup('T')
#       cells = find_cells_with_shape_and_value(notes_wb, CJK_SHAPE_LTR, '口', 0)
#       for cell in cells:
#           show_definition(cell)
