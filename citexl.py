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
import json
import re
import sys
from collections import Counter
from io import BytesIO

from openpyxl import load_workbook
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.workbook.defined_name import DefinedName

from os import path
from enum import IntEnum    # Backported to python 2.7 by https://pypi.org/project/enum34

from pprint import pprint   # Pretty printing of lists, tuples, etc.

import sys
if sys.version_info.major == 2:
    from cjklib import characterlookup

import ccdict

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
COL_HDR_TOPIC      = '題'
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
    '三圖',
    '三十一', '三十二', '三十三', '三十四', '三十五', '三十六', '三十七', '三十八', '三十九', '四十',
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
CJK_SHAPE_LTR   = u'\u2ff0'         # ⿰    Left to right
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

#################################################
#
#################################################
FILE_TERMS_GROUP = "GRP:"


#############################
# Cantonese dictionary object
#############################
canto_dict =  ccdict.CantoDict("cite_dict.db")


###############################################################################
def header_row(ws):
    # type (Worksheet) -> Tuple
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
    # type (Worksheet) -> Dict
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
    # type (Worksheet, str) -> str
    """
    Returns the letter corresponding to the column with the given name in the
    specified worksheet.

    :param ws:          The worksheet
    :param col_name:    The column name
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
    above the specified row and the row's rank among those sharing that value,
    e.g. if row == 5, and the nearest non-empty value is in row 2, rank = 4

    :param  ws:         The worksheet
    :param  col_letter: Column letter
    :param  row:        Row number (1-based)
    :returns: The value for a given column and row, and the row's rank among
              those sharing a value for that column.
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
    page number, line number of the citation, plus its rank amongst citations
    for that chapter/page/line number combination.

    :param  ws:     The worksheet
    :param  cell:   A cell in the worksheet
    :returns: The defined name and label for referring to the cell.
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
    citations.
    This requires building a reference to the first occurrence of each phrase
    in a previous chapter worksheet.

    :param  wb:         A citation workbook
    :param  ws_name:    A citation name
    :param  overwrite:  If True, overwrite any existing content in referring
                        cells if these don't already refer to the referenced
                        cells
    :param  audit_only: If True, show/print the actions for building references
                        without modifying any data
    :returns Nothing
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
            referenced_rows = find_matches(wb, [phrase_cell.value], COL_HDR_PHRASE, False, CellType.CT_DEFN, 1)
            referenced_row = referenced_rows[0] if len(referenced_rows) > 0 else None
            referenced_cell = referenced_row[COL_HDR_PHRASE] if referenced_row else None

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
                    overwrite       = False,
                    audit_only      = False):
    # type: (Worksheet, Cell, Worksheet, Cell, bool, bool) -> None
    """
    Builds a reference to one cell in another.
    The referencing is achieved by creating a defined name that includes
    the referenced cell, and linking to this name from the referring cell.

    :param  referenced_ws       The referenced worksheet
    :param  referenced_cell     The referenced cell
    :param  referring_ws        The referring worksheet
    :param  referring_cell      The referring cell
    :param  overwrite           If True, overwrite any existing content in the
                                referring cell (if this isn't already a label
                                for the referenced cell)
    :param  audit_only          If True, show/print the actions for building
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
def get_citation_sheets(wb):
    # type: (Workbook) -> List[Worksheet]
    """
    Returns a workbook's citation worksheets

    :param  wb: A citation workbook
    :returns a list of the citation worksheets
    """
    return [wb.get_sheet_by_name(name) for name in CITATION_SHEETS if name in wb.sheetnames]
###############################################################################


###############################################################################
def get_row(ws,
            row):
    # type: (Worksheet, int) -> Dict
    """
    Retrieves a citation row, formatted as a column name to Cell mapping.

    :param  ws:     A citation worksheet
    :param  row:    A 1-based row number
    :returns a column name to Cell mapping for the specified row.
    """
    return dict(zip([header_cell.value for header_cell in header_row(ws)],
                    ws[row]))
###############################################################################


###############################################################################
def find_matches(wb,
                 search_terms,
                 col_name,
                 do_re_search       = False,
                 cell_type          = CellType.CT_ALL,
                 max_instances      = -1):
    # type: (Workbook, str, List[str], bool, CellType, int) -> List[Dict]
    """
    Finds citation rows matching conditions on a given column.

    :param  wb:             A citation workbook
    :param  search_terms:   Search term/s to be matched
    :param  col_name:       Name of the column to be searched
    :param  do_re_search:   If True, treat search terms as regular expressions
    :param  cell_type:      Type of cells to search for
    :param  max_instances:  Maximum number of matched citations to return
    :returns a list of column name to Cell mappings for the matching rows.
    """

    #
    # Initialise match list
    #
    matching_rows = list()

    #
    # Perform search over all citation sheets
    #
    citation_sheets = get_citation_sheets(wb)
    max_ws_instances = max_instances
    for ws in citation_sheets:
        #
        # Check whether the maximum instances limit has been reached
        #
        if max_ws_instances == 0:
            break

        #
        # Find matches in the latest worksheet
        #
        ws_matches = find_matches_in_sheet(ws,
                                           search_terms,
                                           col_name,
                                           do_re_search,
                                           cell_type,
                                           max_ws_instances)
        if len(ws_matches) > 0:
            #
            # Add matches to return list and update the next worksheet's limit
            #
            matching_rows += ws_matches

            if max_instances > 0:
                max_ws_instances -= len(ws_matches)

    return matching_rows
###############################################################################


###############################################################################
def find_matches_in_sheet(ws,
                          search_terms,
                          col_name,
                          do_re_search       = False,
                          cell_type          = CellType.CT_ALL,
                          max_instances      = -1):
    # type: (Workbook, str, List[str], bool, CellType, int) -> List[Dict]
    """
    Finds citation rows matching conditions on a given column.

    :param  wb:             A citation worksheet
    :param  search_terms:   Search term/s to be matched, converted to a list if
                            necessary
    :param  col_name:       Name of the column to be searched
    :param  do_re_search:   If True, treat search terms as regular expressions
    :param  cell_type:      Type of cells to search for
    :param  max_instances:  Maximum number of matched cells to return
    :returns a list of column name to Cell mappings for the matching rows.
    """

    #
    # Correct formatting of search_terms
    #
    if isinstance(search_terms, str):
        search_terms = [search_terms]

    #
    # Replace the empty string with None to allow searching for blank column values
    #
    search_terms = [search_term if search_term else None for search_term in search_terms]

    #
    # Build the list of cells in the worksheet matching the search terms
    #
    search_col  = get_col_id(ws, col_name)
    matching_cells = ws[search_col][1:]
    if do_re_search:
        matching_cells  = [cell for cell in matching_cells if cell.value and any(re.search(term, cell.value) for term in search_terms)]
    else:
        matching_cells  = [cell for cell in matching_cells if cell.value in search_terms]

    #
    # Retrieve the corresponding rows
    #
    matching_rows = [get_row(ws, cell.row) for cell in matching_cells]

    #
    # Filter based on cell type
    #
    if cell_type == CellType.CT_DEFN:
        matching_rows = [r for r in matching_rows if r[COL_HDR_DEFN] and r[COL_HDR_DEFN].style == STYLE_GENERAL]
    elif cell_type == CellType.CT_REFERRING:
        matching_rows = [r for r in matching_rows if r[COL_HDR_DEFN] and r[COL_HDR_DEFN].style == STYLE_LINK]

    #
    # Trim the results list to the maximum instances limit
    #
    if (max_instances > 0):
        matching_rows = matching_rows[:max_instances]

    return matching_rows
###############################################################################


###############################################################################
def format_cell_value(cell,
                      col_name):
    # type: (Cell, str) -> (str, str)
    """
    Retrieves the formatted value and trailing delimiter for a given cell.

    :param  cell:           A cell in a citation worksheet
    :param  col_name:       The name of the column to be displayed
    :returns the formatted column value and trailing delimiter.
    """
    display_value = ""
    display_delim = ""
    if not cell is None:
        col_value = cell.value
        if col_name == COL_HDR_CITATION:
            display_value = "\"{}\"".format(col_value) if col_value else "-"
            display_delim = ' '
        elif col_name == COL_HDR_CATEGORY:
            display_value = '<{}>'.format(col_value if col_value else '-')
            display_delim = ' '
        elif col_name == COL_HDR_TOPIC:
            display_value = "[{}]".format(col_value) if col_value else ""
            display_delim = " "
        elif col_name == COL_HDR_PHRASE or col_name == COL_HDR_DEFN:
            display_value = col_value if col_value else ''
            display_delim = '\t'
        elif col_name == COL_HDR_JYUTPING:
            display_value = '({})'.format(col_value) if col_value else ''
            display_delim = '\t'
    return display_value, display_delim
###############################################################################


###############################################################################
def get_formatted_citation(citation_row,
                           cols_to_show = [COL_HDR_CATEGORY, COL_HDR_PHRASE,
                                           COL_HDR_JYUTPING, COL_HDR_DEFN],
                           show_cell_ref = True):
    # type: (Cell) -> None
    """
    Retrieves a formatted string corresponding to a citation

    :param  citation_row:   A row from a citation worksheet
    :param  cols_to_show:   Columns to include
    :param  show_cell_ref:  If True, prefix the string with the cell label
    :returns the citation as a formatted string
    """
    formatted_defn = None

    phrase_cell = citation_row[COL_HDR_PHRASE]
    ws  = phrase_cell.parent
    row = phrase_cell.row
    _, cell_label = get_def_name_id_and_label(ws, phrase_cell)

    formatted_defn = ''
    if show_cell_ref:
        formatted_defn += '[{}!{}]\t'.format(cell_label, row)
    for col_to_show in cols_to_show:
        col_display_value, delim = format_cell_value(citation_row[col_to_show],
                                                     col_to_show)
        formatted_defn += '{}{}'.format(col_display_value, delim)

    return formatted_defn
###############################################################################


###############################################################################
def display_matches(wb,
                    search_terms,
                    col_name        = COL_HDR_PHRASE,
                    do_re_search    = False,
                    cell_type       = CellType.CT_DEFN,
                    max_instances   = -1,
                    cols_to_show    = [COL_HDR_CATEGORY, COL_HDR_PHRASE,
                                       COL_HDR_JYUTPING, COL_HDR_DEFN],
                    show_cell_ref   = True):
    # type: (Workbook, List, bool, str, CellType, int, List[str], bool) -> None
    """
    Displays the matches for one or more search terms in a citations workbook

    :param  wb:             A citation workbook
    :param  search_terms:   Search term/s to be matched
    :param  do_re_search:   If True, treat search terms as regular expressions
    :param  col_name:       Name of the column to be searched
    :param  cell_type:      Type of cells to search for
    :param  max_instances:  Maximum number of matched cells to return
    :param  cols_to_show:   Columns to be displayed
    :param  show_cell_ref:  If True, prefix each displayed row with the cell label
    :returns: Nothing
    """
    matching_rows = find_matches(wb, search_terms, col_name, do_re_search, cell_type, max_instances)
    for row in matching_rows:
        print(get_formatted_citation(row, cols_to_show, show_cell_ref))
###############################################################################


###############################################################################
def show_definedname_cells(wb):
    # type: (Workbook) -> None
    """
    Displays the definition of workbook cells associated with a defined name.

    :param  wb: A citation workbook
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
            print(get_formatted_citation(get_row(ws, ws[cell_loc].row)))
###############################################################################


###############################################################################
def def_specified(cell):
    # type (Cell) -> bool
    """
    Checks if a given cell is associated with a definition.

    :param  cell:   The cell
    :returns True if the cell matches up with a definition
    """
    if not cell is None:
        def_loc = '{}{}'.format(get_col_id(cell.parent, COL_HDR_DEFN), cell.row)
        return cell.parent[def_loc].style == STYLE_GENERAL
    return False
###############################################################################


###############################################################################
def find_citations_with_no_def(ws,
                               min_num_chars = 1,
                               max_num_chars = 1):
    # type (Worksheet, int, int) -> List[Dict]
    """
    Finds citation rows with no associated definition

    :param ws:              A citation worksheet
    :param min_num_chars:   Minimum number of characters in the phrase
    :param max_num_chars:   Maximum number of characters in the phrase,
                            if 0 no upper limit is imposed on the phrase length
    :returns the list of citation rows with no definition
    """
    citation_rows = find_matches_in_sheet(ws, '', COL_HDR_DEFN)
    if max_num_chars > 0:
        citation_rows = [row for row in citation_rows if row[COL_HDR_PHRASE].value and
                                                         len(row[COL_HDR_PHRASE].value) >= min_num_chars and
                                                         len(row[COL_HDR_PHRASE].value) <= max_num_chars]
    else:
        citation_rows = [row for row in citation_rows if row[COL_HDR_PHRASE].value and
                                                         len(row[COL_HDR_PHRASE].value) >= min_num_chars]
    return citation_rows
###############################################################################


###############################################################################
def get_links(ws):
    # type (Worksheet) -> List[str]
    """
    Finds links (potentially defined names in other worksheets) in citation
    worksheet definitions

    :param ws:  A citation worksheet
    :returns the list of link targets in the worksheet
    """
    links = [c.hyperlink.location for c in ws[get_col_id(ws, COL_HDR_DEFN)] if c.hyperlink and c.hyperlink.location]
    return links
###############################################################################


###############################################################################
def get_link_counts(wb):
    # type (Workbook) -> Counter
    """

    :param wb:  A citation workbook
    :returns a Counter, mapping links to number of occurrences
    """

    #
    # Build the list of links across all citation worksheets
    #
    links = list()
    citation_sheets = get_citation_sheets(wb)
    for ws in citation_sheets:
        links.extend(get_links(ws))

    #
    # Return the mapping between links and occurrences
    #
    return Counter(links)
###############################################################################


###############################################################################
def show_multiply_used_defns(wb):
    # type (Workbook) -> None
    """
    Shows the definition of phrases that are recorded multiple times in a
    citation workbook.

    :param wb:  A citation workbook
    :returns Nothing
    """

    #
    # Retrieve mapping between links and occurrences and traverse it in
    # descending number of occurrences.
    #
    link_counter = get_link_counts(wb)
    for link_name, ref_count in link_counter.most_common():
        defined_name = notes_wb.defined_names.get(link_name)
        ws_name, cell_loc = defined_name.attr_text.split('!')
        ws = wb.get_sheet_by_name(ws_name)
        print("({}) {}".format(ref_count + 1, get_formatted_citation(get_row(ws, ws[cell_loc].row))))
###############################################################################


###############################################################################
def show_char_decomposition(c):
    # type (char) -> Nothing
    """
    Shows the CJK shape decomposition of a character

    :param  c:  A character
    """
    cjk = characterlookup.CharacterLookup('T')
    decs = cjk.getDecompositionEntries(c)
    for dec in decs:
        print(dec)

###############################################################################


###############################################################################
def find_cells_with_shape_and_value(wb,
                                    shape,
                                    value,
                                    pos):
    # type (Workbook, str, str, int) -> List[Cell]
    """
    Finds phrase cells in a citation workbook that fit CJK shape conditions

    :param  wb:     A citation workbook
    :param  shape:  Ideographic shape to match
    :param  value:  Value to match within the shape
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
def find_matches_for_file(wb,
                          search_terms_filename,
                          do_re_search    = False,
                          col_name        = COL_HDR_PHRASE,
                          cell_type       = CellType.CT_DEFN,
                          cols_to_show    = [COL_HDR_CATEGORY, COL_HDR_PHRASE,
                                             COL_HDR_JYUTPING, COL_HDR_DEFN],
                          show_cell_ref   = True):
    # type (Workbook, str, bool, str, CellType, List[str], bool) -> None
    """
    Find and display matches in a citation workbook for terms listed in a file

    :param  wb:                     A citation workbook
    :param  search_terms_filename:  Name of the file containing the terms
    :param  do_re_search:           If True, treat search terms as regular expressions
    :param  col_name:               Name of the column to be searched
    :param  cols_to_show:           Columns to be displayed
    :param  show_cell_ref:          If True, prefix each displayed row with the
                                    cell label
    :returns Nothing
    """

    #
    # Retrieve a mapping between links and occurrences, for use in the
    # output...
    #
    link_counter = get_link_counts(wb)

    with open(search_terms_filename) as search_terms_file:
        search_term = search_terms_file.readline()
        while search_term:
            if (search_term != '\n'):
                search_term = re.sub('\n$', '', search_term)
                if re.match('^' + FILE_TERMS_GROUP, search_term):
                    group_name = re.sub('^' + FILE_TERMS_GROUP + '\s+', '', search_term)
                    print(group_name)
                else:
                    matches = find_matches(wb,
                                           search_term,
                                           col_name,
                                           do_re_search     = do_re_search,
                                           cell_type        = cell_type,
                                           max_instances    = 1)
                    if len(matches) == 1:
                        row = matches[0]
                        cell_name, _ = get_def_name_id_and_label(row[COL_HDR_PHRASE].parent,
                                                                 row[COL_HDR_PHRASE])
                        occurrences = link_counter[cell_name] + 1
                        formatted_citation = get_formatted_citation(row,
                                                                    cols_to_show,
                                                                    show_cell_ref)
                        print("\t({}) {}".format(occurrences, formatted_citation))
            else:
                print()
            search_term = search_terms_file.readline()
###############################################################################


###############################################################################
def fill_defn(citation_row,
              overwrite = False,
              audit_only = False):
    # type: (Dict, bool, bool) -> bool
    """
    Fills in the definition (including Jyutping transcription) of a citation
    row.

    :param  citation_row:   A row from a citation worksheet
    :param  overwrite:      If True, overwrites existing definition information
    :param  audit_only:     If True, print the definition data without
                            modifying the citation worksheet
    :returns True if definition data was found
    """

    #
    # Check for the existence of a valid phrase cell before proceeding
    #
    phrase_cell = citation_row[COL_HDR_PHRASE]
    if not phrase_cell or not phrase_cell.value:
        return False

    #
    # Retrieve the worksheet and row via the phrase_cell
    #
    ws          = phrase_cell.parent
    row_number  = phrase_cell.row

    INTRA_DEF_SEP   = ", "
    INTER_DEF_SEP   = ";\n"

    jsonDecoder = json.JSONDecoder()

    #
    # Each search result bundles up a list of English definitions and
    # Jyutping transcriptions corresponding to the phrase
    #
    dict_search_res = canto_dict.search(phrase_cell.value)
    defn_vals = list()
    jyutping_vals = list()
    for search_res in dict_search_res:
        #
        # Generate English and Jyutping strings
        #
        defn_list = jsonDecoder.decode(search_res[ccdict.DE_ENGLISH])
        defn_list = list(filter(None, defn_list))
        if len(defn_list) != 0:
            defn_vals.append(INTRA_DEF_SEP.join(defn_list))
        jyutping_list = jsonDecoder.decode(search_res[ccdict.DE_JYUTPING])
        jyutping_list = list(filter(None, jyutping_list))
        if len(jyutping_list) == 0:
            jyutping_list.append("?")
        jyutping_vals.append(INTRA_DEF_SEP.join(jyutping_list))

    jyut_cell = citation_row[COL_HDR_JYUTPING]
    defn_cell = citation_row[COL_HDR_DEFN]

    print("{}:\t{}".format(phrase_cell.row, phrase_cell.value))
    print(INTER_DEF_SEP.join(["\t{}".format(jyutping) for jyutping in jyutping_vals]))
    print(INTER_DEF_SEP.join(["\t{}".format(defn) for defn in defn_vals]))

    if not audit_only:
        if not jyut_cell.value or overwrite:
            jyut_cell.value = INTER_DEF_SEP.join(jyutping_vals)
            assign_style(jyut_cell)
        if not defn_cell.value or overwrite:
            defn_cell.value = INTER_DEF_SEP.join(defn_vals)
            assign_style(defn_cell)

    return len(defn_vals) > 0 or len(jyutping_vals) > 0
###############################################################################


###############################################################################
def fill_in_sheet(wb,
                  ws_name,
                  overwrite = False):
    # type: (Workbook) -> None
    """
    Fills in a citation worksheet by:
        1. Building links to prior citations that define the same phrase
        2. Performing a dictionary lookup for first occurrences of phrases
        3. Displaying cited phrases that still require a definitions after 1/2.

    :param  wb:         A citation workbook
    :parm   ws_name:    Name of a citation worksheet
    :param  overwrite:  If True, overwrite any existing content in cells to be
                        filled out
    :returns Nothing
    """
    if ws_name in CITATION_SHEETS and ws_name in wb.sheetnames:
        ws = wb.get_sheet_by_name(ws_name)
        get_refs_for_ws_phrases(wb, ws_name, overwrite, False)

        citations_with_no_def = list()

        #
        # Attempt to fill in definitions/Jyutping for single character phrases
        #
        citations_to_fill = find_citations_with_no_def(ws)
        for citation_row in citations_to_fill:
            if not fill_defn(citation_row, overwrite, False):
                citations_with_no_def.append(citation_row)

        #
        # Repeat for multi-character phrases
        #
        citations_to_fill = find_citations_with_no_def(ws, 2, -1)
        for citation_row in citations_to_fill:
            if not fill_defn(citation_row, overwrite, False):
                citations_with_no_def.append(citation_row)


        print("Definition still required...")
        for citation_row in citations_with_no_def:
            phrase_cell         = citation_row[COL_HDR_PHRASE]
            citation_value, _   = find_closest_value(ws, get_col_id(ws, COL_HDR_CITATION), phrase_cell.row)
            print("{}:\t{}".format(phrase_cell.row, phrase_cell.value))
            print("\t\t{}".format(citation_value))
###############################################################################


###############################################################################
def fill_in_last_sheet(wb,
                       overwrite = False):
    # type: (Workbook) -> None
    """
    Fills in a workbook's latest citation worksheet based on existing
    definitions, then shows the phrases that require definitions.

    :param  wb:         A citation workbook
    :param  overwrite:  If True, overwrite any existing content in cells to be
                        filled out
    :returns Nothing
    """
    citation_sheet_names = [c for c in CITATION_SHEETS if c in wb.sheetnames]
    last_sheet_name = citation_sheet_names[-1]

    fill_in_sheet(wb, last_sheet_name, overwrite)
###############################################################################


###############################################################################
def save_changes(wb,
                 filename = DEST_FILE):
    # type: (Workbook, str) -> None
    """
    Saves the workbook to the chosen file
    :param  wb          A citation workbook
    :param  filename    The destination filename
    :returns Nothing
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

#   find_matches_for_file(notes_wb, 'confounds_list',
#                         cols_to_show = [COL_HDR_PHRASE, COL_HDR_JYUTPING, COL_HDR_DEFN],
#                         show_cell_ref = False)

#   for sheet_name in CITATION_SHEETS:
#       fill_in_sheet(notes_wb, sheet_name, True)
#   fill_in_last_sheet(notes_wb)
#   save_changes(notes_wb)

    if  sys.version_info.major ==  2:
        show_char_decomposition('彆')
#       cjk = characterlookup.CharacterLookup('T')
#       cells = find_cells_with_shape_and_value(notes_wb, CJK_SHAPE_LTR, '口', 0)
#       for cell in cells:
#           print(get_formatted_citation(cell.parent, cell.row))

#   display_matches(notes_wb, "..武", COL_HDR_TOPIC, do_re_search = True, cols_to_show = [COL_HDR_CITATION, COL_HDR_CATEGORY, COL_HDR_TOPIC, COL_HDR_PHRASE, COL_HDR_DEFN])

#   ws = notes_wb.get_sheet_by_name(CITATION_SHEETS[0])
#   citation_row = get_row(ws, 604)
#   print(get_formatted_citation(citation_row, [COL_HDR_CITATION, COL_HDR_CATEGORY, COL_HDR_TOPIC, COL_HDR_PHRASE, COL_HDR_DEFN]))
