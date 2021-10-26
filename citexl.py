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
from openpyxl.cell.cell import Cell

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

#########################################################################
# Citation fields names:
# values are retrieved directly or generated from a citation worksheet
#########################################################################
CITE_FLD_CHAPTER        = "回"                  # Generated
CITE_FLD_PAGE           = "頁"                  # Retrieved/generated
CITE_FLD_LINE           = "直行"                # Retrieved/generated
CITE_FLD_LINE_INSTANCE  = "instance_number"     # Generated
CITE_FLD_CITE_TEXT      = "引句"                # Retrieved
CITE_FLD_CATEGORY       = "範疇"                # Retrieved
CITE_FLD_TOPIC          = "題"                  # Retrieved
CITE_FLD_PHRASE         = "字詞"                # Retrieved
CITE_FLD_JYUTPING       = "粵拼"                # Retrieved
CITE_FLD_DEFN           = "定義"                # Retrieved/generated
CITE_FLD_ID             = "cite_id"             # Generated
CITE_FLD_LBL_SHORT      = "cite_label_short"    # Generated
CITE_FLD_LBL_VERBOSE    = "cite_label_verbose"  # Generated
CITE_FLD_COUNT          = "citation_count"      # Generated

#
# Get a list of all citation fields
#
CITE_FLDS = [eval(fld) for fld in list(locals().keys()) if  re.match("CITE_FLD_.*", fld)]

###############################################
# Mapping between citation and CantoDict fields
###############################################
CiteFldToDictFld = {
    CITE_FLD_PHRASE:     ccdict.DE_TRAD,
    CITE_FLD_JYUTPING:   ccdict.DE_JYUTPING,
    CITE_FLD_DEFN:       ccdict.DE_ENGLISH
}


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
# Field names for JSON search files
#################################################
MSEARCH_NAME        = "search_name"
MSEARCH_TERMS       = "search_terms"
MSEARCH_TERM_VALUE  = "search_value"
MSEARCH_TERM_FIELD  = "search_field"
MSEARCH_TERM_USE_RE = "use_re_search"


#############################
# Cantonese dictionary object
#############################
canto_dict =  ccdict.CantoDict("cite_dict.db")


###############################################################################
# Miscellaneous spreadsheet helper functions
###############################################################################

###############################################################################
def find_closest_value(ws,
                       col_letter,
                       row_number):
    # type: (Worksheet, str, int) -> (str, int)
    """
    Finds the first non-empty value that appears in the given column, at or
    above the specified row and the row's rank among those sharing that value,
    e.g. if row_number == 5, and the nearest non-empty value is in row 2, rank = 4

    :param  ws:         The worksheet
    :param  col_letter: Column letter
    :param  row_number: Row number (1-based)
    :returns: The value for a given column and row, and the row's rank among
              those sharing a value for that column.
    """

    # Identify non-empty cells at or above the specified row
    non_empty_cells = [c for c in ws[col_letter][:row_number] if not c.value is None]

    if len(non_empty_cells) != 0:
        return non_empty_cells[-1].value, (row_number - non_empty_cells[-1].row + 1)

    return None, None
###############################################################################


###############################################################################
def header_row(ws):
   # type (Worksheet) -> Tuple
   """
   Returns the header row of a worksheet, assumed to contain column names.

   :param ws:  A worksheet
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
def get_named_row(ws,
                  row_number):
    # type: (Worksheet, int) -> Dict
    """
    Retrieves a worksheet row, formatted as a column/field name to Cell mapping.

    :param  ws:         A worksheet
    :param  row_number: A 1-based row number
    :returns a mapping between the column names and Cells of the specified row
    """
    return dict(zip([header_cell.value for header_cell in header_row(ws)],
                    ws[row_number]))
###############################################################################


###############################################################################
# A class that wraps up functionality for managing a citation spreadsheet
###############################################################################
class CitationWB(object):
    def __init__(self,
                 source_file            = SOURCE_FILE,
                 citation_sheet_names   = CITATION_SHEETS):
        """
        Citation workbook constructor

        :param  source_file:            Path to the source spreadsheet file
        :param  citation_sheet_names:   Citation worksheet names
        """
        self.source_file = source_file
        self.citation_sheet_names = citation_sheet_names
        self.wb = None
        self.wb_links = None
        self.reload()


    ###########################################################################
    def reload(self):
        """
        (Re)loads the citation workbook
        """
        if self.wb:
            self.wb.close()
            self.wb_links = None
        self.wb = load_workbook(self.source_file)
    ###########################################################################


    ###########################################################################
    def save_changes(self, filename = DEST_FILE):
        # type: (str) -> None
        """
        Saves the workbook to the chosen file
        :param  filename    The destination filename
        :returns Nothing
        """
        self.wb.save(filename)
    ###########################################################################


    ###########################################################################
    @staticmethod
    def get_defn_links(ws):
        # type (Worksheet) -> List[str]
        """
        Finds links (potentially defined names in other worksheets) in citation
        worksheet definitions

        :param ws:  A citation worksheet
        :returns the list of link targets in the worksheet
        """
        links = [defn_cell.hyperlink.location for defn_cell
                    in ws[get_col_id(ws, CITE_FLD_DEFN)]
                    if defn_cell.hyperlink and defn_cell.hyperlink.location]
        return links
    ###########################################################################


    ###############################################################################
    @staticmethod
    def format_cit_value(cit_value,
                         fld_name):
        # type: (str, str) -> (str, str)
        """
        Generates a formatted version of a citation field value, plus its trailing
        delimiter

        :param  cit_value:  Raw citation field value
        :param  fld_name:   Name of the citation field
        :returns the formatted field value and trailing delimiter.
        """
        display_value = ""
        display_delim = ""
        if fld_name == CITE_FLD_CITE_TEXT:
            display_value = "\"{}\"".format(cit_value) if cit_value else "-"
            display_delim = ' '
        elif fld_name == CITE_FLD_CATEGORY:
            display_value = "<{}>".format(cit_value if cit_value else "-")
            display_delim = " "
        elif fld_name == CITE_FLD_TOPIC:
            display_value = "[{}]".format(cit_value) if cit_value else ""
            display_delim = " "
        elif fld_name == CITE_FLD_PHRASE or fld_name == CITE_FLD_DEFN:
            display_value = cit_value if cit_value else ""
            display_delim = "\t"
        elif fld_name == CITE_FLD_JYUTPING:
            display_value = "({})".format(cit_value) if cit_value else ""
            display_delim = "\t"
        elif fld_name == CITE_FLD_COUNT:
            display_value = "({})".format(cit_value) if cit_value is not None else ""
            display_delim = " "
        elif fld_name == CITE_FLD_LBL_VERBOSE:
            display_value = cit_value if cit_value else ""
            display_delim = "\t"
        else:
            display_value = cit_value if cit_value else ""
            display_delim = " "
        return display_value, display_delim
    ###############################################################################


    ###############################################################################
    @staticmethod
    def format_cit_values(cit_values,
                          cit_fields = [CITE_FLD_CATEGORY, CITE_FLD_PHRASE,
                                        CITE_FLD_JYUTPING, CITE_FLD_DEFN]):
        # type: (Dict, List) -> str
        """
        Retrieves a formatted string corresponding to a citation

        :param  cit_values: Field to value mapping for the citation
        :param  cit_fields: Fields to include
        :returns the selected citation values as formatted string
        """
        cit_str = ""
        for cit_field in cit_fields:
            fld_display_value, delim = CitationWB.format_cit_value(cit_values.get(cit_field, None),
                                                                   cit_field)
            cit_str += "{}{}".format(fld_display_value, delim)
        return cit_str
    ###############################################################################


    ###########################################################################
    def get_citation_sheets(self):
        # type: None -> List[Worksheet]
        """
        Retrieves the workbook's citation worksheets

        :returns a list of the citation worksheets
        """
        return [self.wb.get_sheet_by_name(ws_name) for ws_name in
                self.citation_sheet_names if ws_name in self.wb.sheetnames]
    ###########################################################################


    ###########################################################################
    def get_link_counts(self):
        # type () -> Counter
        """
        :returns a Counter, mapping links to number of occurrences
        """

        if not self.wb_links:
            #
            # Build the list of links across all citation worksheets
            #
            self.wb_links = list()
            for ws in self.get_citation_sheets():
                self.wb_links.extend(CitationWB.get_defn_links(ws))

        #
        # Return the mapping between links and occurrences
        #
        return Counter(self.wb_links)
    ###########################################################################


    ###########################################################################
    def style_workbook(self):
        # type: () -> None
        """
        Assigns the appropriate style to all citation worksheets in the workbook
        :param  wb  The workbook
        :returns: Nothing
        """
        for ws in self.get_citation_sheets():
            style_citation_sheet(ws)
    ###########################################################################


    ###########################################################################
    def get_cit_values(self,
                       cit_row,
                       fields_to_fill = [CITE_FLD_CHAPTER, \
                                         CITE_FLD_PAGE, \
                                         CITE_FLD_LINE, \
                                         CITE_FLD_LINE_INSTANCE, \
                                         CITE_FLD_ID, \
                                         CITE_FLD_LBL_SHORT, \
                                         CITE_FLD_LBL_VERBOSE]):
        # type: (Worksheet, int) -> Dict
        """
        Retrieves a citation row's values, formatted as field name to value mappings

        :param  cit_row:        The citation row
        :param  fields_to_fill: Fields that should be generated/filled in if the
                                worksheet does not provide a (direct) value
        :returns a field name to value for the specified row.
        """
        ws          = cit_row[CITE_FLD_PHRASE].parent
        row_number  = cit_row[CITE_FLD_PHRASE].row
        cit_values  = dict([(cit_data[0], cit_data[1].value) for cit_data in cit_row.items()])

        cit_chapter = ws.title
        cit_page, _             = find_closest_value(ws, get_col_id(ws, CITE_FLD_PAGE), row_number)
        cit_line, cit_instance  = find_closest_value(ws, get_col_id(ws, CITE_FLD_LINE), row_number)
        cit_label_short         = "{}{}{}{}{}".format(cit_chapter, REF_LABEL_SEP, cit_page, REF_LABEL_SEP, cit_line)
        cit_id                  = "{}{}{:02d}{}{:02d}{}{:02d}".format(
                                      cit_chapter, DEF_NAME_ID_SEP,
                                      cit_page, DEF_NAME_ID_SEP,
                                      cit_line, DEF_NAME_ID_SEP, cit_instance)

        cit_values[CITE_FLD_CHAPTER] = ws.title
        if CITE_FLD_PAGE in fields_to_fill:
            cit_values[CITE_FLD_PAGE] = cit_page
        if CITE_FLD_LINE in fields_to_fill:
            cit_values[CITE_FLD_LINE] = cit_line
        if CITE_FLD_LINE_INSTANCE in fields_to_fill:
            cit_values[CITE_FLD_LINE_INSTANCE] = cit_instance
        if CITE_FLD_ID in fields_to_fill:
            cit_values[CITE_FLD_ID] = cit_id
        if CITE_FLD_LBL_SHORT in fields_to_fill:
            cit_values[CITE_FLD_LBL_SHORT] = cit_label_short
        if CITE_FLD_LBL_VERBOSE in fields_to_fill:
            cit_values[CITE_FLD_LBL_VERBOSE] = "{}!{}".format(cit_label_short, row_number)
        if CITE_FLD_COUNT in fields_to_fill:
            link_counter = self.get_link_counts()
            cit_values[CITE_FLD_COUNT] = link_counter[cit_id] + 1
        if CITE_FLD_DEFN in fields_to_fill:
            wb  = ws.parent
            defn_cell = cit_row.get(CITE_FLD_DEFN, None)
            if defn_cell and defn_cell.hyperlink and \
               defn_cell.hyperlink.location in wb.defined_names:
                defn_source_defined_name = list(wb.defined_names[defn_cell.hyperlink.location].destinations)[0]
                defn_source_ws = wb.get_sheet_by_name(defn_source_defined_name[0])
                defn_source = defn_source_ws[defn_source_defined_name[1]]
                defn_source_row = get_named_row(defn_source.parent, defn_source.row)
                cit_values[CITE_FLD_DEFN] = defn_source_row[CITE_FLD_DEFN].value
                cit_values[CITE_FLD_JYUTPING] = defn_source_row[CITE_FLD_JYUTPING].value

        return cit_values
    ###########################################################################


    ###############################################################################
    def format_cit_row(self,
                       cit_row,
                       cit_fields = [CITE_FLD_LBL_VERBOSE, CITE_FLD_CATEGORY,
                                     CITE_FLD_PHRASE,
                                     CITE_FLD_JYUTPING, CITE_FLD_DEFN]):
        # type: (Dict, List) -> str
        """
        Retrieves a formatted string corresponding to a citation

        :param  cit_row:    A citation row
        :param  cit_fields: Fields to include
        :returns the citation as a formatted string
        """

        #
        # Retrieve the mapping between citation field names and values
        #
        cit_values = self.get_cit_values(cit_row, fields_to_fill = cit_fields)
        return CitationWB.format_cit_values(cit_values, cit_fields)
    ###############################################################################


    ###########################################################################
    def build_reference(self,
                        referenced_row,
                        referring_row,
                        overwrite       = False,
                        audit_only      = False):
        # type: (Dict, Dict, bool, bool) -> None
        """
        Builds a reference between two citation rows.
        The referencing is achieved by creating a defined name that includes
        the referenced row's phrase cell, and linking to this name from the
        referring row's phrase cell.

        :param  referenced_row  The referenced row
        :param  referring_row   The referring row
        :param  overwrite       If True, overwrite any existing content in the
                                referring definition cell (if this isn't already the
                                label for the referenced phrase cell)
        :param  audit_only      If True, show/print the actions for building the
                                reference, but do not modify the data
        :returns: Nothing
        """

        #
        # Tracks actions required in order to build the reference
        #
        audit_log = list()

        #
        # Retrieve referenced and referring cells and worksheets
        #
        referenced_cell = referenced_row[CITE_FLD_PHRASE]
        referenced_ws   = referenced_cell.parent
        referring_cell  = referring_row[CITE_FLD_DEFN]
        referring_ws    = referring_cell.parent

        #
        # Retrieve an identifier for the defined name, and the label for referring
        # to the referenced cell (to be displayed in the referring cell)
        #
        referenced_cit_vals = self.get_cit_values(referenced_row, CITE_FLDS)
        def_name_id = referenced_cit_vals[CITE_FLD_ID]
        label = referenced_cit_vals[CITE_FLD_LBL_SHORT]

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
                audit_log.append("Create defined name: {}: {}".format(def_name_destination, def_name_id))

                if not audit_only:
                    workbook.defined_names.append(def_name)

        write_needed = referring_cell.value is None or (overwrite and referring_cell.value != label)
        referring_cell_loc = referring_cell.coordinate
        if write_needed:
            audit_log.append("{}!{} current value = {}".format(referring_ws.title,
                                                               referring_cell_loc,
                                                               referring_cell.value))
            if not audit_only:
                referring_cell.value = label
                referring_cell.hyperlink = Hyperlink(ref = referring_cell_loc,
                                                     location = def_name_id)

                # Clear Jyutping cell contents
                jyutping_cell       = referring_row[CITE_FLD_JYUTPING]
                jyutping_cell.value = None

                # Assign styles
                assign_style(referring_cell)
                assign_style(jyutping_cell)

        if len(audit_log) != 0:
            print("{}!{} ({}) --> {}!{}".format(
                  referring_ws.title, referring_cell.coordinate, referring_cell.value,
                  referenced_ws.title, referenced_cell.coordinate))
            for audit_msg in audit_log:
                print("\t{}".format(audit_msg))
    ###########################################################################


    ###########################################################################
    def find_matches(self,
                     search_terms,
                     fld_name,
                     do_re_search       = False,
                     cell_type          = CellType.CT_ALL,
                     max_instances      = -1):
        # type: (Workbook, str/List[str], str, bool, CellType, int) -> List[Dict]
        """
        Finds citation rows matching conditions on a given column.

        :param  search_terms:   Search term/s to be matched
        :param  fld_name:       Name of the field to be searched
        :param  do_re_search:   If True, treat search terms as regular expressions
        :param  cell_type:      Type of cells to search for
        :param  max_instances:  Maximum number of matched citations to return
        :returns a list of column name to Cell mappings for the matching rows.
        """
        wb = self.wb

        #
        # Initialise match list
        #
        matching_rows = list()

        #
        # Perform search over all citation sheets
        #
        citation_sheets = self.get_citation_sheets()
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
                                               fld_name,
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
    ###########################################################################


    ###########################################################################
    def find_matches_for_file(self,
                              search_terms_filename,
                              cell_type                 = CellType.CT_DEFN):
        # type (str, CellType) -> Dict
        """
        Find matches in the workbook for search terms specified in a JSON file.
        For search terms with no hits in the citation workbook, revert to a
        dictionary lookup.

        :param  search_terms_filename:  Name of the file containing the terms
        :param  cell_type:              Type of cells to search for
        :returns a dictionary mapping search groups (provided by the JSON file)
                 to search results
        """

        #
        # Initialise the return value
        #
        file_matches = dict()

        with open(search_terms_filename) as search_terms_file:
            search_groups = json.load(search_terms_file)
            for search_group in search_groups:
                search_name = search_group[MSEARCH_NAME]
                search_terms = search_group[MSEARCH_TERMS]
                for search_term in search_terms:
                    #
                    # Retrieve matches for this search group
                    #
                    group_matches = file_matches.get(search_name, list())

                    search_value = search_term.get(MSEARCH_TERM_VALUE)
                    search_field = search_term.get(MSEARCH_TERM_FIELD, CITE_FLD_PHRASE)
                    use_re_search = search_term.get(MSEARCH_TERM_USE_RE, False)
                    citation_matches = self.find_matches(search_value,
                                                         search_field,
                                                         use_re_search,
                                                         max_instances = 1)
                    if len(citation_matches) == 1:
                        cit_values = self.get_cit_values(citation_matches[0], CITE_FLDS)
                        group_matches.append(cit_values)
                    else:
                        #
                        # Generate an artificial citation mapping via a dictionary
                        # search
                        #
                        dict_matches = canto_dict.search_dict(search_value)
                        for dict_match in dict_matches:
                            cit_values = dict()
                            for cite_fld in CiteFldToDictFld.keys():
                                ccdict_field = CiteFldToDictFld.get(cite_fld)
                                if ccdict_field:
                                    cit_values[cite_fld] = dict_match.get(ccdict_field)
                            cit_values[CITE_FLD_COUNT] = 0
                            group_matches.append(cit_values)
                    file_matches[search_name] = group_matches
            return file_matches
    ###########################################################################


    ###########################################################################
    def find_cells_with_shape_and_value(self,
                                        shape,
                                        value,
                                        pos):
        # type (Workbook, str, str, int) -> List[Cell]
        """
        Finds phrase cells in a citation workbook that fit CJK shape conditions

        :param  shape:  Ideographic shape to match
        :param  value:  Value to match within the shape
        :param  pos:
        :returns: list
        """
        matching_cells = list()

        cjk = characterlookup.CharacterLookup('T')
        sheets = self.get_citation_sheets()
        for ws in sheets:
            matches = []

            phrase_cells = [cell for cell in ws[get_col_id(ws, CITE_FLD_PHRASE)]
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
    ###########################################################################


    ###########################################################################
    def display_cit_values_list(self,
                                cit_values_list,
                                cit_fields      = CITE_FLDS,
                                prefix          = ""):
        # type: (List(Dict), List(str), str) -> None
        """
        Displays a list of citation row values

        :param cit_values_list: A list of citation rows, formatted as field
                                to value mappings
        :param cit_fields:      Fields to display
        :param prefix:          A string to prefix to each citation line
        """
        for cit_values in cit_values_list:
            print("{}{}".format(prefix, CitationWB.format_cit_values(cit_values, cit_fields)))
    ###########################################################################


    ###########################################################################
    def display_matches(self,
                        search_terms,
                        fld_name        = CITE_FLD_PHRASE,
                        do_re_search    = False,
                        cell_type       = CellType.CT_DEFN,
                        max_instances   = -1,
                        cit_fields      = [CITE_FLD_LBL_VERBOSE, CITE_FLD_CATEGORY,
                                           CITE_FLD_PHRASE, CITE_FLD_JYUTPING,
                                           CITE_FLD_DEFN]):
        # type: (str/List(str), str, bool, CellType, int, List[str]) -> None
        """
        Displays the matches for one or more search terms in the citations workbook

        :param  search_terms:   Search term/s to be matched
        :param  fld_name:       Name of the field to be searched
        :param  do_re_search:   If True, treat search terms as regular expressions
        :param  cell_type:      Type of cells to search for
        :param  max_instances:  Maximum number of matched cells to return
        :param  cit_fields:     Fields to display
        :returns nothing
        """
        matching_rows = self.find_matches(search_terms, fld_name, do_re_search, cell_type, max_instances)
        cit_values_list = [self.get_cit_values(cit_row, cit_fields)
                               for cit_row
                               in self.find_matches(search_terms, fld_name, do_re_search, cell_type, max_instances)]
        self.display_cit_values_list(cit_values_list, cit_fields)
    ###########################################################################


    ###############################################################################
    def display_matches_for_file(self,
                                 search_terms_filename,
                                 cell_type              = CellType.CT_DEFN,
                                 cit_fields             = [CITE_FLD_LBL_VERBOSE,
                                                           CITE_FLD_COUNT,
                                                           CITE_FLD_CATEGORY,
                                                           CITE_FLD_PHRASE,
                                                           CITE_FLD_JYUTPING,
                                                           CITE_FLD_DEFN]):
        # type (str, bool, str, CellType, List[str], bool) -> None
        """
        Find and display matches in the workbook for search terms specified in
        a JSON file.
        For search terms with no hits in the workbook, revert to a dictionary
        lookup.

        :param  search_terms_filename:  Name of the file containing the terms
        :param  cell_type:              Type of cells to search for
        :param  cit_fields:             Citation fields to show
        :returns Nothing
        """
        file_matches = self.find_matches_for_file(search_terms_filename, cell_type)
        for search_name, citations in file_matches.items():
            print(search_name)
            self.display_cit_values_list(citations, cit_fields, "\t")
    ###############################################################################


    ############################################################################
    def display_multiply_used_defns(self,
                                    cit_fields = [CITE_FLD_COUNT,
                                                  CITE_FLD_PHRASE,
                                                  CITE_FLD_DEFN]):
        # type () -> None
        """
        Shows the definition of phrases that are recorded multiple times in the
        citation workbook.

        :param  cit_fields: Fields to display
        :returns Nothing
        """

        wb = self.wb
        cit_values_list = list()

        #
        # Traverse the link counter in descending number of occurrences,
        # collecting the citations associated with the links for display
        #
        link_counter = self.get_link_counts()
        for link_name, ref_count in link_counter.most_common():
            defined_name = wb.defined_names.get(link_name)
            ws_name, coordinate = defined_name.attr_text.split('!')
            ws = wb.get_sheet_by_name(ws_name)
            cit_row = get_named_row(ws, ws[coordinate].row)
            cit_values_list.append(self.get_cit_values(cit_row, cit_fields))
        self.display_cit_values_list(cit_values_list, cit_fields)
    ############################################################################


    ############################################################################
    def get_refs_for_ws_phrases(self,
                                ws_name,
                                overwrite,
                                audit_only):
        # type: (Workbook, str, bool, bool) -> None
        """
        Fills in the references for the phrase column in a citation worksheet.
        This requires building a reference to the first occurrence of each phrase
        in the workbook.

        :param  ws_name:    A citation name
        :param  overwrite:  If True, overwrite any existing content in referring
                            cells if these don't already refer to the referenced
                            cells
        :param  audit_only: If True, show/print the actions for building references
                            without modifying any data
        :returns Nothing
        """
        if ws_name in self.citation_sheet_names and ws_name in self.wb.sheetnames:
            ws = self.wb.get_sheet_by_name(ws_name)
            #
            # Iterate through all non-empty phrase cells of the chapter worksheet
            #
            for phrase_cell in [c for c in ws[get_col_id(ws, CITE_FLD_PHRASE)][1:] if c.value]:
                #
                # Find the first cell to define the phrase
                #
                referenced_rows = self.find_matches([phrase_cell.value], CITE_FLD_PHRASE, False, CellType.CT_DEFN, 1)
                referenced_row  = referenced_rows[0] if len(referenced_rows) > 0 else None
                referenced_cell = referenced_row[CITE_FLD_PHRASE] if referenced_row else None

                if referenced_cell and not referenced_cell == phrase_cell:
                    #
                    # Ensure this cell isn't the one providing the definition!
                    #
                    referring_row = get_named_row(ws, phrase_cell.row)
                    self.build_reference(referenced_row, referring_row, overwrite, audit_only)
    ###########################################################################


    ###########################################################################
    def fill_in_sheet(self,
                      ws_name,
                      overwrite = False):
        # type: (str, bool) -> None
        """
        Fills in a citation worksheet by:
            1. Building links to prior citations that define the same phrase
            2. Performing a dictionary lookup for first occurrences of phrases
            3. Displaying cited phrases that still require a definitions after 1/2.

        :parm   ws_name:    Name of a citation worksheet
        :param  overwrite:  If True, overwrite any existing content in cells to be
                            filled out
        :returns Nothing
        """
        if ws_name in self.citation_sheet_names and ws_name in self.wb.sheetnames:
            ws = self.wb.get_sheet_by_name(ws_name)
            self.get_refs_for_ws_phrases(ws_name, overwrite, False)

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
                phrase_cell         = citation_row[CITE_FLD_PHRASE]
                citation_value, _   = find_closest_value(ws, get_col_id(ws, CITE_FLD_CITE_TEXT), phrase_cell.row)
                print("{}:\t{}".format(phrase_cell.row, phrase_cell.value))
                print("\t\t{}".format(citation_value))
    ###########################################################################


    ###########################################################################
    def fill_in_last_sheet(self,
                           overwrite = False):
        # type: (bool) -> None
        """
        Fills in the latest citation worksheet based on existing definitions,
        then shows the phrases that require definitions.

        :param  overwrite:  If True, overwrite any existing content in cells
                            to be filled out
        :returns Nothing
        """
        valid_citation_sheet_names = [s for s in self.citation_sheet_names if s in self.wb.sheetnames]
        last_sheet_name = valid_citation_sheet_names[-1]

        self.fill_in_sheet(last_sheet_name, overwrite)
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
# TODO: STATIC METHOD?
def find_matches_in_sheet(ws,
                          search_terms,
                          fld_name,
                          do_re_search       = False,
                          cell_type          = CellType.CT_ALL,
                          max_instances      = -1):
    # type: (Worksheet, str/List[str], str, bool, CellType, int) -> List[Dict]
    """
    Finds citation rows matching conditions on a given column.

    :param  wb:             A citation worksheet
    :param  search_terms:   Search term/s to be matched, converted to a list if
                            necessary
    :param  fld_name:       Name of the field to be searched
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
    # Add None to allow searching for blank column values if appropriate
    #
    if "" in search_terms and None not in search_terms:
        search_terms.append(None)

    #
    # Build the list of cells in the worksheet matching the search terms
    #
    search_col  = get_col_id(ws, fld_name)
    matching_cells = ws[search_col][1:]
    if do_re_search:
        matching_cells  = [cell for cell in matching_cells if cell.value and any(re.search(term, cell.value) for term in search_terms)]
    else:
        matching_cells  = [cell for cell in matching_cells if cell.value in search_terms]

    #
    # Retrieve the corresponding rows
    #
    matching_rows = [get_named_row(ws, cell.row) for cell in matching_cells]

    #
    # Filter based on cell type
    #
    if cell_type == CellType.CT_DEFN:
        matching_rows = [r for r in matching_rows if r[CITE_FLD_DEFN] and r[CITE_FLD_DEFN].style == STYLE_GENERAL]
    elif cell_type == CellType.CT_REFERRING:
        matching_rows = [r for r in matching_rows if r[CITE_FLD_DEFN] and r[CITE_FLD_DEFN].style == STYLE_LINK]

    #
    # Trim the results list to the maximum instances limit
    #
    if (max_instances > 0):
        matching_rows = matching_rows[:max_instances]

    return matching_rows
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
            print(format_cit_row(get_named_row(ws, ws[cell_loc].row)))
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
        def_loc = '{}{}'.format(get_col_id(cell.parent, CITE_FLD_DEFN), cell.row)
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
    citation_rows = find_matches_in_sheet(ws, "", CITE_FLD_DEFN)
    if max_num_chars > 0:
        citation_rows = [row for row in citation_rows if row[CITE_FLD_PHRASE].value and
                                                         len(row[CITE_FLD_PHRASE].value) >= min_num_chars and
                                                         len(row[CITE_FLD_PHRASE].value) <= max_num_chars]
    else:
        citation_rows = [row for row in citation_rows if row[CITE_FLD_PHRASE].value and
                                                         len(row[CITE_FLD_PHRASE].value) >= min_num_chars]
    return citation_rows
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
    phrase_cell = citation_row[CITE_FLD_PHRASE]
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
    dict_search_res = canto_dict.search_dict(phrase_cell.value)
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

    jyut_cell = citation_row[CITE_FLD_JYUTPING]
    defn_cell = citation_row[CITE_FLD_DEFN]

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
def main():
    """
    TODO

    :returns: None
    """

###############################################################################


if __name__ == "__main__":
    main()
#   notes_wb = load_workbook(SOURCE_FILE)

    citewb = CitationWB()
    citewb.fill_in_last_sheet(True)
#   for sheet_name in citewb.citation_sheet_names:
#       citewb.fill_in_sheet(sheet_name, True)
#   citewb.save_changes()

    citewb.display_matches_for_file("confounds.match",
                                    cit_fields = CITE_FLDS)

#   citewb.display_matches("..武", CITE_FLD_TOPIC, do_re_search = True, cit_fields = CITE_FLDS)
    citewb.display_matches("囉", cit_fields = CITE_FLDS)
#   citewb.display_matches("..武", CITE_FLD_TOPIC, do_re_search = True, cit_fields = [CITE_FLD_LBL_VERBOSE, CITE_FLD_CITE_TEXT, CITE_FLD_CATEGORY, CITE_FLD_TOPIC, CITE_FLD_PHRASE, CITE_FLD_DEFN])
#   citewb.display_matches("..武", CITE_FLD_TOPIC, do_re_search = True, cit_fields = [CITE_FLD_CITE_TEXT, CITE_FLD_LBL_VERBOSE, CITE_FLD_CATEGORY, CITE_FLD_TOPIC, CITE_FLD_PHRASE, CITE_FLD_DEFN])

    citewb.display_multiply_used_defns()

    if  sys.version_info.major ==  2:
        show_char_decomposition('彆')
#       cjk = characterlookup.CharacterLookup('T')
#       cells = citewb.find_cells_with_shape_and_value(CJK_SHAPE_LTR, '口', 0)
#       for cell in cells:
#           print(format_cit_row(cell.parent, cell.row))
