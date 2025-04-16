#!/usr/bin/env python
# -*- coding: utf8 -*-

"""
Script for import significant data from Credit Advice documents from Porsche
Variant A (created 30.03.2025)
"""

__author__ = "Richard Linhart"
__copyright__ = ""
__credits__ = [""]
__license__ = "GPL"
__version__ = "1.0.0"
__maintainer__ = "Richard Linhart"
__email__ = "OK1CTR@gmail.com"
__status__ = "Development"

from PdfExtractor import Extractor
import sys

# Configuration ------------------------------------------------------------------------------------------------------
# Default *.pdf input file name for test purposes only
default_pdf_name = 'data/Porsche/Lieferant 60411 - Beleg 5125143267  FI 6619663053.PDF'
# --------------------------------------------------------------------------------------------------------------------

if __name__ == '__main__':
    ext = Extractor()
    file_list = ext.find_files(default_pdf_name)

    for file_name in file_list:
        print(f'Processing \'{file_name}\'')
        ext.open(file_name)
        ext.open_page(1)
        ext.by_colon('Analysis period')
        ext.by_colon('Purchase document/item')
        ext.by_colon('Material number/EAN')
        ext.by_colon('Credit amount', num=True)
        ext.under_key_on_pos('Receipt number/date', 16)

        page_num = 1
        while True:
            ext.label(f'Pg: {page_num}')
            page_num += 1
            if not ext.open_page(page_num):
                break
            ret = ext.tab_after('Doc. No/fiscal', 2)
            if not ret:
                ext.label('Nothing more.')
                break
        ext.close()

    sys.exit(0)
