#!/usr/bin/env python
# -*- coding: utf8 -*-

"""
Script for import significant data from Corrected Self Billed Invoice from BMW
Variant A (created 16.04.2025)
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
# Common output file name with path
output_file = 'data/output.csv'

# Default *.pdf input file name for test purposes only
default_pdf_name = 'data/bmw/SND_GFB_A1_15054410_101550881_de96a04a-4790-4b71-a3a9-4e7a1a8f16d6.pdf'
#default_pdf_name = 'data/bmw/SND_GFB_A1_15054410_101511933_e14f0337-a893-41d3-b21a-e60740bd3ea1.pdf'
#default_pdf_name = 'data/bmw/SND_GFB_A1_15054410_101370251_e1579c4c-8da4-4845-b9db-198358268929_VZOR.pdf'
# --------------------------------------------------------------------------------------------------------------------

def correct_correction(item_in):
    """
    Invert sign of the price difference and the total amount due to direction of the price change
    :param item_in: List of items of the price difference table row
    :return: Corrected price difference table row
    """
    try:
        old_price = float(item_in[1].replace(',', '.'))
        new_price = float(item_in[2].replace(',', '.'))
        diff = float(item_in[3].replace(',', '.'))
        total = float(item_in[4].replace(',', '.'))
    except [ValueError, IndexError]:
        print('Invalid data in price difference table!')
        sys.exit(0)
    if old_price > new_price:
        diff *= -1
        total *= -1
    s_diff = f'{diff:.02f}'.replace('.', ',')
    s_total = f'{total:.02f}'.replace('.', ',')
    return [item_in[0], item_in[1], item_in[2], s_diff, s_total]


if __name__ == '__main__':
    ext = Extractor()
    file_list = ext.find_files(default_pdf_name)
    ext.common_output_open(output_file)

    for file_name in file_list:
        print(f'Processing \'{file_name}\'')
        ext.open(file_name, make_csv=False)
        ext.open_page(1)
        row = []
        #ext.drop_page()

        # document number
        ln = ext.grep('Document number') + 2
        items = ext.lines[ln].split('\t')
        row += [items[0]]
        # print(items[-1])

        # date
        ln = ext.grep('Buchhaltung')
        items = ext.lines[ln].split('\t')
        row += [items[-1]]
        # print(items[-1])

        # number
        ln = ext.grep('Number')
        items = ext.lines[ln].split('\t')
        row += [items[-1]]
        # print(items[-1])

        # correction table
        ln = ext.grep('Service date') + 2
        items = ext.lines[ln].split('\t')
        items = [ext.no_dot(x) for x in items]
        row += correct_correction(items[1:6])
        #print(items[1:6])
        ln += 2
        items = ext.lines[ln].split('\t')
        row += [items[0]]
        # print(items[0])
        ln += 2
        items = ext.lines[ln].split('\t')
        if len(items) == 1:
            items = ext.lines[ln + 1].split('\t')
        items = items[1].split(' ')
        row += [items[0]]
        # print(items[-1])

        # due date
        ext.open_page(2)
        ln = ext.grep('Due date')
        items = ext.lines[ln].split('\t')
        row += [items[-1]]
        # print(items[1])

        ext.common_output_write(row)
        ext.close()

    ext.common_output_close()
    sys.exit(0)
