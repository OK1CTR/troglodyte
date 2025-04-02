#!/usr/bin/env python
# -*- coding: utf8 -*-

"""
Class for extract significant data from pdf office documents
"""

__author__ = "Richard Linhart"
__copyright__ = ""
__credits__ = [""]
__license__ = "GPL"
__version__ = "1.0.0"
__maintainer__ = "Richard Linhart"
__email__ = "OK1CTR@gmail.com"
__status__ = "Development"
__all__ = ["Extractor"]


import pymupdf
import re
import sys
from os import listdir
from os.path import isfile, join
from pymupdf4llm.helpers.get_text_lines import get_text_lines

# Configuration ------------------------------------------------------------------------------------------------------
# Item sparator in output files
sep = ';'
# --------------------------------------------------------------------------------------------------------------------

class Extractor:
    def __init__(self):
        self.pdf = None
        self.csv = None
        self.page = None
        self.text = None
        self.page_number = None

    @classmethod
    def find_files(cls, def_file_name):
        """
        In accordance with 1st program argument prepare the list of files to process
        :param def_file_name: Default file name for test
        :return: List of file names to process
        """
        # no argument - default file name
        file_list = [def_file_name]
        if len(sys.argv) > 1:
            if sys.argv[1].find('.pdf') != -1 or sys.argv[1].find('.PDF') != -1:
                # argument is *.pdf or *.PDF file
                file_list = [sys.argv[1]]
            else:
                # argument is directory
                directory = sys.argv[1]
                file_list = [x for x in listdir(directory) if isfile(join(directory, x))]
                file_list = [x for x in file_list if x.find('.pdf') != -1 or x.find('.PDF') != -1]
                file_list = [join(directory, x) for x in file_list]
        return file_list

    @classmethod
    def num(cls, in_str):
        """
        In the input is a number, convert it form US format to CZ format
        :param in_str: Input string containing number in US format
        :return: Output string containing number in CZ format
        """
        if in_str.count('.') != 2:
            x = in_str.replace(',', '')
            out_str = x.replace('.', ',')
            return out_str
        else:
            return in_str

    def open(self, file_name):
        """
        Load the selected input *.pdf file into the extractor and open the output *.csv file
        :param file_name:
        :return: True if successful
        """
        f_name_parts = file_name.split( '.')
        if len(f_name_parts) != 2:
            print(f'Wrong file name \'{file_name}\'!')
            return False
        f_name = f_name_parts[0] + '.csv'
        try:
            self.pdf = pymupdf.open(file_name)
            self.csv = open(f_name, 'w')
        except FileNotFoundError:
            print(f'File {file_name} not found!')
            return False
        return True

    def close(self):
        """
        Close the destination output *.csv file
        :return: True
        """
        assert self.csv is not None
        assert self.pdf is not None
        self.csv.close()
        self.pdf.close()
        return True

    def open_page(self, page_number):
        """
        Find the requested page in the input file opened
        :param page_number:
        :return: True
        """
        assert self.pdf is not None
        page_number = page_number - 1
        if page_number not in self.pdf:
            print(f'Cannot find page {page_number}')
            return False
        self.page = self.pdf[page_number]
        self.text = get_text_lines(self.page)
        self.page_number = page_number
        return True

    def by_colon(self, key, num=False):
        """
        Extract pair Key: Value separated by a colon
        :param key: Phrase to find
        :param num: If true, reformat numbers US to CZ
        :return: True if ket found
        """
        assert self.csv is not None
        assert self.text is not None

        lines = self.text.split('\n')
        key_find = key + ':'
        ret = False

        for line in lines:
            if line.find(key_find) != -1:
                line_parts = line.split(':')
                value = line_parts[1]
                value = value.strip()
                if num:
                    self.csv.write(key + sep + self.num(value) + '\n')
                else:
                    self.csv.write(key + sep + value + '\n')
                ret = True
                break
        return ret

    def under_key_on_pos(self, key, pos):
        """
        Extract string on position under the line containing the Key
        :param key: Phrase to find
        :param pos: Position of value on next line
        :return: True if key found
        """
        assert self.csv is not None
        assert self.text is not None

        lines = self.text.split('\n')
        line_num = 0
        ret = False

        for line in lines:
            if line.find(key) != -1:
                ret = True
                break
            line_num += 1

        if not ret:
            return False

        line = lines[line_num + 1]
        line = line[pos:]
        self.csv.write(key + sep + line.strip() + '\n')
        return ret

    def tab_after(self, header, num):
        """
        Extract table after a header
        :param header: String in line, after which a table is following
        :param num: Number of text lines belonging one table line
        :return: True if header found
        """
        assert self.csv is not None
        assert self.text is not None

        lines = self.text.split('\n')
        export = False
        counter = num
        output = ''

        for line in lines:

            # find the header
            if line.find(header) != -1:
                export = True
                continue

            # export the table
            if export:
                # terminate export
                if len(line) > 0 and line[0].isalpha():
                    break

                # export line
                line_parts = re.split('[\t ]', line)
                line_parts = [x for x in line_parts if x]
                line_parts = [x.strip() for x in line_parts]
                line_parts = [x.replace('\n', '') for x in line_parts]
                line_parts = [self.num(x) for x in line_parts]
                if len(line_parts) > 0:
                    output += line_parts[0]
                    del line_parts[0]
                else:
                    continue
                for part in line_parts:
                    output += sep + part

                counter -= 1
                if counter > 0:
                    output += sep
                    continue
                else:
                    self.csv.write(output + '\n')
                    output = ''
                    counter = num
        return export

    def space(self):
        """
        Insert a blank line into the output file
        :return: True
        """
        assert self.csv is not None
        self.csv.write('\n')
        return True

    def label(self, text):
        """
        Insert a label line into the output file
        :param text: A label text
        :return: True
        """
        assert self.csv is not None
        self.csv.write(text + '\n')
        return True
