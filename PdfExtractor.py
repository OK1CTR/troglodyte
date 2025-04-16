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
from os.path import isfile, join, dirname
from pymupdf4llm.helpers.get_text_lines import get_text_lines

# Configuration ------------------------------------------------------------------------------------------------------
# Item separator in output files
sep = ';'
# --------------------------------------------------------------------------------------------------------------------

class Extractor:
    def __init__(self):
        self.pdf = None
        self.csv = None
        self.page = None
        self.lines = None
        self.page_number = None
        self.common_output = None
        self.path = None

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
        If the input is a number, convert it form US format to CZ format
        :param in_str: Input string containing number in US format
        :return: Output string containing number in CZ format
        """
        if in_str.count('.') != 2:
            x = in_str.replace(',', '')
            out_str = x.replace('.', ',')
            return out_str
        else:
            return in_str

    @classmethod
    def no_dot(cls, in_str):
        """
        If the input string is not a date, removes dots form it
        :param in_str: Input string containing number
        :return: Output string containing number
        """
        if in_str.count('.') != 2:
            out_str = in_str.replace('.', '')
            return out_str
        else:
            return in_str

    def open(self, file_name, make_csv=True):
        """
        Load the selected input *.pdf file into the extractor and open the output *.csv file
        :param file_name:
        :param make_csv: If true, also a *.csv file of the same name will be created (default True)
        :return: True if successful
        """
        # load the *.pdf file
        try:
            self.pdf = pymupdf.open(file_name)
        except FileNotFoundError:
            print(f'File {file_name} not found!')
            return False

        # create the corresponding *.csv file
        if make_csv:
            file_name_parts = file_name.split( '.')
            if len(file_name_parts) != 2:
                print(f'Wrong file name \'{file_name}\'!')
                return False
            csv_file_name = file_name_parts[0] + '.csv'
            try:
                self.csv = open(csv_file_name, 'w')
            except FileNotFoundError:
                print(f'Cannot create {csv_file_name} file!')
                return False

        # save project path
        self.path = dirname(file_name)
        return True

    def close(self):
        """
        Close the destination output *.csv file
        :return: True
        """
        assert self.pdf is not None
        if self.csv is not None:
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
        self.lines = get_text_lines(self.page).split('\n')
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
        assert self.lines is not None

        key_find = key + ':'
        ret = False

        for line in self.lines:
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
        assert self.lines is not None

        line_num = 0
        ret = False

        for line in self.lines:
            if line.find(key) != -1:
                ret = True
                break
            line_num += 1

        if not ret:
            return False

        line = self.lines[line_num + 1]
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
        assert self.lines is not None

        export = False
        counter = num
        output = ''

        for line in self.lines:

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

    def grep(self, phrase):
        """
        Return line index containing given phrase on page
        :param phrase: Phrase to search
        :return: Line index
        """
        count = 0
        for line in self.lines:
            if line.find(phrase) != -1:
                return count
            count += 1
        return None

    def common_output_open(self, file_name):
        """
        Open common output *.csv file
        :param file_name: Output file name
        :return: True if successful
        """
        try:
            self.common_output = open(file_name, 'w')
        except FileNotFoundError:
            print(f'Cannot create {file_name} file!')
            return False
        return True

    def common_output_write(self, data):
        """
        Write row into the common output file
        :param data: Input data list
        :return: True if successful
        """
        if len(data) > 0:
            self.common_output.write(f'{data[0]}')
        for item in data[1:]:
            self.common_output.write(f'{sep}{item}')
        self.common_output.write('\n')

    def common_output_close(self):
        """
        Close common output file
        :return: True
        """
        assert self.common_output is not None
        self.common_output.close()
        return True

    def drop_page(self, page_number=None):
        """
        Save selected page to a text file for manual analysis
        :param page_number: Number of page to drop or None for currently open page
        :return:
        """
        # obtain the text
        if page_number is None:
            lines = self.lines
            page_number = self.page_number
        else:
            page_number = page_number - 1
            if page_number not in self.pdf:
                print(f'Cannot find page {page_number}')
                return False
            lines = get_text_lines(self.pdf[page_number]).split('\n')

        # make file name
        file_name = f'{self.path}/page_{(page_number + 1):03d}.txt'
        try:
            with open(file_name, 'w') as fw:
                for line in lines:
                    fw.write('%s\n' % line)
        except FileNotFoundError:
            print(f'Cannot create {file_name} file!')
            return False
        return True
