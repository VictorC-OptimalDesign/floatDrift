# === IMPORTS ==================================================================
import enum
import glob
import math
import input
import operator
import os
import re
import shutil
import string
import typing
import xlsxwriter


# === CLASSES ==================================================================

class gyro:
    
    class Row(enum.Enum):
        Header = 0
        Data = 1
        
    class Col(enum.Enum):
        X = 0
        Y = 1
        Z = 2
        
    __COL_PER = Col.Z.value + 1
    
    def __init__(self):
        __FILE_NAME: str = 'gyro'
        __EXTENSION: str = '.xlsx'
        __DATA_SHEET: str = 'data'
        
        self.wb: xlsxwriter.Workbook = xlsxwriter.Workbook(self.__FILE_NAME + self.__EXTENSION)
        self.dataSheet: xlsxwriter.Workbook.worksheet_class = self.wb.add_worksheet(self.__DATA_SHEET)
        self.count: int = 0
        
    def __parseLine(self, line: str) -> typing.List[int]:
        __START_PATTERN: str = 'gyro: (x,y,z)'
        values: typing.List[int] = []
        hits: typing.List[str] = []
        if line.find(__START_PATTERN) > 0:
            # https://stackoverflow.com/questions/12231193/python-split-string-by-start-and-end-characters
            pattern = re.compile(r'\(([^)]+)\)')
            hits = pattern.findall(line)
            if len(hits) is 1:
                hits = hits.split(',')
                # Remove empty strings.
                hits = [i for i in hits if i]
                # Strip each member of leading/trailing whitespace.
                hits = [x.strip() for x in hits]
                values = [int(x, 16) for x in hits]
                values = [x if x < 2^15 else x - 2^16 for x in values]
        return values
    
    def __convertValues(self, values: typing.List[int]) -> typing.List[float]:
        floats: typing.List[float] = [(500.0 * x) / 2^16 for x in values]
        return floats
        
    def process(self):
        __FILE_SEARCH_PATTERN: str = '*.txt'
        __FILE_REPLACE_PATTERN: str = __FILE_SEARCH_PATTERN.replace('*', '')
        
        for fileName in glob.glob(self.__FILE_SEARCH_PATTERN):
            print('processing {0}...'.format(fileName))
            name: str = fileName.replace(self.__FILE_REPLACE_PATTERN, '')
            fullPath: str = os.path.join(os.getcwd(), fileName)
            with open(fullPath, 'r') as file:
                readLines = file.readlines()
            file.close()
            row = self.Row.Header.value
            colOffset = self.count * self.__COL_PER
            self.dataSheet.write(row, colOffset, name)
            row += 1
            for line in readLines:
                values: typing.List[int] = self.__parseLine(line)
                if len(values) > 0:
                    floats: typing.List[float] = self.__convertValues(values)
                    self.dataSheet.write_number(row, colOffset + self.Col.X.value, floats[0])
                    self.dataSheet.write_number(row, colOffset + self.Col.Y.value, floats[0])
                    self.dataSheet.write_number(row, colOffset + self.Col.Z.value, floats[0])
                    row += 1
            self.count += 1
            
    def finalize(self):
        self.wb.close()
        