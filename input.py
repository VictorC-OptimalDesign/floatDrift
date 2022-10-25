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
from statistics import mean
from statistics import stdev


# === CLASSES ==================================================================

class gyro:
    
    class Row(enum.Enum):
        Header = 0
        Ave = 1
        Stdev = 2
        Data = 3
        
    class Col(enum.Enum):
        Label = 0
        X = 1
        Y = 2
        Z = 3
        
    class SummaryRow(enum.Enum):
        Header = 0
        Data = 1
        
    class SummaryCol(enum.Enum):
        Name = 0
        AveX = 1
        AveY = 2
        AveZ = 3
        StdevX = 4
        StdevY = 5
        StdevZ = 6
        
    __COL_PER = Col.Z.value + 1
    
    def __init__(self):
        __FILE_NAME: str = 'gyro'
        __EXTENSION: str = '.xlsx'
        __SUMMARY_SHEET: str = 'summary'
        
        self.wb: xlsxwriter.Workbook = xlsxwriter.Workbook(__FILE_NAME + __EXTENSION)
        self.summary: xlsxwriter.Workbook.worksheet_class = self.wb.add_worksheet(__SUMMARY_SHEET)
        #self.sheets: typing.List[xlsxwriter.Workbook.worksheet_class] = []
        self.count: int = 0
        
    def __parseLine(self, line: str) -> typing.List[int]:
        __START_PATTERN: str = 'gyro: (x,y,z)'
        values: typing.List[int] = []
        hits: typing.List[str] = []
        if line.find(__START_PATTERN) < 0:
            # https://stackoverflow.com/questions/12231193/python-split-string-by-start-and-end-characters
            pattern = re.compile(r'\(([^)]+)\)')
            hits = pattern.findall(line)
            if len(hits) == 1:
                hits = hits[0].split(',')
                # Remove empty strings.
                hits = [i for i in hits if i]
                # Strip each member of leading/trailing whitespace.
                hits = [x.strip() for x in hits]
                values = [int(x, 16) for x in hits]
                values = [x if (x < 2**15) else (x - 2**16) for x in values]
        return values
    
    def __convertValues(self, values: typing.List[int]) -> typing.List[float]:
        floats: typing.List[float] = [(500.0 * float(x)) / 2**16 for x in values]
        return floats
        
    def process(self):
        __FILE_SEARCH_PATTERN: str = '*.log'
        __FILE_REPLACE_PATTERN: str = __FILE_SEARCH_PATTERN.replace('*', '')
        
        self.summary.write(self.SummaryRow.Header.value, self.SummaryCol.Name.value, "NAME")
        self.summary.write(self.SummaryRow.Header.value, self.SummaryCol.AveX.value, "X[AVE]")
        self.summary.write(self.SummaryRow.Header.value, self.SummaryCol.AveY.value, "Y[AVE]")
        self.summary.write(self.SummaryRow.Header.value, self.SummaryCol.AveZ.value, "Z[AVE]")
        self.summary.write(self.SummaryRow.Header.value, self.SummaryCol.StdevX.value, "X[STDEV]")
        self.summary.write(self.SummaryRow.Header.value, self.SummaryCol.StdevY.value, "Y[STDEV]")
        self.summary.write(self.SummaryRow.Header.value, self.SummaryCol.StdevZ.value, "Z[STDEV]")
        
        for fileName in glob.glob(__FILE_SEARCH_PATTERN):
            print('processing {0}...'.format(fileName))
            name: str = fileName.replace(__FILE_REPLACE_PATTERN, '')
            fullPath: str = os.path.join(os.getcwd(), fileName)
            with open(fullPath, 'r') as file:
                readLines = file.readlines()
            file.close()
            row = self.Row.Header.value
            sheet: xlsxwriter.Workbook.worksheet_class = self.wb._add_sheet(name)
            sheet.write(row, self.Col.X.value, 'X')
            sheet.write(row, self.Col.Y.value, 'Y')
            sheet.write(row, self.Col.Z.value, 'Z')
            row += 1
            sheet.write(row, self.Col.Label.value, 'AVE')
            row += 1
            sheet.write(row, self.Col.Label.value, 'STDEV')
            row += 1
            xs: typing.List[float] = []
            ys: typing.List[float] = []
            zs: typing.List[float] = []
            for line in readLines:
                values: typing.List[int] = self.__parseLine(line)
                if len(values) > 0:
                    floats: typing.List[float] = self.__convertValues(values)
                    x: float = floats[0]
                    y: float = floats[1]
                    z: float = floats[2]
                    xs.append(x)
                    ys.append(y)
                    zs.append(z)
                    sheet.write_number(row, self.Col.X.value, x)
                    sheet.write_number(row, self.Col.Y.value, y)
                    sheet.write_number(row, self.Col.Z.value, z)
                    row += 1
            aveX = mean(xs)
            aveY = mean(ys)
            aveZ = mean(zs)
            stdevX = stdev(xs)
            stdevY = stdev(ys)
            stdevZ = stdev(zs)
            sheet.write_number(self.Row.Ave.value, self.Col.X.value, aveX)
            sheet.write_number(self.Row.Ave.value, self.Col.Y.value, aveY)
            sheet.write_number(self.Row.Ave.value, self.Col.Z.value, aveZ)
            sheet.write_number(self.Row.Stdev.value, self.Col.X.value, stdevX)
            sheet.write_number(self.Row.Stdev.value, self.Col.Y.value, stdevY)
            sheet.write_number(self.Row.Stdev.value, self.Col.Z.value, stdevZ)
            summaryRow = self.SummaryRow.Data.value + self.count
            self.summary.write(summaryRow, self.SummaryCol.Name.value, name)
            self.summary.write_number(summaryRow, self.SummaryCol.AveX.value, aveX)
            self.summary.write_number(summaryRow, self.SummaryCol.AveY.value, aveY)
            self.summary.write_number(summaryRow, self.SummaryCol.AveZ.value, aveZ)
            self.summary.write_number(summaryRow, self.SummaryCol.StdevX.value, stdevX)
            self.summary.write_number(summaryRow, self.SummaryCol.StdevY.value, stdevY)
            self.summary.write_number(summaryRow, self.SummaryCol.StdevZ.value, stdevZ)
            self.count += 1
            
    def finalize(self):
        self.wb.close()
        