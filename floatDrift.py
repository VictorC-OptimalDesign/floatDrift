# === IMPORTS ==================================================================
import enum
import glob
import math
import input
import operator
import os
import shutil
import string
import typing
import xlsxwriter

# === GLOBAL CONSTANTS =========================================================

DATA_FOLDER = '_DATA'

# === FUNCTIONS ================================================================
    
def __process():
    data: input.gyro = input.gyro()
    data.process()
    data.finalize()


# === MAIN =====================================================================

if __name__ == "__main__":
    __process()
else:
    print("ERROR: floatDrift needs to be the calling python module!")
    