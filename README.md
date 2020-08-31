# ExcelMerge

Merge rows in two excel files, with exact or fuzzy-matching algorithms.

## Features

1. GUI for easy use
2. Fuzzy matching with various algorithms, to match data with typos or from different sources.
3. Join types - Outer Join, Left join and Inner join supported.
4. Export to Excel
5. Shows warning is multiple rows match with fuzzy matching.

## Example

![TODO: add image]()

## Installation

### Windows

Download the [single EXE file here](https://github.com/Kartikay26/excelmerge/releases/download/v1.0/ExcelMerge.exe)

### Linux

Clone the code, run `python3 main.py`

Requirements:

1. Python 3 and pip `sudo apt install python3 python3-pip`
2. PyQt5 `pip install PyQt5`
3. OpenPyXL `pip install openpyxl`
4. FuzzyWuzzy `pip install fuzzywuzzy`