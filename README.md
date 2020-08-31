# ExcelMerge

Merge rows in two excel files, with exact or fuzzy-matching algorithms.

## Features

1. GUI for easy use
2. Fuzzy matching with various algorithms, to match data with typos or from different sources.
3. Join types - Outer Join, Left join and Inner join supported.
5. Shows warning if multiple rows match with fuzzy matching.
4. Export to Excel. (`*.xlsx`)

## Example

![Example screenshot with two files](https://i.imgur.com/OyIosSS.png)

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


## Known Issues

Feel free to open a PR!

1. Merging large files makes the UI slow.
2. Exe is too big and takes time to load.
3. Fuzzy algorithms could be improved / expanded.
