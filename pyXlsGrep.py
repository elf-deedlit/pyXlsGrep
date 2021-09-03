#!/usr/bin/env python3
# vim: set ts=4 sw=4 et smartindent ignorecase fileencoding=utf8:
import argparse
import re
import os
# https://openpyxl.readthedocs.io/en/stable/
import openpyxl
from openpyxl.utils.exceptions import InvalidFileException

BASEPATH = os.path.dirname(os.path.abspath(__file__))
FNMATCH = re.compile('(?s:.*\.xlsx?)$')

def in_value(fs: str, value: str) -> bool:
    if not isinstance(value, str):
        value = repr(value)
    return fs.lower() in value.lower()

def search_xlsx(filename: str, fs: str) -> None:
    try:
        # https://openpyxl.readthedocs.io/en/stable/api/openpyxl.reader.excel.html?highlight=load_workbook
        wb = openpyxl.load_workbook(filename, read_only=True, data_only=True)
    except InvalidFileException:
        print(f'{filename}の形式はサポートしていません')
        return
    except PermissionError as err:
        print(f'{filename}が開けません。: {err.strerror}({err.errno})')
        return
    for sheetname in wb.sheetnames:
        sheet = wb[sheetname]
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value and in_value(fs, cell.value):
                    print(f'[{sheetname}]({cell.coordinate})={cell.value}')
    return

def find_xls(path: str, fs: str) -> None:
    for root, dirs, files in os.walk(path):
        for f in files:
            if re.match(FNMATCH, f):
                fullpath = os.path.join(root, f)
                print(f'searching: {fullpath}')
                search_xlsx(fullpath, fs)

def option_parse() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description='エクセルファイルをgrepする')
    parser.add_argument('--basepath', type=str, default=BASEPATH, help='検索パス')
    parser.add_argument('findstr', type=str, help='検索文字列')
    return parser.parse_args()

def main() -> None:
    args = option_parse()
    find_xls(args.basepath, args.findstr)

if __name__ == '__main__':
    main()