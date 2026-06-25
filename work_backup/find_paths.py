import os, sys, io, shutil, csv, re, subprocess, openpyxl
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

BASE = os.getcwd()
files = {f: os.path.join(BASE, f) for f in os.listdir(BASE)}

def find_file(keyword):
    for fname, fpath in files.items():
        if keyword.lower() in fname.lower():
            return fpath
    return None

FILE_V5      = find_file('v5.xlsx')
FILE_V6      = os.path.join(BASE, [f for f in files if 'v5.xlsx' in f][0].replace('v5', 'v6'))
FILE_JISSEKI = find_file('6.24')
FILE_CSV     = find_file('2026実績.csv')
FILE_BACKEND = find_file('backend_calc.py')

print('v5:', FILE_V5, '->', os.path.exists(FILE_V5) if FILE_V5 else False)
print('実績:', FILE_JISSEKI, '->', os.path.exists(FILE_JISSEKI) if FILE_JISSEKI else False)
print('CSV:', FILE_CSV, '->', os.path.exists(FILE_CSV) if FILE_CSV else False)
print('backend:', FILE_BACKEND, '->', os.path.exists(FILE_BACKEND) if FILE_BACKEND else False)
