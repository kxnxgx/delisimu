import os, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
BASE = os.getcwd()
# ファイル名一覧からパターンマッチ
files = os.listdir(BASE)
for f in sorted(files):
    print(repr(f))
