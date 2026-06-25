import os, sys, io, csv
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
BASE = os.getcwd()
files = {f: os.path.join(BASE, f) for f in os.listdir(BASE)}
FILE_CSV = next(v for k,v in files.items() if "2026実績" in k)
for enc in ["utf-8","utf-8-sig","cp932"]:
    try:
        with open(FILE_CSV, encoding=enc, newline="") as f:
            reader = csv.DictReader(f)
            fn = reader.fieldnames
            first = next(reader)
        print("エンコード:", enc)
        print("フィールド名:", fn)
        print("先頭行:", dict(first))
        break
    except Exception as e:
        print(f"  {enc} 失敗: {e}")
