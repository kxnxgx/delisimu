import os, sys, io, re, subprocess, openpyxl
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

BASE = os.getcwd()
fmap = {f: os.path.join(BASE, f) for f in os.listdir(BASE)}
def fp(kw):
    for k,v in fmap.items():
        if kw in k: return v
    return None

FILE_V6 = fp("v6.xlsx")
FILE_BACKEND = fp("backend_calc.py")

def n(val,d=0):
    if isinstance(val,(int,float)): return float(val)
    try: return float(str(val).replace(",",""))
    except: return float(d)

wb = openpyxl.load_workbook(FILE_V6)

# == 全体サマリー内のPLC行を探す ==
ws_s = wb["全体サマリー"]
print("=== 全体サマリー C列スキャン (5-35行) ===")
for ri in range(5, 36):
    v = ws_s.cell(ri, 3).value
    if v: print(f"  Row{ri}: '{v}'")
