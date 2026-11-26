import sys, glob, re
from openpyxl import load_workbook
pat = re.compile(r"tian\s*wei|tan\s*tian|tianwei|tiān", re.I)
for f in sorted(glob.glob("mcp-5461be80-*-blob-17955146*.xlsx")):
    try:
        wb = load_workbook(f, read_only=True, data_only=True)
    except Exception as e:
        print(f[-14:], "ERR", e); continue
    n=0
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            line = " | ".join(str(c) for c in row if c is not None)
            if pat.search(line):
                n+=1
                if n<=6: print(f[-12:], ws.title, "::", line[:200])
    print(f[-12:], "hits", n)
