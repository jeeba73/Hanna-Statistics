import base64,sys
b64=sys.stdin.read().strip()
open(r"c:/Users/riccardo.gibertini/Desktop/Hanna Statistics/extract_excel.py","w",encoding="utf-8").write(base64.b64decode(b64).decode("utf-8"))
print("Script written")
