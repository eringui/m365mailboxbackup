import os
import sys
os.environ["PYTHONUTF8"]="1"
os.environ["PYTHONIOENCODING"]="utf-8"
for stream in (getattr(sys,"stdout",None),getattr(sys,"stderr",None)):
    reconfigure=getattr(stream,"reconfigure",None)
    if callable(reconfigure):
        try: reconfigure(encoding="utf-8",errors="backslashreplace",line_buffering=True)
        except Exception: pass
