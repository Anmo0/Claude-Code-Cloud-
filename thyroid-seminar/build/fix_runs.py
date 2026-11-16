"""Remove mid-paragraph <a:pPr> elements pptxgenjs emits after a run (they cancel bullets on mixed-format lines)."""
import re, sys, zipfile, shutil, os
src = sys.argv[1]; tmp = src + ".tmp"
pat = re.compile(r'(</a:r>)<a:pPr\b[^>]*?(?:/>|>.*?</a:pPr>)', re.S)
n = 0
with zipfile.ZipFile(src) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename.startswith("ppt/slides/slide") and item.filename.endswith(".xml"):
            x = data.decode("utf-8"); x, k = pat.subn(r"\1", x); n += k; data = x.encode("utf-8")
        zout.writestr(item, data)
shutil.move(tmp, src); print("removed", n, "stray pPr")
