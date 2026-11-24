"""Builds apps-script/Index.html from the main index.html (drops the GitHub Pages-only PWA parts)."""
import re, pathlib
root = pathlib.Path(__file__).resolve().parent.parent
s = (root / "index.html").read_text(encoding="utf-8")
s = re.sub(r"<!--deploy-only-->.*?<!--/deploy-only-->", "", s, flags=re.S)
s = s.replace('<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n', "")
(root / "apps-script" / "Index.html").write_text(s, encoding="utf-8")
print("wrote apps-script/Index.html", len(s), "bytes")
