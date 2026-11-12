"""Makes index.html (the dashboard on this site) from the Apps Script page Dashboard.html.
Use:  python tools/build_index.py path/to/Dashboard.html
Only three things change: the Google-frame line goes, the page reads its figures through gas-shim.js,
and the call links open the dialer directly (no Google frame here)."""
import sys, os

src = sys.argv[1] if len(sys.argv) > 1 else os.path.join("..", "apps_script", "Dashboard.html")
here = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
html = open(src, encoding="utf-8").read()

def swap(old, new, count=1):
    global html
    n = html.count(old)
    if n != count:
        sys.exit("Expected %d of %r in Dashboard.html, found %d. Nothing was written." % (count, old, n))
    html = html.replace(old, new)

swap('<base target="_top">\n', '')
swap("href:'tel:' + tel, target:'_blank', rel:'noopener', ", "href:'tel:' + tel, ")
swap('<script>\n', '<script src="config.js"></script>\n<script src="gas-shim.js"></script>\n<script>\n')
open(os.path.join(here, "index.html"), "w", encoding="utf-8", newline="\n").write(html)
print("index.html written (%d bytes)." % len(html.encode("utf-8")))
