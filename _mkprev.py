import json, re, os
os.chdir("/Users/nacho/Desktop/ClaudeGOAT/AppsClaude/porra-mundial-2026")
def R(p): return open(p, encoding="utf-8").read()

html = R("RondaEliminatoria/index.html")
data = {
    "data/predictions.json":     json.loads(R("RondaEliminatoria/data/predictions.json")),
    "data/results.json":         json.loads(R("RondaEliminatoria/data/results.json")),
    "../data/team-codes.json":   json.loads(R("data/team-codes.json")),
    "../data/predictions.json":  json.loads(R("data/predictions.json")),
    "../data/results.json":      json.loads(R("data/results.json")),
}
shim = "<script>window.__DATA__=" + json.dumps(data) + ";</script>\n"
html = html.replace("<script>\nlet PRED=null", shim + "<script>\nlet PRED=null")
html = re.sub(
    r"async function getJSON\(url\)\{.*?\n\}",
    ("async function getJSON(url){\n"
     "  const key=url.split('?')[0];\n"
     "  if(window.__DATA__ && key in window.__DATA__) return window.__DATA__[key];\n"
     "  throw new Error('no data for '+key);\n}"),
    html, count=1, flags=re.S)
sel = os.environ.get("SEL", "Nacho A")
html = html.replace("let selected=null;", f'let selected="{sel}";')
open("RondaEliminatoria/_standalone.html", "w", encoding="utf-8").write(html)
print("ok", len(html), "sel=", sel)
