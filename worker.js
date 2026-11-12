/* Runs the report program (py/outage_report_app.py, the same file as the Windows app) inside the browser.
   Nothing leaves the computer from here: the page itself sends the finished report to Google. */
importScripts("pyodide/pyodide.js");

let py = null, app = null;

async function ready() {
  if (app) return app;
  post({ kind: "progress", text: "Getting the report program ready (first time takes a few seconds) ..." });
  py = await loadPyodide({ indexURL: new URL("pyodide/", self.location.href).href });
  await py.loadPackage("sqlite3");
  const src = await (await fetch("py/outage_report_app.py", { cache: "no-cache" })).text();
  py.FS.mkdirTree("/app");
  py.FS.writeFile("/app/outage_report_app.py", src);
  py.runPython(`
import sys, types
if "ssl" not in sys.modules:
    _s = types.ModuleType("ssl"); _s.SSLError = type("SSLError", (OSError,), {}); sys.modules["ssl"] = _s
sys.path.insert(0, "/app")
import outage_report_app as app
`);
  app = py.globals.get("app");
  return app;
}

function post(m) { self.postMessage(m); }

function clearDir(dir) {
  try {
    for (const n of py.FS.readdir(dir)) {
      if (n === "." || n === "..") continue;
      const p = dir + "/" + n;
      if (py.FS.isDir(py.FS.stat(p).mode)) { clearDir(p); py.FS.rmdir(p); } else py.FS.unlink(p);
    }
  } catch (e) { /* not there yet */ }
}

function putFiles(files) {
  clearDir("/in");
  py.FS.mkdirTree("/in");
  return files.map((f, i) => {
    const dir = "/in/" + i;
    py.FS.mkdirTree(dir);
    const path = dir + "/" + f.name.replace(/[\\/]/g, "_");
    py.FS.writeFile(path, new Uint8Array(f.bytes));
    return [f.discom, f.company, path];
  });
}

self.onmessage = async (ev) => {
  const m = ev.data;
  try {
    await ready();
    if (m.cmd === "warm") { post({ kind: "ready", id: m.id }); return; }

    if (m.cmd === "guess") {                      // the report date written in the files, as the Windows app offers it
      const inputs = putFiles(m.files);
      py.globals.set("PATHS", py.toPy(inputs.map(x => x[2])));
      const d = py.runPython(`
import datetime as _dt
_g = app.guess_date(list(PATHS))
if isinstance(_g, int): _g = _dt.date.fromordinal(_g)
_g.strftime("%Y-%m-%d") if _g else ""`);
      post({ kind: "guess", id: m.id, date: d });
      return;
    }

    if (m.cmd === "make") {
      const inputs = putFiles(m.files);
      clearDir("/out"); py.FS.mkdirTree("/out");
      py.globals.set("INPUTS", py.toPy(inputs));
      py.globals.set("DATE", m.date);
      py.globals.set("TYPES", m.types ? py.toPy(m.types) : null);   // null: the shared list could not be read
      py.globals.set("JS_LOG", (k, t) => post({ kind: "log", level: k, text: t }));
      py.globals.set("JS_PROGRESS", (f, t) => post({ kind: "progress", fraction: f, text: t }));
      const res = py.runPython(`
import json, threading, os
if TYPES is None:
    app.SHARED_TYPE_ROWS = None
    app.WEB_APP_URL = "PASTE_YOUR_WEB_APP_LINK_HERE"   # no second try from here: the starting list built into the program is used
else:
    app.SHARED_TYPE_ROWS = [tuple(x) for x in TYPES]
_err = None
try:
    _main, _info = app.make_report([tuple(x) for x in INPUTS], "/out", app.parse_date(DATE), False,
                                   lambda k, t: JS_LOG(k, t), lambda f, t: JS_PROGRESS(f, t), threading.Event())
except app.UserError as e:
    _err = str(e)
if _err:
    _out = {"ok": False, "message": _err}
else:
    _t = _info["summary"]
    _out = {"ok": True, "main": os.path.basename(_main),
            "parts": [[d, os.path.basename(p)] for d, p in _info["parts"]],
            "files": sorted(os.listdir("/out")),
            "head": [c[0] for c in _t.cols], "summary": _t.rows,
            "agri_like": _info.get("agri_like", []), "days": _info["days"],
            "date": app.fmt_date(_info["d_max"])}
json.dumps(_out, default=str)`);
      const out = JSON.parse(res);
      if (out.ok) out.blobs = out.files.map(n => ({ name: n, bytes: py.FS.readFile("/out/" + n) }));
      post({ kind: "made", id: m.id, result: out });
      return;
    }

    if (m.cmd === "body") {                       // the text to send to Google for one DISCOM file
      let path = "/out/" + m.name;
      if (m.bytes) {                              // a report file made earlier, chosen on the page
        clearDir("/sent"); py.FS.mkdirTree("/sent");
        path = "/sent/" + m.name.replace(/[\\/]/g, "_");
        py.FS.writeFile(path, new Uint8Array(m.bytes));
      }
      py.globals.set("NAME", path);
      py.globals.set("PW", m.password);
      const res = py.runPython(`
import json
try:
    _d, _b, _n, _s = app.upload_body(NAME, PW)
    _r = {"ok": True, "discom": _d, "body": _b, "rows": _n, "sheets": _s}
except app.UserError as e:
    _r = {"ok": False, "message": str(e)}
json.dumps(_r)`);
      post({ kind: "body", id: m.id, result: JSON.parse(res) });
      py.globals.set("PW", "");
      return;
    }
  } catch (e) {
    post({ kind: "error", id: m.id, message: String((e && e.message) || e).split("\n").slice(-6).join("\n") });
  }
};
