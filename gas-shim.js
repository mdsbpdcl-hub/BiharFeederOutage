/* Lets the dashboard page (made for Google Apps Script) run on GitHub Pages: when the page is not inside Google,
   google.script.run.getDashboardData() is answered by reading the web app's data link instead. */
(function () {
  if (window.google && window.google.script && window.google.script.run) return;
  function runner(ok, bad) {
    return {
      withSuccessHandler: function (f) { return runner(f, bad); },
      withFailureHandler: function (f) { return runner(ok, f); },
      getDashboardData: function () {
        var url = (window.OUTAGE_CONFIG || {}).WEB_APP_URL || "";
        if (!/^https:\/\/script\.google\.com\/.+\/exec$/.test(url)) { if (bad) bad(new Error("The web app link is not set in config.js")); return; }
        fetch(url + "?api=data&t=" + Date.now(), { method: "GET", cache: "no-store", redirect: "follow" })
          .then(function (r) { if (!r.ok) throw new Error("HTTP " + r.status); return r.json(); })
          .then(function (p) { if (ok) ok(p); })
          .catch(function (e) { if (bad) bad(e); });
      }
    };
  }
  window.google = window.google || {};
  window.google.script = { run: runner(null, null) };
})();
