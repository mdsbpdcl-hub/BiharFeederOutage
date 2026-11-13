/* Lets the dashboard page (made for Google Apps Script) run on GitHub Pages: when the page is not inside Google,
   google.script.run.getDashboardData(fresh) is answered by reading the web app's data link instead.
   The figures are asked for at once, while the rest of the page is still loading. */
(function () {
  if (window.google && window.google.script && window.google.script.run) return;
  var url = ((window.OUTAGE_CONFIG || {}).WEB_APP_URL || "").trim();
  var linkOk = /^https:\/\/script\.google\.com\/.+\/exec$/.test(url);

  function fetchData(fresh) {
    if (!linkOk) return Promise.reject(new Error("The web app link is not set in config.js"));
    return fetch(url + "?api=data" + (fresh ? "&fresh=1" : "") + "&t=" + Date.now(), { method: "GET", cache: "no-store", redirect: "follow" })
      .then(function (r) { if (!r.ok) throw new Error("HTTP " + r.status); return r.text(); })
      .then(function (t) {
        try { return JSON.parse(t); }
        catch (e) { throw new Error("The data link sent a page instead of figures: deploy the new Code.gs as a new version, with access 'Anyone'"); }
      });
  }
  var early = linkOk ? fetchData(false) : null;   // started before the page script runs
  if (early) early.catch(function () {});

  function runner(ok, bad) {
    return {
      withSuccessHandler: function (f) { return runner(f, bad); },
      withFailureHandler: function (f) { return runner(ok, f); },
      getDashboardData: function (fresh) {
        var p = (!fresh && early) ? early : fetchData(!!fresh);
        early = null;
        p.then(function (d) { if (ok) ok(d); }, function (e) { try { console.warn("[dashboard]", e.message); } catch (x) {} if (bad) bad(e); });
      }
    };
  }
  window.google = window.google || {};
  window.google.script = { run: runner(null, null) };
})();
