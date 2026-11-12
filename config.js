/* The link of the Google Apps Script web app (Deploy > Manage deployments, the link ending with /exec).
   The dashboard reads its figures from it, and the update page sends reports and the Agriculture list to it.
   The upload password is never written here: it is asked each time and checked by the script. */
window.OUTAGE_CONFIG = {
  WEB_APP_URL: "https://script.google.com/macros/s/AKfycbxzXRkOmfOwReMdvNAW8Baz2h4ppDr7sTP3mO8KvACq6tS7eeDEhXWPHgLqGMhNaWxoXw/exec"
};
