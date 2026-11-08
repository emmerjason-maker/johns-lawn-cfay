/** @OnlyCurrentDoc */
// Game Night Scoreboard loader: paste this ONCE into Apps Script and deploy it as a web app.
// It runs the latest scoreboard code from the website, so future updates need no pasting
// or redeploying (changes go live within about 10 minutes).
const CODE_URL = "https://emmerjason-maker.github.io/johns-lawn-cfay/pushy-dos-apps-script.gs";

function app_() {
  const cache = CacheService.getScriptCache();
  let src = cache.get("scoreboard-src");
  if (!src) {
    src = UrlFetchApp.fetch(CODE_URL + "?t=" + Date.now()).getContentText();
    cache.put("scoreboard-src", src, 600);
  }
  return new Function(src + "\nreturn { doGet: doGet, doPost: doPost };")();
}

function doGet(e) { return app_().doGet(e); }
function doPost(e) { return app_().doPost(e); }
