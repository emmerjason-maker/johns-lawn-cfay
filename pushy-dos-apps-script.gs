/**
 * Pushy Dos Scoreboard: Google Apps Script web app
 *
 * Lets pushy-dos.html read the scores live and write back to the Google Sheet:
 * record a win, undo the last game, and set a player's avatar.
 *
 * SETUP (once):
 * 1. Open the "Pushy Dos Scoreboard" Google Sheet.
 * 2. Extensions → Apps Script. Delete what's there and paste this whole file. Save.
 * 3. Deploy → New deployment → gear icon → Web app.
 *      Execute as: Me
 *      Who has access: Anyone
 *    Click Deploy, approve access, and copy the Web app URL (ends in /exec).
 * 4. Paste that URL into CONFIG.scriptUrl in pushy-dos.html.
 *
 * If you change this code later: Deploy → Manage deployments → edit (pencil) →
 * Version: New version → Deploy. The URL stays the same.
 */

const STANDINGS = "Standings";
const GAMES = "Games";
const MAX_PLAYERS = 10; // Standings rows 2 to 11

function doGet() {
  return json_(readAll_());
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const players = readPlayers_();
    const names = players.map(p => p.name);
    const games = ss_().getSheetByName(GAMES);

    if (req.action === "addWin") {
      if (names.indexOf(req.winner) < 0) return json_({ error: "Unknown player." });
      const row = lastGameRow_(games) + 1;
      const today = Utilities.formatDate(new Date(), ss_().getSpreadsheetTimeZone(), "yyyy-MM-dd");
      games.getRange(row, 1, 1, 2).setValues([[today, req.winner]]);
      games.getRange(row, 1).setNumberFormat("yyyy-mm-dd");
    } else if (req.action === "undo") {
      const row = lastGameRow_(games);
      if (row < 2) return json_({ error: "There are no games to undo." });
      games.getRange(row, 1, 1, games.getLastColumn()).clearContent();
    } else if (req.action === "setAvatar") {
      const p = players.filter(x => x.name === req.player)[0];
      if (!p) return json_({ error: "Unknown player." });
      const col = headerIndex_(STANDINGS, "avatar");
      if (col < 0) return json_({ error: "The Standings tab has no Avatar column." });
      ss_().getSheetByName(STANDINGS).getRange(p.row, col + 1).setValue(cleanAvatar_(req.avatar));
    } else {
      return json_({ error: "Unknown action." });
    }

    SpreadsheetApp.flush();
    return json_(readAll_());
  } catch (err) {
    return json_({ error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function readAll_() {
  return {
    players: readPlayers_().map(p => ({ name: p.name, wins: p.wins, status: p.status, avatar: p.avatar })),
    games: readGames_()
  };
}

function readPlayers_() {
  const sh = ss_().getSheetByName(STANDINGS);
  const vals = sh.getRange(1, 1, MAX_PLAYERS + 1, sh.getLastColumn()).getDisplayValues();
  const head = vals[0].map(h => String(h).trim().toLowerCase());
  const iName = head.indexOf("player"), iWins = head.indexOf("wins");
  const iStatus = head.indexOf("status"), iAvatar = head.indexOf("avatar");
  return vals.slice(1).map((r, i) => ({
    row: i + 2,
    name: String(r[iName]).trim(),
    wins: Number(r[iWins]) || 0,
    status: iStatus >= 0 ? String(r[iStatus]).trim() : "",
    avatar: iAvatar >= 0 ? String(r[iAvatar]).trim() : ""
  })).filter(p => p.name);
}

function readGames_() {
  const sh = ss_().getSheetByName(GAMES);
  const last = lastGameRow_(sh);
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 2).getDisplayValues()
    .filter(r => String(r[1]).trim())
    .map(r => ({ date: r[0], winner: String(r[1]).trim() }));
}

// Last row in the Games tab with a winner filled in (1 = header only)
function lastGameRow_(sh) {
  const v = sh.getRange(1, 2, sh.getMaxRows(), 1).getValues();
  for (let i = v.length - 1; i >= 1; i--) if (String(v[i][0]).trim()) return i + 1;
  return 1;
}

function headerIndex_(tab, name) {
  const sh = ss_().getSheetByName(tab);
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0]
    .map(h => String(h).trim().toLowerCase()).indexOf(name);
}

// Allowed avatars: small photo (data URL), https image link, or a short emoji/text
function cleanAvatar_(a) {
  a = String(a || "").trim();
  if (!a) return "";
  if (/^data:image\/(jpeg|png|webp);base64,[A-Za-z0-9+/=]+$/.test(a) && a.length <= 45000) return a;
  if (/^https:\/\/\S+$/.test(a) && a.length <= 1000) return a;
  if (a.length <= 16 && !/^[=+\-@]/.test(a) && !/[<>]/.test(a)) return a;
  throw new Error("That avatar isn't allowed. Use a photo, an image link, or an emoji.");
}

function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
