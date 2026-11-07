/**
 * Pushy Dos Scoreboard: Google Apps Script web app
 *
 * Lets pushy-dos.html read the scores live and write back to the Google Sheet:
 * record a win, undo the last game, set a player's avatar, check in
 * who's playing tonight, add a player, and mark a player MIA or back.
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
const GAMES = "Games";        // A: Date (game night), B: Winner, C: Recorded (time entered)
const NIGHTS = "Nights";      // A: Night, B: Players (comma-separated). Created automatically.
const MAX_PLAYERS = 10;       // Standings rows 2 to 11
const NIGHT_CUTOFF_HOURS = 4; // a game night runs until 4 a.m.
const UNDO_HOURS = 3;         // undo only works this long after a win

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
      const now = new Date();
      games.getRange(row, 3).clearDataValidations();
      games.getRange(row, 1, 1, 3).setValues([[nightKey_(now), req.winner, now]]);
      games.getRange(row, 1).setNumberFormat("yyyy-mm-dd");
      games.getRange(row, 3).setNumberFormat("yyyy-mm-dd hh:mm");
    } else if (req.action === "undo") {
      const row = lastGameRow_(games);
      if (row < 2) return json_({ error: "There are no games to undo." });
      const t = games.getRange(row, 3).getValue();
      if (!(t instanceof Date) || Date.now() - t.getTime() > UNDO_HOURS * 3600 * 1000) {
        return json_({ error: "Undo only works for " + UNDO_HOURS + " hours after a win. To fix an older game, edit the Games tab in the sheet." });
      }
      games.getRange(row, 1, 1, 3).clearContent();
    } else if (req.action === "setAvatar") {
      const p = players.filter(x => x.name === req.player)[0];
      if (!p) return json_({ error: "Unknown player." });
      const col = headerIndex_(STANDINGS, "avatar");
      if (col < 0) return json_({ error: "The Standings tab has no Avatar column." });
      ss_().getSheetByName(STANDINGS).getRange(p.row, col + 1).setValue(cleanAvatar_(req.avatar));
    } else if (req.action === "addPlayer") {
      const name = cleanName_(req.name);
      if (names.some(n => n.toLowerCase() === name.toLowerCase())) return json_({ error: name + " is already on the scoreboard." });
      const sh = ss_().getSheetByName(STANDINGS);
      const col = sh.getRange(2, 1, MAX_PLAYERS, 1).getDisplayValues();
      let row = -1;
      for (let i = 0; i < col.length; i++) if (!String(col[i][0]).trim()) { row = i + 2; break; }
      if (row < 0) return json_({ error: "The scoreboard is full (" + MAX_PLAYERS + " players)." });
      sh.getRange(row, 1).setValue(name);
    } else if (req.action === "setStatus") {
      const p = players.filter(x => x.name === req.player)[0];
      if (!p) return json_({ error: "Unknown player." });
      const col = headerIndex_(STANDINGS, "status");
      if (col < 0) return json_({ error: "The Standings tab has no Status column." });
      const status = String(req.status || "").trim().toUpperCase() === "MIA" ? "MIA" : "";
      ss_().getSheetByName(STANDINGS).getRange(p.row, col + 1).setValue(status);
    } else if (req.action === "checkIn") {
      const list = (Array.isArray(req.players) ? req.players : []).filter(n => names.indexOf(n) >= 0);
      const sh = nightsSheet_();
      const key = nightKey_(new Date());
      const last = sh.getLastRow();
      let row = -1;
      if (last >= 2) {
        const keys = sh.getRange(2, 1, last - 1, 1).getDisplayValues();
        for (let i = 0; i < keys.length; i++) if (String(keys[i][0]).trim() === key) row = i + 2;
      }
      if (row < 0) row = last + 1;
      sh.getRange(row, 1).setNumberFormat("@");
      sh.getRange(row, 1, 1, 2).setValues([[key, list.join(", ")]]);
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
    games: readGames_(),
    nights: readNights_()
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
  const disp = sh.getRange(2, 1, last - 1, 2).getDisplayValues();
  const raw = sh.getRange(2, 3, last - 1, 1).getValues();
  const out = [];
  for (let i = 0; i < disp.length; i++) {
    const winner = String(disp[i][1]).trim();
    if (!winner) continue;
    const t = raw[i][0];
    out.push({ date: String(disp[i][0]).trim(), winner: winner, ts: t instanceof Date ? t.getTime() : null });
  }
  return out;
}

function readNights_() {
  const sh = ss_().getSheetByName(NIGHTS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues()
    .filter(r => String(r[0]).trim())
    .map(r => ({
      night: String(r[0]).trim(),
      players: String(r[1]).split(",").map(s => s.trim()).filter(Boolean)
    }));
}

function nightsSheet_() {
  let sh = ss_().getSheetByName(NIGHTS);
  if (!sh) {
    sh = ss_().insertSheet(NIGHTS);
    sh.getRange(1, 1, 1, 2).setValues([["Night", "Players"]]);
    sh.setFrozenRows(1);
  }
  return sh;
}

// The game night a moment belongs to: anything before 4 a.m. counts as the night before
function nightKey_(d) {
  const shifted = new Date(d.getTime() - NIGHT_CUTOFF_HOURS * 3600 * 1000);
  return Utilities.formatDate(shifted, ss_().getSpreadsheetTimeZone(), "yyyy-MM-dd");
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

// Player names: 1 to 20 characters, no formulas or markup
function cleanName_(n) {
  n = String(n || "").trim().replace(/\s+/g, " ");
  if (!n || n.length > 20 || /^[=+\-@]/.test(n) || /[<>,]/.test(n)) {
    throw new Error("Use a name of 1 to 20 letters (no commas).");
  }
  return n;
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
