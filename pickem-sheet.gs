/**
 * Pick'em → Google Sheets
 *
 * Setup, once:
 *   1. Make a new Google Sheet. Extensions → Apps Script.
 *   2. Delete whatever's in Code.gs, paste this in, save.
 *   3. Deploy → New deployment → type "Web app".
 *        Execute as:      Me
 *        Who has access:  Anyone            <-- must be "Anyone", not "Anyone with a Google account"
 *   4. Copy the /exec URL it gives you into the pick'em page (DEFAULT_EP in
 *      pickem.html), or into Commissioner tools to point one browser at it.
 *
 * "Anyone" is what keeps your friends from having to sign in. The URL is
 * long and unguessable, but treat it as semi-public: anyone who gets it can
 * append a row. That's fine for a pick'em pool and not fine for anything else.
 *
 * If you ever edit this file, you have to Deploy → Manage deployments →
 * edit → New version, or the live URL keeps running the old code.
 *
 * Three tabs, made as needed:
 *   Picks  one row per player per card, newest submission wins
 *   Cards  the card each week's picks are supposed to be against
 *   Log    what went wrong, and what went right after going wrong
 */

var SHEET_NAME = 'Picks';
var CARDS_NAME = 'Cards';
var LOG_NAME   = 'Log';
var MAX_GAMES  = 12;
var LOG_LIMIT  = 300;   // rows handed back to the commissioner's Trouble panel

/* ---------------------------------------------------------------- write */

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return json_({ ok: false, error: 'busy' });
  }
  try {
    var data = JSON.parse(e.postData.contents);
    return String(data.action || 'picks') === 'register'
      ? registerCard_(data)
      : storePicks_(data);
  } catch (err) {
    log_('error', '', '', 'could not read the submission: ' + err, '', '');
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function storePicks_(data) {
  var week = String(data.week || '').trim();
  var name = String(data.name || '').trim();
  var fp   = String(data.fp || '').trim();
  var games = data.games || [];

  if (!name) {
    log_('error', week, '', 'a card arrived with no name on it', fp, '');
    return json_({ ok: false, error: 'no name' });
  }

  var ref = reference_(week);
  var flag = '';

  if (!ref.fp && fp) {
    // nobody has registered this week's card, so the first one in sets it
    setReference_(week, fp, games, 'first submission (' + name + ')');
    ref = reference_(week);
    log_('info', week, name, 'no card was registered, so this one became the reference', fp, fp);
  }

  if (ref.fp && fp && fp !== ref.fp) {
    flag = 'MISMATCH';
    log_('warn', week, name, 'picked a different card — ' + diff_(ref.games, games), fp, ref.fp);
  }

  if (data.retryOf) {
    log_('info', week, name, 'this card failed to send at ' + data.retryOf + ' and went through on a retry', fp, ref.fp);
  }

  var sh = picksSheet_();
  removeExisting_(sh, week, name);

  var picks = (data.picks || []).slice(0, MAX_GAMES).map(String);
  while (picks.length < MAX_GAMES) picks.push('');

  sh.appendRow([new Date(), week, name].concat(picks).concat([flag, fp]));
  return json_({ ok: true, flagged: !!flag, reference: ref.fp });
}

/** The commissioner says which card this week's picks are supposed to be against. */
function registerCard_(data) {
  var week = String(data.week || '').trim();
  var fp   = String(data.fp || '').trim();
  if (!week || !fp) return json_({ ok: false, error: 'need a card name and fingerprint' });

  var had = reference_(week);
  setReference_(week, fp, data.games || [], 'commissioner');
  log_('info', week, '', had.fp && had.fp !== fp
      ? 'the registered card was replaced'
      : 'card registered', fp, had.fp || fp);

  // anything already in the sheet under a different card is worth knowing about
  var stale = reflag_(week, fp);
  if (stale) log_('warn', week, '', stale + ' row(s) already in the sheet are against a different card', fp, had.fp || '');

  return json_({ ok: true, reflagged: stale });
}

/* ----------------------------------------------------------------- read */

function doGet(e) {
  try {
    var p = (e && e.parameter) || {};
    if (String(p.log || '') === '1') return json_({ ok: true, log: readLog_() });

    var week = String(p.week || '').trim().toLowerCase();
    var sh = picksSheet_();
    if (sh.getLastRow() < 2) return json_({ ok: true, entries: [], log: [] });

    var width = 3 + MAX_GAMES + 2;
    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, width).getValues();
    var entries = [];
    rows.forEach(function (r) {
      if (!String(r[2]).trim()) return;
      if (week && String(r[1]).trim().toLowerCase() !== week) return;
      entries.push({
        name: String(r[2]),
        picks: r.slice(3, 3 + MAX_GAMES)
                .filter(function (v) { return String(v).trim() !== ''; })
                .map(String),
        flag: String(r[3 + MAX_GAMES] || '')
      });
    });
    return json_({ ok: true, entries: entries });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function readLog_() {
  var sh = logSheet_();
  if (sh.getLastRow() < 2) return [];
  var n = Math.min(LOG_LIMIT, sh.getLastRow() - 1);
  var rows = sh.getRange(sh.getLastRow() - n + 1, 1, n, 6).getValues();
  return rows.map(function (r) {
    return {
      at: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
      level: String(r[1]), card: String(r[2]), who: String(r[3]),
      detail: String(r[4]), fp: String(r[5])
    };
  }).reverse();
}

/* ------------------------------------------------------------ the card */

function reference_(week) {
  var sh = cardsSheet_();
  if (sh.getLastRow() < 2) return { fp: '', games: [] };
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  for (var i = rows.length - 1; i >= 0; i--) {
    if (String(rows[i][0]).trim().toLowerCase() === week.toLowerCase()) {
      return { fp: String(rows[i][1]), games: String(rows[i][2] || '').split('|').filter(String) };
    }
  }
  return { fp: '', games: [] };
}

function setReference_(week, fp, games, source) {
  var sh = cardsSheet_();
  if (sh.getLastRow() >= 2) {
    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    for (var i = rows.length - 1; i >= 0; i--) {
      if (String(rows[i][0]).trim().toLowerCase() === week.toLowerCase()) sh.deleteRow(i + 2);
    }
  }
  sh.appendRow([week, fp, (games || []).join('|'), new Date(), source]);
}

/** Says what is actually different, rather than just that something is. */
function diff_(refGames, gotGames) {
  if (!refGames.length || !gotGames.length) return 'no game list to compare';
  if (refGames.length !== gotGames.length) {
    return 'their card has ' + gotGames.length + ' games, the registered one has ' + refGames.length;
  }
  var off = [];
  for (var i = 0; i < refGames.length; i++) {
    if (String(refGames[i]) !== String(gotGames[i])) {
      off.push('#' + (i + 1) + ' ' + gotGames[i] + ' instead of ' + refGames[i]);
    }
  }
  if (!off.length) return 'same games in the same order — the fingerprints should not differ';
  return off.length + ' of ' + refGames.length + ' differ: ' + off.slice(0, 3).join('; ') + (off.length > 3 ? '; …' : '');
}

/** Marks rows already in the sheet that were against some other card. */
function reflag_(week, fp) {
  var sh = picksSheet_();
  if (sh.getLastRow() < 2) return 0;
  var col = 3 + MAX_GAMES;                       // Flag
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, col + 2).getValues();
  var n = 0;
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][1]).trim().toLowerCase() !== week.toLowerCase()) continue;
    var rowFp = String(rows[i][col + 1] || '');
    var want = (rowFp && rowFp !== fp) ? 'MISMATCH' : '';
    if (String(rows[i][col] || '') !== want) {
      sh.getRange(i + 2, col + 1).setValue(want);
      if (want) n++;
    } else if (want) { n++; }
  }
  return n;
}

/* --------------------------------------------------------------- sheets */

/** A resubmitted card replaces the earlier one instead of stacking up. */
function removeExisting_(sh, week, name) {
  if (sh.getLastRow() < 2) return;
  var rows = sh.getRange(2, 2, sh.getLastRow() - 1, 2).getValues();
  for (var i = rows.length - 1; i >= 0; i--) {
    var sameWeek = String(rows[i][0]).trim().toLowerCase() === week.toLowerCase();
    var sameName = String(rows[i][1]).trim().toLowerCase() === name.toLowerCase();
    if (sameWeek && sameName) sh.deleteRow(i + 2);
  }
}

function picksSheet_() {
  var head = ['Submitted', 'Card', 'Player'];
  for (var i = 1; i <= MAX_GAMES; i++) head.push('Game ' + i);
  head.push('Flag', 'Card id');
  return sheet_(SHEET_NAME, head);
}
function cardsSheet_() { return sheet_(CARDS_NAME, ['Card', 'Card id', 'Games', 'Registered', 'By']); }
function logSheet_()   { return sheet_(LOG_NAME,   ['When', 'Level', 'Card', 'Player', 'What happened', 'Card id']); }

function sheet_(name, head) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(head);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, head.length).setFontWeight('bold');
  } else if (sh.getLastColumn() < head.length) {
    // an older sheet from before these columns existed
    var from = sh.getLastColumn() + 1;
    var add = head.slice(from - 1);
    sh.getRange(1, from, 1, add.length).setValues([add]).setFontWeight('bold');
  }
  return sh;
}

function log_(level, card, who, detail, fp, ref) {
  try {
    logSheet_().appendRow([new Date(), level, card, who,
      detail + (ref && fp && ref !== fp ? ' (registered card ' + ref + ')' : ''), fp]);
  } catch (err) { /* a log that breaks the submission is worse than no log */ }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
