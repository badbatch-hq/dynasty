/**
 * BAD BATCH HQ — PROPHECIES BACKEND (v2)
 * ----------------------------------------
 * A free Google Apps Script "Web App" that uses a Google Sheet as a database
 * for league predictions ("prophecies"), and auto-validates the structured
 * ones against the Sleeper API on a schedule.
 *
 * SETUP:
 * 1. Rename your EXISTING "Prophecies" tab to "Prophecies Archive" first —
 *    this script creates a fresh "Prophecies" tab for live tracking and
 *    won't touch/overwrite your old data.
 * 2. Extensions > Apps Script. Paste this file in as Code.gs.
 * 3. Fill in CONFIG below (league ID, a real admin key).
 * 4. Deploy > New deployment > Web app. Execute as "Me", access "Anyone".
 * 5. Run `getSheet` once manually from the editor to authorize + create headers.
 * 6. Run `importLegacyProphecies` once to pull your ~41 old predictions in
 *    from "Prophecies Archive" (safe to re-run — it skips rows it already imported).
 * 7. (Optional) Triggers > Add trigger > autoValidatePredictions > time-driven,
 *    daily, so pending prophecies get checked automatically.
 */

const CONFIG = {
  // The Sleeper league ID for the CURRENT season. Same one your dashboard uses.
  CURRENT_LEAGUE_ID: '1353201498326044672',
  CURRENT_SEASON: '2026',

  // Any string you want. Required to call /validate from outside (not needed
  // for normal use — the daily trigger calls autoValidatePredictions directly).
  ADMIN_KEY: 'kneeldowns-helped-me-win2021',

  SHEET_NAME: 'Prophecies',
  ARCHIVE_SHEET_NAME: 'Prophecies Archive',
  POLLS_SHEET_NAME: 'Polls',
  POLL_VOTES_SHEET_NAME: 'PollVotes',

  // Playoff bracket size — used to decide which roster_ids get a real placement
  // vs. get ranked by regular-season record for "finishes top N" predictions
  // where N is bigger than the playoff field.
  PLAYOFF_TEAMS: 6,
};

const HEADERS = [
  'id', 'username', 'dateSubmitted', 'season', 'week', 'type', 'category',
  'checkType', 'checkParams', 'predictionText', 'status', 'resolvedDate', 'resolvedNote',
];

// Anyone can post a poll through the /createpoll endpoint below (same
// open-submission spirit as Prophecies) — no admin key, no approval step.
// closesAt is optional: blank means the poll never auto-closes on a deadline.
// A poll can still close before then two other ways: submitPollVote() below
// auto-closes it the moment vote count reaches getLeagueSize() (everyone's
// voted), and `active` can always be hand-flipped to FALSE in the sheet for
// an early manual close — no in-app button for that by design.
const POLL_HEADERS = ['pollId', 'question', 'options', 'active', 'createdAt', 'closesAt', 'createdBy', 'anonymousVoting'];
const POLL_VOTE_HEADERS = ['pollId', 'voterName', 'option', 'timestamp'];

// ---------- HTTP entry points ----------

function doGet(e) {
  const action = (e.parameter.action || 'list').toLowerCase();

  if (action === 'list') {
    return jsonResponse({ ok: true, predictions: getAllPredictions() });
  }

  if (action === 'polls') {
    return jsonResponse({ ok: true, polls: getAllPolls(), votes: getAllPollVotes() });
  }

  if (action === 'validate') {
    requireAdmin(e);
    const summary = autoValidatePredictions();
    return jsonResponse({ ok: true, summary });
  }

  return jsonResponse({ ok: true, message: 'Bad Batch HQ Prophecies API is alive.' });
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse({ ok: false, error: 'Invalid JSON body' }, 400);
  }

  const action = (body.action || 'submit').toLowerCase();

  if (action === 'submit') {
    return jsonResponse(submitPrediction(body));
  }

  if (action === 'vote') {
    return jsonResponse(submitPollVote(body));
  }

  if (action === 'createpoll') {
    return jsonResponse(submitPoll(body));
  }

  return jsonResponse({ ok: false, error: 'Unknown action. Resolve freeform predictions directly in the sheet (status column).' }, 400);
}

function requireAdmin(e) {
  if (e.parameter.key !== CONFIG.ADMIN_KEY) {
    throw new Error('Unauthorized');
  }
}

function jsonResponse(obj) {
  const output = ContentService.createTextOutput(JSON.stringify(obj));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

// ---------- Sheet helpers ----------

// Shared by getSheet()/getPollsSheet()/getPollVotesSheet() below.
// getSheetByName() does an exact, CASE-SENSITIVE lookup, but Google Sheets
// enforces tab-name uniqueness case-INSENSITIVELY — so a tab that ended up
// named e.g. "pollvotes" or "Poll Votes" silently hides from our lookup,
// and insertSheet() then throws "a sheet with this name already exists"
// when it collides with that same name in a different case. Falling back to
// a case-insensitive scan (both before AND after a failed insert, in case
// two requests raced to create the same sheet) avoids failing the whole
// request over a naming quirk.
function findSheetCaseInsensitive(ss, name) {
  return ss.getSheets().find(s => s.getName().toLowerCase() === name.toLowerCase()) || null;
}

function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name) || findSheetCaseInsensitive(ss, name);
  if (!sheet) {
    try {
      sheet = ss.insertSheet(name);
    } catch (e) {
      sheet = findSheetCaseInsensitive(ss, name);
      if (!sheet) throw e; // genuinely nothing there — surface the real error
    }
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getSheet() {
  return getOrCreateSheet(CONFIG.SHEET_NAME, HEADERS);
}

function getAllPredictions() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  return values.map(rowToObject).filter(p => p.id); // skip blank rows
}

function rowToObject(row) {
  const obj = {};
  HEADERS.forEach((h, i) => { obj[h] = row[i]; });
  try { obj.checkParams = obj.checkParams ? JSON.parse(obj.checkParams) : {}; } catch (e) { obj.checkParams = {}; }
  return obj;
}

function findRowIndexById(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) return i + 2; // sheet row (1-indexed, +1 for header)
  }
  return -1;
}

// ---------- Submit ----------
// Resolving is intentionally NOT exposed over the web (no admin key sitting in
// public JS). Auto-validation handles structured predictions; freeform ones
// get resolved by typing "hit"/"miss" (+ optional note) straight into the sheet.

function submitPrediction(body) {
  const required = ['username', 'season', 'type', 'predictionText'];
  for (const field of required) {
    if (!body[field]) return { ok: false, error: `Missing field: ${field}` };
  }

  const sheet = getSheet();
  const id = Utilities.getUuid();
  const row = [
    id,
    body.username,
    new Date().toISOString(),
    body.season,
    body.week || 'season',
    body.type,                    // 'structured' | 'freeform'
    body.category || '',
    body.checkType || 'none',     // standings_rank | team_wins | matchup_winner | player_stat | trade_happens | none
    JSON.stringify(body.checkParams || {}),
    body.predictionText,
    'pending',
    '',
    '',
  ];
  sheet.appendRow(row);
  return { ok: true, id };
}

// ---------- Polls ----------
// Votes are upserted one row per (pollId, voterName) in "PollVotes" — changing
// your vote overwrites your existing row instead of piling up duplicates.
// Tallying happens client-side in index.html from the raw vote rows, same as
// how the Prophecies hit-rate stats are computed.

function getPollsSheet() {
  return getOrCreateSheet(CONFIG.POLLS_SHEET_NAME, POLL_HEADERS);
}

function getPollVotesSheet() {
  return getOrCreateSheet(CONFIG.POLL_VOTES_SHEET_NAME, POLL_VOTE_HEADERS);
}

// Polls sheet columns: pollId, question, options (comma-separated), active
// (TRUE/FALSE), createdAt, closesAt (optional ISO timestamp), createdBy.
// Rows normally come in through submitPoll() below, but hand-editing a row
// (e.g. to flip `active` to FALSE, or edit a typo) works fine too.
function getAllPolls() {
  const sheet = getPollsSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, POLL_HEADERS.length).getValues();
  return values
    .map(row => {
      const obj = {};
      POLL_HEADERS.forEach((h, i) => { obj[h] = row[i]; });
      obj.options = String(obj.options || '').split(',').map(o => o.trim()).filter(Boolean);
      obj.active = obj.active === true || String(obj.active).toLowerCase() === 'true';
      obj.closesAt = obj.closesAt ? new Date(obj.closesAt).toISOString() : '';
      return obj;
    })
    .filter(p => p.pollId);
}

// Open to anyone — no admin key, mirroring how Prophecies submissions work.
// options: array of option strings (need at least 2). durationDays: number
// of days until the poll auto-closes, or '' / omitted for no deadline.
function submitPoll(body) {
  const required = ['question', 'options'];
  for (const field of required) {
    if (!body[field]) return { ok: false, error: `Missing field: ${field}` };
  }

  const options = (Array.isArray(body.options) ? body.options : String(body.options).split(','))
    .map(o => String(o).trim()).filter(Boolean);
  if (options.length < 2) return { ok: false, error: 'Need at least 2 options' };

  const sheet = getPollsSheet();
  const id = Utilities.getUuid();
  const now = new Date();

  let closesAt = '';
  const days = Number(body.durationDays);
  if (days > 0) {
    closesAt = new Date(now.getTime() + days * 24 * 60 * 60 * 1000).toISOString();
  }

  sheet.appendRow([
    id,
    body.question,
    options.join(','),
    true, // active
    now.toISOString(),
    closesAt,
    body.createdBy || '',
    !!body.anonymousVoting,
  ]);
  return { ok: true, id };
}

function getAllPollVotes() {
  const sheet = getPollVotesSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, POLL_VOTE_HEADERS.length).getValues();
  return values
    .map(row => {
      const obj = {};
      POLL_VOTE_HEADERS.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    })
    .filter(v => v.pollId && v.voterName);
}

function findPollVoteRowIndex(sheet, pollId, voterName) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const rows = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); // pollId, voterName
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === pollId && rows[i][1] === voterName) return i + 2;
  }
  return -1;
}

function findPollRowIndex(sheet, pollId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues(); // pollId only
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === pollId) return i + 2;
  }
  return -1;
}

// Number of teams in the current-season league — the "everyone's voted"
// threshold for auto-closing a poll at full participation. Pulled live from
// Sleeper rather than a hardcoded 12, so it stays correct if the league ever
// changes size.
function getLeagueSize() {
  const rosters = fetchJson(`https://api.sleeper.app/v1/league/${CONFIG.CURRENT_LEAGUE_ID}/rosters`);
  return rosters.length;
}

function submitPollVote(body) {
  const required = ['pollId', 'voterName', 'option'];
  for (const field of required) {
    if (!body[field]) return { ok: false, error: `Missing field: ${field}` };
  }

  // Real gatekeeper on closed polls — a browser with a stale poll list could
  // otherwise still fire off a vote after the deadline (or after Joe flips
  // `active` off by hand), so re-check against the sheet here rather than
  // trusting whatever the client last saw.
  const poll = getAllPolls().find(p => p.pollId === body.pollId);
  if (!poll) return { ok: false, error: 'Poll not found' };
  if (!poll.active) return { ok: false, error: 'This poll is closed' };
  if (poll.closesAt && new Date(poll.closesAt).getTime() < Date.now()) {
    return { ok: false, error: 'This poll is closed' };
  }

  const sheet = getPollVotesSheet();
  const rowIndex = findPollVoteRowIndex(sheet, body.pollId, body.voterName);
  const row = [body.pollId, body.voterName, body.option, new Date().toISOString()];
  const isNewVoter = rowIndex === -1;

  if (isNewVoter) {
    sheet.appendRow(row);
  } else {
    sheet.getRange(rowIndex, 1, 1, POLL_VOTE_HEADERS.length).setValues([row]);
  }

  // Auto-close once every manager has voted. Only a brand-new voter can push
  // the vote count up (changing an existing vote overwrites a row rather than
  // adding one), so this only ever needs checking on that path. This counts
  // raw vote rows, not confirmed real managers — an anonymous-voting poll's
  // rows are per-browser ids rather than real names, so "12 votes" and "12
  // managers" aren't strictly guaranteed to be the same 12 people there, same
  // honor-system caveat as the rest of Polls/Prophecies. Wrapped in try/catch
  // so a hiccup fetching the roster count (a transient Sleeper API error)
  // never breaks the vote itself — worst case, the poll just waits for
  // closesAt or a manual close instead of closing this instant.
  if (isNewVoter) {
    try {
      const voteCount = getAllPollVotes().filter(v => v.pollId === body.pollId).length;
      if (voteCount >= getLeagueSize()) {
        const pollsSheet = getPollsSheet();
        const pollRow = findPollRowIndex(pollsSheet, body.pollId);
        if (pollRow !== -1) {
          pollsSheet.getRange(pollRow, POLL_HEADERS.indexOf('active') + 1).setValue(false);
        }
      }
    } catch (e) {
      // Non-fatal — see comment above. The vote itself already saved.
    }
  }

  return { ok: true };
}

// Optional housekeeping: flips `active` to FALSE in the sheet for any poll
// whose closesAt deadline has passed. Not required for correctness —
// submitPollVote already rejects votes on expired polls regardless of this
// column — but keeps the sheet itself honest at a glance instead of showing
// TRUE forever. Wire it up the same way as the Prophecies auto-validation:
// Triggers > Add trigger > autoClosePolls > time-driven (hourly works well).
function autoClosePolls() {
  const sheet = getPollsSheet();
  const polls = getAllPolls();
  const activeCol = POLL_HEADERS.indexOf('active') + 1;
  const now = Date.now();
  let closed = 0;

  polls.forEach((p, i) => {
    if (p.active && p.closesAt && new Date(p.closesAt).getTime() < now) {
      sheet.getRange(i + 2, activeCol).setValue(false);
      closed++;
    }
  });

  return { closed };
}

// ---------- One-time legacy import ----------
// Pulls rows from "Prophecies Archive" (your old Who?/What?/Did It Happen?/Comment
// sheet) into the new structured format. Safe to re-run: skips rows whose
// predictionText it's already imported (tagged via the resolvedNote prefix).
function importLegacyProphecies() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const archive = ss.getSheetByName(CONFIG.ARCHIVE_SHEET_NAME);
  if (!archive) {
    throw new Error(`No sheet named "${CONFIG.ARCHIVE_SHEET_NAME}" found. Rename your old Prophecies tab to that first.`);
  }

  const sheet = getSheet();
  const existing = getAllPredictions();
  const alreadyImported = new Set(
    existing.filter(p => p.category === 'legacy-import').map(p => p.predictionText)
  );

  const lastRow = archive.getLastRow();
  if (lastRow < 2) return { imported: 0 };

  // Columns A-D: Who?, What?, Did It Happen?, Comment
  const values = archive.getRange(2, 1, lastRow - 1, 4).getValues();
  let imported = 0;

  values.forEach(([who, what, didItHappen, comment]) => {
    if (!what) return; // skip blank rows
    const predictionText = String(what).trim();
    if (alreadyImported.has(predictionText)) return; // already imported, skip

    const answer = String(didItHappen || '').trim().toLowerCase();
    const status = answer === 'yes' ? 'hit' : answer === 'no' ? 'miss' : 'pending';

    const id = Utilities.getUuid();
    const row = [
      id,
      who || 'Unknown',
      '', // dateSubmitted unknown — left blank rather than guessed
      'legacy', // season unknown, tagged explicitly rather than guessed
      'season',
      'freeform',
      'legacy-import',
      'none',
      '{}',
      predictionText,
      status,
      status !== 'pending' ? 'imported' : '',
      comment ? String(comment).trim() : '',
    ];
    sheet.appendRow(row);
    imported++;
  });

  return { imported };
}

// ---------- Auto-validation against Sleeper ----------

function autoValidatePredictions() {
  const sheet = getSheet();
  const predictions = getAllPredictions();
  const summary = { checked: 0, resolved: 0, hit: 0, miss: 0, errors: [] };

  predictions.forEach(p => {
    if (p.type !== 'structured' || p.status !== 'pending' || !p.checkType || p.checkType === 'none') return;

    summary.checked++;
    try {
      const result = evaluateCheck(p.checkType, p.checkParams, p.season);
      if (result.resolved) {
        const rowIndex = findRowIndexById(sheet, p.id);
        sheet.getRange(rowIndex, 11).setValue(result.status);
        sheet.getRange(rowIndex, 12).setValue(new Date().toISOString());
        sheet.getRange(rowIndex, 13).setValue(result.note || '');
        summary.resolved++;
        if (result.status === 'hit') summary.hit++;
        if (result.status === 'miss') summary.miss++;
      }
    } catch (err) {
      summary.errors.push(`${p.id}: ${err.message}`);
    }
  });

  return summary;
}

function evaluateCheck(checkType, params, season) {
  switch (checkType) {
    case 'standings_rank': return checkStandingsRank(params, season);
    case 'team_wins': return checkTeamWins(params, season);
    case 'matchup_winner': return checkMatchupWinner(params, season);
    case 'player_stat': return checkPlayerStat(params, season);
    case 'trade_happens': return checkTradeHappens(params, season);
    default: return { resolved: false };
  }
}

// Resolves "at least/at most/exactly" predictions against a value that can only
// move in one direction over a season (wins, counting stats). Lets an early
// HIT land the moment it's mathematically locked in for >=/> , an early MISS
// land the moment it's mathematically locked in for <=/< , and otherwise waits
// for the season to actually be over (covers '==' always, and the "still could
// go either way" side of >=/<=).
function resolveMonotonic(currentValue, operator, target, seasonComplete) {
  const hitNow = compare(currentValue, operator, target);

  if ((operator === '>=' || operator === '>') && hitNow) {
    return { resolved: true, status: 'hit' };
  }
  if ((operator === '<=' || operator === '<') && !hitNow) {
    return { resolved: true, status: 'miss' };
  }
  if (!seasonComplete) return { resolved: false };
  return { resolved: true, status: hitNow ? 'hit' : 'miss' };
}

// params: { rosterId, operator ('<='|'=='|'>='), rank }
// Uses the REAL playoff bracket for placements 1..PLAYOFF_TEAMS, and
// regular-season record (same tiebreak as the main dashboard) for the rest.
function checkStandingsRank(params, season) {
  const leagueId = getLeagueIdForSeason(season);
  if (!leagueId) return { resolved: false };

  const placements = getPlayoffPlacements(leagueId);
  const hasPlayoffData = Object.keys(placements).length > 0;
  if (!hasPlayoffData) return { resolved: false }; // season's playoffs haven't concluded yet

  const rosters = fetchJson(`https://api.sleeper.app/v1/league/${leagueId}/rosters`);
  const nonPlayoff = rosters
    .filter(r => !(r.roster_id in placements))
    .sort((a, b) => {
      const wa = a.settings.wins || 0, wb = b.settings.wins || 0;
      if (wb !== wa) return wb - wa;
      return (b.settings.fpts || 0) - (a.settings.fpts || 0);
    });

  const finalRank = {};
  Object.entries(placements).forEach(([rid, place]) => { finalRank[rid] = place; });
  nonPlayoff.forEach((r, i) => { finalRank[r.roster_id] = Object.keys(placements).length + i + 1; });

  const rank = finalRank[params.rosterId];
  if (rank == null) return { resolved: false };

  const hit = compare(rank, params.operator, params.rank);
  return { resolved: true, status: hit ? 'hit' : 'miss', note: `Finished rank ${rank}` };
}

// Decodes Sleeper's winners_bracket into { roster_id: place }, same logic as
// the main dashboard (index.html's getPlayoffPlacements).
function getPlayoffPlacements(leagueId) {
  const bracket = fetchJson(`https://api.sleeper.app/v1/league/${leagueId}/winners_bracket`);
  const placements = {};
  (bracket || []).forEach(game => {
    if (game.p == null) return;
    if (game.w != null) placements[game.w] = game.p;
    if (game.l != null) placements[game.l] = game.p + 1;
  });
  return placements;
}

// params: { rosterId, operator, value }
function checkTeamWins(params, season) {
  const leagueId = getLeagueIdForSeason(season);
  if (!leagueId) return { resolved: false };

  const rosters = fetchJson(`https://api.sleeper.app/v1/league/${leagueId}/rosters`);
  const roster = rosters.find(r => String(r.roster_id) === String(params.rosterId));
  if (!roster) return { resolved: false };

  const wins = roster.settings.wins || 0;
  const result = resolveMonotonic(wins, params.operator, params.value, isSeasonComplete(leagueId));
  if (result.resolved) result.note = `Currently ${wins} wins`;
  return result;
}

// params: { week, rosterId (predicted winner) }
function checkMatchupWinner(params, season) {
  const leagueId = getLeagueIdForSeason(season);
  if (!leagueId) return { resolved: false };

  const matchups = fetchJson(`https://api.sleeper.app/v1/league/${leagueId}/matchups/${params.week}`);
  const mine = matchups.find(m => String(m.roster_id) === String(params.rosterId));
  if (!mine || !mine.matchup_id) return { resolved: false };

  const opponents = matchups.filter(m => m.matchup_id === mine.matchup_id && String(m.roster_id) !== String(params.rosterId));
  if (opponents.length === 0) return { resolved: false };
  const opp = opponents[0];

  if ((mine.points || 0) === 0 && (opp.points || 0) === 0) return { resolved: false }; // not played yet

  const hit = (mine.points || 0) > (opp.points || 0);
  return { resolved: true, status: hit ? 'hit' : 'miss', note: `${mine.points} - ${opp.points}` };
}

// params: { week ('season' for full-season totals), playerId, statKey, operator, value }
function checkPlayerStat(params, season) {
  const leagueId = getLeagueIdForSeason(season);
  const url = params.week === 'season'
    ? `https://api.sleeper.app/v1/stats/nfl/regular/${season}`
    : `https://api.sleeper.app/v1/stats/nfl/regular/${season}/${params.week}`;
  const stats = fetchJson(url);
  const playerStats = stats[params.playerId];
  if (!playerStats || playerStats[params.statKey] === undefined) return { resolved: false };

  const value = playerStats[params.statKey];

  if (params.week === 'season') {
    return { ...resolveMonotonic(value, params.operator, params.value, leagueId ? isSeasonComplete(leagueId) : false), note: `${params.statKey}: ${value}` };
  }

  // Weekly checks: only resolve once that week is actually over.
  const state = fetchJson('https://api.sleeper.app/v1/state/nfl');
  const weekOver = Number(state.week) > Number(params.week) || state.season_type !== 'regular';
  if (!weekOver) return { resolved: false };

  const hit = compare(value, params.operator, params.value);
  return { resolved: true, status: hit ? 'hit' : 'miss', note: `${params.statKey}: ${value}` };
}

// params: { rosterIdA, rosterIdB }
function checkTradeHappens(params, season) {
  const leagueId = getLeagueIdForSeason(season);
  if (!leagueId) return { resolved: false };
  const state = fetchJson('https://api.sleeper.app/v1/state/nfl');
  const maxWeek = Math.min(Number(state.week) || 18, 18);

  for (let week = 1; week <= maxWeek; week++) {
    let txns;
    try {
      txns = fetchJson(`https://api.sleeper.app/v1/league/${leagueId}/transactions/${week}`);
    } catch (e) { continue; }
    const found = txns.some(t =>
      t.type === 'trade' &&
      t.status === 'complete' &&
      t.roster_ids.map(String).includes(String(params.rosterIdA)) &&
      t.roster_ids.map(String).includes(String(params.rosterIdB))
    );
    if (found) return { resolved: true, status: 'hit', note: `Trade found in week ${week}` };
  }

  if (isSeasonComplete(leagueId)) {
    return { resolved: true, status: 'miss', note: 'Season ended, no trade occurred' };
  }
  return { resolved: false };
}

// ---------- Sleeper utilities ----------

function fetchJson(url) {
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) {
    throw new Error(`Sleeper API error ${resp.getResponseCode()} for ${url}`);
  }
  return JSON.parse(resp.getContentText());
}

// Walks previous_league_id backward from CONFIG.CURRENT_LEAGUE_ID to find the
// league object for a given season. Caches the season->leagueId map.
function getLeagueIdForSeason(season) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('season_league_map');
  let map = cached ? JSON.parse(cached) : null;

  if (!map || !map[season]) {
    map = {};
    let leagueId = CONFIG.CURRENT_LEAGUE_ID;
    let guard = 0;
    while (leagueId && guard < 20) {
      const league = fetchJson(`https://api.sleeper.app/v1/league/${leagueId}`);
      map[league.season] = leagueId;
      leagueId = league.previous_league_id;
      guard++;
    }
    cache.put('season_league_map', JSON.stringify(map), 21600); // 6 hours
  }
  return map[season];
}

function isSeasonComplete(leagueId) {
  const league = fetchJson(`https://api.sleeper.app/v1/league/${leagueId}`);
  const state = fetchJson('https://api.sleeper.app/v1/state/nfl');
  if (String(state.league_season) !== String(league.season)) return true;
  const playoffStart = league.settings.playoff_week_start || 15;
  return Number(state.week) > playoffStart + 3;
}

function compare(actual, operator, target) {
  actual = Number(actual);
  target = Number(target);
  switch (operator) {
    case '<=': return actual <= target;
    case '>=': return actual >= target;
    case '<': return actual < target;
    case '>': return actual > target;
    case '==': return actual === target;
    default: return false;
  }
}
