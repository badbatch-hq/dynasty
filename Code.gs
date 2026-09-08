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
  POLL_PARTICIPANTS_SHEET_NAME: 'PollParticipants',
  COMMISSIONER_CALLOUTS_SHEET_NAME: 'CommissionerCallouts',
  WEEKLY_REPORTS_SHEET_NAME: 'WeeklyReports',

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
// A real name, no option — deliberately separate from PollVotes so that even
// reading the raw sheet can't match a name to a choice on an anonymous poll.
// Only written for anonymous-voting polls (a named poll's real votes already
// carry a real name directly); see submitPollVote() below. Sheet-only —
// nothing in index.html ever fetches or renders this data.
//
// Deliberately NO timestamp column, and rows are kept sorted alphabetically
// by name rather than in the order people voted (see submitPollVote()) — an
// insertion-order or timestamp field here would let someone line this sheet
// up against PollVotes' own timestamps/row order and infer which anonymous
// vote belongs to which real name purely from *when* each was recorded, even
// though the two sheets share no other key. Sorting/dropping the timestamp
// closes that side channel.
const POLL_PARTICIPANT_HEADERS = ['pollId', 'voterName'];

// ---------- HTTP entry points ----------

function doGet(e) {
  const action = (e.parameter.action || 'list').toLowerCase();

  if (action === 'list') {
    return jsonResponse({ ok: true, predictions: getAllPredictions() });
  }

  if (action === 'polls') {
    return jsonResponse({ ok: true, polls: getAllPolls(), votes: getAllPollVotes() });
  }

  // League News: commissioner callouts + weekly reports. ?preview=1 also
  // returns draft rows (both callouts and reports), rendered by index.html
  // through the exact same code path as the published view — true WYSIWYG
  // for Joe to check a row before flipping it to `published` in the sheet.
  if (action === 'news') {
    const preview = e.parameter.preview === '1';
    return jsonResponse({ ok: true, callouts: getAllCallouts(preview), reports: getAllWeeklyReports(preview) });
  }

  if (action === 'validate') {
    requireAdmin(e);
    const summary = autoValidatePredictions();
    return jsonResponse({ ok: true, summary });
  }

  // Manual/backfill trigger for a specific week's report, without waiting on
  // the scheduled trigger: ?action=generatereport&week=N&key=ADMIN_KEY
  if (action === 'generatereport') {
    return jsonResponse(adminGenerateWeeklyReport(e));
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
//
// "PollParticipants" is a separate, sheet-only record of real names that have
// voted on an anonymous-voting poll — no option column, no shared key with
// the actual (anonymous-id-keyed) PollVotes row for that same vote. It exists
// purely so Joe can tell who's voted on an anonymous poll without index.html
// (or Joe reading the raw sheet) ever being able to tie a name to a choice.

function getPollsSheet() {
  return getOrCreateSheet(CONFIG.POLLS_SHEET_NAME, POLL_HEADERS);
}

function getPollVotesSheet() {
  return getOrCreateSheet(CONFIG.POLL_VOTES_SHEET_NAME, POLL_VOTE_HEADERS);
}

function getPollParticipantsSheet() {
  return getOrCreateSheet(CONFIG.POLL_PARTICIPANTS_SHEET_NAME, POLL_PARTICIPANT_HEADERS);
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

function getAllPollParticipants() {
  const sheet = getPollParticipantsSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, POLL_PARTICIPANT_HEADERS.length).getValues();
  return values
    .map(row => {
      const obj = {};
      POLL_PARTICIPANT_HEADERS.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    })
    .filter(p => p.pollId && p.voterName);
}

function findParticipantRowIndex(sheet, pollId, voterName) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const rows = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); // pollId, voterName
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === pollId && rows[i][1] === voterName) return i + 2;
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

  // For an anonymous-voting poll, also record (separately, in PollParticipants)
  // that this real name has participated — never in the same row as their
  // choice, never joinable back to it. `voterName` above stays the random
  // per-browser id for the actual vote; `participantName` here is the real
  // "Voting as" name index.html now also collects for an anonymous poll
  // specifically so this can exist. Sheet-only bookkeeping for Joe — nothing
  // in the app ever reads this back. Wrapped so a hiccup here can never
  // block the vote itself, same defensive pattern as the auto-close check
  // below.
  //
  // No timestamp is stored, and a newly-added row is immediately re-sorted
  // alphabetically by name rather than left in the order votes arrived —
  // both writes (this one and the PollVotes upsert above) happen in the same
  // request, so if this sheet were left in append order, its row-for-row
  // timing would line up with PollVotes' own append order/timestamps closely
  // enough to "math out" which real name cast which anonymous vote, even
  // though the two rows share no key. Sorting by name (not append order)
  // severs that.
  let isNewParticipant = false;
  if (poll.anonymousVoting && body.participantName) {
    try {
      const pSheet = getPollParticipantsSheet();
      const pRowIndex = findParticipantRowIndex(pSheet, body.pollId, body.participantName);
      isNewParticipant = pRowIndex === -1;
      const pRow = [body.pollId, body.participantName];
      if (isNewParticipant) {
        pSheet.appendRow(pRow);
        const pLastRow = pSheet.getLastRow();
        if (pLastRow > 2) {
          pSheet.getRange(2, 1, pLastRow - 1, POLL_PARTICIPANT_HEADERS.length)
            .sort([{ column: 1, ascending: true }, { column: 2, ascending: true }]);
        }
      } else {
        pSheet.getRange(pRowIndex, 1, 1, POLL_PARTICIPANT_HEADERS.length).setValues([pRow]);
      }
    } catch (e) {
      // Non-fatal — the vote itself already saved above.
    }
  }

  // Auto-close once every manager has voted. For a named poll, a brand-new
  // voter row (isNewVoter) is the only thing that can push the count up —
  // changing an existing vote overwrites a row rather than adding one. For
  // an anonymous poll, count real participants instead of raw anon-id vote
  // rows now that PollParticipants exists — a manager voting from a second
  // browser/device only ever gets counted once there, since it's keyed by
  // their real name rather than a fresh random id each time. Falls back to
  // counting raw vote rows if no participantName came through at all (e.g.
  // a client that hasn't picked up this change yet) so closing still works,
  // just with the older, less precise signal. Wrapped in try/catch so a
  // hiccup fetching the roster count (a transient Sleeper API error) never
  // breaks the vote itself — worst case, the poll just waits for closesAt or
  // a manual close instead of closing this instant.
  const usingParticipantCount = poll.anonymousVoting && !!body.participantName;
  const shouldCheckClose = usingParticipantCount ? isNewParticipant : isNewVoter;
  if (shouldCheckClose) {
    try {
      const count = usingParticipantCount
        ? getAllPollParticipants().filter(p => p.pollId === body.pollId).length
        : getAllPollVotes().filter(v => v.pollId === body.pollId).length;
      if (count >= getLeagueSize()) {
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

// ============================================================
// LEAGUE NEWS — commissioner callouts + automated weekly reports
// ------------------------------------------------------------
// Two new sheet tabs, both auto-created via getOrCreateSheet():
//   CommissionerCallouts — Joe types rows in by hand. status column
//     (draft/published) gates visibility; index.html's ?preview=1 shows
//     drafts through the same render path as production.
//   WeeklyReports — one row per (season, week), appended by
//     generateWeeklyReport() below (wire it up: Triggers > Add trigger >
//     generateWeeklyReport > time-driven > week timer, the day after your
//     last games finish e.g. Tuesday morning). Always written as `draft`
//     first so a bad week's computation can't go live without a look.
//
// See claude/league-news-plan.md in the project for the full design
// rationale (why Overachiever/Underachiever was cut, why NFL opponent
// comes from a hand-maintained schedule array instead of Sleeper's API,
// the start-week guards, etc).
// ============================================================

const CALLOUT_HEADERS = ['date', 'title', 'body', 'status'];
const WEEKLY_REPORT_HEADERS = [
  'season', 'week', 'generatedAt', 'status',
  'trophiesJson', 'playersOfWeekJson', 'playoffPictureJson', 'powerRankingsJson',
];

// Start-week guards (see plan doc "Open items" — Joe signed off on these).
// Weekly trophies / players-of-the-week: from Week 1, no guard needed.
const POWER_RANKINGS_START_WEEK = 3;   // needs a real 3-game trailing window
const PLAYOFF_PICTURE_START_WEEK = 5;  // early standings are too thin before this

function getCalloutsSheet() {
  return getOrCreateSheet(CONFIG.COMMISSIONER_CALLOUTS_SHEET_NAME, CALLOUT_HEADERS);
}

function getWeeklyReportsSheet() {
  return getOrCreateSheet(CONFIG.WEEKLY_REPORTS_SHEET_NAME, WEEKLY_REPORT_HEADERS);
}

// ---------- NFL schedule (hand-maintained, see nfl.com/schedules) ----------
// Format: NFL_SCHEDULE[season][week][TEAM_ABBR] = 'vs OPP' | '@ OPP' | 'BYE'.
// Update once per season (~15 minutes, copy straight from the official NFL
// schedule) — only used for the opponent line on Players/Benchwarmers of the
// Week cards, nothing else depends on it. A missing season/week/team just
// means that card renders without an opponent line rather than erroring.
const NFL_SCHEDULE = {
  '2026': {
    1: {ARI:'@ LAC', ATL:'@ PIT', BAL:'@ IND', BUF:'@ HOU', CAR:'vs CHI', CHI:'@ CAR', CIN:'vs TB', CLE:'@ JAX', DAL:'@ NYG', DEN:'@ KC', DET:'vs NO', GB:'@ MIN', HOU:'vs BUF', IND:'vs BAL', JAX:'vs CLE', KC:'vs DEN', LAC:'vs ARI', LAR:'vs SF', LV:'vs MIA', MIA:'@ LV', MIN:'vs GB', NE:'@ SEA', NO:'@ DET', NYG:'vs DAL', NYJ:'@ TEN', PHI:'vs WAS', PIT:'vs ATL', SEA:'vs NE', SF:'@ LAR', TB:'@ CIN', TEN:'vs NYJ', WAS:'@ PHI'},
    2: {ARI:'vs SEA', ATL:'vs CAR', BAL:'vs NO', BUF:'vs DET', CAR:'@ ATL', CHI:'vs MIN', CIN:'@ HOU', CLE:'@ TB', DAL:'vs WAS', DEN:'vs JAX', DET:'@ BUF', GB:'@ NYJ', HOU:'vs CIN', IND:'@ KC', JAX:'@ DEN', KC:'vs IND', LAC:'vs LV', LAR:'vs NYG', LV:'@ LAC', MIA:'@ SF', MIN:'@ CHI', NE:'vs PIT', NO:'@ BAL', NYG:'@ LAR', NYJ:'vs GB', PHI:'@ TEN', PIT:'@ NE', SEA:'@ ARI', SF:'vs MIA', TB:'vs CLE', TEN:'vs PHI', WAS:'@ DAL'},
    3: {ARI:'@ SF', ATL:'@ GB', BAL:'@ DAL', BUF:'vs LAC', CAR:'@ CLE', CHI:'vs PHI', CIN:'@ PIT', CLE:'vs CAR', DAL:'vs BAL', DEN:'vs LAR', DET:'vs NYJ', GB:'vs ATL', HOU:'@ IND', IND:'vs HOU', JAX:'vs NE', KC:'@ MIA', LAC:'@ BUF', LAR:'@ DEN', LV:'@ NO', MIA:'vs KC', MIN:'@ TB', NE:'@ JAX', NO:'vs LV', NYG:'vs TEN', NYJ:'@ DET', PHI:'@ CHI', PIT:'vs CIN', SEA:'@ WAS', SF:'vs ARI', TB:'vs MIN', TEN:'@ NYG', WAS:'vs SEA'},
    4: {ARI:'@ NYG', ATL:'@ NO', BAL:'vs TEN', BUF:'vs NE', CAR:'vs DET', CHI:'vs NYJ', CIN:'vs JAX', CLE:'vs PIT', DAL:'@ HOU', DEN:'@ SF', DET:'@ CAR', GB:'@ TB', HOU:'vs DAL', IND:'@ WAS', JAX:'@ CIN', KC:'@ LV', LAC:'@ SEA', LAR:'@ PHI', LV:'vs KC', MIA:'@ MIN', MIN:'vs MIA', NE:'@ BUF', NO:'vs ATL', NYG:'vs ARI', NYJ:'@ CHI', PHI:'vs LAR', PIT:'@ CLE', SEA:'vs LAC', SF:'vs DEN', TB:'vs GB', TEN:'@ BAL', WAS:'vs IND'},
    5: {ARI:'vs DET', ATL:'vs BAL', BAL:'@ ATL', BUF:'@ LAR', CAR:'BYE', CHI:'@ GB', CIN:'@ MIA', CLE:'@ NYJ', DAL:'vs TB', DEN:'@ LAC', DET:'@ ARI', GB:'vs CHI', HOU:'@ TEN', IND:'@ PIT', JAX:'vs PHI', KC:'BYE', LAC:'vs DEN', LAR:'vs BUF', LV:'@ NE', MIA:'vs CIN', MIN:'@ NO', NE:'vs LV', NO:'vs MIN', NYG:'@ WAS', NYJ:'vs CLE', PHI:'@ JAX', PIT:'vs IND', SEA:'vs SF', SF:'@ SEA', TB:'@ DAL', TEN:'vs HOU', WAS:'vs NYG'},
    6: {ARI:'@ LAR', ATL:'vs CHI', BAL:'@ CLE', BUF:'@ LV', CAR:'@ PHI', CHI:'@ ATL', CIN:'BYE', CLE:'vs BAL', DAL:'@ GB', DEN:'vs SEA', DET:'BYE', GB:'vs DAL', HOU:'@ JAX', IND:'vs TEN', JAX:'vs HOU', KC:'vs LAC', LAC:'@ KC', LAR:'vs ARI', LV:'vs BUF', MIA:'BYE', MIN:'BYE', NE:'vs NYJ', NO:'@ NYG', NYG:'vs NO', NYJ:'@ NE', PHI:'vs CAR', PIT:'@ TB', SEA:'@ DEN', SF:'vs WAS', TB:'vs PIT', TEN:'@ IND', WAS:'@ SF'},
    7: {ARI:'vs DEN', ATL:'vs SF', BAL:'vs CIN', BUF:'BYE', CAR:'vs TB', CHI:'vs NE', CIN:'@ BAL', CLE:'@ TEN', DAL:'@ PHI', DEN:'@ ARI', DET:'vs GB', GB:'@ DET', HOU:'vs NYG', IND:'@ MIN', JAX:'BYE', KC:'@ SEA', LAC:'BYE', LAR:'@ LV', LV:'vs LAR', MIA:'@ NYJ', MIN:'vs IND', NE:'@ CHI', NO:'vs PIT', NYG:'@ HOU', NYJ:'vs MIA', PHI:'vs DAL', PIT:'@ NO', SEA:'vs KC', SF:'@ ATL', TB:'@ CAR', TEN:'vs CLE', WAS:'BYE'},
    8: {ARI:'@ DAL', ATL:'@ TB', BAL:'@ BUF', BUF:'vs BAL', CAR:'@ GB', CHI:'@ SEA', CIN:'vs TEN', CLE:'@ PIT', DAL:'vs ARI', DEN:'vs KC', DET:'vs MIN', GB:'vs CAR', HOU:'BYE', IND:'@ JAX', JAX:'vs IND', KC:'@ DEN', LAC:'@ LAR', LAR:'vs LAC', LV:'@ NYJ', MIA:'vs NE', MIN:'@ DET', NE:'@ MIA', NO:'BYE', NYG:'BYE', NYJ:'vs LV', PHI:'@ WAS', PIT:'vs CLE', SEA:'vs CHI', SF:'BYE', TB:'vs ATL', TEN:'@ CIN', WAS:'vs PHI'},
    9: {ARI:'@ SEA', ATL:'vs CIN', BAL:'vs JAX', BUF:'@ MIN', CAR:'vs DEN', CHI:'vs TB', CIN:'@ ATL', CLE:'@ NO', DAL:'@ IND', DEN:'@ CAR', DET:'@ MIA', GB:'@ NE', HOU:'@ LAC', IND:'vs DAL', JAX:'@ BAL', KC:'vs NYJ', LAC:'vs HOU', LAR:'@ WAS', LV:'@ SF', MIA:'vs DET', MIN:'vs BUF', NE:'vs GB', NO:'vs CLE', NYG:'@ PHI', NYJ:'@ KC', PHI:'vs NYG', PIT:'BYE', SEA:'vs ARI', SF:'vs LV', TB:'@ CHI', TEN:'BYE', WAS:'vs LAR'},
    10: {ARI:'vs LAR', ATL:'vs KC', BAL:'vs LAC', BUF:'@ NYJ', CAR:'@ NO', CHI:'BYE', CIN:'vs PIT', CLE:'vs HOU', DAL:'vs SF', DEN:'BYE', DET:'vs NE', GB:'vs MIN', HOU:'@ CLE', IND:'vs MIA', JAX:'@ TEN', KC:'@ ATL', LAC:'@ BAL', LAR:'@ ARI', LV:'vs SEA', MIA:'@ IND', MIN:'@ GB', NE:'@ DET', NO:'vs CAR', NYG:'vs WAS', NYJ:'vs BUF', PHI:'BYE', PIT:'@ CIN', SEA:'@ LV', SF:'@ DAL', TB:'BYE', TEN:'vs JAX', WAS:'@ NYG'},
    11: {ARI:'@ KC', ATL:'BYE', BAL:'@ CAR', BUF:'vs MIA', CAR:'vs BAL', CHI:'vs NO', CIN:'@ WAS', CLE:'BYE', DAL:'vs TEN', DEN:'vs LV', DET:'vs TB', GB:'BYE', HOU:'vs IND', IND:'@ HOU', JAX:'@ NYG', KC:'vs ARI', LAC:'vs NYJ', LAR:'BYE', LV:'@ DEN', MIA:'@ BUF', MIN:'@ SF', NE:'BYE', NO:'@ CHI', NYG:'vs JAX', NYJ:'@ LAC', PHI:'vs PIT', PIT:'@ PHI', SEA:'BYE', SF:'vs MIN', TB:'@ DET', TEN:'@ DAL', WAS:'vs CIN'},
    12: {ARI:'vs WAS', ATL:'@ MIN', BAL:'@ HOU', BUF:'vs KC', CAR:'@ TB', CHI:'@ DET', CIN:'vs NO', CLE:'vs LV', DAL:'vs PHI', DEN:'@ PIT', DET:'vs CHI', GB:'@ LAR', HOU:'vs BAL', IND:'vs NYG', JAX:'vs TEN', KC:'@ BUF', LAC:'vs NE', LAR:'vs GB', LV:'@ CLE', MIA:'vs NYJ', MIN:'vs ATL', NE:'@ LAC', NO:'@ CIN', NYG:'@ IND', NYJ:'@ MIA', PHI:'@ DAL', PIT:'vs DEN', SEA:'@ SF', SF:'vs SEA', TB:'vs CAR', TEN:'@ JAX', WAS:'@ ARI'},
    13: {ARI:'vs PHI', ATL:'vs DET', BAL:'BYE', BUF:'@ NE', CAR:'@ MIN', CHI:'vs JAX', CIN:'@ CLE', CLE:'vs CIN', DAL:'@ SEA', DEN:'vs MIA', DET:'@ ATL', GB:'@ NO', HOU:'@ PIT', IND:'BYE', JAX:'@ CHI', KC:'@ LAR', LAC:'@ TB', LAR:'vs KC', LV:'BYE', MIA:'@ DEN', MIN:'vs CAR', NE:'vs BUF', NO:'vs GB', NYG:'vs SF', NYJ:'BYE', PHI:'@ ARI', PIT:'vs HOU', SEA:'vs DAL', SF:'@ NYG', TB:'vs LAC', TEN:'vs WAS', WAS:'@ TEN'},
    14: {ARI:'@ BUF', ATL:'@ CLE', BAL:'vs PIT', BUF:'vs ARI', CAR:'vs NO', CHI:'@ MIA', CIN:'BYE', CLE:'vs ATL', DAL:'vs GB', DEN:'@ NYJ', DET:'vs TEN', GB:'@ DAL', HOU:'BYE', IND:'@ NYG', JAX:'BYE', KC:'vs LV', LAC:'@ SF', LAR:'BYE', LV:'@ KC', MIA:'vs CHI', MIN:'@ NE', NE:'vs MIN', NO:'@ CAR', NYG:'vs IND', NYJ:'vs DEN', PHI:'@ WAS', PIT:'@ BAL', SEA:'BYE', SF:'vs LAC', TB:'BYE', TEN:'@ DET', WAS:'vs PHI'},
    15: {ARI:'@ LAR', ATL:'@ JAX', BAL:'vs DAL', BUF:'vs NYJ', CAR:'@ WAS', CHI:'@ MIN', CIN:'@ NO', CLE:'@ PIT', DAL:'@ BAL', DEN:'@ LAC', DET:'@ GB', GB:'vs DET', HOU:'vs TB', IND:'@ TEN', JAX:'vs ATL', KC:'@ LV', LAC:'vs DEN', LAR:'vs ARI', LV:'vs KC', MIA:'vs NE', MIN:'vs CHI', NE:'@ MIA', NO:'vs CIN', NYG:'vs PHI', NYJ:'@ BUF', PHI:'@ NYG', PIT:'vs CLE', SEA:'@ SF', SF:'vs SEA', TB:'@ HOU', TEN:'vs IND', WAS:'vs CAR'},
    16: {ARI:'vs SEA', ATL:'vs BAL', BAL:'@ ATL', BUF:'vs NE', CAR:'vs MIA', CHI:'vs NYJ', CIN:'@ PIT', CLE:'vs TEN', DAL:'@ NYG', DEN:'vs LAC', DET:'@ SF', GB:'vs MIN', HOU:'@ KC', IND:'vs JAX', JAX:'@ IND', KC:'vs HOU', LAC:'@ DEN', LAR:'vs LV', LV:'@ LAR', MIA:'@ CAR', MIN:'@ GB', NE:'@ BUF', NO:'vs TB', NYG:'vs DAL', NYJ:'@ CHI', PHI:'vs WAS', PIT:'vs CIN', SEA:'@ ARI', SF:'vs DET', TB:'@ NO', TEN:'@ CLE', WAS:'@ PHI'},
    17: {ARI:'@ SEA', ATL:'vs BUF', BAL:'@ JAX', BUF:'@ ATL', CAR:'vs TB', CHI:'@ DET', CIN:'vs CLE', CLE:'@ CIN', DAL:'vs PHI', DEN:'@ LV', DET:'vs CHI', GB:'@ MIN', HOU:'vs TEN', IND:'vs NYG', JAX:'vs BAL', KC:'vs PIT', LAC:'vs WAS', LAR:'vs SF', LV:'vs DEN', MIA:'@ NE', MIN:'vs GB', NE:'vs MIA', NO:'vs NYJ', NYG:'@ IND', NYJ:'@ NO', PHI:'@ DAL', PIT:'@ KC', SEA:'vs ARI', SF:'@ LAR', TB:'@ CAR', TEN:'@ HOU', WAS:'@ LAC'},
  },
};

function nflOpponentFor(season, week, teamAbbr) {
  if (!teamAbbr) return '';
  const seasonSched = NFL_SCHEDULE[season];
  if (!seasonSched || !seasonSched[week]) return '';
  return seasonSched[week][teamAbbr] || '';
}

// ---------- Commissioner callouts ----------

// Sheets auto-parses a typed date (e.g. "9/7/2026") into a real Date cell,
// and getValues() hands that back as a JS Date — which, once JSON-serialized
// for the API response, turns into a full ISO timestamp with a time-of-day
// component nobody typed (see claude/status.md for the report that caught
// this). Format down to date-only for display; sorting still happens on the
// real Date value in getAllCallouts() below, before this runs, so this never
// risks breaking newest-first ordering. Falls back to the raw value as typed
// if it isn't a parseable date at all (e.g. Joe leaves it as free text).
function formatCalloutDate(raw) {
  if (!raw) return '';
  const d = raw instanceof Date ? raw : new Date(raw);
  if (isNaN(d.getTime())) return String(raw);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MM-dd-yyyy');
}

function getAllCallouts(includeDrafts) {
  const sheet = getCalloutsSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, CALLOUT_HEADERS.length).getValues();
  return values
    .map(row => {
      const obj = {};
      CALLOUT_HEADERS.forEach((h, i) => { obj[h] = row[i]; });
      obj.status = String(obj.status || 'draft').toLowerCase() || 'draft';
      // Keep the real Date (or unparsed raw value) for sorting, format a
      // separate display string for the API response.
      obj._sortDate = obj.date instanceof Date ? obj.date : new Date(obj.date || 0);
      obj.date = formatCalloutDate(obj.date);
      return obj;
    })
    .filter(c => c.title || c.body)
    .filter(c => includeDrafts || c.status === 'published')
    .sort((a, b) => b._sortDate - a._sortDate)
    .map(c => { delete c._sortDate; return c; });
}

// ---------- Weekly reports ----------

function getAllWeeklyReports(includeDrafts) {
  const sheet = getWeeklyReportsSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, WEEKLY_REPORT_HEADERS.length).getValues();
  return values
    .map(row => {
      const obj = {};
      WEEKLY_REPORT_HEADERS.forEach((h, i) => { obj[h] = row[i]; });
      obj.status = String(obj.status || 'draft').toLowerCase() || 'draft';
      ['trophiesJson', 'playersOfWeekJson', 'playoffPictureJson', 'powerRankingsJson'].forEach(k => {
        try { obj[k.replace('Json', '')] = obj[k] ? JSON.parse(obj[k]) : null; } catch (e) { obj[k.replace('Json', '')] = null; }
        delete obj[k];
      });
      return obj;
    })
    .filter(r => r.season && r.week)
    .filter(r => includeDrafts || r.status === 'published')
    .sort((a, b) => (Number(b.season) - Number(a.season)) || (Number(b.week) - Number(a.week)));
}

function findWeeklyReportRowIndex(sheet, season, week) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const rows = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); // season, week
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === String(season) && Number(rows[i][1]) === Number(week)) return i + 2;
  }
  return -1;
}

// ---------- Data plumbing shared by the report builder ----------

// Full Sleeper player directory, fetched fresh each run (~5MB) rather than
// cached — CacheService/PropertiesService are both far too small to hold it,
// and this only runs on a weekly trigger, not per page view.
function getPlayerInfoGS() {
  const all = fetchJson('https://api.sleeper.app/v1/players/nfl');
  const info = {};
  Object.keys(all).forEach(id => {
    const p = all[id];
    if (!p) return;
    info[id] = {
      name: p.full_name || `${p.first_name || ''} ${p.last_name || ''}`.trim() || `Player #${id}`,
      pos: p.position || null,
      team: p.team || null,
    };
  });
  return info;
}

function getCurrentManagersGS() {
  const leagueId = CONFIG.CURRENT_LEAGUE_ID;
  const rosters = fetchJson(`https://api.sleeper.app/v1/league/${leagueId}/rosters`);
  const users = fetchJson(`https://api.sleeper.app/v1/league/${leagueId}/users`);
  const usersById = {};
  users.forEach(u => { usersById[u.user_id] = u; });

  const managers = {};
  rosters.forEach(r => {
    const u = usersById[r.owner_id] || null;
    const teamName = (u && u.metadata && u.metadata.team_name) || (u && u.display_name) || `Team ${r.roster_id}`;
    const managerName = (u && u.display_name) || `Team ${r.roster_id}`;
    managers[r.roster_id] = { rosterId: r.roster_id, teamName, managerName };
  });
  return managers;
}

// Same optimal-lineup math as index.html's computeOptimalLineupScore(), just
// taking a plain playerInfo map instead of the client's module-level lookup.
function computeOptimalLineupScoreGS(playersPoints, rosterPositions, playerInfo) {
  const startSlots = (rosterPositions || []).filter(p => p !== 'BN' && p !== 'IR' && p !== 'TAXI');
  const strictSlots = startSlots.filter(p => p !== 'FLEX' && p !== 'SUPER_FLEX');
  const flexCount = startSlots.filter(p => p === 'FLEX').length;
  const superFlexCount = startSlots.filter(p => p === 'SUPER_FLEX').length;

  const byPos = {};
  Object.keys(playersPoints || {}).forEach(pid => {
    const info = playerInfo[pid];
    const pos = info && info.pos;
    if (!pos) return;
    (byPos[pos] = byPos[pos] || []).push({ pid, pts: playersPoints[pid] || 0 });
  });
  Object.values(byPos).forEach(arr => arr.sort((a, b) => b.pts - a.pts));

  const used = new Set();
  let total = 0;

  const strictCounts = {};
  strictSlots.forEach(s => { strictCounts[s] = (strictCounts[s] || 0) + 1; });
  Object.keys(strictCounts).forEach(pos => {
    const count = strictCounts[pos];
    const pool = byPos[pos] || [];
    let filled = 0;
    for (const p of pool) {
      if (filled >= count) break;
      if (used.has(p.pid)) continue;
      used.add(p.pid); total += p.pts; filled++;
    }
  });

  if (flexCount > 0) {
    const pool = ['RB', 'WR', 'TE'].reduce((acc, pos) => acc.concat(byPos[pos] || []), [])
      .filter(p => !used.has(p.pid)).sort((a, b) => b.pts - a.pts);
    for (let i = 0; i < flexCount && i < pool.length; i++) { used.add(pool[i].pid); total += pool[i].pts; }
  }
  if (superFlexCount > 0) {
    const pool = ['QB', 'RB', 'WR', 'TE'].reduce((acc, pos) => acc.concat(byPos[pos] || []), [])
      .filter(p => !used.has(p.pid)).sort((a, b) => b.pts - a.pts);
    for (let i = 0; i < superFlexCount && i < pool.length; i++) { used.add(pool[i].pid); total += pool[i].pts; }
  }

  return total;
}

// Pairs up a raw /matchups/{week} response into [{a, b}] entries, skipping the
// both-sides-zero placeholder Sleeper returns for a week nobody's played yet.
function pairUpMatchups(rawEntries) {
  const byId = {};
  (rawEntries || []).forEach(entry => {
    const mid = entry.matchup_id;
    if (mid == null) return;
    (byId[mid] = byId[mid] || []).push(entry);
  });
  const pairs = [];
  Object.values(byId).forEach(entries => {
    if (entries.length < 2) return;
    const [a, b] = entries;
    if ((a.points || 0) === 0 && (b.points || 0) === 0) return;
    pairs.push({ a, b });
  });
  return pairs;
}

// Cumulative regular-season standings through (and including) throughWeek —
// wins/losses/points-for per roster, same sort convention as the live
// Standings table (wins, then points-for). Also returns each roster's points
// for just the trailing 3 weeks (or fewer, early in the season), for the
// power-ranking formula.
function computeStandingsThroughWeek(leagueId, throughWeek, playoffWeekStart) {
  const lastRegularWeek = Math.min(throughWeek, playoffWeekStart - 1);
  const standings = {}; // rosterId -> { wins, losses, pf }
  const weeklyPointsByRoster = {}; // rosterId -> [points per week, in week order]

  for (let week = 1; week <= lastRegularWeek; week++) {
    let raw;
    try {
      raw = fetchJson(`https://api.sleeper.app/v1/league/${leagueId}/matchups/${week}`);
    } catch (e) { continue; }
    const pairs = pairUpMatchups(raw);
    pairs.forEach(({ a, b }) => {
      [a, b].forEach(side => {
        if (!standings[side.roster_id]) standings[side.roster_id] = { wins: 0, losses: 0, pf: 0 };
        standings[side.roster_id].pf += side.points || 0;
        (weeklyPointsByRoster[side.roster_id] = weeklyPointsByRoster[side.roster_id] || []).push(side.points || 0);
      });
      const aWon = (a.points || 0) > (b.points || 0);
      if ((a.points || 0) !== (b.points || 0)) {
        standings[aWon ? a.roster_id : b.roster_id].wins++;
        standings[aWon ? b.roster_id : a.roster_id].losses++;
      }
    });
  }

  return { standings, weeklyPointsByRoster };
}

// ---------- Weekly report generation ----------

// Run manually from the Apps Script editor (or via the one-time trigger
// setup) with no arguments to report on "the most recently fully completed
// week" — same last_scored_leg convention index.html itself uses to avoid
// treating a still-live week as final. Pass a week number to backfill or
// re-run a specific week on purpose.
function generateWeeklyReport(weekOverride) {
  const league = fetchJson(`https://api.sleeper.app/v1/league/${CONFIG.CURRENT_LEAGUE_ID}`);
  const season = league.season;
  const playoffWeekStart = league.settings.playoff_week_start || 15;
  const playoffTeams = league.settings.playoff_teams || CONFIG.PLAYOFF_TEAMS;
  const rosterPositions = league.roster_positions || [];

  const reportWeek = weekOverride || Math.max(1, (league.settings.last_scored_leg || 1) - 1);
  if (reportWeek < 1) return { skipped: true, reason: 'No completed week yet this season.' };

  const reportsSheet = getWeeklyReportsSheet();
  if (findWeeklyReportRowIndex(reportsSheet, season, reportWeek) !== -1) {
    return { skipped: true, reason: `A report for ${season} Week ${reportWeek} already exists — not overwriting it.` };
  }

  const managers = getCurrentManagersGS();
  const playerInfo = getPlayerInfoGS();
  const rawWeek = fetchJson(`https://api.sleeper.app/v1/league/${CONFIG.CURRENT_LEAGUE_ID}/matchups/${reportWeek}`);
  const pairs = pairUpMatchups(rawWeek);

  if (pairs.length === 0) {
    return { skipped: true, reason: `No completed matchups found for ${season} Week ${reportWeek}.` };
  }

  const mgrName = rid => (managers[rid] && managers[rid].managerName) || `Team ${rid}`;
  const teamName = rid => (managers[rid] && managers[rid].teamName) || `Team ${rid}`;

  // ---- Highest / Lowest Scorer, Highest Pts in Loss / Lowest Pts in Win,
  //      Biggest Blowout / Narrow Victory, Most/Least Efficient Manager ----
  const allSides = []; // { rosterId, points, won, oppRosterId, oppPoints, optimal, efficiency }
  const games = []; // { winner, loser, winnerPts, loserPts, marginPct }

  pairs.forEach(({ a, b }) => {
    const aWon = (a.points || 0) > (b.points || 0);
    const winner = aWon ? a : b, loser = aWon ? b : a;
    const margin = Math.abs((a.points || 0) - (b.points || 0));
    const avg = ((a.points || 0) + (b.points || 0)) / 2 || 1;
    games.push({
      winnerRosterId: winner.roster_id, loserRosterId: loser.roster_id,
      winnerPts: winner.points || 0, loserPts: loser.points || 0,
      marginPct: (margin / avg) * 100,
    });

    [{ side: a, opp: b }, { side: b, opp: a }].forEach(({ side, opp }) => {
      const optimal = computeOptimalLineupScoreGS(side.players_points, rosterPositions, playerInfo);
      allSides.push({
        rosterId: side.roster_id,
        points: side.points || 0,
        won: (side.points || 0) > (opp.points || 0),
        oppRosterId: opp.roster_id,
        oppPoints: opp.points || 0,
        optimal,
        efficiency: optimal > 0 ? ((side.points || 0) / optimal) * 100 : 0,
      });
    });
  });

  const withEpsilonMax = (arr, valueFn) => {
    const best = Math.max(...arr.map(valueFn));
    return arr.filter(x => Math.abs(valueFn(x) - best) < 1e-9);
  };
  const withEpsilonMin = (arr, valueFn) => {
    const best = Math.min(...arr.map(valueFn));
    return arr.filter(x => Math.abs(valueFn(x) - best) < 1e-9);
  };
  const sideTag = s => ({ rosterId: s.rosterId, manager: mgrName(s.rosterId), team: teamName(s.rosterId), points: Math.round(s.points * 100) / 100 });

  const highestScorers = withEpsilonMax(allSides, s => s.points).map(sideTag);
  const lowestScorers = withEpsilonMin(allSides, s => s.points).map(sideTag);

  const losers = allSides.filter(s => !s.won);
  const winners = allSides.filter(s => s.won);
  const highestPtsInLoss = losers.length ? withEpsilonMax(losers, s => s.points).map(s => ({ ...sideTag(s), lostTo: mgrName(s.oppRosterId) })) : [];
  const lowestPtsInWin = winners.length ? withEpsilonMin(winners, s => s.points).map(s => ({ ...sideTag(s), beat: mgrName(s.oppRosterId) })) : [];

  const biggestBlowout = withEpsilonMax(games, g => g.marginPct).map(g => ({
    winner: mgrName(g.winnerRosterId), loser: mgrName(g.loserRosterId),
    winnerPts: Math.round(g.winnerPts * 100) / 100, loserPts: Math.round(g.loserPts * 100) / 100,
    marginPct: Math.round(g.marginPct * 10) / 10,
  }));
  const narrowVictory = withEpsilonMin(games, g => g.marginPct).map(g => ({
    winner: mgrName(g.winnerRosterId), loser: mgrName(g.loserRosterId),
    winnerPts: Math.round(g.winnerPts * 100) / 100, loserPts: Math.round(g.loserPts * 100) / 100,
    marginPct: Math.round(g.marginPct * 10) / 10,
  }));

  // Shootout / Dud of the Week — same idea as the Trophy Case's season-long
  // "Shootout of the Century"/"Dud of the Century", scoped to just this
  // week's games. Reuses the same `games` array as the blowout/narrow-
  // victory trophies above (no new fetches), just ranked by combined score
  // instead of margin. Added 2026-09-08, per Joe, to fill out the Weekly
  // Trophies grid to an even 12 cards (was 10, leaving an unbalanced last row).
  const shootoutOfWeek = withEpsilonMax(games, g => g.winnerPts + g.loserPts).map(g => ({
    winner: mgrName(g.winnerRosterId), loser: mgrName(g.loserRosterId),
    winnerPts: Math.round(g.winnerPts * 100) / 100, loserPts: Math.round(g.loserPts * 100) / 100,
    totalPts: Math.round((g.winnerPts + g.loserPts) * 100) / 100,
  }));
  const dudOfWeek = withEpsilonMin(games, g => g.winnerPts + g.loserPts).map(g => ({
    winner: mgrName(g.winnerRosterId), loser: mgrName(g.loserRosterId),
    winnerPts: Math.round(g.winnerPts * 100) / 100, loserPts: Math.round(g.loserPts * 100) / 100,
    totalPts: Math.round((g.winnerPts + g.loserPts) * 100) / 100,
  }));

  const mostEfficient = withEpsilonMax(allSides, s => s.efficiency).map(s => ({ ...sideTag(s), efficiency: Math.round(s.efficiency * 10) / 10, optimal: Math.round(s.optimal * 100) / 100 }));
  const leastEfficient = withEpsilonMin(allSides, s => s.efficiency).map(s => ({ ...sideTag(s), efficiency: Math.round(s.efficiency * 10) / 10, optimal: Math.round(s.optimal * 100) / 100 }));

  // ---- Players of the Week / Benchwarmers of the Week, by position ----
  const POSITIONS = ['QB', 'RB', 'WR', 'TE', 'K', 'DEF'];
  const startedBest = {}, benchedBest = {}; // pos -> best entry seen so far
  POSITIONS.forEach(p => { startedBest[p] = null; benchedBest[p] = null; });

  rawWeek.forEach(entry => {
    const starters = new Set((entry.starters || []).filter(id => id && id !== '0'));
    const pp = entry.players_points || {};
    Object.keys(pp).forEach(pid => {
      const info = playerInfo[pid];
      const pos = info && info.pos;
      if (!POSITIONS.includes(pos)) return;
      const pts = pp[pid] || 0;
      const candidate = {
        playerId: pid,
        name: info.name,
        pos,
        team: info.team,
        opponent: nflOpponentFor(season, reportWeek, info.team),
        points: Math.round(pts * 100) / 100,
        rosterId: entry.roster_id,
        manager: mgrName(entry.roster_id),
      };
      const bucket = starters.has(pid) ? startedBest : benchedBest;
      if (!bucket[pos] || pts > bucket[pos].points) bucket[pos] = candidate;
    });
  });

  const playersOfWeek = { started: startedBest, benched: benchedBest };

  // ---- Standings-to-date (feeds Streak Watch, Closest to the Cutoff,
  //      playoff picture, and power rankings) ----
  const { standings, weeklyPointsByRoster } = computeStandingsThroughWeek(CONFIG.CURRENT_LEAGUE_ID, reportWeek, playoffWeekStart);
  const rosterIds = Object.keys(standings).map(Number);
  const sortedStandings = rosterIds.slice().sort((r1, r2) => {
    const s1 = standings[r1], s2 = standings[r2];
    if (s2.wins !== s1.wins) return s2.wins - s1.wins;
    return s2.pf - s1.pf;
  });

  // ---- Bonus trophy: Streak Watch (longest active win or loss streak) ----
  // computeStandingsThroughWeek() above only totals wins/losses/PF, not
  // per-roster chronological results — walk regular-season weeks once more,
  // tracking each roster's current streak as we go (resets on a W/L flip).
  const streaks = {}; // rosterId -> { type: 'W'|'L', length }
  (function buildStreaks() {
    const lastRegularWeek = Math.min(reportWeek, playoffWeekStart - 1);
    const current = {}; // rosterId -> { type, length }
    for (let week = 1; week <= lastRegularWeek; week++) {
      let raw;
      try {
        raw = fetchJson(`https://api.sleeper.app/v1/league/${CONFIG.CURRENT_LEAGUE_ID}/matchups/${week}`);
      } catch (e) { continue; }
      pairUpMatchups(raw).forEach(({ a, b }) => {
        if ((a.points || 0) === (b.points || 0)) return; // no real ties expected, guard anyway
        const aWon = (a.points || 0) > (b.points || 0);
        [{ rid: a.roster_id, won: aWon }, { rid: b.roster_id, won: !aWon }].forEach(({ rid, won }) => {
          const type = won ? 'W' : 'L';
          if (current[rid] && current[rid].type === type) current[rid].length++;
          else current[rid] = { type, length: 1 };
        });
      });
    }
    Object.assign(streaks, current);
  })();

  const streakLeaders = rosterIds.filter(rid => streaks[rid]);
  let streakWatch = null;
  if (streakLeaders.length) {
    const best = streakLeaders.reduce((a, b) => (streaks[b].length > streaks[a].length ? b : a));
    streakWatch = {
      rosterId: best, manager: mgrName(best), team: teamName(best),
      type: streaks[best].type === 'W' ? 'win' : 'loss', length: streaks[best].length,
    };
  }

  // ---- Bonus trophy: Closest to the Cutoff (first team out, vs. the last team in) ----
  let closestToCutoff = null;
  if (sortedStandings.length > playoffTeams) {
    const lastIn = sortedStandings[playoffTeams - 1];
    const firstOut = sortedStandings[playoffTeams];
    const gamesBack = (standings[lastIn].wins - standings[firstOut].wins + (standings[firstOut].losses - standings[lastIn].losses)) / 2;
    closestToCutoff = {
      rosterId: firstOut, manager: mgrName(firstOut), team: teamName(firstOut),
      lastInRosterId: lastIn, lastInManager: mgrName(lastIn),
      gamesBack: Math.round(gamesBack * 10) / 10,
      pointsBack: Math.round((standings[lastIn].pf - standings[firstOut].pf) * 100) / 100,
    };
  }

  const trophies = {
    highestScorers, lowestScorers, highestPtsInLoss, lowestPtsInWin,
    biggestBlowout, narrowVictory, shootoutOfWeek, dudOfWeek,
    mostEfficient, leastEfficient, streakWatch, closestToCutoff,
  };

  // ---- Playoff picture (Week 5+) ----
  let playoffPicture = null;
  if (reportWeek >= PLAYOFF_PICTURE_START_WEEK) {
    const bubbleSize = 2; // 7th-8th, per the confirmed framing
    const toRow = rid => ({
      rosterId: rid, manager: mgrName(rid), team: teamName(rid),
      wins: standings[rid].wins, losses: standings[rid].losses, pf: Math.round(standings[rid].pf * 100) / 100,
    });
    playoffPicture = {
      in: sortedStandings.slice(0, playoffTeams).map(toRow),
      bubble: sortedStandings.slice(playoffTeams, playoffTeams + bubbleSize).map(toRow),
      out: sortedStandings.slice(playoffTeams + bubbleSize).map(toRow),
    };
  }

  // ---- Power rankings (Week 3+) ----
  let powerRankings = null;
  if (reportWeek >= POWER_RANKINGS_START_WEEK) {
    const winPct = rid => { const s = standings[rid]; const t = s.wins + s.losses; return t ? s.wins / t : 0; };
    const seasonPpg = rid => { const wk = weeklyPointsByRoster[rid] || []; return wk.length ? wk.reduce((a, b) => a + b, 0) / wk.length : 0; };
    const last3Ppg = rid => { const wk = weeklyPointsByRoster[rid] || []; const last3 = wk.slice(-3); return last3.length ? last3.reduce((a, b) => a + b, 0) / last3.length : 0; };

    const rankOf = (ids, valueFn) => {
      const sorted = ids.slice().sort((a, b) => valueFn(b) - valueFn(a));
      const ranks = {};
      sorted.forEach((rid, i) => { ranks[rid] = i + 1; });
      return ranks;
    };
    const winPctRank = rankOf(rosterIds, winPct);
    const ppgRank = rankOf(rosterIds, seasonPpg);
    const last3Rank = rankOf(rosterIds, last3Ppg);

    const blended = {};
    rosterIds.forEach(rid => {
      blended[rid] = 0.45 * winPctRank[rid] + 0.35 * last3Rank[rid] + 0.20 * ppgRank[rid];
    });
    const finalOrder = rosterIds.slice().sort((a, b) => blended[a] - blended[b]);

    // Trend vs. the most recent earlier report for this season that actually
    // has power rankings on it — not strictly reportWeek-1, so a one-off gap
    // (a missed trigger run, or simply the first report at/after the Week 3
    // start guard) doesn't wipe out trend tracking for every week after it.
    const priorReport = getAllWeeklyReports(true)
      .filter(r => String(r.season) === String(season) && Number(r.week) < reportWeek && r.powerRankings)
      .sort((a, b) => Number(b.week) - Number(a.week))[0] || null;
    const priorRanks = {};
    if (priorReport && priorReport.powerRankings && priorReport.powerRankings.rankings) {
      priorReport.powerRankings.rankings.forEach(r => { priorRanks[r.rosterId] = r.rank; });
    }

    powerRankings = {
      rankings: finalOrder.map((rid, i) => {
        const rank = i + 1;
        const prior = priorRanks[rid];
        return {
          rosterId: rid, manager: mgrName(rid), team: teamName(rid), rank,
          trend: prior == null ? 'new' : (prior > rank ? 'up' : prior < rank ? 'down' : 'same'),
          priorRank: prior == null ? null : prior,
          winPct: Math.round(winPct(rid) * 1000) / 10,
          seasonPpg: Math.round(seasonPpg(rid) * 100) / 100,
          last3Ppg: Math.round(last3Ppg(rid) * 100) / 100,
        };
      }),
    };
  }

  reportsSheet.appendRow([
    season, reportWeek, new Date().toISOString(), 'draft',
    JSON.stringify(trophies), JSON.stringify(playersOfWeek),
    JSON.stringify(playoffPicture), JSON.stringify(powerRankings),
  ]);

  return { ok: true, season, week: reportWeek };
}

// Manual/admin trigger for testing or backfilling a specific week without
// waiting for the scheduled trigger: GET ?action=generatereport&week=N&key=...
function adminGenerateWeeklyReport(e) {
  requireAdmin(e);
  const week = e.parameter.week ? Number(e.parameter.week) : undefined;
  return generateWeeklyReport(week);
}
