/**
 * Interview Recorder - server (Google Apps Script), v0.1.0
 *
 * What it does
 *  - Hands the question list / school list / interviewer list to the phone app
 *  - Receives the audio recording (in small pieces) and saves it to Google Drive
 *  - Writes the interview details and per-question notes into this spreadsheet
 *
 * One-time setup: see README (run setup() once, then Deploy > New deployment > Web app).
 */

// ===== EDIT THESE TWO LINES =====
const PASSCODE = 'CHANGE-ME';   // shared secret typed (or sent by link) to the team's phones
const FOLDER_ID = '';            // Drive folder ID for recordings. Leave '' to auto-create "DS Teacher Interviews" in your My Drive.
// ================================

const FOLDER_NAME = 'DS Teacher Interviews';
const SHEETS = {Q: 'Questions', S: 'Schools', W: 'Interviewers', I: 'Interviews', N: 'Notes'};
const HEAD_Q = ['No', 'Section', 'Stage', 'Question', 'Probe', 'Required'];
const HEAD_I = ['interview_id', 'date', 'start_time', 'school', 'interviewee_id', 'interviewer', 'duration', 'consent_confirmed',
                'questions_with_notes', 'audio_file', 'audio_url', 'overall_impressions', 'demo', 'app_version', 'device', 'uploaded_at'];
const HEAD_N = ['interview_id', 'date', 'school', 'interviewee_id', 'interviewer', 'q_no', 'stage', 'section', 'question',
                'first_asked_at', 'notes', 'marks_at'];

/* ---------- web app entry points ---------- */

function doGet() {
  return ContentService.createTextOutput('Interview Recorder server is running (v0.1.0).');
}

function doPost(e) {
  try {
    const b = JSON.parse(e.postData.contents);
    if (b.key !== PASSCODE) return out_({ok: false, error: 'Wrong passcode.'});
    switch (b.action) {
      case 'config':        return out_({ok: true, data: getConfig_()});
      case 'audioChunk':    return out_(audioChunk_(b));
      case 'audioFinalize': return out_(audioFinalize_(b));
      case 'saveInterview': return out_(saveInterview_(b.record));
      default:              return out_({ok: false, error: 'Unknown action.'});
    }
  } catch (err) {
    return out_({ok: false, error: String(err && err.message || err)});
  }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- one-time setup (run from the editor) ---------- */

function setup() {
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open this script from the spreadsheet: Extensions > Apps Script.');
  props.setProperty('SS_ID', ss.getId());

  const q = sheet_(ss, SHEETS.Q, HEAD_Q);
  if (q.getLastRow() < 2) {
    q.getRange(2, 1, 6, 6).setValues([
      [1, 'A. Observation', 'Observation', 'Think about the lesson you observed. What did you notice about the learners?', 'What did you see or hear that made you say that?', 'Y'],
      [2, 'A. Observation', 'Observation', 'What did the student-teacher do well in that lesson?', 'Can you give one specific moment?', 'Y'],
      [3, 'B. Feedback', 'Feedback', 'What feedback did you give after the lesson?', 'When? Face to face or written? How long did it take?', 'Y'],
      [4, 'B. Feedback', 'Feedback', 'Why did you choose that feedback?', 'What did you want the student-teacher to understand?', ''],
      [5, 'C. Next action', 'Advice', 'What advice or next action did you suggest?', 'What exactly should they try in the next lesson?', 'Y'],
      [6, 'D. Follow-up', 'Re-observation', 'Did you observe the student-teacher again after the advice?', 'If yes: what changed? If no: what got in the way?', 'Y']
    ]);
  }
  q.getRange('A1').setNote('Question number (1, 2, 3...). Rows are shown to interviewers in this order.');
  q.getRange('C1').setNote('Optional label shown as a small tag: Observation / Feedback / Advice / Re-observation ...');
  q.getRange('E1').setNote('Optional follow-up prompt shown under the question.');
  q.getRange('F1').setNote('Y = required, blank = optional.');

  const s = sheet_(ss, SHEETS.S, ['School']);
  if (s.getLastRow() < 2) s.getRange(2, 1, 2, 1).setValues([['Sample School A (edit me)'], ['Sample School B (edit me)']]);
  const w = sheet_(ss, SHEETS.W, ['Interviewer']);
  if (w.getLastRow() < 2) w.getRange(2, 1, 3, 1).setValues([['Kin'], ['Local member 1 (edit me)'], ['Local member 2 (edit me)']]);
  sheet_(ss, SHEETS.I, HEAD_I);
  sheet_(ss, SHEETS.N, HEAD_N);
  const first = ss.getSheetByName('Sheet1') || ss.getSheetByName('シート1');
  if (first && ss.getSheets().length > 1 && first.getLastRow() === 0) ss.deleteSheet(first);

  const f = mainFolder_();
  Logger.log('Setup done. Recordings folder: ' + f.getUrl());
  Logger.log('Now deploy: Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).');
}

/* ---------- helpers ---------- */

function sheet_(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground('#e0f2fe');
    sh.setFrozenRows(1);
  }
  return sh;
}

function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  if (id) return SpreadsheetApp.openById(id);
  const a = SpreadsheetApp.getActiveSpreadsheet();
  if (!a) throw new Error('Run setup() once first.');
  return a;
}

function mainFolder_() {
  if (FOLDER_ID) return DriveApp.getFolderById(FOLDER_ID);
  const props = PropertiesService.getScriptProperties();
  const saved = props.getProperty('AUTO_FOLDER_ID');
  if (saved) { try { return DriveApp.getFolderById(saved); } catch (e) { /* recreate below */ } }
  const f = DriveApp.createFolder(FOLDER_NAME);
  props.setProperty('AUTO_FOLDER_ID', f.getId());
  return f;
}

function tmpFolder_() {
  const main = mainFolder_();
  const it = main.getFoldersByName('_tmp');
  return it.hasNext() ? it.next() : main.createFolder('_tmp');
}

function col1_(sh) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, 1).getValues().map(r => String(r[0]).trim()).filter(String);
}

function p_(n, w) { let s = String(n); while (s.length < w) s = '0' + s; return s; }
function mmss_(ms) {
  const s = Math.max(0, Math.round(ms / 1000));
  const h = Math.floor(s / 3600), m = Math.floor(s % 3600 / 60), x = s % 60;
  return (h ? h + ':' + p_(m, 2) : p_(m, 2)) + ':' + p_(x, 2);
}

/* ---------- config for the phones ---------- */

function getConfig_() {
  const ss = ss_();
  const qs = ss.getSheetByName(SHEETS.Q);
  const n = Math.max(0, qs.getLastRow() - 1);
  const rows = n ? qs.getRange(2, 1, n, 6).getValues() : [];
  const questions = rows
    .filter(r => String(r[3]).trim() !== '')
    .map((r, i) => ({
      no: (r[0] === '' ? i + 1 : r[0]),
      section: String(r[1]).trim(),
      stage: String(r[2]).trim(),
      question: String(r[3]).trim(),
      probe: String(r[4]).trim(),
      required: /^(y|yes|true|1|required)$/i.test(String(r[5]).trim())
    }));
  return {questions: questions, schools: col1_(ss.getSheetByName(SHEETS.S)), interviewers: col1_(ss.getSheetByName(SHEETS.W))};
}

/* ---------- audio ---------- */

function audioChunk_(b) {
  const tmp = tmpFolder_();
  const name = b.id + '__' + p_(b.index, 4);
  const old = tmp.getFilesByName(name);
  while (old.hasNext()) old.next().setTrashed(true);
  const bytes = Utilities.base64Decode(b.data);
  tmp.createFile(Utilities.newBlob(bytes, 'application/octet-stream', name));
  return {ok: true};
}

function audioFinalize_(b) {
  const main = mainFolder_();
  const marker = 'interview_id=' + b.id;
  const baseName = String(b.filename || (b.id + '.webm')).replace(/[\\/:*?"<>|]+/g, '_');

  // Already finalized earlier (a retry after a lost reply)? Return that file.
  const same = main.getFilesByName(baseName);
  while (same.hasNext()) {
    const f = same.next();
    if (f.getDescription() === marker && !f.isTrashed()) return {ok: true, url: f.getUrl(), name: f.getName()};
  }

  const tmp = tmpFolder_();
  let all = [];
  const pieces = [];
  for (let i = 0; i < b.total; i++) {
    const it = tmp.getFilesByName(b.id + '__' + p_(i, 4));
    if (!it.hasNext()) return {ok: false, error: 'Audio piece ' + (i + 1) + ' of ' + b.total + ' is missing on the server. Please send again.'};
    const f = it.next();
    pieces.push(f);
    all = all.concat(f.getBlob().getBytes());
  }
  // Pick a free file name (another interview may already use it)
  let name = baseName, k = 2;
  const dot = baseName.lastIndexOf('.');
  const stem = dot > 0 ? baseName.slice(0, dot) : baseName, ext = dot > 0 ? baseName.slice(dot) : '';
  while (main.getFilesByName(name).hasNext()) { name = stem + '_' + k + ext; k++; }
  const file = main.createFile(Utilities.newBlob(all, b.mime || 'audio/webm', name));
  file.setDescription(marker);
  pieces.forEach(f => f.setTrashed(true));
  return {ok: true, url: file.getUrl(), name: file.getName()};
}

/* ---------- interview details + notes ---------- */

function saveInterview_(r) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = ss_();
    const sI = ss.getSheetByName(SHEETS.I), sN = ss.getSheetByName(SHEETS.N);
    const withNotes = (r.questions || []).filter(q => String((r.notes || {})[String(q.no)] || '').trim()).length;
    const row = [r.id, r.date, r.startTime, r.school, r.interviewee, r.interviewer, mmss_((r.durationSec || 0) * 1000),
                 r.consent ? 'Yes' : 'No', withNotes + '/' + (r.questions || []).length, r.audioName || '', r.audioUrl || '',
                 r.general || '', r.demo ? 'DEMO' : '', r.appVersion || '', r.device || '', new Date().toISOString()];
    upsertRow_(sI, r.id, row);

    // replace this interview's note rows
    const last = sN.getLastRow();
    if (last > 1) {
      const ids = sN.getRange(2, 1, last - 1, 1).getValues();
      for (let i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]) === String(r.id)) sN.deleteRow(i + 2);
    }
    const rows = (r.questions || []).map(q => {
      const v = (r.visits || []).filter(x => String(x.q) === String(q.no))[0];
      const marks = (r.flags || []).filter(x => String(x.q) === String(q.no)).map(x => mmss_(x.t)).join(', ');
      return [r.id, r.date, r.school, r.interviewee, r.interviewer, q.no, q.stage || '', q.section || '', q.question,
              v ? mmss_(v.t) : '', String((r.notes || {})[String(q.no)] || ''), marks];
    });
    if (rows.length) {
      const start = sN.getLastRow() + 1;
      const rg = sN.getRange(start, 1, rows.length, HEAD_N.length);
      rg.setNumberFormat('@');          // store as plain text so notes starting with "=" are not treated as formulas
      rg.setValues(rows);
    }
    return {ok: true};
  } finally {
    lock.releaseLock();
  }
}

function upsertRow_(sh, id, row) {
  const last = sh.getLastRow();
  let target = last + 1;
  if (last > 1) {
    const ids = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) { target = i + 2; break; }
  }
  const rg = sh.getRange(target, 1, 1, row.length);
  rg.setNumberFormat('@');
  rg.setValues([row]);
}
