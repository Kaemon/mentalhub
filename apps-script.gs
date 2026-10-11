/**
 * MentalHub: backend for the check-in (Google Apps Script + Google Sheets)
 *
 * Setup:
 * 1. Create a new Google Sheet (e.g. "MentalHub – Responses").
 * 2. Extensions > Apps Script. Delete the sample code and paste this whole file. Save.
 * 3. Deploy > New deployment > type "Web app".
 *      Execute as: Me
 *      Who has access: Anyone
 *    Click Deploy and authorise it.
 * 4. Copy the Web app URL (ends in /exec) and paste it into CONFIG.endpoint in index.html.
 *
 * If you edit this script later: Deploy > Manage deployments > Edit > Version: New version.
 * The URL stays the same.
 *
 * Sheets created:
 *   Responses – one row per completed check-in (scores + every answer)
 *   Events    – one row per support contact / article click (for the "used support" metric)
 */

const PSS_REVERSED = [3, 4, 6, 7]; // PSS-10 items 4, 5, 7, 8 (0-based)
const YEARS = ['Foundation / Diploma', 'Year 1', 'Year 2', 'Year 3', 'Other / prefer not to say'];
const LEN = { pss: 10, gad: 7, work: 4 };
const MAXV = { pss: 4, gad: 3, work: 4 };

function sheet_(name, header) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(header);
    sh.setFrozenRows(1);
  }
  return sh;
}
function responses_() {
  const h = ['timestamp', 'session', 'year', 'level', 'stress', 'anxiety', 'workload', 'overwhelm'];
  for (let i = 1; i <= 10; i++) h.push('pss' + i);
  for (let i = 1; i <= 7; i++) h.push('gad' + i);
  for (let i = 1; i <= 4; i++) h.push('work' + i);
  return sheet_('Responses', h);
}
function events_() {
  return sheet_('Events', ['timestamp', 'session', 'type', 'resource']);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function band_(area, s) {
  if (area === 'stress') return s <= 13 ? 'low' : s <= 26 ? 'mod' : 'high';
  if (area === 'anxiety') return s <= 4 ? 'low' : s <= 9 ? 'mod' : 'high';
  if (area === 'workload') return s <= 5 ? 'low' : s <= 10 ? 'mod' : 'high';
  return s <= 4 ? 'low' : s <= 8 ? 'mod' : 'high'; // overwhelm
}

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false }); }
  const session = String(d.session || '').slice(0, 64);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (d.type === 'quiz') {
      const a = d.answers || {};
      for (const k in LEN) {
        if (!Array.isArray(a[k]) || a[k].length !== LEN[k] ||
            a[k].some(v => !Number.isInteger(v) || v < 0 || v > MAXV[k])) {
          return json_({ ok: false, error: 'bad answers' });
        }
      }
      // Recalculate scores here so the sheet never trusts the browser's numbers.
      const stress = a.pss.reduce((s, v, i) => s + (PSS_REVERSED.includes(i) ? 4 - v : v), 0);
      const anxiety = a.gad.reduce((s, v) => s + v, 0);
      const workload = a.work.reduce((s, v) => s + v, 0);
      const overwhelm = a.pss[1] + a.pss[5] + a.pss[9];
      const rank = { low: 0, mod: 1, high: 2 };
      const level = [band_('stress', stress), band_('anxiety', anxiety), band_('workload', workload), band_('overwhelm', overwhelm)]
        .reduce((m, l) => rank[l] > rank[m] ? l : m, 'low');
      const year = YEARS.includes(d.year) ? d.year : 'Other / prefer not to say';
      responses_().appendRow([new Date(), session, year, level, stress, anxiety, workload, overwhelm]
        .concat(a.pss, a.gad, a.work));
    } else if (['help_click', 'contact_click', 'article_click'].includes(d.type)) {
      events_().appendRow([new Date(), session, d.type, String(d.resource || '').slice(0, 40)]);
    }
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true });
}

function doGet() {
  const rows = responses_().getDataRange().getValues().slice(1);
  const ev = events_().getDataRange().getValues().slice(1);
  const helpSessions = new Set(ev.filter(r => r[2] === 'help_click' || r[2] === 'contact_click').map(r => r[1]));

  const levels = { low: 0, mod: 0, high: 0 };
  const years = {};
  YEARS.forEach(y => years[y] = 0);
  const sums = { stress: 0, anxiety: 0, workload: 0, overwhelm: 0 };
  let helped = 0;
  rows.forEach(r => {
    years[r[2]] = (years[r[2]] || 0) + 1;
    levels[r[3]] = (levels[r[3]] || 0) + 1;
    sums.stress += Number(r[4]); sums.anxiety += Number(r[5]);
    sums.workload += Number(r[6]); sums.overwhelm += Number(r[7]);
    if (helpSessions.has(r[1])) helped++;
  });

  const total = rows.length;
  const avg = {};
  for (const k in sums) avg[k] = total ? sums[k] / total : 0;
  return json_({ total: total, levels: levels, years: years, avg: avg, helpRate: total ? helped / total : 0 });
}
