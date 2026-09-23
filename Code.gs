'use strict';

const SOURCE_ID = '1a-v3gtk5KOUcyfFlz6hk2qGPTapiwkCfilEoQUXYijY';
const CACHE_KEY = 'duty-feed-v1';

const DutyParser = (() => {
  const clean = value => String(value ?? '').replace(/\s+/g, ' ').trim();

  const months = {
    jan: 1, feb: 2, mar: 3, apr: 4,
    may: 5, jun: 6, jul: 7, aug: 8,
    sep: 9, oct: 10, nov: 11, dec: 12
  };

  function dateOf(value) {
    const match = clean(value).match(
      /^(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]+),?\s+(\d{4})$/
    );

    if (!match) return null;

    const month = months[match[2].slice(0, 3).toLowerCase()];
    if (!month) return null;

    const day = Number(match[1]);
    const year = Number(match[3]);
    const date = new Date(Date.UTC(year, month - 1, day));

    if (
      date.getUTCDate() !== day ||
      date.getUTCMonth() !== month - 1
    ) {
      return null;
    }

    return (
      year + '-' +
      String(month).padStart(2, '0') + '-' +
      String(day).padStart(2, '0')
    );
  }

  function column(index) {
    let result = '';

    for (index++; index; index = Math.floor((index - 1) / 26)) {
      result = String.fromCharCode(65 + (index - 1) % 26) + result;
    }

    return result;
  }

  function parse(sheets) {
    const master = sheets.find(sheet => sheet.title === 'Master Sheet');

    if (!master) {
      throw new Error('The Master Sheet tab is missing.');
    }

    const roster = new Map();
    let department = '';

    for (const row of master.values) {
      const name = clean(row[1]);
      const code = clean(row[2]);

      if (name.startsWith('DEPARTMENT OF ')) {
        department = name.slice(14).trim();
      }

      if (code && code !== 'Code') {
        if (roster.has(code) || !department) {
          throw new Error('The master roster layout has changed.');
        }

        roster.set(code, {
          code,
          name,
          department,
          category: clean(row[3])
        });
      }
    }

    if (!roster.size) {
      throw new Error('No staff roster was found.');
    }

    const duties = [];
    const dates = [];
    const notes = [];

    for (const sheet of sheets) {
      const date = dateOf(sheet.title);
      if (!date) continue;

      dates.push(date);

      let building = '';
      const rows = sheet.values;

      for (let r = 0; r < rows.length; r++) {
        const label = clean(rows[r]?.[1]);
        const buildingMatch = label.match(/Building\s*-\s*(\d+)/i);

        if (buildingMatch) {
          building = buildingMatch[1];
        }

        const sessionMatch = label.match(/^(S-\d+)\s+(.+)$/);
        if (!sessionMatch) continue;

        const teachers = [];

        for (
          let k = r + 1;
          k < rows.length &&
          /^Teacher\s+\d+$/i.test(clean(rows[k]?.[1]));
          k++
        ) {
          teachers.push(k);
        }

        const width = Math.max(
          0,
          ...teachers.map(k => rows[k].length)
        );

        for (let c = 2; c < width; c++) {
          const people = teachers
            .map(k => ({
              code: clean(rows[k][c]),
              cell: column(c) + (k + 1)
            }))
            .filter(person => person.code);

          if (!people.length) continue;

          const room = clean(rows[r - 1]?.[c]);

          if (!room || !building) {
            throw new Error(
              'An assigned room or building could not be identified.'
            );
          }

          let time = sessionMatch[2].replace(/\s*[-–]\s*/g, ' – ');

          if (
            new Date(date + 'T12:00:00Z').getUTCDay() === 5 &&
            sessionMatch[1] === 'S-1'
          ) {
            time = '10:00 AM – 12:00 PM';
          }

          for (const person of people) {
            if (!roster.has(person.code)) {
              const inferredDepartment =
                person.code.match(/\(([^)]+)\)$/)?.[1]?.toUpperCase() ||
                'Unlisted';

              roster.set(person.code, {
                code: person.code,
                name: person.code,
                department: inferredDepartment,
                category: 'Unlisted'
              });

              notes.push(
                person.code +
                ': assigned in the dated schedule but missing from the master roster.'
              );
            }

            duties.push({
              faculty: person.code,
              department: roster.get(person.code).department,
              date,
              slot: sessionMatch[1],
              time,
              room,
              building,
              coInvigilators: people
                .filter(other => other.cell !== person.cell)
                .map(other => other.code),
              source: sheet.title + '!' + person.cell
            });
          }
        }
      }
    }

    if (!dates.length) {
      throw new Error('No dated schedule tabs were found.');
    }

    const counts = new Map();

    for (const duty of duties) {
      const key = duty.faculty + '|' + duty.date;
      counts.set(key, (counts.get(key) || 0) + 1);
    }

    const headers =
      master.values.find(row => clean(row[2]) === 'Code') || [];

    const discrepancies = [];

    for (const row of master.values) {
      const code = clean(row[2]);
      if (!roster.has(code)) continue;

      for (let c = 5; c < headers.length; c++) {
        const date = dateOf(headers[c]);

        if (!date || !dates.includes(date)) continue;

        const raw = clean(row[c]);

        if (raw && !/^\d+(?:\.0+)?$/.test(raw)) {
          notes.push(
            code + ': summary count for ' + date + ' could not be checked.'
          );
          continue;
        }

        const expected = Number(raw || 0);
        const actual = counts.get(code + '|' + date) || 0;

        if (expected !== actual) {
          discrepancies.push([code, date, expected, actual]);
        }
      }
    }

    duties.sort((a, b) =>
      a.date.localeCompare(b.date) ||
      a.slot.localeCompare(b.slot) ||
      a.faculty.localeCompare(b.faculty)
    );

    return {
      faculty: [...roster.values()],
      dates: dates.sort(),
      duties,
      discrepancies,
      notes
    };
  }

  return { dateOf, parse };
})();

function loadSchedule_() {
  const book = SpreadsheetApp.openById(SOURCE_ID);

  const sheets = book.getSheets().filter(sheet =>
    sheet.getName() === 'Master Sheet' ||
    DutyParser.dateOf(sheet.getName())
  );

  const parsed = DutyParser.parse(
    sheets.map(sheet => {
      const rowCount = sheet.getLastRow();

      return {
        title: sheet.getName(),
        values: rowCount
          ? sheet.getRange(
              1,
              1,
              rowCount,
              Math.min(26, sheet.getMaxColumns())
            ).getDisplayValues()
          : []
      };
    })
  );

  parsed.faculty = parsed.faculty.map(teacher => ({
    code: teacher.code,
    name: teacher.name,
    department: teacher.department
  }));

  return {
    ok: true,
    schemaVersion: 1,
    fetchedAt: new Date().toISOString(),
    data: parsed
  };
}

function readScheduleCache_() {
  try {
    const text = CacheService.getScriptCache().get(CACHE_KEY);
    if (!text) return null;

    const saved = JSON.parse(
      Utilities.ungzip(
        Utilities.newBlob(Utilities.base64Decode(text))
      ).getDataAsString()
    );

    const age = Date.now() - Date.parse(saved.fetchedAt);

    return saved.ok && age >= 0 && age < 60000
      ? saved
      : null;
  } catch (error) {
    return null;
  }
}

function cachedSchedule_() {
  let saved = readScheduleCache_();
  if (saved) return saved;

  const lock = LockService.getScriptLock();

  if (!lock.tryLock(10000)) {
    throw new Error('Schedule refresh is busy.');
  }

  try {
    saved = readScheduleCache_();
    if (saved) return saved;

    const fresh = loadSchedule_();

    const compressed = Utilities.base64Encode(
      Utilities.gzip(
        Utilities.newBlob(JSON.stringify(fresh), 'application/json')
      ).getBytes()
    );

    if (compressed.length < 95000) {
      try {
        CacheService.getScriptCache().put(
          CACHE_KEY,
          compressed,
          60
        );
      } catch (error) {
        // The fresh schedule remains usable if caching fails.
      }
    }

    return fresh;
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  const saved = readScheduleCache_();

  const json = JSON.stringify(saved)
    .replace(/</g, '\\u003c')
    .replace(/\u2028/g, '\\u2028')
    .replace(/\u2029/g, '\\u2029');

  const page = PAGE_HTML_.replace(
    '/*DUTY_BOOT*/null',
    () => json
  );

  return HtmlService.createHtmlOutput(page)
    .setTitle('Exam Duty Finder')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getDuties() {
  try {
    return cachedSchedule_();
  } catch (error) {
    console.error(String(error));

    throw new Error(
      'Unable to read the schedule. Please try Refresh duties. ' +
      'If this continues, ask the administrator to run setup again.'
    );
  }
}

function setup() {
  const result = loadSchedule_();

  console.log(
    'READY: ' +
    result.data.faculty.length + ' faculty/staff, ' +
    result.data.duties.length + ' assignments.'
  );
}

const PAGE_HTML_ = [
  '<!doctype html>',
  '<html lang="en">',
  '<head>',
  '<meta charset="utf-8">',
  '<meta name="viewport" content="width=device-width, initial-scale=1">',
  '<title>Exam Duty Finder | University of Scholars</title>',
  '<base target="_top">',
  '<style>',
  ':root{font-family:system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;color:#152b3e;background:#f3f6fa;font-size:16px;line-height:1.5}',
  '*{box-sizing:border-box}',
  'body{margin:0}',
  'header{background:#102b43;color:white;padding:22px max(5vw,20px);border-bottom:4px solid #18b4bd;display:flex;justify-content:space-between;align-items:center;gap:20px}',
  '.brand{font-weight:700}.brand small{display:block;font-size:14px;font-weight:400;color:#c2d2e0}',
  '.term{font-size:13px;letter-spacing:1px}',
  'main{max-width:1280px;margin:auto;padding:32px 24px}',
  '.title-row{display:flex;justify-content:space-between;align-items:center;gap:20px;margin-bottom:24px}',
  'h1{font-size:clamp(28px,4vw,42px);line-height:1.2;margin:0}',
  '.intro{color:#54697b}',
  'button,select{font:inherit;min-height:46px;border:1px solid #c5d1db;border-radius:6px;background:white;color:#152b3e;padding:10px 14px}',
  'button{cursor:pointer;font-weight:600}',
  'button:hover{background:#e8f3f5}',
  '.primary{background:#102b43;color:white;border-color:#102b43}',
  '.primary:hover{background:#214a69}',
  'button:disabled{opacity:.5;cursor:default}',
  'button:focus-visible,select:focus-visible,summary:focus-visible{outline:3px solid #008c98;outline-offset:3px}',
  '.connection{background:white;border:1px solid #d9e1e9;border-radius:8px;padding:18px;margin-bottom:20px}',
  '.connection p{margin:10px 0 0;font-size:14px;color:#536779}',
  '#sync-status{font-weight:650}.sync-error{color:#8a4c00!important}',
  '.filters{background:white;border:1px solid #d9e1e9;padding:20px;display:grid;grid-template-columns:1fr 2fr 1.2fr 1fr auto;gap:16px;align-items:end;border-radius:10px}',
  'label{display:flex;flex-direction:column;gap:8px;font-size:14px;font-weight:650;min-width:0}',
  'select{width:100%;min-width:0;font-weight:400}',
  '.help{font-size:13px;color:#607385}',
  '.summary{display:flex;justify-content:space-between;gap:12px;margin:24px 0 16px;font-size:14px;color:#54697b}',
  '#count{font-weight:650;color:#152b3e}',
  'aside{border-left:4px solid #bd7c13;background:#fff4d9;padding:14px 18px;margin-bottom:22px;font-size:14px;color:#634410}',
  '#show-all{margin-bottom:16px}',
  '.day{margin-bottom:28px}',
  '.day-header{display:flex;align-items:baseline;gap:14px;margin-bottom:10px}',
  '.day h2{font-size:20px;margin:0}',
  '.day-header span{font-size:14px;color:#54697b}',
  '.table-wrap{border:1px solid #d9e1e9;border-radius:8px;overflow:hidden;background:white}',
  'table{border-collapse:collapse;width:100%;text-align:left}',
  'th{font-size:13px;text-transform:uppercase;background:#eaf0f5;color:#536779;padding:12px 18px}',
  'td{padding:18px;vertical-align:top;border-top:1px solid #e4eaf0}',
  'td small{display:block;font-size:13px;color:#607385;margin-top:4px}',
  '.staff{font-weight:650}.room{font-weight:750;font-size:20px}',
  '.session{white-space:nowrap;font-size:14px}.team{font-size:14px;line-height:1.8}',
  '.conflict{color:#8a4c00!important;font-weight:650}',
  '.empty{padding:40px 24px;text-align:center;background:white;border:1px solid #d9e1e9;border-radius:8px;margin-bottom:24px}',
  '.empty h2{font-size:22px}.empty p{color:#54697b}',
  'footer{border-top:1px solid #d3dde7;padding-top:22px;font-size:14px;color:#607385}',
  'summary{cursor:pointer;color:#214a69}',
  '.print-selection{display:none}',
  '[hidden]{display:none!important}',
  '@media(max-width:850px){.filters{grid-template-columns:1fr 1fr}.term{display:none}main{padding:24px 16px}.summary{flex-direction:column}.session{white-space:normal}}',
  '@media(max-width:600px){.title-row{flex-direction:column;align-items:stretch}.filters{grid-template-columns:1fr}.table-wrap{border:0;background:transparent;overflow:visible}table,tbody{display:block}thead{display:none}tr{display:grid;grid-template-columns:1fr 1fr;border:1px solid #d9e1e9;border-radius:8px;background:white;margin-bottom:12px;padding:6px}td{display:block;border:0;padding:10px}td:before{content:attr(data-label);display:block;text-transform:uppercase;font-size:12px;color:#607385;margin-bottom:4px}td:first-child,td:last-child{grid-column:1/-1}.day-header{flex-wrap:wrap}}',
  '@media print{@page{size:A4 landscape;margin:12mm}body{background:white;color:black}header{background:white;color:black;border-bottom:2px solid black;padding:0 0 12px}.brand small{color:#333}main{padding:18px 0;max-width:none}h1{font-size:25px}.intro,.filters,button,.help,footer details{display:none!important}.print-selection{display:block;font-weight:650}.table-wrap{overflow:visible;border-radius:0}table{display:table}tbody{display:table-row-group}thead{display:table-header-group}tr{display:table-row;break-inside:avoid}th,td{display:table-cell;padding:8px;font-size:12px}td:before{display:none}td small,.team,.session{font-size:11px}.room{font-size:15px}.day-header{break-after:avoid}footer{font-size:10px}}',
  '</style>',
  '</head>',
  '<body>',
  '<header><div class="brand">University of Scholars<small>Exam Duty Finder · All departments</small></div><span class="term">SPRING 2026 / SET A</span></header>',
  '<main>',
  '<div class="title-row"><div><h1>Find your exam duties.</h1><p class="intro">Your date, room and invigilation team, in one place.</p></div><button id="print" class="primary" disabled>Print duties</button></div>',
  '<section class="connection" aria-label="Schedule updates">',
  '<button id="refresh" class="primary">Refresh duties</button>',
  '<p id="sync-status" role="status" aria-live="polite">Loading the latest schedule…</p>',
  '<p id="last-sync">Not synced yet</p>',
  '<p class="help">Updates automatically every 2 minutes while this tab is visible. No sign-in required.</p>',
  '</section>',
  '<section class="filters" aria-label="Filter duties">',
  '<label>Department<select id="department"><option value="">All departments</option></select></label>',
  '<label>Faculty / staff<select id="faculty"><option value="">All faculty / staff</option></select></label>',
  '<label>Exam date<select id="date"><option value="">All dates</option></select></label>',
  '<label>Session<select id="slot"><option value="">All sessions</option><option value="S-1">Morning · Slot 1</option><option value="S-2">Afternoon · Slot 2</option></select></label>',
  '<button id="reset">Clear filters</button>',
  '</section>',
  '<p class="help">Teacher counts show past/total assigned duties across the full schedule. Past means dates before today in Bangladesh, not confirmed attendance. Each date and session counts once, even if multiple rooms are listed.</p>',
  '<div class="summary"><span id="count" aria-live="polite"></span><span>All times are Bangladesh time (UTC+6)</span></div>',
  '<p id="selection" class="print-selection"></p>',
  '<button id="show-all" hidden>Show all duties</button>',
  '<aside id="notice" hidden></aside>',
  '<div id="results"><section class="empty"><h2>Loading duties</h2><p>The latest published schedule will appear here.</p></section></div>',
  '<noscript>Please enable JavaScript to display and filter the duties.</noscript>',
  '<footer>',
  '<strong>Spring 2026 · Final examinations · Set A</strong>',
  '<p>Live source: the university’s main Google Sheet. Friday morning: 10:00 AM–12:00 PM. The sync time shows when the source was read, not when it was edited. Follow any revised notice from the exam office.</p>',
  '<details><summary>Data checks &amp; source notes</summary><p>Assignments come from the dated room schedules. Co-invigilators share the same room and session.</p><div id="checks"></div></details>',
  '<details><summary>About this schedule</summary><p>The original spreadsheet remains restricted. Phone numbers, email addresses and edit history are not included. Chief and co-chief supervisory headings are not counted as room assignments.</p></details>',
  '</footer>',
  '</main>',
  '<script>',
  '(' + dutyFinderClient_.toString() + ')(/*DUTY_BOOT*/null);',
  '</script>',
  '</body>',
  '</html>'
].join('\n');

function dutyFinderClient_(initialPayload) {
  'use strict';

  const $ = id => document.getElementById(id);

  let data = null;
  let names = new Map();
  let departments = new Map();
  let clashes = new Map();
  let busy = false;
  let showAll = false;

  const dateFormatter = new Intl.DateTimeFormat('en-GB', {
    day: 'numeric',
    month: 'long',
    year: 'numeric',
    timeZone: 'Asia/Dhaka'
  });

  const weekdayFormatter = new Intl.DateTimeFormat('en-GB', {
    weekday: 'long',
    timeZone: 'Asia/Dhaka'
  });

  const stampFormatter = new Intl.DateTimeFormat('en-GB', {
    dateStyle: 'medium',
    timeStyle: 'medium',
    timeZone: 'Asia/Dhaka'
  });

  const fullDate = date =>
    dateFormatter.format(new Date(date + 'T12:00:00+06:00'));

  const dayName = date =>
    weekdayFormatter.format(new Date(date + 'T12:00:00+06:00'));

  const escapeHtml = value => String(value).replace(
    /[&<>"']/g,
    character => ({
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#39;'
    })[character]
  );

  function options(id, items, placeholder, wanted) {
    const element = $(id);
    element.innerHTML = '';
    element.add(new Option(placeholder, ''));

    for (const [label, value] of items) {
      element.add(new Option(label, value));
    }

    element.value = [...element.options].some(
      option => option.value === wanted
    ) ? wanted : '';
  }

  function dutyCounts(duties) {
    const today = new Date(Date.now() + 21600000)
      .toISOString()
      .slice(0, 10);

    const counts = new Map();
    const seen = new Set();

    for (const duty of duties) {
      const key = JSON.stringify([
        duty.faculty,
        duty.date,
        duty.slot
      ]);

      if (seen.has(key)) continue;
      seen.add(key);

      const count = counts.get(duty.faculty) || {
        past: 0,
        total: 0
      };

      count.total++;

      if (duty.date < today) {
        count.past++;
      }

      counts.set(duty.faculty, count);
    }

    return counts;
  }

  function populateFaculty(wanted = $('faculty').value) {
    const counts = dutyCounts(data.duties);

    const teachers = [...data.faculty]
      .filter(teacher =>
        !$('department').value ||
        teacher.department === $('department').value
      )
      .sort((a, b) => a.name.localeCompare(b.name))
      .map(teacher => {
        const count = counts.get(teacher.code) || {
          past: 0,
          total: 0
        };

        return [
          teacher.name + ' (' + teacher.department + ') — ' +
          count.past + '/' + count.total + ' assigned',
          teacher.code
        ];
      });

    options('faculty', teachers, 'All faculty / staff', wanted);
  }

  function status(message, error = false) {
    $('sync-status').textContent = message;
    $('sync-status').classList.toggle('sync-error', error);
  }

  function enabled(ready) {
    for (const id of [
      'department', 'faculty', 'date', 'slot', 'reset', 'print'
    ]) {
      $(id).disabled = !ready;
    }
  }

  function filterDuties(faculty, date, slot) {
    const department = $('department').value;

    return data.duties.filter(duty =>
      (!department || duty.department === department) &&
      (!faculty || duty.faculty === faculty) &&
      (!date || duty.date === date) &&
      (!slot || duty.slot === slot)
    );
  }

  function render() {
    if (!data) return;

    const faculty = $('faculty').value;
    const date = $('date').value;
    const slot = $('slot').value;

    const rows = filterDuties(faculty, date, slot);

    $('count').textContent =
      rows.length + ' assignment' + (rows.length === 1 ? '' : 's') +
      ' · ' + new Set(rows.map(duty => duty.date)).size + ' exam dates';

    $('selection').textContent = [
      $('department').value || 'All departments',
      names.get(faculty) || 'All faculty / staff',
      date ? fullDate(date) : 'All dates',
      slot || 'All sessions'
    ].join(' · ');

    const warnings = data.discrepancies.filter(entry =>
      (
        !$('department').value ||
        departments.get(entry[0]) === $('department').value
      ) &&
      (!faculty || entry[0] === faculty) &&
      (!date || entry[1] === date)
    );

    $('notice').hidden = !warnings.length;

    $('notice').textContent = warnings.length
      ? warnings.length +
        ' master-summary mismatch(es) affect this selection. ' +
        'Detailed room assignments are shown; see Data checks below ' +
        'and confirm with the exam office.'
      : '';

    const waiting =
      !showAll &&
      !faculty &&
      !date &&
      !slot &&
      !$('department').value;

    $('show-all').hidden = !waiting;
    $('print').disabled = waiting;

    if (waiting) {
      $('results').innerHTML =
        '<section class="empty">' +
        '<h2>Select your department or name</h2>' +
        '<p>Your duties will appear here. To view the complete ' +
        'schedule, select Show all duties.</p></section>';
      return;
    }

    if (!rows.length) {
      $('results').innerHTML =
        '<section class="empty"><h2>No duties listed</h2>' +
        '<p>No assignments match these filters. Try another date ' +
        'or clear the filters.</p></section>';
      return;
    }

    const groups = new Map();

    for (const duty of rows) {
      if (!groups.has(duty.date)) {
        groups.set(duty.date, []);
      }

      groups.get(duty.date).push(duty);
    }

    $('results').innerHTML = [...groups].map(([date, items]) => {
      const tableRows = items.map(duty => {
        const conflict =
          clashes.get([
            duty.faculty,
            duty.date,
            duty.slot
          ].join('|')) > 1;

        const team = duty.coInvigilators.length
          ? duty.coInvigilators
              .map(code => escapeHtml(names.get(code) || code))
              .join('<br>')
          : 'No co-invigilator listed';

        return (
          '<tr>' +
          '<td data-label="Faculty / staff">' +
          '<span class="staff">' +
          escapeHtml(names.get(duty.faculty) || duty.faculty) +
          '</span><small>' + escapeHtml(duty.faculty) + '</small>' +
          (
            conflict
              ? '<small class="conflict">Multiple rooms in this ' +
                'session — confirm assignment</small>'
              : ''
          ) +
          '</td>' +
          '<td data-label="Time" class="session">' +
          escapeHtml(duty.time) +
          '<small>' +
          (duty.slot === 'S-1' ? 'Morning' : 'Afternoon') +
          ' · ' + escapeHtml(duty.slot) +
          '</small></td>' +
          '<td data-label="Room"><span class="room">' +
          escapeHtml(duty.room) +
          '</span><small>Building ' +
          escapeHtml(duty.building) +
          '</small></td>' +
          '<td data-label="Co-invigilators" class="team">' +
          team +
          '<small>Source: ' + escapeHtml(duty.source) +
          '</small></td>' +
          '</tr>'
        );
      }).join('');

      return (
        '<section class="day">' +
        '<div class="day-header"><h2>' +
        escapeHtml(fullDate(date)) +
        '</h2><span>' +
        escapeHtml(dayName(date)) +
        '</span></div>' +
        '<div class="table-wrap"><table><thead><tr>' +
        '<th scope="col">Faculty / staff</th>' +
        '<th scope="col">Time</th>' +
        '<th scope="col">Room</th>' +
        '<th scope="col">Co-invigilators</th>' +
        '</tr></thead><tbody>' +
        tableRows +
        '</tbody></table></div></section>'
      );
    }).join('');
  }

  function loadFeed() {
    if (initialPayload) {
      const payload = initialPayload;
      initialPayload = null;

      const age = Date.now() - Date.parse(payload.fetchedAt);

      if (payload.ok && age >= 0 && age < 60000) {
        return Promise.resolve(payload);
      }
    }

    return new Promise((resolve, reject) => {
      let settled = false;

      const timer = setTimeout(() => {
        settled = true;

        reject(new Error(
          'The schedule is taking too long to load. ' +
          'Please press Refresh duties.'
        ));
      }, 45000);

      function finish(callback, value) {
        if (settled) return;

        settled = true;
        clearTimeout(timer);
        callback(value);
      }

      google.script.run
        .withSuccessHandler(value => finish(resolve, value))
        .withFailureHandler(() => finish(
          reject,
          new Error(
            'Unable to read the schedule. Try Refresh duties; ' +
            'if it continues, ask the administrator to run setup again.'
          )
        ))
        .getDuties();
    });
  }

  function validPayload(payload) {
    if (!payload?.ok) {
      throw new Error(
        payload?.message || 'The schedule feed reported an error.'
      );
    }

    const result = payload.data;

    if (
      payload.schemaVersion !== 1 ||
      !Number.isFinite(Date.parse(payload.fetchedAt)) ||
      !result ||
      !['faculty', 'duties', 'dates', 'discrepancies', 'notes']
        .every(key => Array.isArray(result[key]))
    ) {
      throw new Error('The schedule feed returned an unsupported format.');
    }

    const validTeachers = result.faculty.every(teacher =>
      typeof teacher.code === 'string' &&
      typeof teacher.name === 'string' &&
      typeof teacher.department === 'string'
    );

    const validDuties = result.duties.every(duty =>
      [
        'faculty', 'department', 'date', 'slot',
        'time', 'room', 'building', 'source'
      ].every(key => typeof duty[key] === 'string') &&
      Array.isArray(duty.coInvigilators)
    );

    if (!validTeachers || !validDuties) {
      throw new Error(
        'The schedule is incomplete. Ask the administrator to check the importer.'
      );
    }

    return result;
  }

  async function sync() {
    if (busy) return;

    if (data) populateFaculty();

    busy = true;
    $('refresh').disabled = true;
    status('Checking the latest schedule…');

    try {
      const payload = await loadFeed();
      const next = validPayload(payload);

      const selected = Object.fromEntries(
        ['department', 'faculty', 'date', 'slot']
          .map(key => [key, $(key).value])
      );

      data = next;

      names = new Map(
        data.faculty.map(teacher => [teacher.code, teacher.name])
      );

      departments = new Map(
        data.faculty.map(teacher => [teacher.code, teacher.department])
      );

      clashes = new Map();

      for (const duty of data.duties) {
        const key = [
          duty.faculty,
          duty.date,
          duty.slot
        ].join('|');

        clashes.set(key, (clashes.get(key) || 0) + 1);
      }

      options(
        'department',
        [...new Set(data.faculty.map(teacher => teacher.department))]
          .sort()
          .map(department => [department, department]),
        'All departments',
        selected.department
      );

      populateFaculty(selected.faculty);

      options(
        'date',
        data.dates.map(date => [fullDate(date), date]),
        'All dates',
        selected.date
      );

      $('slot').value = ['', 'S-1', 'S-2'].includes(selected.slot)
        ? selected.slot
        : '';

      $('checks').innerHTML = data.discrepancies
        .map(([teacher, date, master, detailed]) =>
          '<p>' +
          escapeHtml(names.get(teacher) || teacher) +
          ' · ' + escapeHtml(fullDate(date)) +
          ': master summary ' + escapeHtml(master) +
          '; detailed schedule ' + escapeHtml(detailed) +
          '.</p>'
        )
        .join('') +
        data.notes.map(note =>
          '<p>' + escapeHtml(note) + '</p>'
        ).join('');

      if (!data.discrepancies.length && !data.notes.length) {
        $('checks').textContent =
          'No master-summary discrepancies detected.';
      }

      $('last-sync').textContent =
        'Last successfully synced from source: ' +
        stampFormatter.format(new Date(payload.fetchedAt)) +
        ' (Bangladesh time)';

      const old =
        Date.now() - Date.parse(payload.fetchedAt) > 5 * 60000;

      status(
        old
          ? 'The returned schedule is more than 5 minutes old. ' +
            'Confirm updates with the exam office.'
          : 'Schedule loaded. Checks every 2 minutes; ' +
            'the feed may cache results for up to 1 minute.',
        old
      );

      enabled(true);
      render();
    } catch (error) {
      status(
        error.message +
        (
          data
            ? ' Displaying the last successful sync; it may be outdated.'
            : ''
        ),
        true
      );

      if (!data) {
        $('results').innerHTML =
          '<section class="empty"><h2>Schedule unavailable</h2>' +
          '<p>Please check the update message above and try ' +
          'Refresh duties.</p></section>';
      }
    } finally {
      busy = false;
      $('refresh').disabled = false;
    }
  }

  for (const key of ['department', 'faculty', 'date', 'slot']) {
    $(key).addEventListener('change', () => {
      if (!data) return;

      if (key === 'department') {
        populateFaculty();
      }

      render();
    });
  }

  $('reset').addEventListener('click', () => {
    if (!data) return;

    showAll = false;

    for (const key of ['department', 'faculty', 'date', 'slot']) {
      $(key).value = '';
    }

    populateFaculty();
    render();
  });

  $('show-all').addEventListener('click', () => {
    showAll = true;
    render();
  });

  $('print').addEventListener('click', () => window.print());
  $('refresh').addEventListener('click', sync);

  document.addEventListener('visibilitychange', () => {
    if (!document.hidden) sync();
  });

  setInterval(() => {
    if (!document.hidden) sync();
  }, 120000);

  enabled(false);
  sync();
}
