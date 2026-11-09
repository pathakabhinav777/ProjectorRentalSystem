/**
 * Projector booking backend (Google Apps Script).
 *
 * Flow:
 *   1. Customer picks a night, pays ₹70 by UPI, enters the 12-digit UTR.
 *   2. Booking is saved as PENDING (the date shows as "pending" to others).
 *      You get an email with a "Review" link.
 *   3. You check the UTR in your UPI app, open the link, press Confirm
 *      (adds it to your Google Calendar) or Reject (frees the date).
 *   Pending bookings never expire.
 *
 * Deploy: Deploy > New deployment > Web app, Execute as: Me, Access: Anyone.
 * Run setup() once from the editor first (see SETUP.md).
 */

// ---------------- Settings ----------------
const TZ = 'Asia/Kolkata';
const PRICE = 70;
const PICKUP_LOCATIONS = ['Block 9 - Room 467', 'Block 9 - Room 149'];
const MAX_DAYS_AHEAD = 60;      // how far ahead people can book
const MAX_PENDING = 10;         // stop new bookings if this many are waiting for you
const MAX_POSTS_PER_10_MIN = 20; // crude site-wide spam throttle
const NIGHT_START = '18:00';    // calendar event start
const NIGHT_END = '06:00';      // calendar event end (next morning)
// ------------------------------------------

const SHEET_NAME = 'Bookings';
const HEADERS = ['ID', 'Created', 'Status', 'Date', 'Name', 'Phone', 'Block', 'Room',
                 'Pickup', 'UTR', 'Amount', 'CalendarEventId'];
const COL = {}; HEADERS.forEach((h, i) => { COL[h] = i; });

/** Run once from the editor. Creates the sheet and the secret admin key. */
function setup() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('ADMIN_KEY')) {
    props.setProperty('ADMIN_KEY', (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, ''));
  }
  const sheet = getSheet_();
  // Touch these so Google asks for permission now, not on first booking.
  CalendarApp.getDefaultCalendar();
  MailApp.getRemainingDailyQuota();
  Logger.log('Setup done. Bookings sheet: ' + sheet.getParent().getUrl());
}

// ================= Public endpoints =================

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'review') return reviewPage_(p.id, p.key);
  return json_({ ok: true, price: PRICE, dates: takenDates_() });
}

function doPost(e) {
  try {
    const p = (e && e.parameter) || {};

    // Honeypot: real users never fill this hidden field.
    if (p.website) return json_({ ok: true });

    if (!throttle_()) return json_({ ok: false, error: 'Too many requests right now. Try again in a few minutes.' });

    const b = validate_(p);
    if (b.error) return json_({ ok: false, error: b.error });

    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      const sheet = getSheet_();
      const rows = readRows_(sheet);
      let pending = 0;
      for (const r of rows) {
        const status = r[COL.Status];
        const active = status === 'PENDING' || status === 'CONFIRMED';
        if (active && toDateStr_(r[COL.Date]) === b.date) {
          return json_({ ok: false, error: 'That night is already taken. Please pick another date.' });
        }
        if (String(r[COL.UTR]) === b.utr) {
          return json_({ ok: false, error: 'This UTR has already been used for a booking.' });
        }
        if (status === 'PENDING') {
          pending++;
          if (String(r[COL.Phone]) === b.phone) {
            return json_({ ok: false, error: 'You already have a booking waiting for confirmation.' });
          }
        }
      }
      if (pending >= MAX_PENDING) {
        return json_({ ok: false, error: 'Too many bookings are waiting for confirmation. Please try again later.' });
      }

      const id = Utilities.getUuid().replace(/-/g, '').slice(0, 12);
      sheet.appendRow([id, new Date(), 'PENDING', b.date, b.name, b.phone, b.block, b.room,
                       b.pickup, b.utr, PRICE, '']);
      SpreadsheetApp.flush();
      notifyOwner_(id, b);
      return json_({ ok: true, id: id });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'Something went wrong. Please try again.' });
  }
}

// ================= Admin actions (called from the review page) =================

function confirmBooking(id, key) {
  checkKey_(key);
  return withRow_(id, (sheet, rowNum, r) => {
    if (r[COL.Status] === 'CONFIRMED') return 'Already confirmed.';
    if (r[COL.Status] !== 'PENDING') return 'This booking was ' + r[COL.Status].toLowerCase() + ', so it cannot be confirmed.';
    const date = toDateStr_(r[COL.Date]);
    const start = new Date(date + 'T' + NIGHT_START + ':00+05:30');
    const end = new Date(addDays_(date, 1) + 'T' + NIGHT_END + ':00+05:30');
    const ev = CalendarApp.getDefaultCalendar().createEvent(
      'Projector: ' + r[COL.Name] + ' (' + r[COL.Block] + ', ' + r[COL.Room] + ')',
      start, end,
      { description: 'Phone: ' + r[COL.Phone] + '\nPickup: ' + r[COL.Pickup] +
                     '\nUTR: ' + r[COL.UTR] + '\nAmount: ₹' + r[COL.Amount] + '\nBooking ID: ' + id });
    sheet.getRange(rowNum, COL.Status + 1).setValue('CONFIRMED');
    sheet.getRange(rowNum, COL.CalendarEventId + 1).setValue(ev.getId());
    return 'Confirmed and added to your calendar.';
  });
}

function rejectBooking(id, key) {
  checkKey_(key);
  return withRow_(id, (sheet, rowNum, r) => {
    if (r[COL.Status] === 'REJECTED') return 'Already rejected.';
    const evId = r[COL.CalendarEventId];
    if (evId) {
      const ev = CalendarApp.getEventById(evId);
      if (ev) ev.deleteEvent();
    }
    sheet.getRange(rowNum, COL.Status + 1).setValue('REJECTED');
    return 'Rejected. The date is free again.';
  });
}

// ================= Helpers =================

function validate_(p) {
  const s = k => String(p[k] || '').trim();
  const b = {
    name: s('name').replace(/\s+/g, ' '),
    phone: s('phone').replace(/[\s-]/g, '').replace(/^(\+?91)(\d{10})$/, '$2'),
    block: s('block').replace(/\s+/g, ' '),
    room: s('room').replace(/\s+/g, ' '),
    pickup: s('pickup'),
    date: s('date'),
    utr: s('utr').replace(/\s/g, ''),
  };
  // Strict whitelists: no HTML, no spreadsheet formulas, no links.
  if (!/^[\p{L}][\p{L} .'-]{1,49}$/u.test(b.name)) return { error: 'Please enter your name (letters only).' };
  if (!/^[6-9]\d{9}$/.test(b.phone)) return { error: 'Please enter a valid 10-digit mobile number.' };
  if (!/^[\p{L}\p{N}][\p{L}\p{N} \/-]{0,19}$/u.test(b.block)) return { error: 'Please enter your block.' };
  if (!/^[\p{L}\p{N}][\p{L}\p{N} \/-]{0,19}$/u.test(b.room)) return { error: 'Please enter your room number.' };
  if (PICKUP_LOCATIONS.indexOf(b.pickup) === -1) return { error: 'Please choose a pickup location.' };
  if (!/^\d{12}$/.test(b.utr)) return { error: 'The UPI transaction ID (UTR) must be 12 digits.' };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(b.date) || toDateStr_(new Date(b.date + 'T12:00:00+05:30')) !== b.date) {
    return { error: 'Please choose a valid date.' };
  }
  const today = todayStr_();
  if (b.date < today) return { error: 'That date is in the past.' };
  if (b.date > addDays_(today, MAX_DAYS_AHEAD)) return { error: 'You can book up to ' + MAX_DAYS_AHEAD + ' days ahead.' };
  return b;
}

/** Only dates and status go public: no names, phones or anything people typed. */
function takenDates_() {
  const today = todayStr_();
  const out = [];
  for (const r of readRows_(getSheet_())) {
    const st = r[COL.Status];
    const d = toDateStr_(r[COL.Date]);
    if ((st === 'PENDING' || st === 'CONFIRMED') && d >= today) {
      out.push({ date: d, status: st === 'CONFIRMED' ? 'booked' : 'pending' });
    }
  }
  out.sort((a, b) => (a.date < b.date ? -1 : 1));
  return out;
}

function reviewPage_(id, key) {
  const page = body => HtmlService.createHtmlOutput(
    '<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<div style="font-family:system-ui,sans-serif;max-width:460px;margin:24px auto;padding:0 16px">' + body + '</div>'
  ).setTitle('Review booking');

  if (!keyOk_(key)) return page('<p>Invalid link.</p>');
  const found = findRow_(getSheet_(), id);
  if (!found) return page('<p>Booking not found.</p>');
  const r = found.row;
  const e = esc_;
  const rows = [
    ['Status', r[COL.Status]], ['Night of', prettyDate_(toDateStr_(r[COL.Date]))],
    ['Name', r[COL.Name]], ['Phone', r[COL.Phone]], ['Room', r[COL.Block] + ', ' + r[COL.Room]],
    ['Pickup', r[COL.Pickup]], ['UTR', r[COL.UTR]], ['Amount', '₹' + r[COL.Amount]],
  ].map(x => '<tr><td style="color:#666;padding:4px 12px 4px 0">' + e(x[0]) + '</td><td><b>' + e(x[1]) + '</b></td></tr>').join('');

  const wa = 'https://wa.me/91' + encodeURIComponent(r[COL.Phone]);
  const btn = 'padding:12px 18px;border:0;border-radius:10px;font-size:16px;color:#fff;margin-right:8px;cursor:pointer';
  return page(
    '<h2>Projector booking</h2><table>' + rows + '</table>' +
    '<p style="color:#666;font-size:14px">Check that UTR <b>' + e(r[COL.UTR]) + '</b> for ₹' + e(r[COL.Amount]) +
    ' is in your UPI app before confirming.</p>' +
    '<p><button id="c" style="' + btn + ';background:#15803d">Confirm</button>' +
    '<button id="r" style="' + btn + ';background:#b91c1c">Reject / cancel</button></p>' +
    '<p id="msg" style="font-weight:600"></p>' +
    '<p><a href="' + e(wa) + '" target="_blank">Message them on WhatsApp</a></p>' +
    '<script>' +
    'var ID=' + JSON.stringify(String(id)) + ',KEY=' + JSON.stringify(String(key)) + ';' +
    'function go(fn,label){if(fn==="rejectBooking"&&!confirm("Reject this booking?"))return;' +
    'document.getElementById("msg").textContent=label+"...";' +
    'google.script.run.withSuccessHandler(function(m){document.getElementById("msg").textContent=m;})' +
    '.withFailureHandler(function(err){document.getElementById("msg").textContent="Error: "+err.message;})[fn](ID,KEY);}' +
    'document.getElementById("c").onclick=function(){go("confirmBooking","Confirming")};' +
    'document.getElementById("r").onclick=function(){go("rejectBooking","Rejecting")};' +
    '</script>'
  );
}

function notifyOwner_(id, b) {
  const owner = Session.getEffectiveUser().getEmail();
  const key = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  const link = ScriptApp.getService().getUrl() + '?action=review&id=' + id + '&key=' + key;
  MailApp.sendEmail(owner,
    'Projector booking to check: ' + prettyDate_(b.date) + ' (' + b.name + ')',
    'New booking waiting for your confirmation.\n\n' +
    'Night of: ' + prettyDate_(b.date) + '\nName: ' + b.name + '\nPhone: ' + b.phone +
    '\nRoom: ' + b.block + ', ' + b.room + '\nPickup: ' + b.pickup +
    '\nUTR: ' + b.utr + '\nAmount: ₹' + PRICE +
    '\n\nCheck the UTR in your UPI app, then open this link to confirm or reject:\n' + link +
    '\n\nDo not forward this email: the link lets anyone confirm or reject bookings.');
}

function withRow_(id, fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = getSheet_();
    const found = findRow_(sheet, id);
    if (!found) return 'Booking not found.';
    return fn(sheet, found.rowNum, found.row);
  } finally {
    lock.releaseLock();
  }
}

function findRow_(sheet, id) {
  if (!id) return null;
  const rows = readRows_(sheet);
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][COL.ID]) === String(id)) return { rowNum: i + 2, row: rows[i] };
  }
  return null;
}

function readRows_(sheet) {
  const n = sheet.getLastRow() - 1;
  if (n < 1) return [];
  return sheet.getRange(2, 1, n, HEADERS.length).getValues();
}

function getSheet_() {
  const props = PropertiesService.getScriptProperties();
  let ss;
  const sheetId = props.getProperty('SHEET_ID');
  if (sheetId) {
    ss = SpreadsheetApp.openById(sheetId);
  } else {
    ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.create('Projector Bookings');
    props.setProperty('SHEET_ID', ss.getId());
  }
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    // Keep dates, phones and UTRs as plain text so Sheets doesn't reformat them.
    sh.getRange(2, COL.Date + 1, sh.getMaxRows() - 1, 1).setNumberFormat('@');
    sh.getRange(2, COL.Phone + 1, sh.getMaxRows() - 1, 1).setNumberFormat('@');
    sh.getRange(2, COL.UTR + 1, sh.getMaxRows() - 1, 1).setNumberFormat('@');
  }
  return sh;
}

function throttle_() {
  const cache = CacheService.getScriptCache();
  const n = Number(cache.get('posts') || 0);
  if (n >= MAX_POSTS_PER_10_MIN) return false;
  cache.put('posts', String(n + 1), 600);
  return true;
}

function keyOk_(key) {
  const real = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  return !!real && typeof key === 'string' && key === real;
}

function checkKey_(key) {
  if (!keyOk_(key)) throw new Error('Not allowed.');
}

function toDateStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v || '').trim();
}

function todayStr_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
}

function addDays_(dateStr, days) {
  const d = new Date(dateStr + 'T12:00:00+05:30');
  d.setTime(d.getTime() + days * 86400000);
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}

function prettyDate_(dateStr) {
  return Utilities.formatDate(new Date(dateStr + 'T12:00:00+05:30'), TZ, 'EEE, d MMM yyyy');
}

function esc_(s) {
  return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
