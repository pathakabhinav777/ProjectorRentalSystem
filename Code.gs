/**
 * Projector booking backend (Google Apps Script).
 *
 * Flow:
 *   1. Customer picks a night (up to 7 days ahead), pays ₹70 by UPI, enters the 12-digit UTR.
 *   2. Booking is saved as PENDING (the night shows as "pending" to others).
 *      You get an email with a link to your owner panel.
 *   3. In the owner panel you Confirm (adds it to Google Calendar) or Reject (frees the night).
 *      You can also pause bookings for tonight, or block any night for someone who booked directly.
 *   Pending bookings never expire.
 *
 * Files in this Apps Script project: Code.gs (this file) and Admin.html (the owner panel).
 * Deploy: Deploy > New deployment > Web app, Execute as: Me, Access: Anyone.
 */

// ---------------- Settings ----------------
const TZ = 'Asia/Kolkata';
const PRICE = 70;
const PICKUP_LOCATIONS = ['Block 9 - Room 467', 'Block 9 - Room 149'];
const BOOKING_WINDOW_DAYS = 7;  // people can book from today up to 7 days ahead
const BLACKOUTS = [             // nights nobody can book (inclusive)
  ['2026-12-05', '2027-01-02'],
];
const MAX_PENDING = 10;          // stop new bookings if this many are waiting for you
const MAX_POSTS_PER_10_MIN = 20; // crude site-wide spam throttle
const NIGHT_START = '18:00';     // calendar event start
const NIGHT_END = '06:00';       // calendar event end (next morning)
// ------------------------------------------

// The live web-app URL (Deploy > Manage deployments). Used for links in emails.
const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycby7ke5FHUyWd0GXTvmwfGdvFm92RFA9yigrpfS6WIz794WGgIPdRNOZ3qXB4I1YV8_Jlg/exec';
const SHEET_NAME = 'Bookings';
const HEADERS = ['ID', 'Created', 'Status', 'Date', 'Name', 'Phone', 'Block', 'Room',
                 'Pickup', 'UTR', 'Amount', 'CalendarEventId'];
const COL = {}; HEADERS.forEach((h, i) => { COL[h] = i; });
const ACTIVE = ['PENDING', 'CONFIRMED', 'BLOCKED']; // statuses that hold a night

/** Run once from the editor. Creates the sheet and the secret admin key, and logs your owner-panel link. */
function setup() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('ADMIN_KEY')) {
    props.setProperty('ADMIN_KEY', (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, ''));
  }
  const sheet = getSheet_();
  CalendarApp.getDefaultCalendar();
  MailApp.getRemainingDailyQuota();
  Logger.log('Setup done. Bookings sheet: ' + sheet.getParent().getUrl());
  Logger.log('Owner panel: ' + adminLink_());
}

// ================= Public endpoints =================

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'admin' || p.action === 'review') return adminPage_(p.key);
  return json_({
    ok: true, price: PRICE, today: todayStr_(), windowDays: BOOKING_WINDOW_DAYS,
    blackouts: BLACKOUTS, dates: takenDates_(),
  });
}

function doPost(e) {
  try {
    const p = (e && e.parameter) || {};
    if (p.website) return json_({ ok: true }); // honeypot
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
        if (ACTIVE.indexOf(status) !== -1 && toDateStr_(r[COL.Date]) === b.date) {
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

      const id = newId_();
      sheet.appendRow([id, new Date(), 'PENDING', b.date, b.name, b.phone, b.block, b.room,
                       b.pickup, b.utr, PRICE, '']);
      SpreadsheetApp.flush();
      notifyOwner_(b);
      return json_({ ok: true, id: id });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'Something went wrong. Please try again.' });
  }
}

// ================= Owner panel (called from Admin.html) =================

function adminLogin(pw) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('login_fail') || 0);
  if (fails >= 5) throw new Error('Too many wrong attempts. Try again in 15 minutes.');
  const props = PropertiesService.getScriptProperties();
  const hash = props.getProperty('ADMIN_PW_HASH');
  if (!hash) throw new Error('No password set yet. Open the owner link from one of your booking emails and set one.');
  if (hashPw_(String(pw || ''), props.getProperty('ADMIN_PW_SALT')) !== hash) {
    cache.put('login_fail', String(fails + 1), 900);
    throw new Error('Wrong password.');
  }
  return props.getProperty('ADMIN_KEY');
}

function adminSetPassword(key, pw) {
  checkKey_(key);
  pw = String(pw || '');
  if (pw.length < 6) throw new Error('Use at least 6 characters.');
  const salt = Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperties({ ADMIN_PW_SALT: salt, ADMIN_PW_HASH: hashPw_(pw, salt) });
  return 'Password saved. You can now log in from the "Owner" link at the bottom of the booking site.';
}

function adminState(key) {
  checkKey_(key);
  const today = todayStr_();
  const rows = readRows_(getSheet_())
    .filter(r => ACTIVE.indexOf(r[COL.Status]) !== -1 && toDateStr_(r[COL.Date]) >= today)
    .map(r => ({
      id: String(r[COL.ID]), status: String(r[COL.Status]), date: toDateStr_(r[COL.Date]),
      name: String(r[COL.Name]), phone: String(r[COL.Phone]), block: String(r[COL.Block]),
      room: String(r[COL.Room]), pickup: String(r[COL.Pickup]), utr: String(r[COL.UTR]),
    }))
    .sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));
  return {
    today: today, windowDays: BOOKING_WINDOW_DAYS, blackouts: BLACKOUTS, rows: rows,
    hasPassword: !!PropertiesService.getScriptProperties().getProperty('ADMIN_PW_HASH'),
    siteUrl: 'https://pathakabhinav777.github.io/ProjectorRentalSystem/',
  };
}

/** Block a night: projector is in use / booked directly. */
function adminBlock(key, date, note) {
  checkKey_(key);
  date = String(date || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Pick a date.');
  if (date < todayStr_()) throw new Error('That date is in the past.');
  note = String(note || '').replace(/[<>]/g, '').replace(/^[=+\-@\s]+/, '').trim().slice(0, 60);

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = getSheet_();
    const clash = readRows_(sheet).find(r => ACTIVE.indexOf(r[COL.Status]) !== -1 && toDateStr_(r[COL.Date]) === date);
    if (clash) throw new Error(prettyDate_(date) + ' already has a booking (' + clash[COL.Name] + '). Reject or cancel it first.');
    const ev = CalendarApp.getDefaultCalendar().createEvent(
      'Projector: booked directly' + (note ? ' (' + note + ')' : ''),
      nightStart_(date), nightEnd_(date), { description: 'Blocked from the owner panel.' });
    sheet.appendRow([newId_(), new Date(), 'BLOCKED', date, note || 'Booked directly', '', '', '', '', '', '', ev.getId()]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return adminState(key);
}

function adminConfirm(key, id) {
  checkKey_(key);
  withRow_(id, (sheet, rowNum, r) => {
    if (r[COL.Status] !== 'PENDING') throw new Error('Only pending bookings can be confirmed.');
    const date = toDateStr_(r[COL.Date]);
    const ev = CalendarApp.getDefaultCalendar().createEvent(
      'Projector: ' + r[COL.Name] + ' (' + r[COL.Block] + ', ' + r[COL.Room] + ')',
      nightStart_(date), nightEnd_(date),
      { description: 'Phone: ' + r[COL.Phone] + '\nPickup: ' + r[COL.Pickup] +
                     '\nUTR: ' + r[COL.UTR] + '\nAmount: ₹' + r[COL.Amount] + '\nBooking ID: ' + id });
    sheet.getRange(rowNum, COL.Status + 1).setValue('CONFIRMED');
    sheet.getRange(rowNum, COL.CalendarEventId + 1).setValue(ev.getId());
  });
  return adminState(key);
}

/** Reject a pending booking, cancel a confirmed one, or remove a block. Frees the night. */
function adminCancel(key, id) {
  checkKey_(key);
  withRow_(id, (sheet, rowNum, r) => {
    const evId = r[COL.CalendarEventId];
    if (evId) {
      try { const ev = CalendarApp.getEventById(evId); if (ev) ev.deleteEvent(); } catch (e) { console.error(e); }
    }
    sheet.getRange(rowNum, COL.Status + 1).setValue(r[COL.Status] === 'PENDING' ? 'REJECTED' : 'CANCELLED');
  });
  return adminState(key);
}

// ================= Helpers =================

function adminPage_(key) {
  const t = HtmlService.createTemplateFromFile('Admin');
  t.key = keyOk_(key) ? key : '';
  return t.evaluate()
    .setTitle('Projector owner panel')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

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
  if (b.date > addDays_(today, BOOKING_WINDOW_DAYS)) return { error: 'You can only book up to ' + BOOKING_WINDOW_DAYS + ' days ahead.' };
  if (isBlackout_(b.date)) return { error: 'The projector isn\'t available on that night.' };
  return b;
}

/** Only dates and status go public: no names, phones or anything people typed. */
function takenDates_() {
  const today = todayStr_();
  const out = [];
  for (const r of readRows_(getSheet_())) {
    const st = r[COL.Status];
    const d = toDateStr_(r[COL.Date]);
    if (ACTIVE.indexOf(st) !== -1 && d >= today) {
      out.push({ date: d, status: st === 'PENDING' ? 'pending' : 'booked' });
    }
  }
  out.sort((a, b) => (a.date < b.date ? -1 : 1));
  return out;
}

function isBlackout_(d) {
  return BLACKOUTS.some(r => d >= r[0] && d <= r[1]);
}

function notifyOwner_(b) {
  MailApp.sendEmail(Session.getEffectiveUser().getEmail(),
    'Projector booking to check: ' + prettyDate_(b.date) + ' (' + b.name + ')',
    'New booking waiting for your confirmation.\n\n' +
    'Night of: ' + prettyDate_(b.date) + '\nName: ' + b.name + '\nPhone: ' + b.phone +
    '\nRoom: ' + b.block + ', ' + b.room + '\nPickup: ' + b.pickup +
    '\nUTR: ' + b.utr + '\nAmount: ₹' + PRICE +
    '\n\nCheck the UTR in your UPI app, then open your owner panel to confirm or reject:\n' + adminLink_() +
    '\n\nDo not forward this email: the link opens your owner panel without a password.');
}

function adminLink_() {
  const key = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  return WEB_APP_URL + '?action=admin&key=' + key;
}

function withRow_(id, fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = getSheet_();
    const rows = readRows_(sheet);
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][COL.ID]) === String(id)) {
        const out = fn(sheet, i + 2, rows[i]);
        SpreadsheetApp.flush();
        return out;
      }
    }
    throw new Error('Booking not found.');
  } finally {
    lock.releaseLock();
  }
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
  if (!keyOk_(key)) throw new Error('Not allowed. Please log in again.');
}

function hashPw_(pw, salt) {
  return Utilities.base64Encode(Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, salt + ':' + pw, Utilities.Charset.UTF_8));
}

function newId_() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 12);
}

function nightStart_(date) { return new Date(date + 'T' + NIGHT_START + ':00+05:30'); }
function nightEnd_(date) { return new Date(addDays_(date, 1) + 'T' + NIGHT_END + ':00+05:30'); }

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

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
