/**
 * Convene backend — Google Apps Script Web App
 *
 * Paste this file into a new Apps Script project bound to the Sheet
 * (Extensions → Apps Script from inside the Sheet), then Deploy → New deployment
 *   • Type: Web app
 *   • Execute as: Me (your Gmail)
 *   • Who has access: Anyone
 * Copy the resulting /exec URL into admin.html and index.html as SCRIPT_URL.
 *
 * Sheet ID is hard-coded below to the sheet you provided.
 * Tabs expected (created automatically by ensureSchema()):
 *   Admins, Bookings, Reminders, MemberTokens, AuditLog, RateLimit
 */

const SHEET_ID = '1yDmTRekkhcUTnB2AY4IHFEaZL9syUOH22Ivhhz_tAVc';
const SEED_ADMIN = 'yomer2013@gmail.com';
const TOKEN_TTL_DAYS = 30;

// Public site URL used inside magic-link emails. Set this to the URL where
// index.html is served (e.g. the Vercel deployment URL). Falls back to the
// script's own /exec URL if left blank — not ideal but keeps things working.
const PUBLIC_SITE_URL = '';

const TABS = {
  Admins:       ['email', 'name', 'active', 'addedAt'],
  Bookings:     ['id', 'memberName', 'memberEmail', 'service', 'dateTime', 'status', 'notes', 'createdAt', 'updatedAt', 'createdBy'],
  Reminders:    ['id', 'title', 'target', 'trigger', 'channel', 'dueDate', 'priority', 'state', 'updatedAt'],
  MemberTokens: ['email', 'token', 'issuedAt', 'expiresAt'],
  AuditLog:     ['timestamp', 'actorEmail', 'action', 'entityType', 'entityId', 'details'],
  RateLimit:    ['key', 'windowStart', 'count'],
};

// ---------- HTTP entry points ----------

function doGet(e) {
  return handle(e, 'GET');
}
function doPost(e) {
  return handle(e, 'POST');
}

function handle(e, method) {
  try {
    ensureSchema();
    const params = (e && e.parameter) || {};
    let body = {};
    if (method === 'POST' && e && e.postData && e.postData.contents) {
      try { body = JSON.parse(e.postData.contents); } catch (_) { body = {}; }
    }
    const action = body.action || params.action || '';
    const req = Object.assign({}, params, body);

    switch (action) {
      // Public / member
      case 'requestBooking':  return json(requestBooking(req));
      case 'myBookings':      return json(myBookings(req));

      // Admin identity
      case 'whoAmI':          return json(whoAmI(req));

      // Admin reads
      case 'listBookings':    return json(withAdmin(req, listBookings));
      case 'listReminders':   return json(withAdmin(req, listReminders));
      case 'dashboard':       return json(withAdmin(req, dashboard));

      // Admin writes
      case 'createBooking':   return json(withAdmin(req, createBookingAdmin));
      case 'updateBooking':   return json(withAdmin(req, updateBooking));
      case 'cancelBooking':   return json(withAdmin(req, cancelBooking));
      case 'createReminder':  return json(withAdmin(req, createReminder));
      case 'updateReminder':  return json(withAdmin(req, updateReminder));
      case 'completeReminder':return json(withAdmin(req, completeReminder));

      default:
        return json({ ok: false, error: 'unknown_action', action: action });
    }
  } catch (err) {
    return json({ ok: false, error: String(err && err.message || err) });
  }
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- Schema ----------

function ensureSchema() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  Object.keys(TABS).forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.appendRow(TABS[name]);
      sh.setFrozenRows(1);
    } else if (sh.getLastRow() === 0) {
      sh.appendRow(TABS[name]);
      sh.setFrozenRows(1);
    }
  });
  // Seed first admin if Admins tab is empty (beyond header)
  const admins = ss.getSheetByName('Admins');
  if (admins.getLastRow() < 2) {
    admins.appendRow([SEED_ADMIN, 'Primary Admin', true, new Date().toISOString()]);
  }
}

function sheet(name) {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(name);
}

function readAll(name) {
  const sh = sheet(name);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const header = values.shift();
  return values.map(function (row) {
    const o = {};
    header.forEach(function (h, i) { o[h] = row[i]; });
    return o;
  });
}

function appendRow(name, obj) {
  const sh = sheet(name);
  const header = TABS[name];
  sh.appendRow(header.map(function (h) { return obj[h] == null ? '' : obj[h]; }));
}

function updateRowById(name, id, updates) {
  const sh = sheet(name);
  const data = sh.getDataRange().getValues();
  const header = data.shift();
  const idCol = header.indexOf('id');
  if (idCol < 0) throw new Error('no id column on ' + name);
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][idCol]) === String(id)) {
      const row = data[i].slice();
      header.forEach(function (h, j) {
        if (updates[h] !== undefined) row[j] = updates[h];
      });
      sh.getRange(i + 2, 1, 1, header.length).setValues([row]);
      return true;
    }
  }
  return false;
}

// ---------- Admin auth ----------

/**
 * Verify a Google ID token (from Google Sign-In in the browser).
 * Returns { email, name } on success, throws on failure.
 */
function verifyGoogleIdToken(idToken) {
  if (!idToken) throw new Error('missing_id_token');
  const resp = UrlFetchApp.fetch(
    'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
    { muteHttpExceptions: true }
  );
  if (resp.getResponseCode() !== 200) throw new Error('invalid_id_token');
  const info = JSON.parse(resp.getContentText());
  if (!info.email || info.email_verified !== 'true' && info.email_verified !== true) {
    throw new Error('email_not_verified');
  }
  return { email: String(info.email).toLowerCase(), name: info.name || '' };
}

function isAdminEmail(email) {
  const admins = readAll('Admins');
  const e = String(email || '').toLowerCase();
  return admins.some(function (a) {
    return String(a.email || '').toLowerCase() === e && (a.active === true || a.active === 'TRUE' || a.active === 'true');
  });
}

function withAdmin(req, fn) {
  const who = verifyGoogleIdToken(req.idToken);
  if (!isAdminEmail(who.email)) {
    return { ok: false, error: 'not_admin' };
  }
  return fn(req, who);
}

function whoAmI(req) {
  try {
    const who = verifyGoogleIdToken(req.idToken);
    return { ok: true, email: who.email, name: who.name, isAdmin: isAdminEmail(who.email) };
  } catch (err) {
    return { ok: false, error: String(err.message || err) };
  }
}

// ---------- Public: request booking + magic link ----------

function requestBooking(req) {
  const name  = String(req.name || '').trim();
  const email = String(req.email || '').trim().toLowerCase();
  const service = String(req.service || '').trim();
  const dateTime = String(req.dateTime || '').trim();
  const notes = String(req.notes || '').trim();

  if (!name || !email || !service || !dateTime) {
    return { ok: false, error: 'missing_fields' };
  }
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) {
    return { ok: false, error: 'bad_email' };
  }
  if (!rateLimit('req:' + email, 5, 60)) {
    return { ok: false, error: 'rate_limited' };
  }

  const id = 'BK-' + shortId();
  const now = new Date().toISOString();
  appendRow('Bookings', {
    id: id, memberName: name, memberEmail: email, service: service,
    dateTime: dateTime, status: 'pending', notes: notes,
    createdAt: now, updatedAt: now, createdBy: 'public',
  });
  audit('public', 'requestBooking', 'Booking', id, JSON.stringify({ service: service, dateTime: dateTime }));

  const token = issueToken(email);
  const link = magicLink(token);
  try {
    MailApp.sendEmail({
      to: email,
      subject: 'Your booking request was received',
      htmlBody:
        '<p>Hi ' + escapeHtml(name) + ',</p>' +
        '<p>We received your request for <b>' + escapeHtml(service) + '</b> at <b>' + escapeHtml(dateTime) + '</b>. ' +
        'You will get another email once it is confirmed.</p>' +
        '<p>To view your bookings anytime, use this private link (valid for ' + TOKEN_TTL_DAYS + ' days):<br>' +
        '<a href="' + link + '">' + link + '</a></p>' +
        '<p>Booking ID: ' + id + '</p>',
    });
  } catch (mailErr) {
    // Do not fail the booking if email quota is hit.
    audit('public', 'mailError', 'Booking', id, String(mailErr));
  }

  return { ok: true, id: id, magicLink: link };
}

function myBookings(req) {
  const token = String(req.token || '').trim();
  const t = lookupToken(token);
  if (!t) return { ok: false, error: 'bad_or_expired_token' };
  const email = String(t.email).toLowerCase();
  const rows = readAll('Bookings').filter(function (b) {
    return String(b.memberEmail || '').toLowerCase() === email;
  });
  return { ok: true, email: email, bookings: rows };
}

function issueToken(email) {
  // Reuse an unexpired token if one already exists for this email.
  const existing = readAll('MemberTokens').find(function (t) {
    return String(t.email || '').toLowerCase() === email &&
           new Date(t.expiresAt).getTime() > Date.now();
  });
  if (existing) return existing.token;

  const token = Utilities.getUuid().replace(/-/g, '') + shortId();
  const now = new Date();
  const exp = new Date(now.getTime() + TOKEN_TTL_DAYS * 24 * 60 * 60 * 1000);
  appendRow('MemberTokens', {
    email: email, token: token,
    issuedAt: now.toISOString(), expiresAt: exp.toISOString(),
  });
  return token;
}

function lookupToken(token) {
  if (!token) return null;
  const rows = readAll('MemberTokens');
  const row = rows.find(function (r) { return String(r.token) === token; });
  if (!row) return null;
  if (new Date(row.expiresAt).getTime() <= Date.now()) return null;
  return row;
}

function magicLink(token) {
  const base = PUBLIC_SITE_URL || ScriptApp.getService().getUrl();
  const sep = base.indexOf('?') >= 0 ? '&' : '?';
  return base + sep + 'token=' + encodeURIComponent(token);
}

// ---------- Admin: bookings ----------

function listBookings(req) {
  return { ok: true, bookings: readAll('Bookings') };
}

function createBookingAdmin(req, who) {
  const id = 'BK-' + shortId();
  const now = new Date().toISOString();
  appendRow('Bookings', {
    id: id,
    memberName: String(req.memberName || '').trim(),
    memberEmail: String(req.memberEmail || '').trim().toLowerCase(),
    service: String(req.service || '').trim(),
    dateTime: String(req.dateTime || '').trim(),
    status: String(req.status || 'confirmed'),
    notes: String(req.notes || ''),
    createdAt: now, updatedAt: now, createdBy: who.email,
  });
  audit(who.email, 'createBooking', 'Booking', id, '');
  return { ok: true, id: id };
}

function updateBooking(req, who) {
  const id = String(req.id || '');
  const updates = {};
  ['memberName','memberEmail','service','dateTime','status','notes'].forEach(function (k) {
    if (req[k] !== undefined) updates[k] = req[k];
  });
  updates.updatedAt = new Date().toISOString();
  const ok = updateRowById('Bookings', id, updates);
  audit(who.email, 'updateBooking', 'Booking', id, JSON.stringify(updates));
  return { ok: ok };
}

function cancelBooking(req, who) {
  const id = String(req.id || '');
  const ok = updateRowById('Bookings', id, { status: 'cancelled', updatedAt: new Date().toISOString() });
  audit(who.email, 'cancelBooking', 'Booking', id, '');
  return { ok: ok };
}

// ---------- Admin: reminders ----------

function listReminders(req) {
  return { ok: true, reminders: readAll('Reminders') };
}

function createReminder(req, who) {
  const id = 'RM-' + shortId();
  appendRow('Reminders', {
    id: id,
    title: String(req.title || ''),
    target: String(req.target || ''),
    trigger: String(req.trigger || 'Manual'),
    channel: String(req.channel || 'email'),
    dueDate: String(req.dueDate || ''),
    priority: String(req.priority || 'medium'),
    state: String(req.state || 'upcoming'),
    updatedAt: new Date().toISOString(),
  });
  audit(who.email, 'createReminder', 'Reminder', id, '');
  return { ok: true, id: id };
}

function updateReminder(req, who) {
  const id = String(req.id || '');
  const updates = {};
  ['title','target','trigger','channel','dueDate','priority','state'].forEach(function (k) {
    if (req[k] !== undefined) updates[k] = req[k];
  });
  updates.updatedAt = new Date().toISOString();
  const ok = updateRowById('Reminders', id, updates);
  audit(who.email, 'updateReminder', 'Reminder', id, JSON.stringify(updates));
  return { ok: ok };
}

function completeReminder(req, who) {
  const id = String(req.id || '');
  const ok = updateRowById('Reminders', id, { state: 'completed', updatedAt: new Date().toISOString() });
  audit(who.email, 'completeReminder', 'Reminder', id, '');
  return { ok: ok };
}

// ---------- Dashboard aggregate ----------

function dashboard(req) {
  const bookings = readAll('Bookings');
  const reminders = readAll('Reminders');
  const now = Date.now();
  const rangeDays = Number(req.rangeDays || 30);
  const since = now - rangeDays * 24 * 60 * 60 * 1000;

  const inRange = bookings.filter(function (b) {
    const t = parseWhen(b.createdAt);
    return t && t >= since;
  });
  const upcoming = bookings.filter(function (b) {
    const t = parseWhen(b.dateTime);
    return t && t >= now && b.status !== 'cancelled';
  });
  const pendingReminders = reminders.filter(function (r) { return r.state === 'upcoming' || r.state === 'overdue'; });
  const completed = bookings.filter(function (b) { return b.status === 'confirmed' || b.status === 'completed'; });
  const total = bookings.length || 1;

  return {
    ok: true,
    kpis: {
      inRange: inRange.length,
      upcoming: upcoming.length,
      pendingReminders: pendingReminders.length,
      completionRate: Math.round((completed.length / total) * 1000) / 10,
    },
    recentBookings: bookings.slice(-20).reverse(),
    overdueOrUpcomingReminders: pendingReminders.slice(0, 10),
  };
}

// ---------- Helpers ----------

function audit(actorEmail, action, entityType, entityId, details) {
  appendRow('AuditLog', {
    timestamp: new Date().toISOString(),
    actorEmail: actorEmail, action: action,
    entityType: entityType, entityId: entityId, details: details || '',
  });
}

function rateLimit(key, maxPerWindow, windowSec) {
  const sh = sheet('RateLimit');
  const data = sh.getDataRange().getValues();
  const header = data.shift();
  const now = Date.now();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === key) {
      const windowStart = Number(data[i][1]) || 0;
      let count = Number(data[i][2]) || 0;
      if (now - windowStart > windowSec * 1000) {
        sh.getRange(i + 2, 1, 1, header.length).setValues([[key, now, 1]]);
        return true;
      }
      if (count >= maxPerWindow) return false;
      sh.getRange(i + 2, 3).setValue(count + 1);
      return true;
    }
  }
  sh.appendRow([key, now, 1]);
  return true;
}

function parseWhen(s) {
  if (!s) return null;
  const t = new Date(s).getTime();
  return isNaN(t) ? null : t;
}

function shortId() {
  return String(Date.now()).slice(-6) + Math.floor(Math.random() * 900 + 100);
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, function (c) {
    return ({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;' })[c];
  });
}
