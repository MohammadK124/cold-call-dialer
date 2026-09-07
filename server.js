// Cold Call Dialer - production server
// Uses Node's built-in SQLite (node:sqlite) so there is nothing native to compile.
import express from 'express';
import cors from 'cors';
import bcrypt from 'bcryptjs';
import jwt from 'jsonwebtoken';
import multer from 'multer';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { DatabaseSync } from 'node:sqlite';
import csvParser from 'csv-parser';
import XLSX from 'xlsx';
import twilio from 'twilio';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const app = express();
const PORT = parseInt(process.env.PORT || '3000', 10);
const SECRET_KEY = process.env.SECRET_KEY || 'cold-dialer-secret-2025';
const TWILIO_SID = process.env.TWILIO_ACCOUNT_SID || '';
const TWILIO_TOKEN = process.env.TWILIO_AUTH_TOKEN || '';
const TWILIO_NUMBER = process.env.TWILIO_PHONE_NUMBER || '';

// ---------- storage locations (use /data volume if Railway mounts one) ----------
const DATA_DIR = fs.existsSync('/data') ? '/data' : path.join(__dirname, 'data');
fs.mkdirSync(DATA_DIR, { recursive: true });
const UPLOAD_DIR = path.join(DATA_DIR, 'uploads');
fs.mkdirSync(UPLOAD_DIR, { recursive: true });
const DB_FILE = process.env.DATABASE_FILE || path.join(DATA_DIR, 'dialer.db');

// ---------- database ----------
const db = new DatabaseSync(DB_FILE);
db.exec(`
  CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY,
    username TEXT UNIQUE NOT NULL,
    password TEXT NOT NULL,
    email TEXT UNIQUE NOT NULL,
    agent_phone TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
  );
  CREATE TABLE IF NOT EXISTS contacts (
    id INTEGER PRIMARY KEY,
    user_id INTEGER NOT NULL,
    name TEXT NOT NULL,
    phone TEXT NOT NULL,
    company TEXT,
    email TEXT,
    notes TEXT,
    custom_fields TEXT,
    imported_at DATETIME DEFAULT CURRENT_TIMESTAMP
  );
  CREATE TABLE IF NOT EXISTS call_history (
    id INTEGER PRIMARY KEY,
    user_id INTEGER NOT NULL,
    contact_id INTEGER NOT NULL,
    status TEXT DEFAULT 'pending',
    call_date DATETIME DEFAULT CURRENT_TIMESTAMP,
    call_duration INTEGER DEFAULT 0,
    notes TEXT,
    follow_up_date DATE,
    tags TEXT,
    call_sid TEXT
  );
`);

// migrations for older databases
for (const col of ['title', 'location']) {
  try { db.exec(`ALTER TABLE contacts ADD COLUMN ${col} TEXT`); } catch {}
}

// default user
const DEFAULT_USER = 'mkhoja@trilix';
if (!db.prepare('SELECT id FROM users WHERE username = ?').get(DEFAULT_USER)) {
  db.prepare('INSERT INTO users (username, email, password) VALUES (?, ?, ?)')
    .run(DEFAULT_USER, 'mkhoja@trilix.ai', bcrypt.hashSync('password123', 10));
  console.log('[db] created default user');
}

// ---------- middleware ----------
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));
const upload = multer({ dest: UPLOAD_DIR });

function auth(req, res, next) {
  const token = (req.headers.authorization || '').split(' ')[1];
  if (!token) return res.status(401).json({ error: 'Not logged in' });
  try {
    req.user = jwt.verify(token, SECRET_KEY);
    next();
  } catch {
    res.status(403).json({ error: 'Session expired, please log in again' });
  }
}

// ---------- auth ----------
app.post('/api/auth/register', (req, res) => {
  const { username, email, password } = req.body || {};
  if (!username || !email || !password) return res.status(400).json({ error: 'Missing required fields' });
  if (password.length < 6) return res.status(400).json({ error: 'Password must be at least 6 characters' });
  try {
    db.prepare('INSERT INTO users (username, email, password) VALUES (?, ?, ?)')
      .run(username.trim(), email.trim(), bcrypt.hashSync(password, 10));
    res.status(201).json({ message: 'User registered successfully' });
  } catch (e) {
    res.status(400).json({ error: /UNIQUE/.test(e.message) ? 'Username or email already exists' : e.message });
  }
});

app.post('/api/auth/login', (req, res) => {
  const { username, password } = req.body || {};
  if (!username || !password) return res.status(400).json({ error: 'Missing username or password' });
  const user = db.prepare('SELECT * FROM users WHERE username = ? OR email = ?').get(username.trim(), username.trim());
  if (!user || !bcrypt.compareSync(password, user.password)) return res.status(401).json({ error: 'Invalid credentials' });
  const token = jwt.sign({ userId: user.id, username: user.username }, SECRET_KEY, { expiresIn: '30d' });
  res.json({ token, user: { id: user.id, username: user.username, email: user.email, agent_phone: user.agent_phone } });
});

app.get('/api/me', auth, (req, res) => {
  const u = db.prepare('SELECT id, username, email, agent_phone FROM users WHERE id = ?').get(req.user.userId);
  res.json({ ...u, twilio_ready: !!(TWILIO_SID && TWILIO_TOKEN && TWILIO_NUMBER), twilio_number: TWILIO_NUMBER });
});

app.post('/api/me/phone', auth, (req, res) => {
  const phone = normalizePhone(req.body?.agent_phone || '');
  if (!phone) return res.status(400).json({ error: 'Enter a valid phone number' });
  db.prepare('UPDATE users SET agent_phone = ? WHERE id = ?').run(phone, req.user.userId);
  res.json({ agent_phone: phone });
});

// ---------- contacts ----------
// Column matching works on any export (Apollo, ZoomInfo, HubSpot, LinkedIn, hand-made sheets...)
const PLACEHOLDER = /^(researching\.*|n\/?a|none|null|unknown|tbd|-+|not available)$/i;
function clean(v) { const s = String(v ?? '').replace(/\s+/g, ' ').trim(); return PLACEHOLDER.test(s) ? '' : s; }
function norm(k) { return String(k).toLowerCase().replace(/[^a-z0-9]/g, ''); }
function isPhone(v) { return (String(v).match(/\d/g) || []).length >= 7; }

// find first non-empty value whose normalised header matches `test` (function) — in priority order
function find(row, ...tests) {
  for (const test of tests) {
    for (const k of Object.keys(row)) {
      const nk = norm(k);
      if (test(nk)) { const v = clean(row[k]); if (v) return v; }
    }
  }
  return '';
}

function normalizePhone(raw) {
  const digits = String(raw || '').replace(/[^\d+]/g, '');
  if (!digits) return '';
  if (digits.startsWith('+')) return digits;
  if (digits.length === 10) return '+1' + digits;
  if (digits.length === 11 && digits.startsWith('1')) return '+' + digits;
  return '+' + digits;
}

function parseRow(row) {
  const has = (nk, s) => nk.includes(s);
  const phoneRaw = find(row,
    nk => (has(nk, 'mobile') || has(nk, 'cell') || has(nk, 'direct')) && !has(nk, 'company'),
    nk => has(nk, 'phone') && !has(nk, 'company') && !has(nk, 'hq'),
    nk => has(nk, 'phone') || has(nk, 'tel') || nk === 'number');
  if (!phoneRaw || !isPhone(phoneRaw)) return null;
  const first = find(row, nk => has(nk, 'firstname') || nk === 'first');
  const last = find(row, nk => has(nk, 'lastname') || nk === 'last');
  const name = find(row,
    nk => (has(nk, 'fullname') || has(nk, 'contactname') || nk === 'name' || nk === 'contact' || nk === 'person') && !has(nk, 'company'),
    nk => has(nk, 'name') && !has(nk, 'company') && !has(nk, 'first') && !has(nk, 'last') && !has(nk, 'domain') && !has(nk, 'file'))
    || [first, last].filter(Boolean).join(' ');
  const company = find(row,
    nk => has(nk, 'companyname') || nk === 'company' || nk === 'organization' || nk === 'organisation' || nk === 'account',
    nk => has(nk, 'company') && !has(nk, 'location') && !has(nk, 'description') && !has(nk, 'website') && !has(nk, 'domain') && !has(nk, 'industry') && !has(nk, 'revenue') && !has(nk, 'staff') && !has(nk, 'founded') && !has(nk, 'phone') && !has(nk, 'size') && !has(nk, 'employees'),
    nk => has(nk, 'employer') || has(nk, 'business'));
  const email = find(row, nk => has(nk, 'email') && !has(nk, 'status') && !has(nk, 'verified'));
  const title = find(row, nk => nk === 'title' || has(nk, 'jobtitle') || nk === 'position' || nk === 'role' || nk === 'designation');
  const location = find(row,
    nk => has(nk, 'contactlocation') || nk === 'location' || nk === 'city' || has(nk, 'contactcity'),
    nk => has(nk, 'contactstate') || nk === 'state' || has(nk, 'companylocation') || has(nk, 'address'));
  const notes = find(row, nk => nk === 'notes' || nk === 'note' || nk === 'comments');
  return { name: name || company || phoneRaw, phone: phoneRaw, company, email, title, location, notes };
}

function saveContacts(userId, rows) {
  const insert = db.prepare('INSERT INTO contacts (user_id, name, phone, company, email, title, location, notes, custom_fields) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)');
  let inserted = 0, skipped = 0;
  for (const row of rows) {
    const p = parseRow(row);
    if (!p) { skipped++; continue; }
    // keep only non-empty extra columns so the UI can show them
    const extras = {};
    for (const k of Object.keys(row)) { const v = clean(row[k]); if (v) extras[String(k).trim()] = v; }
    insert.run(userId, p.name, p.phone, p.company, p.email, p.title, p.location, p.notes, JSON.stringify(extras));
    inserted++;
  }
  return { inserted, skipped };
}

app.post('/api/upload', auth, upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  const ext = path.extname(req.file.originalname).toLowerCase();
  try {
    let rows = [];
    if (ext === '.csv' || ext === '.txt') {
      rows = await new Promise((resolve, reject) => {
        const out = [];
        fs.createReadStream(req.file.path).pipe(csvParser({ mapHeaders: ({ header }) => header.replace(/^﻿/, '').trim() }))
          .on('data', r => out.push(r)).on('end', () => resolve(out)).on('error', reject);
      });
    } else if (ext === '.xlsx' || ext === '.xls') {
      const wb = XLSX.readFile(req.file.path);
      rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' });
    } else {
      return res.status(400).json({ error: 'Unsupported file. Upload a CSV or Excel file.' });
    }
    const { inserted, skipped } = saveContacts(req.user.userId, rows);
    if (!inserted) return res.status(400).json({ error: 'No contacts with phone numbers found. Make sure there is a "phone" column.' });
    res.json({ message: `${inserted} contacts imported${skipped ? ` (${skipped} skipped — no phone number)` : ''}`, count: inserted, skipped });
  } catch (e) {
    res.status(500).json({ error: e.message });
  } finally {
    fs.unlink(req.file.path, () => {});
  }
});

app.get('/api/contacts', auth, (req, res) => {
  const rows = db.prepare(`
    SELECT c.*,
      (SELECT status FROM call_history WHERE contact_id = c.id ORDER BY call_date DESC LIMIT 1) AS last_call_status,
      (SELECT call_date FROM call_history WHERE contact_id = c.id ORDER BY call_date DESC LIMIT 1) AS last_call_date,
      (SELECT COUNT(*) FROM call_history WHERE contact_id = c.id) AS total_calls
    FROM contacts c WHERE c.user_id = ? ORDER BY c.id`).all(req.user.userId);
  res.json(rows);
});

app.delete('/api/contacts', auth, (req, res) => {
  db.prepare('DELETE FROM call_history WHERE user_id = ?').run(req.user.userId);
  db.prepare('DELETE FROM contacts WHERE user_id = ?').run(req.user.userId);
  res.json({ message: 'All contacts cleared' });
});

app.get('/api/contacts/:id/calls', auth, (req, res) => {
  res.json(db.prepare('SELECT * FROM call_history WHERE contact_id = ? AND user_id = ? ORDER BY call_date DESC')
    .all(req.params.id, req.user.userId));
});

// ---------- calls ----------
app.post('/api/calls', auth, (req, res) => {
  const { contact_id, status, call_duration, notes, follow_up_date, tags, call_sid } = req.body || {};
  if (!contact_id) return res.status(400).json({ error: 'contact_id required' });
  const r = db.prepare(`INSERT INTO call_history (user_id, contact_id, status, call_duration, notes, follow_up_date, tags, call_sid)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?)`)
    .run(req.user.userId, contact_id, status || 'called', call_duration || 0, notes || '', follow_up_date || null, tags || '', call_sid || null);
  res.json({ id: Number(r.lastInsertRowid), message: 'Call logged successfully' });
});

// Twilio click-to-call: rings your phone first, then bridges to the contact (recorded).
app.post('/api/twilio/call', auth, async (req, res) => {
  if (!(TWILIO_SID && TWILIO_TOKEN && TWILIO_NUMBER)) return res.status(400).json({ error: 'Twilio is not configured on the server' });
  const contact = db.prepare('SELECT * FROM contacts WHERE id = ? AND user_id = ?').get(req.body?.contact_id, req.user.userId);
  if (!contact) return res.status(404).json({ error: 'Contact not found' });
  const me = db.prepare('SELECT agent_phone FROM users WHERE id = ?').get(req.user.userId);
  const agent = normalizePhone(req.body?.agent_phone || me?.agent_phone);
  if (!agent) return res.status(400).json({ error: 'Save your own phone number first (Settings)' });
  if (req.body?.agent_phone) db.prepare('UPDATE users SET agent_phone = ? WHERE id = ?').run(agent, req.user.userId);
  const to = normalizePhone(contact.phone);
  const safeName = String(contact.name || 'your contact').replace(/[<>&"']/g, '');
  const twiml = `<Response><Say voice="alice">Connecting you to ${safeName}.</Say><Dial callerId="${TWILIO_NUMBER}" record="record-from-answer-dual" timeout="30">${to}</Dial></Response>`;
  try {
    const call = await twilio(TWILIO_SID, TWILIO_TOKEN).calls.create({ to: agent, from: TWILIO_NUMBER, twiml });
    res.json({ sid: call.sid, message: `Calling your phone (${agent}) now. Answer it to be connected to ${contact.name}.` });
  } catch (e) {
    res.status(500).json({ error: `Twilio error: ${e.message}` });
  }
});

app.get('/api/stats', auth, (req, res) => {
  const u = req.user.userId;
  res.json(db.prepare(`SELECT
    (SELECT COUNT(*) FROM contacts WHERE user_id = ?) AS total_contacts,
    (SELECT COUNT(*) FROM call_history WHERE user_id = ?) AS total_calls,
    (SELECT COUNT(*) FROM call_history WHERE user_id = ? AND status = 'interested') AS interested_count,
    (SELECT COUNT(*) FROM call_history WHERE user_id = ? AND status = 'not_interested') AS not_interested_count,
    (SELECT COUNT(*) FROM call_history WHERE user_id = ? AND status = 'callback') AS callback_count,
    (SELECT AVG(call_duration) FROM call_history WHERE user_id = ?) AS avg_call_duration`).get(u, u, u, u, u, u));
});

app.get('/api/export/calls', auth, (req, res) => {
  const rows = db.prepare(`SELECT c.name, c.phone, c.company, c.email, h.status, h.call_date, h.call_duration, h.notes, h.follow_up_date
    FROM call_history h JOIN contacts c ON c.id = h.contact_id WHERE h.user_id = ? ORDER BY h.call_date DESC`).all(req.user.userId);
  const q = v => `"${String(v ?? '').replace(/"/g, '""')}"`;
  const csv = ['Name,Phone,Company,Email,Status,Date,Duration,Notes,Follow-up',
    ...rows.map(r => [r.name, r.phone, r.company, r.email, r.status, r.call_date, r.call_duration, r.notes, r.follow_up_date].map(q).join(','))].join('\n');
  res.setHeader('Content-Type', 'text/csv');
  res.setHeader('Content-Disposition', 'attachment; filename="call-history.csv"');
  res.send(csv);
});

app.get('/api/health', (req, res) => res.json({ status: 'ok', twilio: !!(TWILIO_SID && TWILIO_TOKEN && TWILIO_NUMBER), time: new Date().toISOString() }));
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));
app.use('/api', (req, res) => res.status(404).json({ error: 'Not found' }));
app.use((err, req, res, next) => { console.error(err); res.status(500).json({ error: err.message || 'Server error' }); });

// ---------- start ----------
app.listen(PORT, '0.0.0.0', () => console.log(`[server] listening on ${PORT}`));
if (PORT !== 3000) app.listen(3000, '0.0.0.0', () => console.log('[server] also listening on 3000')).on('error', () => {});
process.on('unhandledRejection', e => console.error('unhandledRejection', e));
