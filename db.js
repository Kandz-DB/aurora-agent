'use strict';
/**
 * Aurora DB — PostgreSQL via Render (or Azure PostgreSQL)
 * Falls back to JSON files if DATABASE_URL not set (local dev)
 */

const fs   = require('fs');
const path = require('path');

const USE_PG = !!process.env.DATABASE_URL;

let pool = null;
if (USE_PG) {
  const { Pool } = require('pg');
  pool = new Pool({
    connectionString: process.env.DATABASE_URL,
    ssl: process.env.NODE_ENV === 'production' ? { rejectUnauthorized: false } : false,
  });
}

// ── Schema ────────────────────────────────────────────────────────────────────
const SCHEMA = `
CREATE TABLE IF NOT EXISTS projects (
  id TEXT PRIMARY KEY,
  client_name TEXT NOT NULL,
  project_name TEXT,
  client_contact TEXT,
  client_email TEXT,
  client_phone TEXT,
  status TEXT DEFAULT 'In Progress',
  phase INTEGER DEFAULT 0,
  value TEXT,
  contract_start TEXT,
  due_date TEXT,
  summary TEXT,
  deliverables TEXT,
  milestones TEXT,
  timeline TEXT,
  invoicing_notes TEXT,
  consultant TEXT,
  consultant_email TEXT,
  flights_required TEXT,
  accommodation_required TEXT,
  notes TEXT,
  sharepoint_url TEXT,
  type TEXT DEFAULT 'standard',
  monday_id TEXT,
  created_at TIMESTAMPTZ DEFAULT NOW(),
  updated_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS documents (
  id TEXT PRIMARY KEY,
  project_id TEXT REFERENCES projects(id) ON DELETE CASCADE,
  name TEXT,
  extract TEXT,
  uploaded_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS drafts (
  id TEXT PRIMARY KEY,
  project_id TEXT,
  client_name TEXT,
  project_name TEXT,
  type TEXT,
  urgency TEXT DEFAULT 'routine',
  to_name TEXT,
  to_email TEXT,
  subject TEXT,
  body TEXT,
  approved BOOLEAN DEFAULT FALSE,
  rejected BOOLEAN DEFAULT FALSE,
  source TEXT DEFAULT 'manual',
  created_at TIMESTAMPTZ DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS spend (
  id SERIAL PRIMARY KEY,
  month TEXT,
  total_usd NUMERIC DEFAULT 0,
  calls INTEGER DEFAULT 0,
  updated_at TIMESTAMPTZ DEFAULT NOW()
);
`;

async function initDB() {
  if (!USE_PG) { console.log('[DB] Using JSON file storage (no DATABASE_URL)'); return; }
  try {
    await pool.query(SCHEMA);
    console.log('[DB] PostgreSQL schema ready');
  } catch (err) {
    console.error('[DB] Schema init failed:', err.message);
  }
}

// ── JSON file fallback ────────────────────────────────────────────────────────
const DATA = path.join(__dirname, 'data');
fs.mkdirSync(DATA, { recursive: true });
fs.mkdirSync(path.join(DATA, 'uploads'), { recursive: true });

function readJSON(file, fallback = []) {
  try {
    const f = path.join(DATA, file);
    return fs.existsSync(f) ? JSON.parse(fs.readFileSync(f, 'utf8')) : fallback;
  } catch { return fallback; }
}

function writeJSON(file, data) {
  fs.writeFileSync(path.join(DATA, file), JSON.stringify(data, null, 2));
}

// ── Project CRUD ──────────────────────────────────────────────────────────────
async function getProjects() {
  if (USE_PG) {
    const r = await pool.query('SELECT * FROM projects ORDER BY updated_at DESC');
    return r.rows.map(dbRowToProject);
  }
  return readJSON('projects.json');
}

async function getProject(id) {
  if (USE_PG) {
    const r = await pool.query('SELECT * FROM projects WHERE id=$1', [id]);
    return r.rows[0] ? dbRowToProject(r.rows[0]) : null;
  }
  return readJSON('projects.json').find(p => p.id === id) || null;
}

async function upsertProject(p) {
  if (USE_PG) {
    await pool.query(`
      INSERT INTO projects (id,client_name,project_name,client_contact,client_email,client_phone,
        status,phase,value,contract_start,due_date,summary,deliverables,milestones,timeline,
        invoicing_notes,consultant,consultant_email,flights_required,accommodation_required,
        notes,sharepoint_url,type,monday_id,updated_at)
      VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18,$19,$20,$21,$22,$23,$24,NOW())
      ON CONFLICT (id) DO UPDATE SET
        client_name=$2,project_name=$3,client_contact=$4,client_email=$5,client_phone=$6,
        status=$7,phase=$8,value=$9,contract_start=$10,due_date=$11,summary=$12,deliverables=$13,
        milestones=$14,timeline=$15,invoicing_notes=$16,consultant=$17,consultant_email=$18,
        flights_required=$19,accommodation_required=$20,notes=$21,sharepoint_url=$22,
        type=$23,monday_id=$24,updated_at=NOW()
    `, [p.id,p.clientName,p.projectName,p.clientContact,p.clientEmail,p.clientPhone||'',
        p.status,p.phase||0,p.value,p.contractStart,p.dueDate,p.summary,p.deliverables,
        p.milestones||'',p.timeline||'',p.invoicingNotes,p.consultant,p.consultantEmail||'',
        p.flightsRequired||'',p.accommodationRequired||'',p.notes,
        p.sharepointUrl,p.type||'standard',p.mondayId]);
    return p;
  }
  const projects = readJSON('projects.json');
  const idx = projects.findIndex(x => x.id === p.id);
  if (idx >= 0) projects[idx] = { ...projects[idx], ...p, updatedAt: new Date().toISOString() };
  else projects.push({ ...p, createdAt: new Date().toISOString(), updatedAt: new Date().toISOString() });
  writeJSON('projects.json', projects);
  return p;
}

async function updateProjectField(id, fields) {
  if (USE_PG) {
    const keys = Object.keys(fields);
    const snakeFields = keys.map(k => toSnake(k));
    const setClause = snakeFields.map((k,i) => `${k}=$${i+2}`).join(',');
    const vals = keys.map(k => fields[k]);
    await pool.query(`UPDATE projects SET ${setClause},updated_at=NOW() WHERE id=$1`, [id, ...vals]);
    return getProject(id);
  }
  const projects = readJSON('projects.json');
  const idx = projects.findIndex(p => p.id === id);
  if (idx >= 0) { Object.assign(projects[idx], fields, { updatedAt: new Date().toISOString() }); writeJSON('projects.json', projects); }
  return projects[idx];
}

async function deleteProject(id) {
  if (USE_PG) { await pool.query('DELETE FROM projects WHERE id=$1', [id]); return; }
  const projects = readJSON('projects.json').filter(p => p.id !== id);
  writeJSON('projects.json', projects);
}

// ── Document CRUD ─────────────────────────────────────────────────────────────
async function getDocuments(projectId) {
  if (USE_PG) {
    const q = projectId
      ? 'SELECT * FROM documents WHERE project_id=$1 ORDER BY uploaded_at DESC'
      : 'SELECT * FROM documents ORDER BY uploaded_at DESC';
    const r = await pool.query(q, projectId ? [projectId] : []);
    return r.rows;
  }
  const docs = readJSON('documents.json');
  return projectId ? docs.filter(d => d.projectId === projectId) : docs;
}

async function saveDocument(doc) {
  if (USE_PG) {
    await pool.query(
      'INSERT INTO documents (id,project_id,name,extract) VALUES ($1,$2,$3,$4) ON CONFLICT (id) DO UPDATE SET extract=$4',
      [doc.id, doc.projectId, doc.name, doc.extract]
    );
    return doc;
  }
  const docs = readJSON('documents.json');
  const idx = docs.findIndex(d => d.id === doc.id);
  if (idx >= 0) docs[idx] = doc; else docs.push(doc);
  writeJSON('documents.json', docs);
  return doc;
}

// ── Draft CRUD ────────────────────────────────────────────────────────────────
async function getDrafts(filter = {}) {
  if (USE_PG) {
    let q = 'SELECT * FROM drafts WHERE approved=FALSE AND rejected=FALSE';
    const vals = [];
    if (filter.projectId) { vals.push(filter.projectId); q += ` AND project_id=$${vals.length}`; }
    q += ' ORDER BY created_at DESC';
    const r = await pool.query(q, vals);
    return r.rows.map(dbRowToDraft);
  }
  let drafts = readJSON('drafts.json').filter(d => !d.approved && !d.rejected);
  if (filter.projectId) drafts = drafts.filter(d => d.projectId === filter.projectId);
  return drafts;
}

async function saveDraft(draft) {
  if (USE_PG) {
    await pool.query(
      `INSERT INTO drafts (id,project_id,client_name,project_name,type,urgency,to_name,to_email,subject,body,source)
       VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11)`,
      [draft.id,draft.projectId,draft.clientName,draft.projectName,draft.type,draft.urgency||'routine',
       draft.toName,draft.toEmail,draft.subject,draft.body,draft.source||'manual']
    );
    return draft;
  }
  const drafts = readJSON('drafts.json');
  drafts.push({ ...draft, createdAt: new Date().toISOString() });
  writeJSON('drafts.json', drafts);
  return draft;
}

async function updateDraft(id, fields) {
  if (USE_PG) {
    const keys = Object.keys(fields);
    const setClause = keys.map((k,i) => `${toSnake(k)}=$${i+2}`).join(',');
    await pool.query(`UPDATE drafts SET ${setClause} WHERE id=$1`, [id, ...keys.map(k => fields[k])]);
    return;
  }
  const drafts = readJSON('drafts.json');
  const idx = drafts.findIndex(d => d.id === id);
  if (idx >= 0) Object.assign(drafts[idx], fields);
  writeJSON('drafts.json', drafts);
}

// ── Spend tracking ────────────────────────────────────────────────────────────
async function getSpend() {
  const month = new Date().toISOString().slice(0, 7);
  if (USE_PG) {
    const r = await pool.query('SELECT * FROM spend WHERE month=$1', [month]);
    if (r.rows[0]) return { month, total: parseFloat(r.rows[0].total_usd), calls: r.rows[0].calls };
    return { month, total: 0, calls: 0 };
  }
  const saved = readJSON('spend.json', { month: '', total: 0, calls: 0 });
  return saved.month === month ? saved : { month, total: 0, calls: 0 };
}

async function recordSpend(cost) {
  const month = new Date().toISOString().slice(0, 7);
  if (USE_PG) {
    await pool.query(`
      INSERT INTO spend (month, total_usd, calls, updated_at) VALUES ($1,$2,1,NOW())
      ON CONFLICT (month) DO UPDATE SET total_usd=spend.total_usd+$2, calls=spend.calls+1, updated_at=NOW()
    `, [month, cost]);
    return;
  }
  const spend = await getSpend();
  spend.total += cost;
  spend.calls += 1;
  writeJSON('spend.json', spend);
}

// Add unique constraint for spend month if using PG
async function ensureSpendConstraint() {
  if (!USE_PG) return;
  try {
    await pool.query('ALTER TABLE spend ADD CONSTRAINT spend_month_unique UNIQUE (month)');
  } catch { /* already exists */ }
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function toSnake(camel) {
  return camel.replace(/[A-Z]/g, c => '_' + c.toLowerCase());
}

function dbRowToProject(row) {
  return {
    id: row.id, clientName: row.client_name, projectName: row.project_name,
    clientContact: row.client_contact, clientEmail: row.client_email,
    clientPhone: row.client_phone,
    status: row.status, phase: row.phase || 0, value: row.value,
    contractStart: row.contract_start, dueDate: row.due_date,
    summary: row.summary, deliverables: row.deliverables,
    milestones: row.milestones, timeline: row.timeline,
    invoicingNotes: row.invoicing_notes,
    consultant: row.consultant, consultantEmail: row.consultant_email,
    flightsRequired: row.flights_required,
    accommodationRequired: row.accommodation_required,
    notes: row.notes, sharepointUrl: row.sharepoint_url,
    type: row.type, mondayId: row.monday_id,
    updatedAt: row.updated_at, createdAt: row.created_at,
  };
}

function dbRowToDraft(row) {
  return {
    id: row.id, projectId: row.project_id, clientName: row.client_name,
    projectName: row.project_name, type: row.type, urgency: row.urgency,
    toName: row.to_name, toEmail: row.to_email, subject: row.subject,
    body: row.body, approved: row.approved, rejected: row.rejected,
    source: row.source, createdAt: row.created_at,
  };
}

module.exports = {
  initDB, ensureSpendConstraint,
  getProjects, getProject, upsertProject, updateProjectField, deleteProject,
  getDocuments, saveDocument,
  getDrafts, saveDraft, updateDraft,
  getSpend, recordSpend,
  DATA, readJSON, writeJSON,
};
