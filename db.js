'use strict';
/**
 * Aurora DB — Azure Blob Storage (primary) with local JSON fallback (dev)
 *
 * Set these env vars for Azure Blob:
 *   AZURE_STORAGE_CONNECTION_STRING  — from your storage account Access Keys
 *   AZURE_STORAGE_CONTAINER          — container name e.g. "aurora-data"
 *
 * Without those vars it falls back to local JSON files (Render / local dev).
 */

const fs   = require('fs');
const path = require('path');

// ── Storage mode detection ────────────────────────────────────────────────────
const USE_BLOB = !!process.env.AZURE_STORAGE_CONNECTION_STRING;

let blobContainer = null;

async function initDB() {
  if (!USE_BLOB) {
    fs.mkdirSync(path.join(__dirname, 'data', 'uploads'), { recursive: true });
    console.log('[DB] Using local JSON file storage (no Azure Blob config)');
    return;
  }

  try {
    const { BlobServiceClient } = require('@azure/storage-blob');
    const serviceClient = BlobServiceClient.fromConnectionString(
      process.env.AZURE_STORAGE_CONNECTION_STRING
    );
    const containerName = process.env.AZURE_STORAGE_CONTAINER || 'aurora-data';
    blobContainer = serviceClient.getContainerClient(containerName);

    // Create container if it doesn't exist
    await blobContainer.createIfNotExists({ access: 'private' });
    console.log(`[DB] Azure Blob Storage ready — container: ${containerName}`);
  } catch (err) {
    console.error('[DB] Azure Blob init failed:', err.message);
    console.log('[DB] Falling back to local JSON files');
    blobContainer = null;
  }
}

// No-op for compatibility with old PG code
async function ensureSpendConstraint() {}

// ── Blob helpers ──────────────────────────────────────────────────────────────
async function blobRead(key, fallback) {
  if (!blobContainer) return fallback;
  try {
    const blob = blobContainer.getBlockBlobClient(key);
    const download = await blob.downloadToBuffer();
    return JSON.parse(download.toString('utf8'));
  } catch (err) {
    if (err.statusCode === 404 || err.code === 'BlobNotFound') return fallback;
    console.error(`[Blob] Read error for ${key}:`, err.message);
    return fallback;
  }
}

async function blobWrite(key, data) {
  if (!blobContainer) return;
  try {
    const blob = blobContainer.getBlockBlobClient(key);
    const json = JSON.stringify(data, null, 2);
    await blob.upload(json, Buffer.byteLength(json), {
      blobHTTPHeaders: { blobContentType: 'application/json' },
    });
  } catch (err) {
    console.error(`[Blob] Write error for ${key}:`, err.message);
    throw err;
  }
}

// ── Local JSON fallback ───────────────────────────────────────────────────────
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

async function read(key, fallback) {
  if (USE_BLOB && blobContainer) return blobRead(key, fallback);
  return readJSON(key, fallback);
}

async function write(key, data) {
  if (USE_BLOB && blobContainer) return blobWrite(key, data);
  writeJSON(key, data);
}

// ── Project CRUD ──────────────────────────────────────────────────────────────
async function getProjects() {
  return read('projects.json', []);
}

async function getProject(id) {
  const projects = await getProjects();
  return projects.find(p => p.id === id) || null;
}

async function upsertProject(p) {
  const projects = await getProjects();
  const idx = projects.findIndex(x => x.id === p.id);
  const now = new Date().toISOString();
  if (idx >= 0) {
    projects[idx] = { ...projects[idx], ...p, updatedAt: now };
  } else {
    projects.push({ ...p, createdAt: now, updatedAt: now });
  }
  await write('projects.json', projects);
  return p;
}

async function updateProjectField(id, fields) {
  const projects = await getProjects();
  const idx = projects.findIndex(p => p.id === id);
  if (idx >= 0) {
    Object.assign(projects[idx], fields, { updatedAt: new Date().toISOString() });
    await write('projects.json', projects);
    return projects[idx];
  }
  return null;
}

async function deleteProject(id) {
  const projects = (await getProjects()).filter(p => p.id !== id);
  await write('projects.json', projects);
}

// ── Document CRUD ─────────────────────────────────────────────────────────────
async function getDocuments(projectId) {
  const docs = await read('documents.json', []);
  return projectId ? docs.filter(d => d.projectId === projectId) : docs;
}

async function saveDocument(doc) {
  const docs = await read('documents.json', []);
  const idx = docs.findIndex(d => d.id === doc.id);
  const now = new Date().toISOString();
  if (idx >= 0) docs[idx] = { ...doc, updatedAt: now };
  else docs.push({ ...doc, uploadedAt: now });
  await write('documents.json', docs);

  // Also save the raw file to Blob if available (for contract files)
  if (USE_BLOB && blobContainer && doc.filePath) {
    try {
      const { BlobServiceClient } = require('@azure/storage-blob');
      const blob = blobContainer.getBlockBlobClient(`uploads/${doc.id}_${doc.name}`);
      const fileData = fs.readFileSync(doc.filePath);
      await blob.upload(fileData, fileData.length);
      console.log(`[Blob] Contract file uploaded: ${doc.name}`);
    } catch (err) {
      console.error('[Blob] File upload error:', err.message);
    }
  }

  return doc;
}

// ── Draft CRUD ────────────────────────────────────────────────────────────────
async function getDrafts(filter = {}) {
  const drafts = await read('drafts.json', []);
  let result = drafts.filter(d => !d.approved && !d.rejected);
  if (filter.projectId) result = result.filter(d => d.projectId === filter.projectId);
  return result;
}

async function saveDraft(draft) {
  const drafts = await read('drafts.json', []);
  drafts.push({ ...draft, createdAt: new Date().toISOString() });
  await write('drafts.json', drafts);
  return draft;
}

async function updateDraft(id, fields) {
  const drafts = await read('drafts.json', []);
  const idx = drafts.findIndex(d => d.id === id);
  if (idx >= 0) Object.assign(drafts[idx], fields);
  await write('drafts.json', drafts);
}

// ── Spend tracking ────────────────────────────────────────────────────────────
async function getSpend() {
  const month = new Date().toISOString().slice(0, 7);
  const saved = await read('spend.json', { month: '', total: 0, calls: 0 });
  return saved.month === month ? saved : { month, total: 0, calls: 0 };
}

async function recordSpend(cost) {
  const spend = await getSpend();
  spend.total = (spend.total || 0) + cost;
  spend.calls = (spend.calls || 0) + 1;
  await write('spend.json', spend);
}

module.exports = {
  initDB, ensureSpendConstraint,
  getProjects, getProject, upsertProject, updateProjectField, deleteProject,
  getDocuments, saveDocument,
  getDrafts, saveDraft, updateDraft,
  getSpend, recordSpend,
  DATA, readJSON, writeJSON,
};

// ── Suggestions CRUD ──────────────────────────────────────────────────────────
async function getSuggestions() {
  const all = await read('suggestions.json', []);
  return all.filter(s => s.status === 'pending');
}

async function saveSuggestion(suggestion) {
  const all = await read('suggestions.json', []);
  // Avoid duplicates — same project + same type within 24hrs
  const recent = all.find(s =>
    s.projectId === suggestion.projectId &&
    s.type === suggestion.type &&
    s.status === 'pending' &&
    (Date.now() - new Date(s.createdAt).getTime()) < 24 * 60 * 60 * 1000
  );
  if (recent) return recent;
  const full = { ...suggestion, createdAt: new Date().toISOString(), status: 'pending' };
  all.push(full);
  await write('suggestions.json', all);
  return full;
}

async function updateSuggestion(id, status) {
  const all = await read('suggestions.json', []);
  const idx = all.findIndex(s => s.id === id);
  if (idx >= 0) {
    all[idx].status = status;
    all[idx].resolvedAt = new Date().toISOString();
    await write('suggestions.json', all);
  }
  return all[idx];
}

module.exports.getSuggestions  = getSuggestions;
module.exports.saveSuggestion  = saveSuggestion;
module.exports.updateSuggestion = updateSuggestion;

// ── Risk Register CRUD ────────────────────────────────────────────────────────
async function getRiskRegister(projectId) {
  return read(`risks_${projectId}.json`, []);
}

async function saveRiskRegister(projectId, risks) {
  await write(`risks_${projectId}.json`, risks);
  return risks;
}

async function updateRisk(projectId, riskId, fields) {
  const risks = await getRiskRegister(projectId);
  const idx = risks.findIndex(r => r.id === riskId);
  if (idx >= 0) {
    Object.assign(risks[idx], fields, { updatedAt: new Date().toISOString() });
    await saveRiskRegister(projectId, risks);
    return risks[idx];
  }
  return null;
}

// ── Deliverables Tracker CRUD ─────────────────────────────────────────────────
async function getDeliverables(projectId) {
  return read(`deliverables_${projectId}.json`, []);
}

async function saveDeliverables(projectId, deliverables) {
  await write(`deliverables_${projectId}.json`, deliverables);
  return deliverables;
}

async function updateDeliverable(projectId, deliverableId, fields) {
  const items = await getDeliverables(projectId);
  const idx = items.findIndex(d => d.id === deliverableId);
  if (idx >= 0) {
    Object.assign(items[idx], fields, { updatedAt: new Date().toISOString() });
    await saveDeliverables(projectId, items);
    return items[idx];
  }
  return null;
}

module.exports.getRiskRegister    = getRiskRegister;
module.exports.saveRiskRegister   = saveRiskRegister;
module.exports.updateRisk         = updateRisk;
module.exports.getDeliverables    = getDeliverables;
module.exports.saveDeliverables   = saveDeliverables;
module.exports.updateDeliverable  = updateDeliverable;
