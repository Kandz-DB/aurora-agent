'use strict';

require('dotenv').config();
const express  = require('express');
const cors     = require('cors');
const fs       = require('fs');
const path     = require('path');
const multer   = require('multer');
const cron     = require('node-cron');
const Anthropic = require('@anthropic-ai/sdk');
const axios    = require('axios');

// ── Setup ─────────────────────────────────────────────────────────────────────
const app    = express();
const PORT   = process.env.PORT || 3001;
const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

// Load Aurora's prompt
const AURORA_PROMPT = fs.readFileSync(path.join(__dirname, 'prompt.txt'), 'utf8');

// Data folder for storing drafts, spend, project cache
const DATA = path.join(__dirname, 'data');
fs.mkdirSync(DATA, { recursive: true });
fs.mkdirSync(path.join(DATA, 'uploads'), { recursive: true });

// ── Middleware ────────────────────────────────────────────────────────────────
app.use(cors({ origin: process.env.FRONTEND_URL || '*', credentials: true }));
app.use(express.json({ limit: '5mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// Serve index.html for all non-API routes (frontend)
app.get(/^(?!\/api).*/, (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// File uploads
const upload = multer({ dest: path.join(DATA, 'uploads'), limits: { fileSize: 20 * 1024 * 1024 } });

// ── Simple file-based storage helpers ────────────────────────────────────────
function readData(filename) {
  try {
    const f = path.join(DATA, filename);
    return fs.existsSync(f) ? JSON.parse(fs.readFileSync(f, 'utf8')) : [];
  } catch { return []; }
}

function writeData(filename, data) {
  fs.writeFileSync(path.join(DATA, filename), JSON.stringify(data, null, 2));
}

function readJSON(filename, fallback) {
  try {
    const f = path.join(DATA, filename);
    return fs.existsSync(f) ? JSON.parse(fs.readFileSync(f, 'utf8')) : fallback;
  } catch { return fallback; }
}

function writeJSON(filename, data) {
  fs.writeFileSync(path.join(DATA, filename), JSON.stringify(data, null, 2));
}

// ── Cost tracking ─────────────────────────────────────────────────────────────
const CAP_USD = parseFloat(process.env.MONTHLY_SPEND_CAP_USD || '20');

function getSpend() {
  const saved = readJSON('spend.json', { month: '', total: 0, calls: 0 });
  const month = new Date().toISOString().slice(0, 7);
  if (saved.month !== month) return { month, total: 0, calls: 0 };
  return saved;
}

function recordSpend(inputTokens, outputTokens, model) {
  const rates = model.includes('haiku')
    ? { in: 0.80, out: 4.00 }
    : { in: 3.00, out: 15.00 };
  const cost = (inputTokens / 1e6) * rates.in + (outputTokens / 1e6) * rates.out;
  const spend = getSpend();
  spend.total += cost;
  spend.calls += 1;
  writeJSON('spend.json', spend);
  if (spend.total >= CAP_USD * 0.8 && spend.total < CAP_USD) {
    console.warn(`[Cost] ⚠ 80% of monthly cap used: $${spend.total.toFixed(4)}`);
  }
  if (spend.total >= CAP_USD) throw new Error('MONTHLY_CAP_REACHED');
  return cost;
}

// ── Core AI call ──────────────────────────────────────────────────────────────
// Haiku for routine tasks, Sonnet only for chat and document analysis
const TASK_MODELS = {
  chat:             'claude-sonnet-4-6',
  document_analysis:'claude-sonnet-4-6',
  status_email:     'claude-haiku-4-5-20251001',
  checkin_email:    'claude-haiku-4-5-20251001',
  escalation_email: 'claude-haiku-4-5-20251001',
  invoice_reminder: 'claude-haiku-4-5-20251001',
  closeout_email:   'claude-haiku-4-5-20251001',
  status_report:    'claude-haiku-4-5-20251001',
  risk_summary:     'claude-haiku-4-5-20251001',
  milestone_report: 'claude-haiku-4-5-20251001',
  portfolio_report: 'claude-haiku-4-5-20251001',
  closeout_report:  'claude-haiku-4-5-20251001',
  invoice_summary:  'claude-haiku-4-5-20251001',
  change_request:   'claude-haiku-4-5-20251001',
};

const TASK_TOKENS = {
  chat: 800, document_analysis: 1500,
  status_email: 400, checkin_email: 350, escalation_email: 450,
  invoice_reminder: 300, closeout_email: 400,
  status_report: 800, risk_summary: 600, milestone_report: 600,
  portfolio_report: 900, closeout_report: 900,
  invoice_summary: 400, change_request: 600,
};

async function aurora(taskType, userMessage, projectContext) {
  const spend = getSpend();
  if (spend.total >= CAP_USD) throw new Error('MONTHLY_CAP_REACHED');

  const model     = TASK_MODELS[taskType] || 'claude-haiku-4-5-20251001';
  const maxTokens = TASK_TOKENS[taskType] || 500;

  const system = projectContext
    ? `${AURORA_PROMPT}\n\nPROJECT CONTEXT:\n${projectContext}`
    : AURORA_PROMPT;

  const response = await client.messages.create({
    model,
    max_tokens: maxTokens,
    system,
    messages: [{ role: 'user', content: userMessage }],
  });

  recordSpend(response.usage.input_tokens, response.usage.output_tokens, model);

  return response.content.filter(b => b.type === 'text').map(b => b.text).join('');
}

// ── Monday.com sync ───────────────────────────────────────────────────────────
const ONGOING_KEYWORDS = ['ongoing', 'ongoing training delivery', 'training delivery', 'retainer'];

function isOngoing(status) {
  if (!status) return false;
  return ONGOING_KEYWORDS.some(k => status.toLowerCase().includes(k));
}

async function syncMonday() {
  const apiKey  = process.env.MONDAY_API_KEY;
  const boardId = process.env.MONDAY_BOARD_ID;
  if (!apiKey || !boardId) return readData('projects.json');

  try {
   const query = `query {
      boards(ids:[${boardId}]) {
        id
        name
        items_page(limit: 100) {
          items {
            id
            name
            column_values {
              id
              type
              text
              value
            }
          }
        }
      }
    }`;

    const res = await axios.post('https://api.monday.com/v2',
      { query },
      { headers: { Authorization: apiKey, 'Content-Type': 'application/json', 'API-Version': '2024-01' }, timeout: 10000 }
    );

   console.log('[Monday]', JSON.stringify(res.data).slice(0, 500));
const board = res.data?.data?.boards?.[0];
const items = board?.items_page?.items || [];
    console.log('[Monday Debug]', JSON.stringify(res.data?.data).slice(0, 800));

    const projects = items.map(item => {
      const col = (title) => item.column_values.find(c => c.id?.toLowerCase().includes(title.toLowerCase()) || c.type?.toLowerCase().includes(title.toLowerCase()))?.text || '';
      const status = col('status');
      const ongoing = isOngoing(status);
      return {
        id: item.id,
        name: item.name,
        clientName: col('client') || col('company') || 'Unknown',
        clientContact: col('contact') || col('owner') || '',
        clientEmail: col('email') || '',
        type: ongoing ? 'ongoing' : 'standard',
        status,
        phase: ongoing ? null : parsePhase(col('phase') || col('stage')),
        progress: ongoing ? null : parseInt(col('progress') || '0'),
        dueDate: ongoing ? null : col('due') || col('deadline'),
        value: col('value') || col('budget') || '',
        notes: col('notes') || '',
        lastSynced: new Date().toISOString(),
      };
    });

    writeData('projects.json', projects);
    return projects;
  } catch (err) {
    console.error('[Monday] Sync failed:', err.message);
    return readData('projects.json');
  }
}

function parsePhase(text) {
  if (!text) return 0;
  const t = text.toLowerCase();
  if (t.includes('kick')) return 0;
  if (t.includes('deploy')) return 1;
  if (t.includes('monitor') || t.includes('review')) return 2;
  if (t.includes('report')) return 3;
  if (t.includes('close')) return 4;
  return 0;
}

function buildContext(project) {
  if (!project || project.type === 'ongoing') return null;
  const phases = ['Kick-off','Deployment','Monitoring & Review','Reporting','Close-out'];

  // Include any stored document extracts for this project
  const docs = readData('documents.json').filter(d => d.projectId === project.id && d.extract);
  const docText = docs.map(d => `Document: ${d.name}\n${d.extract}`).join('\n\n');

  return [
    `Client: ${project.clientName}`,
    `Project: ${project.name}`,
    `Contact: ${project.clientContact}${project.clientEmail ? ` (${project.clientEmail})` : ''}`,
    `Phase: ${phases[project.phase] || 'Unknown'}`,
    `Progress: ${project.progress}%`,
    `Due: ${project.dueDate || 'TBC'}`,
    `Value: ${project.value || 'TBC'}`,
    `Notes: ${project.notes || 'None'}`,
    docText ? `\nDocuments:\n${docText}` : '',
  ].filter(Boolean).join('\n');
}

// ── Daily batch (6am AEST) ────────────────────────────────────────────────────
async function runBatch() {
  console.log('[Batch] Starting daily batch...');
  const projects = await syncMonday();
  const standard = projects.filter(p => p.type === 'standard');
  console.log(`[Batch] ${standard.length} standard projects | ${projects.length - standard.length} ongoing (skipped)`);

  for (const p of standard) {
    try {
      const context = buildContext(p);
      const today = new Date();

      // Weekly status email on Mondays
      if (today.getDay() === 1) {
        const phases = ['Kick-off','Deployment','Monitoring & Review','Reporting','Close-out'];
        const text = await aurora('status_email',
          `Draft a short weekly status update email to ${p.clientContact || 'the client'} at ${p.clientName} for the ${p.name} project. Current phase: ${phases[p.phase]}. Progress: ${p.progress}%. Keep it to 3-4 sentences.`,
          context
        );
        saveDraft({ projectId: p.id, clientName: p.clientName, projectName: p.name, type: 'status_email', to: p.clientContact, toEmail: p.clientEmail, subject: `${p.name} — Weekly update`, body: text, source: 'batch' });
      }

      // At-risk or behind — draft an escalation
      if (p.progress < 35 && p.dueDate) {
        const text = await aurora('escalation_email',
          `Draft a professional email to ${p.clientContact || 'the client'} at ${p.clientName} about a schedule concern on the ${p.name} project. Progress is at ${p.progress}% with a deadline of ${p.dueDate}. Honest but constructive. Propose a brief call. Max 5 sentences.`,
          context
        );
        saveDraft({ projectId: p.id, clientName: p.clientName, projectName: p.name, type: 'escalation_email', urgency: 'urgent', to: p.clientContact, toEmail: p.clientEmail, subject: `${p.name} — Project update`, body: text, source: 'batch' });
      }

    } catch (err) {
      if (err.message === 'MONTHLY_CAP_REACHED') { console.error('[Batch] Cap reached — stopping'); break; }
      console.error(`[Batch] Error on ${p.name}:`, err.message);
    }
  }
  console.log('[Batch] Done');
}

// 6am AEST = 8pm UTC
cron.schedule('0 20 * * *', () => runBatch().catch(console.error), { timezone: 'UTC' });

// ── Draft store ───────────────────────────────────────────────────────────────
function saveDraft(draft) {
  const drafts = readData('drafts.json');
  const full = { id: `d_${Date.now()}`, ...draft, approved: false, createdAt: new Date().toISOString() };
  drafts.push(full);
  writeData('drafts.json', drafts);
  return full;
}

// ── API Routes ────────────────────────────────────────────────────────────────

// Health
app.get('/api/health', (req, res) => res.json({ status: 'ok', service: 'Aurora R2S' }));

// Projects
app.get('/api/projects', async (req, res) => {
  try { res.json({ projects: await syncMonday() }); }
  catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/projects/sync', async (req, res) => {
  try {
    // Clear cache and force fresh sync
    writeData('projects.json', []);
    res.json({ projects: await syncMonday() });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// Drafts
app.get('/api/drafts', (req, res) => {
  const drafts = readData('drafts.json').filter(d => !d.approved && !d.rejected);
  res.json({ drafts });
});

app.post('/api/drafts/generate', async (req, res) => {
  try {
    const { projectId, taskType, prompt } = req.body;
    const projects = await syncMonday();
    const project  = projects.find(p => p.id === projectId);

    if (!project) return res.status(404).json({ error: 'Project not found' });
    if (project.type === 'ongoing') return res.status(400).json({ error: 'Ongoing projects do not use Aurora automation' });

    const text = await aurora(taskType || 'status_email', prompt, buildContext(project));
    const draft = saveDraft({ projectId, clientName: project.clientName, projectName: project.name, type: taskType, to: project.clientContact, toEmail: project.clientEmail, body: text, source: 'manual' });
    res.json({ draft });
  } catch (e) {
    if (e.message === 'MONTHLY_CAP_REACHED') return res.status(429).json({ error: 'Monthly spend cap reached' });
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/drafts/:id/approve', (req, res) => {
  const drafts = readData('drafts.json');
  const draft  = drafts.find(d => d.id === req.params.id);
  if (!draft) return res.status(404).json({ error: 'Not found' });
  draft.approved = true;
  draft.approvedAt = new Date().toISOString();
  writeData('drafts.json', drafts);
  // Outlook send happens here once Microsoft Graph is connected
  res.json({ success: true, draft });
});

app.post('/api/drafts/:id/reject', (req, res) => {
  const drafts = readData('drafts.json');
  const draft  = drafts.find(d => d.id === req.params.id);
  if (!draft) return res.status(404).json({ error: 'Not found' });
  draft.rejected = true;
  writeData('drafts.json', drafts);
  res.json({ success: true });
});

app.put('/api/drafts/:id', (req, res) => {
  const drafts = readData('drafts.json');
  const draft  = drafts.find(d => d.id === req.params.id);
  if (!draft) return res.status(404).json({ error: 'Not found' });
  if (req.body.body)    draft.body    = req.body.body;
  if (req.body.subject) draft.subject = req.body.subject;
  writeData('drafts.json', drafts);
  res.json({ draft });
});

// Chat
app.post('/api/chat', async (req, res) => {
  try {
    const { message, projectId, history } = req.body;
    if (!message) return res.status(400).json({ error: 'message required' });

    let context = null;
    if (projectId) {
      const projects = await syncMonday();
      const project  = projects.find(p => p.id === projectId);
      if (project && project.type !== 'ongoing') context = buildContext(project);
    }

    // Build conversation from history (last 6 turns)
    const turns = (history || []).slice(-6);
    const fullMessage = turns.length
      ? turns.map(t => `${t.role === 'user' ? 'User' : 'Aurora'}: ${t.content}`).join('\n') + `\nUser: ${message}`
      : message;

    const reply = await aurora('chat', fullMessage, context);
    res.json({ reply });
  } catch (e) {
    if (e.message === 'MONTHLY_CAP_REACHED') return res.status(429).json({ error: 'Monthly spend cap reached' });
    res.status(500).json({ error: e.message });
  }
});

// Documents
app.get('/api/documents', (req, res) => res.json({ documents: readData('documents.json') }));

app.post('/api/documents/upload', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
    const { projectId, clientId } = req.body;

    const projects = await syncMonday();
    const project  = projects.find(p => p.id === projectId);
    if (project?.type === 'ongoing') return res.status(400).json({ error: 'Ongoing projects do not use the document library' });

    // Read file text (basic — works for txt files; PDF/DOCX need extra libs if needed later)
    let rawText = '';
    try { rawText = fs.readFileSync(req.file.path, 'utf8').slice(0, 12000); } catch { rawText = ''; }

    // Analyse with Sonnet — one time only
    let extract = null;
    if (rawText.trim().length > 100) {
      extract = await aurora('document_analysis',
        `Extract key project management info from this document:\n\n${rawText}\n\nSummarise: scope, deliverables, timeline, fees, contacts, key conditions. Be concise.`,
        buildContext(project)
      );
    }

    const docs = readData('documents.json');
    const doc  = { id: `doc_${Date.now()}`, name: req.file.originalname, projectId, clientId, uploadedAt: new Date().toISOString(), extract, archived: false };
    docs.push(doc);
    writeData('documents.json', docs);

    res.json({ document: doc });
  } catch (e) {
    if (e.message === 'MONTHLY_CAP_REACHED') return res.status(429).json({ error: 'Monthly spend cap reached' });
    res.status(500).json({ error: e.message });
  }
});

// Reports
app.post('/api/reports', async (req, res) => {
  try {
    const { reportType, projectId } = req.body;
    const projects = await syncMonday();

    const targets = projectId
      ? projects.filter(p => p.id === projectId && p.type !== 'ongoing')
      : projects.filter(p => p.type !== 'ongoing');

    if (!targets.length) return res.status(404).json({ error: 'No projects found' });

    const contextBlock = targets.map(buildContext).filter(Boolean).join('\n\n---\n\n');

    const prompts = {
      status:    'Write a concise project status report for each project. Phase, progress, RAG status, key activities this week, any blockers. Factual, no padding.',
      milestones:'List all project milestones. For each: project, milestone name, due date, status. Simple table format.',
      risks:     'Summarise the risk register. Format: project | risk | level | mitigation. Short and factual.',
      closeout:  'Write a project close-out report. What was delivered, outcomes, 2-3 forward recommendations.',
      invoices:  'Summarise invoice status. Format: client | project | amount | due date | status.',
      portfolio: 'Write a one-page portfolio overview. RAG status per project, key dates, anything needing attention at the top.',
    };

    const taskType = reportType === 'portfolio' ? 'portfolio_report' : reportType === 'closeout' ? 'closeout_report' : 'status_report';
    const content  = await aurora(taskType, prompts[reportType] || prompts.status, contextBlock);

    res.json({ content, reportType, generatedAt: new Date().toISOString() });
  } catch (e) {
    if (e.message === 'MONTHLY_CAP_REACHED') return res.status(429).json({ error: 'Monthly spend cap reached' });
    res.status(500).json({ error: e.message });
  }
});

// Cost summary
app.get('/api/cost', (req, res) => {
  const spend = getSpend();
  res.json({
    month: spend.month,
    totalUSD: spend.total.toFixed(4),
    totalAUD: (spend.total * 1.55).toFixed(2),
    calls: spend.calls,
    capUSD: CAP_USD,
    percentUsed: ((spend.total / CAP_USD) * 100).toFixed(1),
  });
});

// Manual batch trigger
app.post('/api/batch', async (req, res) => {
  try { await runBatch(); res.json({ success: true }); }
  catch (e) { res.status(500).json({ error: e.message }); }
});

// ── Start ─────────────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`\nAurora R2S — running on port ${PORT}`);
  console.log(`Spend cap: $${CAP_USD} USD/month`);
});
