'use strict';
require('dotenv').config();

const express   = require('express');
const cors      = require('cors');
const fs        = require('fs');
const path      = require('path');
const multer    = require('multer');
const cron      = require('node-cron');
const Anthropic = require('@anthropic-ai/sdk');
const axios     = require('axios');
const nodemailer= require('nodemailer');
const db        = require('./db');

// ── Invoice storage helpers (since db.read/write may not be exported) ─────────
async function readInvoices(projectId) {
  try {
    if (db.read) return await db.read(`invoices_${projectId}.json`, []);
    const f = require('path').join(db.DATA, `invoices_${projectId}.json`);
    return require('fs').existsSync(f) ? JSON.parse(require('fs').readFileSync(f,'utf8')) : [];
  } catch { return []; }
}

async function writeInvoices(projectId, invoices) {
  if (db.write) return await db.write(`invoices_${projectId}.json`, invoices);
  require('fs').writeFileSync(
    require('path').join(db.DATA, `invoices_${projectId}.json`),
    JSON.stringify(invoices, null, 2)
  );
}

// ── Setup ─────────────────────────────────────────────────────────────────────
const app    = express();
const PORT   = process.env.PORT || 3001;
const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
const AURORA_PROMPT = fs.readFileSync(path.join(__dirname, 'prompt.txt'), 'utf8');

const INTERNAL_EMAILS = [
  process.env.INTERNAL_EMAIL_1 || 'diane.k@risk2solution.com',
  process.env.INTERNAL_EMAIL_2 || 'info@risk2solution.com',
];

// ── Middleware ────────────────────────────────────────────────────────────────
app.use(cors({ origin: process.env.FRONTEND_URL || '*', credentials: true }));
app.use(express.json({ limit: '10mb' }));

// ── Auth — simple password protection ────────────────────────────────────────
const AURORA_PASSWORD = process.env.AURORA_PASSWORD || 'r2s-aurora-2026';

app.post('/api/auth/login', (req, res) => {
  const { password } = req.body;
  if (password === AURORA_PASSWORD) {
    const token = Buffer.from('aurora:' + AURORA_PASSWORD + ':' + Date.now()).toString('base64');
    res.json({ success: true, token });
  } else {
    res.status(401).json({ success: false, error: 'Incorrect password' });
  }
});

app.use('/api', (req, res, next) => {
  // Public endpoints — no auth required
  const publicPaths = ['/auth/login', '/health'];
  if (publicPaths.some(p => req.path === p || req.path.startsWith(p))) return next();

  const authHeader = req.headers['authorization'] || '';
  const token = authHeader.replace('Bearer ', '').trim();
  if (!token) return res.status(401).json({ error: 'Not authenticated' });
  try {
    const decoded = Buffer.from(token, 'base64').toString('utf8');
    if (!decoded.startsWith('aurora:' + AURORA_PASSWORD)) throw new Error('Invalid');
    next();
  } catch (e) {
    res.status(401).json({ error: 'Invalid session — please log in again' });
  }
});

app.use(express.static(path.join(__dirname, 'public')));
app.get(/^(?!\/api).*/, (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

const upload = multer({
  dest: path.join(db.DATA, 'uploads'),
  limits: { fileSize: 25 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const ok = ['.pdf', '.docx', '.doc', '.txt', '.md'].includes(
      path.extname(file.originalname).toLowerCase()
    );
    cb(null, ok);
  },
});

// ── Cost tracking ─────────────────────────────────────────────────────────────
const CAP_USD = parseFloat(process.env.MONTHLY_SPEND_CAP_USD || '20');

const TASK_MODELS = {
  chat:              'claude-haiku-4-5-20251001',
  document_analysis: 'claude-haiku-4-5-20251001',
  contract_extract:  'claude-sonnet-4-6',
  internal_extract:  'claude-sonnet-4-6',
  status_email:      'claude-haiku-4-5-20251001',
  checkin_email:     'claude-haiku-4-5-20251001',
  escalation_email:  'claude-haiku-4-5-20251001',
  invoice_reminder:  'claude-haiku-4-5-20251001',
  reminder_email:    'claude-haiku-4-5-20251001',
  closeout_email:    'claude-haiku-4-5-20251001',
  status_report:     'claude-haiku-4-5-20251001',
  portfolio_report:  'claude-haiku-4-5-20251001',
  closeout_report:   'claude-haiku-4-5-20251001',
};

const TASK_TOKENS = {
  chat: 800, document_analysis: 1500, contract_extract: 2000,
  internal_extract: 2000,
  status_email: 400, checkin_email: 350, escalation_email: 450,
  invoice_reminder: 300, reminder_email: 300, closeout_email: 400,
  status_report: 800, portfolio_report: 900, closeout_report: 900,
};

async function aurora(taskType, userMessage, context) {
  const spend = await db.getSpend();
  if (spend.total >= CAP_USD) throw new Error('MONTHLY_CAP_REACHED');

  const model     = TASK_MODELS[taskType] || 'claude-haiku-4-5-20251001';
  const maxTokens = TASK_TOKENS[taskType] || 500;
  const system    = context ? `${AURORA_PROMPT}\n\nPROJECT CONTEXT:\n${context}` : AURORA_PROMPT;

  const response = await client.messages.create({
    model, max_tokens: maxTokens,
    system,
    messages: [{ role: 'user', content: userMessage }],
  });

  const rates = model.includes('haiku') ? { in: 0.80, out: 4.00 } : { in: 3.00, out: 15.00 };
  const cost  = (response.usage.input_tokens / 1e6) * rates.in + (response.usage.output_tokens / 1e6) * rates.out;
  await db.recordSpend(cost);

  if (spend.total + cost >= CAP_USD * 0.8 && spend.total < CAP_USD * 0.8) {
    await sendInternalEmail('⚠ Aurora spend alert', `Monthly API spend has reached 80% of the $${CAP_USD} USD cap. Current: $${(spend.total + cost).toFixed(2)}`);
  }

  return response.content.filter(b => b.type === 'text').map(b => b.text).join('');
}

// ── Email (Outlook via Graph API) ─────────────────────────────────────────────
async function getOutlookToken() {
  const tenantId     = process.env.OUTLOOK_TENANT_ID;
  const clientId     = process.env.OUTLOOK_CLIENT_ID;
  const clientSecret = process.env.OUTLOOK_CLIENT_SECRET;
  if (!tenantId || !clientId || !clientSecret) {
    console.error('[Email] Missing Outlook env vars — OUTLOOK_TENANT_ID:', !!tenantId, 'OUTLOOK_CLIENT_ID:', !!clientId, 'OUTLOOK_CLIENT_SECRET:', !!clientSecret);
    return null;
  }
  try {
    const res = await axios.post(
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
      new URLSearchParams({ client_id: clientId, client_secret: clientSecret, scope: 'https://graph.microsoft.com/.default', grant_type: 'client_credentials' }),
      { headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, timeout: 10000 }
    );
    return res.data.access_token;
  } catch (err) {
    console.error('[Email] Token failed:', err.response?.data?.error_description || err.message);
    return null;
  }
}

async function sendEmail(to, subject, body, isInternal = false, cc = [], isHtml = false) {
  const token = await getOutlookToken();
  const fromMailbox = process.env.OUTLOOK_SHARED_MAILBOX || 'info@risk2solution.com';
  // Auto-detect HTML if not explicitly set
  const htmlContent = isHtml || body.trimStart().startsWith('<!DOCTYPE') || body.trimStart().startsWith('<html') || body.trimStart().startsWith('<div');

  if (token) {
    try {
      const toArray = Array.isArray(to) ? to : [to];
      const ccArray = Array.isArray(cc) ? cc : (cc ? [cc] : []);
      const message = {
        subject,
        body: { contentType: htmlContent ? 'HTML' : 'Text', content: body },
        toRecipients: toArray.map(addr => ({ emailAddress: { address: addr } })),
      };
      if (ccArray.length > 0) {
        message.ccRecipients = ccArray.map(addr => ({ emailAddress: { address: addr } }));
      }
      await axios.post(
        `https://graph.microsoft.com/v1.0/users/${fromMailbox}/sendMail`,
        { message },
        { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }, timeout: 15000 }
      );
      console.log(`[Email] ✓ Sent to ${toArray.join(', ')}${ccArray.length ? ' CC: '+ccArray.join(', ') : ''}: ${subject}`);
      return true;
    } catch (err) {
      console.error('[Email] Send failed:', err.response?.data || err.message);
    }
  }
  // Fallback — log only
  console.log(`[Email] [LOGGED - no Outlook config] To: ${to}${cc?' CC: '+cc:''} | Subject: ${subject}`);
  return false;
}

async function sendInternalEmail(subject, body) {
  return sendEmail(INTERNAL_EMAILS, subject, body, true);
}

async function saveDraftEmail(draft) {
  const DIANE = 'diane.k@risk2solution.com';
  const token = await getOutlookToken();
  const fromMailbox = process.env.OUTLOOK_SHARED_MAILBOX || 'info@risk2solution.com';
  if (token) {
    try {
      // Always CC Diane on all external drafts unless she IS the recipient
      const ccList = [];
      if (draft.toEmail && draft.toEmail.toLowerCase() !== DIANE.toLowerCase()) {
        ccList.push({ emailAddress: { address: DIANE, name: 'Diane Kruger' } });
      }
      if (draft.ccEmail && draft.ccEmail.toLowerCase() !== DIANE.toLowerCase()) {
        ccList.push({ emailAddress: { address: draft.ccEmail } });
      }

      await axios.post(
        `https://graph.microsoft.com/v1.0/users/${fromMailbox}/messages`,
        {
          subject: draft.subject,
          body: { contentType: 'Text', content: draft.body },
          toRecipients: [{ emailAddress: { address: draft.toEmail || fromMailbox } }],
          ...(ccList.length > 0 ? { ccRecipients: ccList } : {}),
        },
        { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }, timeout: 15000 }
      );
      console.log(`[Email] Draft saved to Outlook: ${draft.subject}`);
    } catch (err) {
      console.error('[Email] Draft save failed:', err.message);
    }
  }
}

// ── Calendar (Outlook) ────────────────────────────────────────────────────────
async function createCalendarReminder(subject, body, reminderDate) {
  const token = await getOutlookToken();
  const mailbox = process.env.OUTLOOK_SHARED_MAILBOX || 'info@risk2solution.com';
  if (!token) { console.log('[Calendar] No token — reminder logged only:', subject); return; }
  try {
    const start = new Date(reminderDate);
    const end   = new Date(start.getTime() + 60 * 60 * 1000);
    await axios.post(
      `https://graph.microsoft.com/v1.0/users/${mailbox}/events`,
      {
        subject,
        body: { contentType: 'Text', content: body },
        start: { dateTime: start.toISOString(), timeZone: 'Australia/Brisbane' },
        end:   { dateTime: end.toISOString(),   timeZone: 'Australia/Brisbane' },
        isReminderOn: true, reminderMinutesBeforeStart: 60 * 24 * 3,
      },
      { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }, timeout: 15000 }
    );
    console.log(`[Calendar] ✓ Reminder set: ${subject} on ${reminderDate}`);
  } catch (err) {
    console.error('[Calendar] Failed:', err.message);
  }
}

// ── Contract text extraction ──────────────────────────────────────────────────
async function extractTextFromFile(filePath, mimeType) {
  const ext = path.extname(filePath).toLowerCase();
  try {
    if (ext === '.pdf' || mimeType === 'application/pdf') {
      const pdfParse = require('pdf-parse');
      const buf  = fs.readFileSync(filePath);
      const data = await pdfParse(buf);
      return data.text;
    }
    if (ext === '.docx') {
      const mammoth = require('mammoth');
      const result  = await mammoth.extractRawText({ path: filePath });
      return result.value;
    }
    return fs.readFileSync(filePath, 'utf8');
  } catch (err) {
    console.error('[Extract] Text extraction failed:', err.message);
    return '';
  }
}

async function analyseContract(rawText, filename) {
  // Use up to 28000 chars to capture full proposals including cost summaries at end
  const fullText = rawText.slice(0, 28000);

  // Also extract a "tail" section — the last 5000 chars often has cost totals
  const tailText = rawText.length > 15000 ? rawText.slice(-5000) : '';

  const combinedText = fullText + (tailText ? '\n\n[END OF DOCUMENT — KEY TOTALS SECTION:]\n' + tailText : '');

  const text = await aurora(
    'contract_extract',
    `You are reading a client contract or proposal for Risk 2 Solution (R2S). Extract ALL of the following information and return it as a valid JSON object with exactly these keys. Be thorough — read the ENTIRE document including the costs summary and commercial offer sections which are often near the end.

IMPORTANT FOR VALUE FIELD: Look for a TOTAL or GRAND TOTAL line in the costs summary table. This is usually the single largest dollar figure in the document. Do NOT use a per-session rate or sub-total. Find the overall total project cost (e.g. TOTAL $63,000).

{
  "organisationName": "full legal organisation/company name of the client",
  "clientName": "organisation name (same as above, used for display)",
  "projectName": "project title or name of the engagement as written in the document",
  "clientContact": "primary client contact person full name",
  "clientEmail": "primary client contact email address",
  "clientPhone": "primary client contact phone number",
  "value": "TOTAL project cost only — the grand total from the costs summary e.g. $63,000. Do NOT list individual line items.",
  "contractStart": "contract start date or engagement commencement date",
  "dueDate": "project completion date, contract end date, or due date",
  "summary": "full description of services R2S is providing — extract the key paragraphs describing what R2S will do for the client",
  "deliverables": "all specific deliverables and stages listed — include phase names and what is delivered in each",
  "milestones": "any key milestones, phases, or stages mentioned with dates or conditions",
  "timeline": "overall project timeline description — start to finish with any phasing or scheduling mentioned",
  "invoicingNotes": "full payment terms, invoicing schedule, milestone payment triggers, and invoicing frequency",
  "consultant": "name(s) of any R2S consultant, trainer, or staff member assigned or mentioned",
  "consultantEmail": "email address of the assigned consultant or trainer if mentioned",
  "flightsRequired": "yes or no — are flights required for this engagement",
  "accommodationRequired": "yes or no — is accommodation required for this engagement",
  "notes": "any special conditions, exclusions, cancellation terms, or important requirements"
}

Return ONLY the JSON object. No markdown, no explanation, no other text. If a field is not found in the document, use an empty string "".

Document: ${filename}
---
${combinedText}`,
    null
  );

  try {
    const clean = text.replace(/```json|```/g, '').trim();
    return JSON.parse(clean);
  } catch (err) {
    console.error('[Contract] JSON parse failed:', err.message);
    return { summary: text, deliverables: '', clientName: '', projectName: filename };
  }
}

// ── Consultant / trainer briefing email ──────────────────────────────────────
async function sendConsultantBriefing(project, extracted, prebuiltContext) {
  const DIANE = 'diane.k@risk2solution.com';

  // Consultant email — use extracted email if available, otherwise fall back to info@ for now
  const consultantEmail = extracted.consultantEmail || process.env.CONSULTANT_DEFAULT_EMAIL || 'info@risk2solution.com';
  const consultantName  = extracted.consultant || project.consultant || 'Team';

  const context = prebuiltContext || buildContext(project);

  const firstName = consultantName.split(' ')[0];
  const briefingBody = await aurora(
    'consultant_briefing',
    `Write a short project assignment email. Plain text only — no asterisks, no bold, no markdown formatting at all.

Start exactly with:
Hi ${firstName},

You've been assigned to a new project for ${project.clientName} — ${project.projectName || project.clientName}. Commencing ${project.contractStart || 'TBC'} and due ${project.dueDate || 'TBC'}.

Then write these four short sections with no section headings — just plain paragraphs and dot points:

1. SCOPE (1-2 sentences only): What R2S is doing for this client. Be concise.

2. KEY DELIVERABLES (3-5 dot points maximum, starting with •): The most important deliverables only. Not the full list.

3. TIMELINE (one sentence only): Just the total duration e.g. "The engagement runs for approximately 6 weeks from commencement." Do not list individual phases.

4. TRAVEL (one line only if flights or accommodation are needed): State simply e.g. "Flights and accommodation are required for this engagement." Skip this section entirely if not required.

Then end with exactly:

Here is the link to the client SharePoint folder: [Diane to insert SP link]

Please ensure all working notes, materials, and deliverables are saved to this folder throughout the engagement in accordance with our File Management SOP.

Please confirm you have received this briefing and are clear on the requirements. Contact Diane if you have any questions.

Kind regards,

Diane Kruger
Corporate Operations Lead
Risk 2 Solution Group
P: 1300 459 970 | M: +61 415 748 747
E: diane.k@risk2solution.com
W: www.risk2solution.com
Queensland, Australia`,
    context
  );

  const subject = `Project briefing: ${project.clientName} — ${project.projectName || project.clientName}`;

  // Save as draft in Outlook shared mailbox (requires approval before sending)
  const draft = {
    id: `d_${Date.now()}_consult`,
    projectId: project.id,
    clientName: project.clientName,
    projectName: project.projectName,
    type: 'consultant_briefing',
    urgency: 'routine',
    toName: consultantName,
    toEmail: consultantEmail,
    ccEmail: DIANE,
    subject,
    body: briefingBody,
    source: 'auto',
  };

  await db.saveDraft(draft);

  // Save to Outlook drafts with CC
  const token = await getOutlookToken();
  const fromMailbox = process.env.OUTLOOK_SHARED_MAILBOX || 'info@risk2solution.com';
  if (token) {
    try {
      await axios.post(
        `https://graph.microsoft.com/v1.0/users/${fromMailbox}/messages`,
        {
          subject,
          body: { contentType: 'Text', content: briefingBody },
          toRecipients: [{ emailAddress: { address: consultantEmail } }],
          ccRecipients: [{ emailAddress: { address: DIANE } }],
        },
        { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }, timeout: 15000 }
      );
      console.log(`[Briefing] ✓ Consultant briefing draft saved for ${consultantName} (${consultantEmail})`);
    } catch (err) {
      console.error('[Briefing] Draft save failed:', err.message);
    }
  } else {
    console.log(`[Briefing] [LOGGED] Consultant briefing for ${consultantName} — ${subject}`);
  }

  // Notify internal team that a briefing has been prepared
  await sendInternalEmail(
    `[Aurora] Consultant briefing ready: ${project.clientName}`,
    `Aurora has prepared a consultant briefing email for ${consultantName} on the ${project.projectName || project.clientName} project.

Please review and approve in the Comms Drafts section of Aurora or in the info@risk2solution.com Outlook shared mailbox drafts folder before sending.

Project: ${project.projectName || project.clientName}
Client: ${project.clientName}
Consultant: ${consultantName}
Due date: ${project.dueDate || 'TBC'}

Aurora
R2S Project Management Intelligence`
  );

  return draft;
}

// ── Project context builder ───────────────────────────────────────────────────
const PHASES = ['Kick-off', 'Deployment', 'Monitoring & Review', 'Reporting', 'Close-out', 'Completed'];

function buildContext(p, docs = []) {
  if (!p || p.type === 'ongoing') return null;
  const docText = docs.map(d => d.extract ? `Contract: ${d.name}\n${d.extract.slice(0, 2000)}` : '').filter(Boolean).join('\n\n');
  return [
    `Organisation: ${p.clientName}`,
    `Project: ${p.projectName || p.clientName}`,
    `Client contact: ${p.clientContact || ''}${p.clientEmail ? ` (${p.clientEmail})` : ''}${p.clientPhone ? ` Ph: ${p.clientPhone}` : ''}`,
    `Phase: ${PHASES[p.phase || 0]}`,
    `Status: ${p.status || 'In Progress'}`,
    `Contract start: ${p.contractStart || 'TBC'}`,
    `Due / completion date: ${p.dueDate || 'TBC'}`,
    `Contract value: ${p.value || 'TBC'}`,
    p.summary ? `Summary of service: ${p.summary}` : '',
    p.deliverables ? `Deliverables: ${p.deliverables}` : '',
    p.milestones ? `Milestones: ${p.milestones}` : '',
    p.timeline ? `Timeline: ${p.timeline}` : '',
    p.invoicingNotes ? `Invoicing terms: ${p.invoicingNotes}` : '',
    p.consultant ? `Consultant/Trainer: ${p.consultant}${p.consultantEmail ? ` (${p.consultantEmail})` : ''}` : '',
    p.flightsRequired ? `Flights required: ${p.flightsRequired}` : '',
    p.accommodationRequired ? `Accommodation required: ${p.accommodationRequired}` : '',
    p.notes ? `Notes / special requirements: ${p.notes}` : '',
    docText ? `\nContract detail:\n${docText}` : '',
  ].filter(Boolean).join('\n');
}

// ── Monday.com sync (backup) ──────────────────────────────────────────────────
const ONGOING_KEYWORDS = ['ongoing training delivery', 'ongoing training', 'ongoing', 'retainer', 'training delivery'];

function isOngoing(status) {
  if (!status) return false;
  return ONGOING_KEYWORDS.some(k => status.toLowerCase().includes(k));
}

async function syncMonday() {
  const apiKey  = process.env.MONDAY_API_KEY;
  const boardId = process.env.MONDAY_BOARD_ID;
  if (!apiKey || !boardId) return [];

  try {
    const query = `query {
      boards(ids:[${boardId}]) {
        items_page(limit: 100) {
          items {
            id name
            column_values { id type text value }
          }
        }
      }
    }`;

    const res = await axios.post('https://api.monday.com/v2', { query },
      { headers: { Authorization: apiKey, 'Content-Type': 'application/json', 'API-Version': '2024-01' }, timeout: 10000 }
    );

    const items = res.data?.data?.boards?.[0]?.items_page?.items || [];

    for (const item of items) {
      const byId   = id   => item.column_values.find(c => c.id === id)?.text || '';
      const byType = type => item.column_values.find(c => c.type === type)?.text || '';

      const status  = byId('color_mks0pnz5') || byType('color') || '';
      const ongoing = isOngoing(status);
      const numCols = item.column_values.filter(c => c.type === 'numbers' && c.text && parseFloat(c.text) > 0);
      const rawVal  = numCols[0]?.text || '';
      const displayValue = rawVal ? '$' + parseFloat(rawVal.replace(/[$,]/g,'')).toLocaleString('en-AU') : '';
      const longTexts = item.column_values.filter(c => c.type === 'long_text' && c.text);
      const dateCols  = item.column_values.filter(c => c.type === 'date' && c.text);
      const spCol     = item.column_values.find(c => c.type === 'link');
      let sharepointUrl = '';
      if (spCol?.value) { try { const v = JSON.parse(spCol.value); sharepointUrl = v.url || ''; } catch { sharepointUrl = spCol.text || ''; } }

      const project = {
        id:            `monday_${item.id}`,
        mondayId:      item.id,
        clientName:    item.name,
        projectName:   byId('text__1') || item.name,
        clientContact: byId('text8__1') || '',
        clientEmail:   byId('client_contact_email__1') || byType('email') || '',
        status,
        type:          ongoing ? 'ongoing' : 'standard',
        phase:         0,
        value:         displayValue,
        summary:       longTexts[0]?.text || '',
        deliverables:  item.column_values.find(c => c.id?.includes('deliver'))?.text || '',
        invoicingNotes:item.column_values.find(c => c.id?.includes('invoic'))?.text || '',
        consultant:    item.column_values.find(c => c.id?.includes('trainer') || c.id?.includes('consultant'))?.text || '',
        dueDate:       dateCols[1]?.text || dateCols[0]?.text || '',
        contractStart: dateCols[0]?.text || '',
        notes:         longTexts[1]?.text || '',
        sharepointUrl,
      };

      await db.upsertProject(project);
    }

    const count = items.length;
    console.log(`[Monday] Synced ${count} projects`);
    return await db.getProjects();
  } catch (err) {
    console.error('[Monday] Sync failed:', err.message);
    return db.getProjects();
  }
}

// ── Due date reminder checker ─────────────────────────────────────────────────
async function checkDueDateReminders() {
  console.log('[Reminders] Checking due dates...');
  const projects = await db.getProjects();
  const standard = projects.filter(p => p.type === 'standard' && !['Completed','Terminated','Closed'].includes(p.status));
  const today    = new Date();

  for (const p of standard) {
    if (!p.dueDate) continue;
    const due  = new Date(p.dueDate);
    if (isNaN(due)) continue;
    const days = Math.round((due - today) / (1000 * 60 * 60 * 24));
    const context = buildContext(p);

    // Send reminders at 14, 7, and 3 days before due date
    if ([14, 7, 3].includes(days)) {
      const urgency = days <= 3 ? 'URGENT' : days <= 7 ? 'Important' : 'Reminder';
      const subject = `[Aurora] ${urgency}: ${p.clientName} — ${p.projectName || 'Project'} due in ${days} day${days !== 1 ? 's' : ''}`;

      const body = await aurora(
        'reminder_email',
        `Draft an internal reminder email to the R2S team. The ${p.projectName || p.clientName} project is due in ${days} days (${p.dueDate}).
Current phase: ${PHASES[p.phase || 0]}.
${p.summary ? `Service: ${p.summary.slice(0, 300)}` : ''}
${p.deliverables ? `Deliverables: ${p.deliverables.slice(0, 200)}` : ''}
Ask the team to review the phase completion status, confirm all deliverables are on track, and flag anything outstanding. Keep it short and direct. This is an internal email only.`,
        context
      );

      await sendInternalEmail(subject, body);
      console.log(`[Reminders] ✓ ${days}-day reminder sent for ${p.clientName}`);
    }

    // Day of due date
    if (days === 0) {
      const subject = `[Aurora] Due today: ${p.clientName} — ${p.projectName || 'Project'}`;
      const body = `The contract end date for ${p.clientName} (${p.projectName || 'project'}) is today.\n\nCurrent phase: ${PHASES[p.phase || 0]}\n\nPlease confirm whether the project is ready for close-out or if the date needs to be updated in Aurora.`;
      await sendInternalEmail(subject, body);
    }

    // Overdue
    if (days < 0 && days >= -3) {
      const subject = `[Aurora] OVERDUE: ${p.clientName} — ${p.projectName || 'Project'} (${Math.abs(days)} days overdue)`;
      const body = `The ${p.clientName} project was due ${Math.abs(days)} days ago and has not been marked complete in Aurora.\n\nPlease review and either update the due date or move to Close-out.`;
      await sendInternalEmail(subject, body);
    }
  }
}

// ── Ensure Outlook category exists ───────────────────────────────────────────
async function ensureOutlookCategory(token, mailbox) {
  try {
    // Check if category already exists
    const res = await axios.get(
      `https://graph.microsoft.com/v1.0/users/${mailbox}/outlook/masterCategories`,
      { headers: { Authorization: `Bearer ${token}` }, timeout: 10000 }
    );
    const exists = (res.data?.value || []).some(c => c.displayName === 'Aurora Processed');
    if (!exists) {
      // Create it with green colour
      await axios.post(
        `https://graph.microsoft.com/v1.0/users/${mailbox}/outlook/masterCategories`,
        { displayName: 'Aurora Processed', color: 'preset5' }, // preset5 = green
        { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }, timeout: 10000 }
      );
      console.log('[Category] Created "Aurora Processed" category in Outlook');
    }
  } catch (err) {
    console.error('[Category] Could not create category:', err.response?.data?.error?.message || err.message);
  }
}

// ── Create tentative calendar booking ────────────────────────────────────────
async function createCalendarBooking(booking, tentative = true) {
  const token = await getOutlookToken();
  if (!token) { console.log('[Calendar] No token — booking logged only'); return null; }
  const mailbox = process.env.OUTLOOK_SHARED_MAILBOX || 'info@risk2solution.com';

  try {
    const start = new Date(booking.startDateTime);
    const end   = new Date(start.getTime() + (booking.durationMinutes || 60) * 60 * 1000);

    // Build attendee list
    const attendees = [];
    if (booking.consultantEmail) attendees.push({ emailAddress: { address: booking.consultantEmail, name: booking.consultantName || booking.consultantEmail }, type: 'required' });
    if (booking.clientEmail)     attendees.push({ emailAddress: { address: booking.clientEmail, name: booking.clientName || booking.clientEmail }, type: 'required' });
    attendees.push({ emailAddress: { address: 'diane.k@risk2solution.com', name: 'Diane Kruger' }, type: 'required' });

    const event = {
      subject: booking.title,
      body: { contentType: 'Text', content: booking.description || '' },
      start: { dateTime: start.toISOString(), timeZone: 'Australia/Brisbane' },
      end:   { dateTime: end.toISOString(),   timeZone: 'Australia/Brisbane' },
      location: { displayName: booking.location || booking.clientName || 'To be confirmed' },
      attendees,
      showAs: tentative ? 'tentative' : 'busy',
      isOnlineMeeting: booking.online || false,
      isReminderOn: true,
      reminderMinutesBeforeStart: 60 * 24, // 24hr reminder
    };

    // Find the training calendar
    const calsRes = await axios.get(
      `https://graph.microsoft.com/v1.0/users/${mailbox}/calendars`,
      { headers: { Authorization: `Bearer ${token}` }, timeout: 10000 }
    );
    const trainingCal = calsRes.data?.value?.find(c =>
      c.name?.toLowerCase().includes('training') || c.name?.toLowerCase().includes('education')
    );
    const calId = trainingCal?.id || 'primary';

    // Create the event
    const eventRes = await axios.post(
      `https://graph.microsoft.com/v1.0/users/${mailbox}/calendars/${calId}/events`,
      event,
      { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }, timeout: 15000 }
    );

    console.log(`[Calendar] ${tentative ? 'Tentative' : 'Confirmed'} booking created: ${booking.title}`);
    return eventRes.data?.id;
  } catch (err) {
    console.error('[Calendar] Booking failed:', err.response?.data?.error?.message || err.message);
    return null;
  }
}

// ── Send meeting invites via Outlook ─────────────────────────────────────────
async function sendMeetingInvite(booking, eventId) {
  const token = await getOutlookToken();
  if (!token || !eventId) return;
  const mailbox = process.env.OUTLOOK_SHARED_MAILBOX || 'info@risk2solution.com';

  try {
    // Confirm the tentative event (changes showAs to 'busy' and sends invites)
    await axios.patch(
      `https://graph.microsoft.com/v1.0/users/${mailbox}/events/${eventId}`,
      { showAs: 'busy' },
      { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }, timeout: 10000 }
    );
    console.log(`[Calendar] Meeting invites sent for: ${booking.title}`);
  } catch (err) {
    console.error('[Calendar] Send invite failed:', err.message);
  }
}

// ── Extract calendar event from email body ────────────────────────────────────
async function extractCalendarEventFromEmail(emailBody, emailSubject, project) {
  try {
    const analysis = await aurora('status_email',
      `Analyse this email to determine if it mentions a specific date and time for a meeting, training session, workshop, or face-to-face deliverable related to the project.

Project: ${project.projectName || project.clientName}
Client: ${project.clientName}
Email subject: ${emailSubject}
Email content: ${emailBody.slice(0, 2000)}

Respond in JSON only:
{
  "hasEvent": true or false,
  "eventType": "training" or "workshop" or "meeting" or "site_visit" or "presentation" or "other",
  "title": "short event title",
  "date": "YYYY-MM-DD or empty string if not found",
  "time": "HH:MM in 24hr format or empty string",
  "durationMinutes": 60,
  "location": "location mentioned or empty string",
  "description": "brief description of what this event is for",
  "requiresConsultant": true or false,
  "requiresClient": true or false
}

Return ONLY the JSON. If no specific date is mentioned, set hasEvent to false.`,
      null
    );

    const clean = analysis.replace(/```json|```/g, '').trim();
    return JSON.parse(clean);
  } catch (err) {
    return { hasEvent: false };
  }
}

// ── Read consultant reply emails from info@ inbox ─────────────────────────────
async function readConsultantReplies() {
  const token = await getOutlookToken();
  if (!token) { console.log('[Poll] No Outlook token — skipping'); return; }
  const mailbox = process.env.OUTLOOK_SHARED_MAILBOX || 'info@risk2solution.com';

  try {
    // Ensure the Aurora Processed category exists in Outlook
    await ensureOutlookCategory(token, mailbox);

    // Get unread emails NOT already tagged with Aurora Processed — last 7 days
    const since = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000).toISOString();

    // Check both inbox and sent items
    const folders = ['inbox', 'sentitems'];
    let allMessages = [];

    for (const folder of folders) {
      try {
        // Fetch recent messages — filter in code to avoid OData type issues
        const url = `https://graph.microsoft.com/v1.0/users/${mailbox}/mailFolders/${folder}/messages?$select=id,subject,from,toRecipients,body,receivedDateTime,sentDateTime,isRead,categories,hasAttachments&$top=50&$orderby=receivedDateTime desc`;
        const res = await axios.get(url, { headers: { Authorization: `Bearer ${token}` }, timeout: 10000 });
        const sinceDate = new Date(since);
        const msgs = (res.data?.value || [])
          .filter(m => {
            const msgDate = new Date(m.receivedDateTime || m.sentDateTime || 0);
            if (msgDate < sinceDate) return false; // too old
            // Skip only if already tagged Aurora Processed
            const cats = m.categories || [];
            if (cats.includes('Aurora Processed')) return false;
            return true; // process ALL emails not yet tagged — read or unread
          })
          .map(m => ({ ...m, folder }));
        allMessages = allMessages.concat(msgs);
        console.log(`[Poll] ${folder}: ${msgs.length} email(s) to process (last 7 days, not yet tagged)`);
      } catch(folderErr) {
        console.error(`[Poll] Error reading ${folder}:`, folderErr.response?.data?.error?.message || folderErr.message);
      }
    }

    if (!allMessages.length) {
      console.log('[Poll] No new emails to process — all recent emails already tagged Aurora Processed');
      return;
    }
    console.log(`[Poll] Processing ${allMessages.length} email(s) total`);

    const projects = await db.getProjects();

    for (const msg of allMessages) {
      const fromEmail = msg.from?.emailAddress?.address || '';
      const fromName  = msg.from?.emailAddress?.name || fromEmail;
      const subject   = msg.subject || '';
      // Use 8000 chars for body — task IDs can be buried deep in quoted reply text
      const bodyText  = msg.body?.content?.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').slice(0, 8000) || '';

      console.log(`[Poll] Checking: "${subject}" from ${fromEmail}`);

      // ── Check for internal task confirmations FIRST — before any skip logic ──
      // Staff reply to Aurora emails from their @risk2solution.com addresses
      // The task ID (chk_ip_xxx) is in the quoted original email body
      const wasInternalConfirmation = await checkInternalTaskConfirmations(bodyText, subject, fromEmail);
      if (wasInternalConfirmation) {
        console.log(`[Poll] Internal task confirmed by ${fromEmail} — tagged and skipped`);
        try {
          await axios.patch(`https://graph.microsoft.com/v1.0/users/${mailbox}/messages/${msg.id}`,
            { isRead: true, categories: ['Aurora Processed'] },
            { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }, timeout: 5000 }
          );
        } catch(e) {}
        continue;
      }

      // Skip Aurora's own emails — prevents infinite feedback loops
      const isXeroInvoice = fromEmail.toLowerCase().includes('xero.com') ||
                            fromEmail.toLowerCase().includes('post.xero.com');

      if (subject.startsWith('[Aurora]') || fromEmail.toLowerCase() === (process.env.OUTLOOK_SHARED_MAILBOX || 'info@risk2solution.com').toLowerCase()) {
        console.log(`[Poll] Skipping — Aurora's own email or self-sent`);
        try {
          await axios.patch(
            `https://graph.microsoft.com/v1.0/users/${mailbox}/messages/${msg.id}`,
            { isRead: true, categories: ['Aurora Processed'] },
            { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }, timeout: 5000 }
          );
        } catch(e) {}
        continue;
      }

      // ── Handle Xero invoice emails BEFORE tagging ─────────────────────────
      if (isXeroInvoice) {
        console.log(`[Poll] Xero invoice detected: ${subject}`);
        const xeroInvNum = subject.match(/Invoice\s+(INV-\d+)/i) || bodyText.match(/invoice\s+(INV-\d+)/i);
        const xeroAmount = bodyText.match(/for\s+\$AUD\s+([\d,]+(?:\.\d{2})?)/i) ||
                           bodyText.match(/\$AUD\s+([\d,]+(?:\.\d{2})?)/i) ||
                           bodyText.match(/AUD\s+([\d,]+(?:\.\d{2})?)/i);
        const xeroClient = subject.match(/for\s+(.+)$/i)?.[1]?.trim();
        console.log(`[Poll] Xero: num=${xeroInvNum?.[1]} amount=${xeroAmount?.[1]} client=${xeroClient}`);

        if (xeroInvNum && xeroAmount && xeroClient) {
          const invoiceNum = xeroInvNum[1].toUpperCase();
          const amount = xeroAmount[1].replace(/,/g,'');
          const xeroProject = projects.find(p =>
            p.type === 'standard' && p.clientName &&
            (xeroClient.toLowerCase().includes(p.clientName.toLowerCase().slice(0,8)) ||
             p.clientName.toLowerCase().includes(xeroClient.toLowerCase().slice(0,8)))
          );
          console.log(`[Poll] Xero project match: ${xeroProject?.clientName || 'NONE'}`);
          if (xeroProject) {
            try {
              const existing = await readInvoices(xeroProject.id);
              if (!existing.find(i => i.invoiceNumber === invoiceNum)) {
                existing.push({ id:`inv_${Date.now()}`, invoiceNumber:invoiceNum, amount, sentDate:new Date().toISOString().slice(0,10), source:'xero', paid:false, createdAt:new Date().toISOString() });
                await writeInvoices(xeroProject.id, existing);
                await db.logActivity(xeroProject.id, { type:'contract', summary:`Invoice ${invoiceNum} sent via Xero — $${parseFloat(amount).toLocaleString('en-AU')} to ${xeroClient}` });
                await sendEmail('diane.k@risk2solution.com',
                  `[Aurora] Invoice recorded: ${invoiceNum} for ${xeroProject.clientName}`,
                  `Aurora detected a Xero invoice sent to ${xeroClient}.\n\nInvoice: ${invoiceNum}\nAmount: $${parseFloat(amount).toLocaleString('en-AU')} AUD\nDue: ${bodyText.match(/due by (.+?)[\.\r\n]/i)?.[1] || 'See Xero'}\n\nThis has been recorded in the ${xeroProject.clientName} project invoicing tab in Aurora.\n\nAurora\nR2S Project Management Intelligence`,
                  true
                );
                console.log(`[Invoice] Xero invoice ${invoiceNum} $${amount} recorded for ${xeroProject.clientName}`);
              } else {
                console.log(`[Invoice] Already recorded: ${invoiceNum}`);
              }
            } catch(xeroErr) { console.error('[Xero Invoice]', xeroErr.message); }
          } else {
            await sendEmail('diane.k@risk2solution.com',
              `[Aurora] Xero invoice — no project match: ${subject}`,
              `Aurora detected a Xero invoice but could not match it to a project.\n\nInvoice: ${xeroInvNum[1]}\nClient in Xero: ${xeroClient}\nAmount: $${parseFloat(amount).toLocaleString('en-AU')} AUD\n\nPlease check the project exists in Aurora with the correct client name.\n\nAurora\nR2S Project Management Intelligence`,
              true
            );
          }
        } else {
          console.log(`[Poll] Xero invoice parse failed — num:${!!xeroInvNum} amount:${!!xeroAmount} client:${!!xeroClient}`);
        }
        // Tag and skip normal processing
        try {
          await axios.patch(`https://graph.microsoft.com/v1.0/users/${mailbox}/messages/${msg.id}`,
            { isRead: true, categories: ['Aurora Processed'] },
            { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }, timeout: 5000 }
          );
        } catch(e) {}
        continue;
      }

      // Tag immediately as seen — prevents re-processing on next poll
      try {
        await axios.patch(
          `https://graph.microsoft.com/v1.0/users/${mailbox}/messages/${msg.id}`,
          { categories: ['Aurora Processed'] },
          { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }, timeout: 5000 }
        );
      } catch(tagErr) {}

      // Determine if this is internal R2S or external (client/other)
      const isInternal = fromEmail.toLowerCase().endsWith('@risk2solution.com') ||
                         fromEmail.toLowerCase().endsWith('@presilience.com');
      const isFromDiane = fromEmail.toLowerCase() === 'diane.k@risk2solution.com';

      // ── Special case: Diane forwarding a new agreement/contract ─────────────
      // Catches: attachment + agreement keywords, OR body mentions a new client/org
      const hasAttachment = msg.hasAttachments || false;
      const mentionsNewProject = /agreement|contract|proposal|signed|new client|new project|new engagement/i.test(subject + ' ' + bodyText.slice(0, 500));

      // Catch agreement emails from Diane — with OR without attachment flag
      // (Graph API hasAttachments can be unreliable for forwarded emails)
      const isDianeAgreement = isFromDiane && mentionsNewProject;

      if (isDianeAgreement) {
        console.log(`[Poll] Agreement email from Diane — attempting to auto-create project from attachment`);
        let projectCreated = false;

        // Try to download and process the PDF attachment automatically
        if (hasAttachment) {
          try {
            // Get attachments list
            const attRes = await axios.get(
              `https://graph.microsoft.com/v1.0/users/${mailbox}/messages/${msg.id}/attachments`,
              { headers: { Authorization: `Bearer ${token}` }, timeout: 15000 }
            );
            const attachments = attRes.data?.value || [];
            const pdfAtt = attachments.find(a =>
              a.contentType?.includes('pdf') ||
              a.name?.toLowerCase().endsWith('.pdf') ||
              a.contentType?.includes('word') ||
              a.name?.toLowerCase().endsWith('.docx')
            );

            if (pdfAtt && pdfAtt.contentBytes) {
              console.log(`[Poll] Found attachment: ${pdfAtt.name} (${pdfAtt.size} bytes)`);
              // Decode base64 attachment
              const pdfBuffer = Buffer.from(pdfAtt.contentBytes, 'base64');

              // Extract text from PDF using pdf-parse
              let rawText = '';
              try {
                const pdfParse = require('pdf-parse');
                const parsed = await pdfParse(pdfBuffer);
                rawText = parsed.text || '';
                console.log(`[Poll] Extracted ${rawText.length} chars from ${pdfAtt.name}`);
              } catch(pdfErr) {
                console.error('[Poll] PDF parse error:', pdfErr.message);
                // Try treating as text if not a real PDF
                rawText = pdfBuffer.toString('utf8').slice(0, 20000);
              }

              if (rawText.length > 100) {
                console.log(`[Poll] Running contract analysis on ${pdfAtt.name}...`);
                // Run contract analysis with timeout protection
                const extracted = await Promise.race([
                  analyseContract(rawText, pdfAtt.name),
                  new Promise((_, reject) => setTimeout(() => reject(new Error('Analysis timeout')), 90000))
                ]);
                if (extracted && extracted.clientName) {
                  // Check for duplicate — must match BOTH client name AND project name
                  // Same client can have multiple different projects
                  const existing = projects.find(p => {
                    const sameClient = p.clientName?.toLowerCase().includes(extracted.clientName.toLowerCase().slice(0,8)) ||
                      extracted.clientName.toLowerCase().includes((p.clientName||'').toLowerCase().slice(0,8));
                    if (!sameClient) return false;
                    // If same client, also check project name similarity
                    const extractedProject = (extracted.projectName || '').toLowerCase();
                    const existingProject = (p.projectName || '').toLowerCase();
                    if (!extractedProject || !existingProject) return false; // can't confirm duplicate without project names
                    // Only flag as duplicate if project names are also similar
                    return extractedProject.slice(0,15) === existingProject.slice(0,15) ||
                      existingProject.includes(extractedProject.slice(0,12)) ||
                      extractedProject.includes(existingProject.slice(0,12));
                  });

                  if (!existing) {
                    const projectId = `p_${Date.now()}`;
                    const project = {
                      id: projectId, type: 'standard',
                      ...extracted,
                      status: 'Active', phase: 0,
                      createdAt: new Date().toISOString(),
                      updatedAt: new Date().toISOString(),
                    };
                    await db.upsertProject(project);
                    await db.logActivity(projectId, { type: 'project_created', summary: `Project auto-created from email attachment: ${pdfAtt.name} sent by Diane` });
                    await draftClientOnboarding(project);
                    projectCreated = true;
                    console.log(`[Poll] ✓ Auto-created project: ${extracted.clientName}`);

                    await sendEmail('diane.k@risk2solution.com',
                      `[Aurora] New project created: ${extracted.clientName}`,
                      `Aurora automatically created a new project from the attachment in your email.\n\nClient: ${extracted.clientName}\nProject: ${extracted.projectName || 'See Aurora'}\nValue: ${extracted.value || 'TBC'}\nConsultant identified: ${extracted.consultant || 'Not specified'}\n\nThe project has been added to Aurora. A client onboarding email draft is ready for your review.\n\n${process.env.FRONTEND_URL || ''}\n\nAurora\nR2S Project Management Intelligence`,
                      true
                    );
                  } else {
                    console.log(`[Poll] Duplicate detected — project already exists: ${existing.clientName}`);
                    await sendEmail('diane.k@risk2solution.com',
                      `[Aurora] Duplicate project detected: ${extracted.clientName}`,
                      `Aurora detected a possible duplicate when processing the attachment from your email.\n\nAttachment: ${pdfAtt.name}\nDetected client: ${extracted.clientName}\nExisting project: ${existing.clientName} (${existing.status})\n\nPlease review in Aurora to confirm.\n\n${process.env.FRONTEND_URL || ''}\n\nAurora\nR2S Project Management Intelligence`,
                      true
                    );
                    projectCreated = true; // Suppress fallback prompt
                  }
                }
              }
            }
          } catch(attErr) {
            console.error('[Poll] Attachment processing error:', attErr.message);
          }
        }

        // Fallback — if no attachment processed, prompt Diane to upload manually
        if (!projectCreated) {
          const clientHint = subject.replace(/new (project|agreement|proposal) (for )?/i,'').trim();
          const valueHint = bodyText.match(/\$([\d,]+)/)?.[0] || '';
          await sendEmail('diane.k@risk2solution.com',
            `[Aurora] New project detected — please upload contract: ${clientHint}`,
            `Hi Diane,\n\nAurora detected a new client agreement but could not automatically process the attachment.\n\nSubject: ${subject}${valueHint ? '\nValue mentioned: ' + valueHint : ''}\n\nPlease upload the contract manually:\n${process.env.FRONTEND_URL || ''}\n→ Projects → New Project → Upload Contract\n\nAurora\nR2S Project Management Intelligence`,
            true
          );
        }

        try {
          await axios.patch(
            `https://graph.microsoft.com/v1.0/users/${mailbox}/messages/${msg.id}`,
            { isRead: true, categories: ['Aurora Processed'] },
            { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }, timeout: 5000 }
          );
        } catch(e) {}
        continue;
      }

      // Find ALL matching projects — check subject, body text, AND email addresses/domains in thread
      const bodyLower    = bodyText.toLowerCase();
      const subjectLower = subject.toLowerCase();
      const fullText     = subjectLower + ' ' + bodyLower;

      // Extract all email addresses from the body (catches forwarded thread recipients)
      const emailsInBody = [...bodyText.matchAll(/[\w.+-]+@[\w-]+\.[\w.]+/g)].map(m => m[0].toLowerCase());
      const domainsInBody = [...new Set(emailsInBody.map(e => e.split('@')[1]).filter(Boolean))];

      const matchedProjects = projects.filter(p => {
        if (p.type !== 'standard' || !p.clientName) return false;
        const clientNameLower = p.clientName.toLowerCase();
        const projectNameLower = (p.projectName || '').toLowerCase();
        const clientEmailLower = (p.clientEmail || '').toLowerCase();
        const clientDomain = clientEmailLower.includes('@') ? clientEmailLower.split('@')[1] : '';

        // Match by client name in full text
        if (fullText.includes(clientNameLower)) return true;
        // Match by project name in full text
        if (projectNameLower.length > 4 && fullText.includes(projectNameLower.slice(0,15))) return true;
        // Match by client email domain appearing in body (catches forwarded threads)
        if (clientDomain && domainsInBody.includes(clientDomain)) return true;
        // Match by exact client email in body
        if (clientEmailLower && emailsInBody.includes(clientEmailLower)) return true;
        return false;
      });

      console.log(`[Poll] Project matches: ${matchedProjects.length > 0 ? matchedProjects.map(p => p.clientName).join(', ') : 'NONE'}  Internal: ${isInternal}`);


      // If no project match — tag and optionally alert Diane
      if (!matchedProjects.length) {
        // Expanded noise filter — anything that's clearly not a project update
        const isNoise = /remittance|payment received|invoice|survey|notification|enquiry form|contact us form|abandoned call|missed call|tender|digest|purchase order|fmclarity|localbuy|vendorpanel|supabase|work order|unsubscribe|auto.?reply|out of office|no.?reply|donotreply|do.not.reply|statement|receipt|confirmation/i.test(subject) ||
          /remittance|payment received|contact us form|enquiry form|abandoned call|missed call|work order/i.test(bodyText.slice(0, 200));

        // Only alert Diane if: internal R2S sender + not noise + looks like genuine project content
        const looksLikeProject = bodyText.length > 150 && (
          /project|client|proposal|contract|deliverable|training|workshop|report|phase|scope|invoice milestone|consultant/i.test(bodyText)
        );

        if (isInternal && !isNoise && looksLikeProject) {
          await sendEmail('diane.k@risk2solution.com',
            `[Aurora] R2S staff email — no project match: ${subject}`,
            `Aurora received an email from ${fromName} (${fromEmail}) that may relate to a project but could not be matched.

Subject: ${subject}

Summary:
${bodyText.slice(0, 400)}

If this relates to a project, please create it in Aurora.

Aurora
R2S Project Management Intelligence`,
            true
          );
        }
        try {
          await axios.patch(
            `https://graph.microsoft.com/v1.0/users/${mailbox}/messages/${msg.id}`,
            { categories: ['Aurora Processed'] },
            { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }, timeout: 5000 }
          );
        } catch(e) {}
        continue;
      }

      // Process EACH matched project — one email may reference multiple projects
      for (const matchedProject of matchedProjects) {
        console.log(`[Poll] Processing for: ${matchedProject.clientName}`);

      // Full autonomous analysis
      try {
        const analysis = await aurora('contract_extract',
          `Analyse this email and return a JSON object. Return ONLY the JSON — no explanation, no preamble, no markdown, just the raw JSON object starting with { and ending with }.

Project context:
- Project: ${matchedProject.projectName || matchedProject.clientName}
- Client: ${matchedProject.clientName}
- Current phase: ${PHASES[matchedProject.phase || 0]} (index: ${matchedProject.phase || 0})
- Consultant: ${matchedProject.consultant || 'Unknown'}
- Deliverables: ${(matchedProject.deliverables || '').slice(0, 200)}

Email:
From: ${fromEmail}
Subject: ${subject}
Content: ${bodyText.slice(0, 2000)}

PHASE CHANGE RULES — be proactive, not conservative:
- If email says work is "commencing", "starting", "about to begin", "wanting to start" → move to Deployment (phase 1)
- If email says work is "underway", "in progress", "delivering", "on site" → move to Deployment (phase 1)  
- If email says report is "submitted", "sent to client", "delivered", "complete", "finalised" → move to Reporting (phase 3)
- If email says "final report sent", "all deliverables complete", "wrapping up" → move to Close-out (phase 4)
- If email says "waiting", "delayed", "on hold", "postponed" → set newStatus to "On Hold"
- Only keep current phase if email has no project progress information

INVOICE RULES — trigger if any of these are mentioned:
- Delivery of any training session or workshop
- Submission of any report or deliverable
- Project commencement (first milestone)
- Project completion

Return this exact JSON (no other text):
{"phaseChange":false,"newPhase":${matchedProject.phase || 0},"newStatus":"","statusSummary":"summary here","completedDeliverables":[],"inProgressDeliverables":[],"invoiceTriggered":false,"invoiceNote":"","needsKickoffScheduling":false,"kickoffNote":"","hasBookableEvent":false,"eventDate":"","eventTime":"","eventType":"","eventTitle":"","eventDuration":60,"requiresAttention":false,"attentionReason":"","activityLogEntry":"summary of what happened"}`,
          buildContext(matchedProject)
        );

        // Extract JSON from response — try multiple strategies
        let parsed;
        try {
          let clean = analysis.trim();
          // Remove markdown code blocks
          clean = clean.replace(/```json\s*/gi, '').replace(/```\s*/g, '').trim();
          // Find JSON object if surrounded by other text
          const jsonMatch = clean.match(/\{[\s\S]*\}/);
          if (jsonMatch) clean = jsonMatch[0];
          console.log(`[Replies] Parsing analysis for ${matchedProject.clientName}:`, clean.slice(0, 150));
          parsed = JSON.parse(clean);
          console.log(`[Replies] Analysis: phaseChange=${parsed.phaseChange}, newPhase=${parsed.newPhase}, invoice=${parsed.invoiceTriggered}, kickoff=${parsed.needsKickoffScheduling}`);
        } catch(jsonErr) {
          console.error('[Replies] JSON parse failed. Raw:', analysis.slice(0, 400));
          // Notify Diane manually
          await sendEmail('diane.k@risk2solution.com',
            `[Aurora] Email received — needs manual review: ${matchedProject.clientName}`,
            `Aurora received an email from ${fromEmail} about ${matchedProject.projectName || matchedProject.clientName} but could not automatically analyse it.\n\nSubject: ${subject}\n\nContent:\n${bodyText.slice(0, 1000)}\n\nPlease review and update Aurora project record manually.\n\nAurora\nR2S Project Management Intelligence`,
            true
          );
          await db.logActivity(matchedProject.id, { type: 'email_processed', source: fromEmail, subject, summary: `Email from ${fromEmail} received — manual review needed` });
          continue;
        }

        console.log(`[Replies] Analysis complete for ${matchedProject.clientName}: phaseChange=${parsed.phaseChange}, newPhase=${parsed.newPhase}, invoiceTriggered=${parsed.invoiceTriggered}, needsKickoff=${parsed.needsKickoffScheduling}, hasEvent=${parsed.hasBookableEvent}`);

        const actions = [];

        // ── 1. Update project phase ───────────────────────────────────────────
        if (parsed.phaseChange && typeof parsed.newPhase === 'number' && parsed.newPhase !== (matchedProject.phase || 0)) {
          await db.updateProjectField(matchedProject.id, { phase: parsed.newPhase });
          matchedProject.phase = parsed.newPhase;
          actions.push(`Phase updated: ${PHASES[parsed.newPhase - 1] || 'Kick-off'} → ${PHASES[parsed.newPhase]}`);
          console.log(`[Replies] Phase updated: ${matchedProject.clientName} → ${PHASES[parsed.newPhase]}`);
          // Add suggestion so Diane can revert if Aurora got it wrong
          await db.saveSuggestion({
            id: `sug_phase_${matchedProject.id}_${Date.now()}`,
            projectId: matchedProject.id,
            clientName: matchedProject.clientName,
            projectName: matchedProject.projectName,
            type: 'phase_advance',
            title: `Aurora moved ${matchedProject.clientName} to ${PHASES[parsed.newPhase]}`,
            reason: `Based on an email from ${fromEmail}, Aurora automatically updated this project to ${PHASES[parsed.newPhase]}. Please confirm this is correct, or dismiss to revert.`,
            action: { phase: parsed.newPhase },
            confirmLabel: 'Confirmed — keep this phase',
            dismissLabel: 'Revert to previous phase',
          });
        }

        // ── 2. Update project status ──────────────────────────────────────────
        if (parsed.newStatus && parsed.newStatus !== matchedProject.status) {
          await db.updateProjectField(matchedProject.id, { status: parsed.newStatus });
          actions.push(`Status updated to: ${parsed.newStatus}`);
        }

        // ── 3. Update deliverable statuses ────────────────────────────────────
        if (parsed.completedDeliverables?.length || parsed.inProgressDeliverables?.length) {
          const deliverables = await db.getDeliverables(matchedProject.id);
          if (deliverables.length > 0) {
            for (const delName of (parsed.completedDeliverables || [])) {
              const match = deliverables.find(d =>
                d.name.toLowerCase().includes(delName.toLowerCase().slice(0,12)) ||
                delName.toLowerCase().includes(d.name.toLowerCase().slice(0,12))
              );
              if (match) {
                await db.updateDeliverable(matchedProject.id, match.id, { status: 'Complete' });
                actions.push(`Deliverable marked complete: ${match.name}`);
              }
            }
            for (const delName of (parsed.inProgressDeliverables || [])) {
              const match = deliverables.find(d =>
                d.name.toLowerCase().includes(delName.toLowerCase().slice(0,12)) ||
                delName.toLowerCase().includes(d.name.toLowerCase().slice(0,12))
              );
              if (match && match.status !== 'Complete') {
                await db.updateDeliverable(matchedProject.id, match.id, { status: 'In Progress' });
                actions.push(`Deliverable in progress: ${match.name}`);
              }
            }
          }
        }

        // ── 4. Invoice trigger ────────────────────────────────────────────────
        if (parsed.invoiceTriggered) {
          actions.push(`Invoice milestone reached: ${parsed.invoiceNote}`);
          await sendEmail('diane.k@risk2solution.com',
            `[Aurora] Invoice milestone: ${matchedProject.clientName}`,
            `Aurora detected an invoice milestone on the ${matchedProject.projectName || matchedProject.clientName} project.\n\nMilestone: ${parsed.invoiceNote}\n\nPlease review the invoicing status in Aurora and issue the relevant invoice to the client.\n\n${process.env.FRONTEND_URL || ''}\n\nAurora\nR2S Project Management Intelligence`,
            true
          );
        }

        // ── 4b. Auto-detect Xero/invoice emails and record them ──────────────────
        const invNumMatch = bodyText.match(/(?:invoice\s*(?:number|no\.?|#:?|#)\s*)([A-Z]{0,5}-?\d{3,})/i) ||
                            bodyText.match(/\b(INV-\d+)\b/i);
        const invAmtMatch = bodyText.match(/\$([\d,]+(?:\.\d{2})?)\s*(?:AUD)?/i) ||
                            bodyText.match(/([\d,]+(?:\.\d{2})?)\s*AUD\b/i);
        if (invNumMatch && invAmtMatch) {
          try {
            const invoiceNum = invNumMatch[1].toUpperCase();
            const amount     = invAmtMatch[1].replace(/,/g,'');
            const existing   = await readInvoices(matchedProject.id);
            if (!existing.find(i => i.invoiceNumber === invoiceNum)) {
              const invoice = { id:`inv_${Date.now()}`, invoiceNumber:invoiceNum, amount, sentDate:new Date().toISOString().slice(0,10), source:'email', paid:false, createdAt:new Date().toISOString() };
              existing.push(invoice);
              await writeInvoices(matchedProject.id, existing);
              actions.push(`Invoice ${invoiceNum} for $${parseFloat(amount).toLocaleString('en-AU')} detected and recorded`);
              await db.logActivity(matchedProject.id, { type:'contract', summary:`Invoice ${invoiceNum} detected — $${parseFloat(amount).toLocaleString('en-AU')} sent to client` });
              console.log(`[Invoice] Auto-recorded: ${invoiceNum} $${amount} for ${matchedProject.clientName}`);
            }
          } catch(invErr) { console.error('[Invoice detect]', invErr.message); }
        }

        // ── 5. Kick-off scheduling ────────────────────────────────────────────
        if (parsed.needsKickoffScheduling) {
          try {
            const availableDates = await getConsultantAvailability();
            let availText = 'Please check the training calendar for available times.';
            if (availableDates?.length) {
              availText = availableDates.map((d, i) => {
                const dateStr = d.toLocaleDateString('en-AU', { weekday: 'long', day: 'numeric', month: 'long' });
                const timeStr = d.toLocaleTimeString('en-AU', { hour: 'numeric', minute: '2-digit', hour12: true });
                return `Option ${i+1}: ${dateStr} at ${timeStr} AEST`;
              }).join('\n');
            }
            await sendEmail('diane.k@risk2solution.com',
              `[Aurora] Kick-off scheduling needed: ${matchedProject.clientName}`,
              `A kick-off meeting needs to be scheduled for the ${matchedProject.projectName || matchedProject.clientName} project.\n\nReason: ${parsed.kickoffNote}\n\nBased on the R2S Training & Education calendar, these slots appear available:\n\n${availText}\n\nPlease confirm with ${matchedProject.consultant || 'the consultant'} and schedule with the client.\n\nAurora\nR2S Project Management Intelligence`,
              true
            );
            actions.push(`Kick-off scheduling options sent to Diane`);
          } catch(kErr) { console.error('[Kickoff]', kErr.message); }
        }

        // ── 6. Calendar event from email ──────────────────────────────────────
        if (parsed.hasBookableEvent && parsed.eventDate) {
          try {
            const booking = {
              title: parsed.eventTitle || `${parsed.eventType} — ${matchedProject.clientName}`,
              description: `${parsed.statusSummary || ''}\n\nSource: Email from ${fromEmail}`,
              startDateTime: `${parsed.eventDate}T${parsed.eventTime || '09:00'}:00`,
              durationMinutes: parsed.eventDuration || 60,
              location: matchedProject.clientName,
              clientName: matchedProject.clientName,
              clientEmail: matchedProject.clientEmail,
              consultantName: matchedProject.consultant,
              consultantEmail: matchedProject.consultantEmail || 'info@risk2solution.com',
              projectId: matchedProject.id,
            };
            const eventId = await createCalendarBooking(booking, true);
            if (eventId) {
              await db.saveSuggestion({
                id: `sug_cal_${matchedProject.id}_${Date.now()}`,
                projectId: matchedProject.id,
                clientName: matchedProject.clientName,
                projectName: matchedProject.projectName,
                type: 'schedule_kickoff',
                title: `Tentative: ${booking.title} on ${parsed.eventDate}`,
                reason: `Aurora detected a scheduled ${parsed.eventType} in an email from ${fromEmail}. A tentative booking has been added to the R2S Training & Education calendar. Approve to send meeting invites.`,
                action: { calendarEventId: eventId, booking },
                confirmLabel: 'Confirm & send invites',
                dismissLabel: 'Cancel tentative booking',
              });
              actions.push(`Tentative calendar booking created: ${booking.title} on ${parsed.eventDate}`);
            }
          } catch(calErr) { console.error('[Calendar]', calErr.message); }
        }

        // ── 7. Log activity on project ────────────────────────────────────────
        await db.logActivity(matchedProject.id, {
          type: 'email_processed',
          source: fromEmail,
          subject,
          summary: parsed.activityLogEntry || parsed.statusSummary,
          actions,
        });

        // ── 8. Notify Diane with full summary ─────────────────────────────────
        const actionsText = actions.length
          ? `\nActions taken by Aurora:\n${actions.map(a => `• ${a}`).join('\n')}\n`
          : '\nNo automatic actions were taken.\n';

        await sendEmail('diane.k@risk2solution.com',
          `[Aurora] Email processed: ${matchedProject.clientName}${actions.length ? ` — ${actions.length} action${actions.length > 1 ? 's' : ''} taken` : ''}`,
          `Aurora has processed an email from ${fromEmail} regarding the ${matchedProject.projectName || matchedProject.clientName} project.\n\nSummary: ${parsed.statusSummary}\n${actionsText}\nPlease log into Aurora to review.\n\n${process.env.FRONTEND_URL || ''}\n\nAurora\nR2S Project Management Intelligence`,
          true
        );

        console.log(`[Replies] ✓ ${matchedProject.clientName}: ${actions.length} actions taken`);

      } catch (err) {
        if (err.message === 'MONTHLY_CAP_REACHED') break;
        console.error('[Replies] Analysis failed:', err.message);
      }

      // Mark as read and tag with Aurora Processed category (always, regardless of analysis result)
      try {
        await axios.patch(
          `https://graph.microsoft.com/v1.0/users/${mailbox}/messages/${msg.id}`,
          { isRead: true, categories: ['Aurora Processed'] },
          { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }, timeout: 5000 }
        );
      } catch(patchErr) {
        console.error('[Poll] Mark read/tag failed:', patchErr.message);
      }

      } // end for each matched project
      // Safety net — ensure email is tagged even if inner loop had issues

    } // end for each message
  } catch (err) {
    console.error('[Replies] Read failed:', err.message);
  }
}

// ── Weekly consultant check-in draft ─────────────────────────────────────────
async function generateWeeklyConsultantCheckins(projects) {
  const active = projects.filter(p =>
    p.type === 'standard' &&
    p.consultant &&
    !['Completed','Terminated','Closed'].includes(p.status) &&
    [1, 2, 3].includes(p.phase) // Deployment, Monitoring, Reporting phases only
  );

  for (const p of active) {
    try {
      const context = buildContext(p);
      const firstName = (p.consultant || '').split(' ')[0];
      const text = await aurora('checkin_email',
        `Draft a short weekly check-in email from Diane (R2S Project Manager) to ${p.consultant}, the assigned consultant/trainer on the ${p.projectName || p.clientName} project.

The email should:
- Be brief and friendly — 3-4 sentences max
- Ask for a quick update on progress against deliverables and timeline
- Ask if there are any issues, blockers, or anything they need from the PM
- Reference the specific project and any relevant deliverables or milestones if known
- Not repeat information they already know

Start with: Hi ${firstName},

Sign off as:
Kind regards,
Diane Kruger
Corporate Operations Lead | Risk 2 Solution Group`,
        context
      );

      const draft = {
        id: `d_${Date.now()}_${Math.random().toString(36).slice(2,6)}`,
        projectId: p.id, clientName: p.clientName, projectName: p.projectName,
        type: 'consultant_checkin', urgency: 'routine',
        toName: p.consultant, toEmail: p.consultantEmail || 'info@risk2solution.com',
        subject: `${p.projectName || p.clientName} — Weekly check-in`,
        body: text, source: 'batch',
      };
      await db.saveDraft(draft);
      await saveDraftEmail(draft);
    } catch (err) {
      if (err.message === 'MONTHLY_CAP_REACHED') break;
      console.error(`[Checkins] Error on ${p.clientName}:`, err.message);
    }
  }
  return active.length;
}

// ── Phase stuck too long detection ────────────────────────────────────────────
async function checkStuckPhases(projects) {
  const MAX_PHASE_DAYS = { 0: 7, 1: 60, 2: 60, 3: 21, 4: 14 }; // days per phase before flagging
  const active = projects.filter(p => p.type === 'standard' && !['Completed','Terminated'].includes(p.status));

  for (const p of active) {
    const updatedAt = p.updatedAt ? new Date(p.updatedAt) : null;
    if (!updatedAt) continue;
    const daysSinceUpdate = Math.round((new Date() - updatedAt) / (1000 * 60 * 60 * 24));
    const maxDays = MAX_PHASE_DAYS[p.phase || 0];
    if (daysSinceUpdate >= maxDays) {
      await sendEmail('diane.k@risk2solution.com',
        `[Aurora] Project phase check: ${p.clientName} — ${PHASES[p.phase||0]}`,
        `The ${p.clientName} project (${p.projectName || ''}) has been in the ${PHASES[p.phase||0]} phase for ${daysSinceUpdate} days without a recorded update in Aurora.

Please log into Aurora and update the project status or phase as appropriate.

Aurora
R2S Project Management Intelligence`,
        true
      );
    }
  }
}

// ── SOP-TRN-001: Materials submission reminder ────────────────────────────────
async function checkMaterialsSubmissionReminders(projects) {
  const active = projects.filter(p =>
    p.type === 'standard' && p.consultant && p.dueDate &&
    p.phase === 1 && // Deployment phase
    !['Completed','Terminated'].includes(p.status)
  );

  for (const p of active) {
    const dueDate = new Date(p.dueDate);
    const days = Math.round((dueDate - new Date()) / (1000 * 60 * 60 * 24));
    // Remind at 4 days before due date (allows 2 days for COO review + 2 buffer)
    if (days === 4) {
      const consultantFirst = (p.consultant || '').split(' ')[0];
      await sendEmail(p.consultantEmail || 'info@risk2solution.com',
        `[Aurora] Materials submission reminder: ${p.clientName}`,
        `Hi ${consultantFirst},

This is a reminder that all training and consulting materials for the ${p.projectName || p.clientName} project must be submitted to the COO for approval no later than 2 business days before delivery.

Project due date: ${p.dueDate}
Deadline for materials submission: Please ensure materials are submitted immediately to allow time for COO review.

Please ensure all materials are:
• Fully customised to ${p.clientName}
• Client-ready (not draft)
• Submitted via the SharePoint project folder

Contact Diane if you have any questions.

Kind regards,
Diane Kruger
Corporate Operations Lead | Risk 2 Solution Group`,
        false,
        'diane.k@risk2solution.com'
      );

      // Also prompt COO (Diane) to expect materials
      await sendEmail('diane.k@risk2solution.com',
        `[Aurora] COO approval needed soon: ${p.clientName}`,
        `Materials for the ${p.clientName} project (${p.projectName || ''}) should be submitted by ${p.consultant} for your approval within the next 1-2 days.

Project due date: ${p.dueDate}

Please allow time in your schedule to review and approve before delivery.

Aurora
R2S Project Management Intelligence`,
        true
      );
    }

    // Session report reminder — 2 days after due date (post-delivery)
    if (days === -2 && p.consultant) {
      const consultantFirst = (p.consultant || '').split(' ')[0];
      await sendEmail(p.consultantEmail || 'info@risk2solution.com',
        `[Aurora] Session report due: ${p.clientName}`,
        `Hi ${consultantFirst},

This is a reminder that your session report for the ${p.projectName || p.clientName} engagement is due today (within 2 business days of delivery).

Your session report should include:
• Session overview
• Key observations
• Identified gaps
• Recommendations
• Any follow-up actions
• Any off-scope items raised
• Any incidents or issues

Please submit your completed report to Diane and save it to the SharePoint project folder: 07 Session Reports

Kind regards,
Diane Kruger
Corporate Operations Lead | Risk 2 Solution Group`,
        false,
        'diane.k@risk2solution.com'
      );
    }
  }
}

// ── Suggestion engine ────────────────────────────────────────────────────────
// Aurora autonomously identifies actions and surfaces them for Diane to approve

const SUGGESTION_TYPES = {
  PHASE_ADVANCE:    'phase_advance',
  PHASE_REGRESS:    'phase_regress',
  STATUS_CHANGE:    'status_change',
  SEND_CLIENT_UPDATE: 'send_client_update',
  ESCALATE:         'escalate',
  CLOSE_OUT:        'close_out',
  SCHEDULE_KICKOFF: 'schedule_kickoff',
  REQUEST_REPORT:   'request_report',
};

async function generateSuggestions() {
  const projects = await db.getProjects();
  const standard = projects.filter(p => p.type === 'standard');
  const suggestions = [];

  for (const p of standard) {
    const phase = p.phase || 0;
    const status = p.status || 'Active';
    const days  = p.dueDate ? Math.round((new Date(p.dueDate) - new Date()) / (1000*60*60*24)) : null;
    const updatedDaysAgo = p.updatedAt ? Math.round((new Date() - new Date(p.updatedAt)) / (1000*60*60*24)) : 999;

    // ── Phase 4 (Close-out) — trigger all close-out actions ──────────────────
    if (phase === 4) {
      // Invoice check
      suggestions.push(await db.saveSuggestion({
        id: `sug_${p.id}_inv_${Date.now()}`,
        projectId: p.id, clientName: p.clientName, projectName: p.projectName,
        type: SUGGESTION_TYPES.SEND_CLIENT_UPDATE,
        title: `Check invoicing status: ${p.clientName}`,
        reason: `${p.clientName} is in Close-out. Please confirm all invoices have been issued as per the payment schedule: ${p.invoicingNotes || 'see contract'}. Aurora will draft a final invoice if needed.`,
        action: { draftInvoice: true },
        confirmLabel: 'Draft final invoice email',
        dismissLabel: 'Invoicing complete',
      }));

      // Client satisfaction email
      suggestions.push(await db.saveSuggestion({
        id: `sug_${p.id}_feedback_${Date.now()}`,
        projectId: p.id, clientName: p.clientName, projectName: p.projectName,
        type: SUGGESTION_TYPES.SEND_CLIENT_UPDATE,
        title: `Send client satisfaction email: ${p.clientName}`,
        reason: `${p.clientName} project is at Close-out. A client satisfaction and feedback email should be sent. Aurora has a draft ready for Diane to review.`,
        action: { draftFeedback: true },
        confirmLabel: 'Draft feedback email',
        dismissLabel: 'Already sent',
      }));

      // Final report check
      suggestions.push(await db.saveSuggestion({
        id: `sug_${p.id}_report_check_${Date.now()}`,
        projectId: p.id, clientName: p.clientName, projectName: p.projectName,
        type: SUGGESTION_TYPES.REQUEST_REPORT,
        title: `Confirm final deliverables sent: ${p.clientName}`,
        reason: `${p.clientName} is in Close-out. Please confirm all reports, materials, and deliverables have been sent to the client and saved in the SharePoint project folder.`,
        action: null,
        confirmLabel: 'Confirmed — all sent',
        dismissLabel: 'Outstanding items remain',
      }));
    }

    // Skip further checks for completed/terminated projects
    if (['Completed','Terminated','Closed'].includes(status)) continue;

    // ── Phase 0 (Kick-off) ────────────────────────────────────────────────────
    if (phase === 0 && updatedDaysAgo >= 5 && p.consultant) {
      suggestions.push(await db.saveSuggestion({
        id: `sug_${p.id}_kickoff_${Date.now()}`,
        projectId: p.id, clientName: p.clientName, projectName: p.projectName,
        type: SUGGESTION_TYPES.PHASE_ADVANCE,
        title: `Move ${p.clientName} to Deployment phase`,
        reason: `This project has been in Kick-off for ${updatedDaysAgo} days. If the kick-off meeting has occurred and work has commenced, it should move to Deployment.`,
        action: { phase: 1 },
        confirmLabel: 'Move to Deployment',
        dismissLabel: 'Keep in Kick-off',
      }));
    }

    // ── Phase 1 (Deployment) → Phase 2 (Monitoring) ──────────────────────────
    if (phase === 1 && updatedDaysAgo >= 45) {
      suggestions.push(await db.saveSuggestion({
        id: `sug_${p.id}_deploy_${Date.now()}`,
        projectId: p.id, clientName: p.clientName, projectName: p.projectName,
        type: SUGGESTION_TYPES.PHASE_ADVANCE,
        title: `Move ${p.clientName} to Monitoring & Review`,
        reason: `This project has been in Deployment for ${updatedDaysAgo} days. If the primary service delivery is complete, it should move to Monitoring & Review.`,
        action: { phase: 2 },
        confirmLabel: 'Move to Monitoring & Review',
        dismissLabel: 'Keep in Deployment',
      }));
    }

    // ── Phase 2 (Monitoring) → Phase 3 (Reporting) ───────────────────────────
    if (phase === 2 && updatedDaysAgo >= 30) {
      suggestions.push(await db.saveSuggestion({
        id: `sug_${p.id}_monitor_${Date.now()}`,
        projectId: p.id, clientName: p.clientName, projectName: p.projectName,
        type: SUGGESTION_TYPES.PHASE_ADVANCE,
        title: `Move ${p.clientName} to Reporting phase`,
        reason: `This project has been in Monitoring & Review for ${updatedDaysAgo} days. If monitoring is complete, move to Reporting to finalise deliverables.`,
        action: { phase: 3 },
        confirmLabel: 'Move to Reporting',
        dismissLabel: 'Keep in Monitoring',
      }));
    }

    // ── Phase 3 (Reporting) → Phase 4 (Close-out) ────────────────────────────
    if (phase === 3 && updatedDaysAgo >= 14) {
      suggestions.push(await db.saveSuggestion({
        id: `sug_${p.id}_report_${Date.now()}`,
        projectId: p.id, clientName: p.clientName, projectName: p.projectName,
        type: SUGGESTION_TYPES.CLOSE_OUT,
        title: `Close out ${p.clientName} project`,
        reason: `This project has been in Reporting for ${updatedDaysAgo} days. If all reports and deliverables are complete, it is ready for Close-out.`,
        action: { phase: 4 },
        confirmLabel: 'Move to Close-out',
        dismissLabel: 'Not ready yet',
      }));
    }

    // ── Due within 7 days and not in Reporting/Close-out ─────────────────────
    if (days !== null && days <= 7 && days >= 0 && phase < 3) {
      suggestions.push(await db.saveSuggestion({
        id: `sug_${p.id}_due_${Date.now()}`,
        projectId: p.id, clientName: p.clientName, projectName: p.projectName,
        type: SUGGESTION_TYPES.ESCALATE,
        title: `${p.clientName} due in ${days} day${days!==1?'s':''}`,
        reason: `This project is due in ${days} days but is still in ${PHASES[phase]} phase. Aurora will draft an escalation email for Diane to review and send.`,
        action: { draftEscalation: true, riskDescription: `Project due in ${days} days but currently in ${PHASES[phase]} phase with deliverables potentially outstanding.` },
        confirmLabel: 'Draft escalation email',
        dismissLabel: 'Acknowledged',
      }));
    }

    // ── On Hold for 14+ days ──────────────────────────────────────────────────
    if (p.status === 'On Hold' && updatedDaysAgo >= 14) {
      suggestions.push(await db.saveSuggestion({
        id: `sug_${p.id}_hold_${Date.now()}`,
        projectId: p.id, clientName: p.clientName, projectName: p.projectName,
        type: SUGGESTION_TYPES.STATUS_CHANGE,
        title: `${p.clientName} has been On Hold for ${updatedDaysAgo} days`,
        reason: `This project has been on hold for ${updatedDaysAgo} days. Consider following up with the client or updating the status.`,
        action: null,
        confirmLabel: 'Draft client follow-up',
        dismissLabel: 'Acknowledged',
        confirmTaskType: 'checkin_email',
      }));
    }

    // ── Phase 4 (Close-out) → Phase 5 (Completed) ───────────────────────────────
    if (phase === 4 && updatedDaysAgo >= 3) {
      suggestions.push(await db.saveSuggestion({
        id: `sug_${p.id}_complete_${Date.now()}`,
        projectId: p.id, clientName: p.clientName, projectName: p.projectName,
        type: SUGGESTION_TYPES.CLOSE_OUT,
        title: `Mark ${p.clientName} as Completed`,
        reason: `This project is in Close-out. Once all close-out actions are done (invoices sent, feedback email sent, documents filed), mark it as Completed to move it to the completed projects archive.`,
        action: { phase: 5, status: 'Completed' },
        confirmLabel: 'Mark as Completed — archive project',
        dismissLabel: 'Still in progress',
      }));
    }

    // ── No consultant assigned and in Deployment/Monitoring ──────────────────
    if (!p.consultant && phase >= 1 && phase <= 3) {
      suggestions.push(await db.saveSuggestion({
        id: `sug_${p.id}_noconsult_${Date.now()}`,
        projectId: p.id, clientName: p.clientName, projectName: p.projectName,
        type: SUGGESTION_TYPES.STATUS_CHANGE,
        title: `No consultant assigned to ${p.clientName}`,
        reason: `This project is in ${PHASES[phase]} phase but has no consultant or trainer assigned. Please assign one in the project record.`,
        action: null,
        confirmLabel: 'Open project to assign',
        dismissLabel: 'Acknowledged',
      }));
    }
  }

  // Filter out null suggestions (duplicates that returned existing)
  return suggestions.filter(Boolean);
}

// ── Apply a suggestion action ─────────────────────────────────────────────────
async function applySuggestion(suggestion) {
  const project = await db.getProject(suggestion.projectId);
  if (!project) { console.error(`[applySuggestion] Project not found: ${suggestion.projectId}`); return; }
  console.log(`[applySuggestion] Applying: ${suggestion.title} for ${project.clientName}`);

  if (suggestion.action?.phase !== undefined) {
    const oldPhase = project.phase || 0;
    await db.updateProjectField(suggestion.projectId, {
      phase: suggestion.action.phase,
      ...(suggestion.action.status ? { status: suggestion.action.status } : {}),
    });
    await db.logActivity(suggestion.projectId, {
      type: 'phase_change',
      summary: `Phase updated to ${PHASES[suggestion.action.phase]} — Diane approved Aurora suggestion: "${suggestion.title}"`,
    });
    console.log(`[applySuggestion] Phase: ${PHASES[oldPhase]} → ${PHASES[suggestion.action.phase]}`);

    // If moving to Close-out or Completed, trigger close-out actions
    if (suggestion.action.phase === 4 || suggestion.action.phase === 5 || suggestion.action.status === 'Completed') {
      try { await draftClientFeedback(project); } catch(e) { console.error('[Feedback]', e.message); }
      const yr1 = new Date(); yr1.setFullYear(yr1.getFullYear() + 1);
      const yr2 = new Date(); yr2.setFullYear(yr2.getFullYear() + 2);
      try { await createCalendarReminder(`1-year follow-up: ${project.clientName}`, `Check in with ${project.clientName} — explore new opportunities.`, yr1.toISOString().slice(0,10)); } catch(e) {}
      try { await createCalendarReminder(`2-year follow-up: ${project.clientName}`, `2-year relationship check-in with ${project.clientName}.`, yr2.toISOString().slice(0,10)); } catch(e) {}
      await sendEmail('diane.k@risk2solution.com',
        `[Aurora] Project ${suggestion.action.phase === 5 ? 'completed' : 'in close-out'}: ${project.clientName}`,
        `The ${project.projectName||project.clientName} project has been moved to ${PHASES[suggestion.action.phase]}.\n\nAurora has drafted a client feedback email for your review and set follow-up calendar reminders.\n\n${process.env.FRONTEND_URL||''}\n\nAurora\nR2S Project Management Intelligence`,
        true
      );
    }
  }

  // If suggestion was to draft an escalation
  if (suggestion.action?.draftEscalation) {
    await draftRiskEscalation(project, suggestion.action.riskDescription || 'Risk identified by Aurora');
  }

  // If suggestion was a close-out action (invoice, feedback, report check)
  if (suggestion.action?.draftInvoice) {
    // Internal email to Diane only — just a prompt to check and send invoices
    await sendEmail('diane.k@risk2solution.com',
      `[Aurora] Action required — check invoicing: ${project.clientName}`,
      `Hi Diane,\n\nThe ${project.projectName || project.clientName} project is now in Close-out. Please review the invoicing status in Aurora and send any outstanding invoices to the client.\n\nProject: ${project.projectName || project.clientName}\nClient: ${project.clientName}\nContract value: ${project.value || 'See project record'}\nInvoicing terms: ${project.invoicingNotes || 'See project record'}\n\nPlease log into Aurora to review and update the invoicing status.\n\n${process.env.FRONTEND_URL || ''}\n\nAurora\nR2S Project Management Intelligence`,
      true
    );
    await db.logActivity(project.id, { type: 'manual_note', summary: 'Invoicing check alert sent to Diane' });
  }

  if (suggestion.action?.draftFeedback) {
    await draftClientFeedback(project);
  }

  // If suggestion was a calendar booking confirmation
  if (suggestion.action?.calendarEventId) {
    await sendMeetingInvite(suggestion.action.booking, suggestion.action.calendarEventId);
    // Notify Diane that invites were sent
    await sendEmail('diane.k@risk2solution.com',
      `[Aurora] Meeting invites sent: ${suggestion.action.booking?.title || 'Event'}`,
      `Meeting invites have been sent to all attendees for:

${suggestion.action.booking?.title || 'Event'}
Date: ${suggestion.action.booking?.startDateTime?.slice(0,10) || 'TBC'}

Aurora
R2S Project Management Intelligence`,
      true
    );
  }

  // If suggestion was to draft an email
  if (suggestion.confirmTaskType) {
    const docs    = await db.getDocuments(suggestion.projectId);
    const context = buildContext(project, docs);
    const text    = await aurora(suggestion.confirmTaskType,
      `Draft a ${suggestion.confirmTaskType.replace(/_/g,' ')} for ${project.projectName||project.clientName} at ${project.clientName}.`,
      context
    );
    const draft = {
      id: `d_${Date.now()}_${Math.random().toString(36).slice(2,6)}`,
      projectId: suggestion.projectId, clientName: project.clientName,
      projectName: project.projectName, type: suggestion.confirmTaskType,
      urgency: suggestion.type === SUGGESTION_TYPES.ESCALATE ? 'urgent' : 'routine',
      toName: project.clientContact, toEmail: project.clientEmail,
      subject: `${project.projectName||project.clientName}`,
      body: text, source: 'suggestion',
    };
    await db.saveDraft(draft);
    await saveDraftEmail(draft);
  }
}


// ── Parse risk register from tab-separated AI output ────────────────────────
function parseRiskRegister(tsvText, projectId) {
  const lines = tsvText.split('\n').filter(l => l.trim() && !l.toLowerCase().startsWith('client'));
  return lines.map((line, i) => {
    const cols = line.split('\t').map(c => c.trim().replace(/"/g,''));
    return {
      id: `risk_${projectId}_${i+1}`,
      projectId,
      number: i + 1,
      description: cols[2] || cols[1] || 'Risk ' + (i+1),
      likelihood: cols[3] || 'Medium',
      impact: cols[4] || 'Medium',
      level: cols[5] || 'Medium',
      mitigation: cols[6] || '',
      owner: cols[7] || 'Diane Kruger',
      status: 'Open',
      triggered: false,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };
  }).filter(r => r.description.length > 2);
}

// ── Parse deliverables tracker from AI output ────────────────────────────────
function parseDeliverables(tsvText, projectId) {
  const lines = tsvText.split('\n').filter(l => l.trim() && !l.toLowerCase().startsWith('client'));
  return lines.map((line, i) => {
    const cols = line.split('\t').map(c => c.trim().replace(/"/g,''));
    return {
      id: `del_${projectId}_${i+1}`,
      projectId,
      number: i + 1,
      name: cols[2] || cols[1] || 'Deliverable ' + (i+1),
      phase: cols[3] || '',
      dueDate: cols[4] || '',
      status: cols[5] || 'Outstanding',
      notes: cols[7] || '',
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };
  }).filter(d => d.name.length > 2);
}

// ── Monitor risks daily ───────────────────────────────────────────────────────
async function monitorRisks(projects) {
  for (const p of projects) {
    if (p.type !== 'standard' || isCompleted(p)) continue;
    const risks = await db.getRiskRegister(p.id);
    if (!risks.length) continue;

    const days = p.dueDate ? Math.round((new Date(p.dueDate) - new Date()) / (1000*60*60*24)) : null;
    const phase = p.phase || 0;
    const status = (p.status||'').toLowerCase();

    for (const risk of risks) {
      if (risk.triggered || risk.status === 'Closed') continue;

      let triggered = false;
      let triggerReason = '';

      const desc = (risk.description||'').toLowerCase();

      // Check if risk conditions are met
      if (desc.includes('delay') || desc.includes('overdue') || desc.includes('timeline')) {
        if (days !== null && days < 0) { triggered = true; triggerReason = 'Project is now overdue.'; }
        else if (days !== null && days <= 7) { triggered = true; triggerReason = `Project due in ${days} days.`; }
      }
      if ((desc.includes('hold') || desc.includes('stall') || desc.includes('block')) && status.includes('hold')) {
        triggered = true; triggerReason = 'Project is currently On Hold.';
      }
      if (desc.includes('consultant') || desc.includes('trainer') || desc.includes('resource')) {
        if (!p.consultant) { triggered = true; triggerReason = 'No consultant assigned to this project.'; }
      }
      if (desc.includes('phase') && phase === 0 && days !== null && days <= 14) {
        triggered = true; triggerReason = `Project still in Kick-off with ${days} days remaining.`;
      }

      if (triggered) {
        await db.updateRisk(p.id, risk.id, { triggered: true, triggeredAt: new Date().toISOString(), triggerReason });

        // Email Diane
        await sendEmail('diane.k@risk2solution.com',
          `[Aurora] Risk triggered: ${p.clientName} — ${risk.description.slice(0,60)}`,
          `Hi Diane,

Aurora has detected that a documented risk has been triggered on the ${p.clientName} project.

Project: ${p.projectName || p.clientName}
Current phase: ${PHASES[phase]}

RISK #${risk.number}: ${risk.description}
Likelihood: ${risk.likelihood} | Impact: ${risk.impact} | Level: ${risk.level}

Trigger reason: ${triggerReason}

DOCUMENTED MITIGATION:
${risk.mitigation || 'No mitigation documented — please review.'}

Risk owner: ${risk.owner || 'Diane Kruger'}

Please review the mitigation plan and take appropriate action. You can view and update the full risk register in Aurora under the project record.

${process.env.FRONTEND_URL ? 'Aurora portal: ' + process.env.FRONTEND_URL : ''}

Aurora
R2S Project Management Intelligence`,
          true
        );
        console.log(`[Risks] Risk triggered for ${p.clientName}: ${risk.description.slice(0,50)}`);
      }
    }
  }
}

// ── Update deliverable status from calendar/emails ────────────────────────────
async function updateDeliverableFromCalendar(projectId, deliverableName, calendarEvent) {
  const items = await db.getDeliverables(projectId);
  const match = items.find(d =>
    d.name.toLowerCase().includes(deliverableName.toLowerCase()) ||
    deliverableName.toLowerCase().includes(d.name.toLowerCase().slice(0,15))
  );
  if (match) {
    const now = new Date();
    const eventDate = new Date(calendarEvent.start);
    const status = eventDate < now ? 'Complete' : 'In Progress';
    await db.updateDeliverable(projectId, match.id, { status, calendarEvent: calendarEvent.subject });
    console.log(`[Deliverables] Updated ${match.name} → ${status}`);
  }
}

function isCompleted(p) {
  return ['Completed','Terminated','Closed'].includes(p.status);
}

// ── Weekly executive status report ───────────────────────────────────────────
async function sendWeeklyExecutiveReport(projects) {
  const RECIPIENTS = [
    'dave.c@risk2solution.com',
    'kandia@risk2solution.com',
    'diane.k@risk2solution.com',
  ];

  const standard = projects.filter(p => p.type === 'standard' && (p.phase || 0) < 5);
  if (!standard.length) return;

  const now = new Date();
  const dateStr = now.toLocaleDateString('en-AU', { weekday:'long', day:'numeric', month:'long', year:'numeric' });

  // ── Build project rows ────────────────────────────────────────────────────
  const phaseColors = ['#6aa3ff','#ff608a','#ffd93d','#a8ff78','#00e8bb','#888'];
  const statusBadge = (p) => {
    if (!p.dueDate) return { label:'No date set', color:'#555' };
    const days = Math.round((new Date(p.dueDate) - now) / (1000*60*60*24));
    if (p.status === 'On Hold') return { label:'On Hold', color:'#ffd93d' };
    if (days < 0) return { label:'Overdue', color:'#ff3860' };
    if (days <= 7) return { label:`Due in ${days}d`, color:'#ff608a' };
    if (days <= 14) return { label:'At Risk', color:'#ffd93d' };
    return { label:'On Track', color:'#00e8bb' };
  };

  // Parse contract value to number
  const parseVal = (v) => {
    if (!v) return 0;
    const m = (v+'').replace(/[$,AUD]/gi,'').match(/[\d.]+/);
    return m ? parseFloat(m[0]) : 0;
  };

  // ── Invoice summary — replaces broken SVG chart ───────────────────────────
  const invoiceProjects = standard.filter(p => parseVal(p.value) > 0);
  const totalPortfolio  = invoiceProjects.reduce((s,p) => s + parseVal(p.value), 0);

  // Load saved invoices for each project to get actual invoiced amounts
  const invoiceData = await Promise.all(invoiceProjects.map(async p => {
    let recorded = [];
    try { recorded = await readInvoices(p.id); } catch {}
    const invoiced = recorded.reduce((s,i) => s + parseFloat(i.amount||0), 0);
    const paid     = recorded.filter(i=>i.paid).reduce((s,i) => s + parseFloat(i.amount||0), 0);
    const val      = parseVal(p.value);
    // Estimate outstanding invoicing by month based on due date
    const dueDate  = p.dueDate ? new Date(p.dueDate) : null;
    const dueMonth = dueDate ? dueDate.toLocaleDateString('en-AU',{month:'short',year:'numeric'}) : 'TBC';
    return { p, val, invoiced, paid, outstanding: Math.max(0, val - invoiced), dueMonth, recorded };
  }));

  const totalInvoiced    = invoiceData.reduce((s,d) => s + d.invoiced, 0);
  const totalPaid        = invoiceData.reduce((s,d) => s + d.paid, 0);
  const totalOutstanding = invoiceData.reduce((s,d) => s + d.outstanding, 0);

  // Group outstanding by due month
  const byMonth = {};
  invoiceData.forEach(d => {
    if (d.outstanding > 0) {
      if (!byMonth[d.dueMonth]) byMonth[d.dueMonth] = 0;
      byMonth[d.dueMonth] += d.outstanding;
    }
  });

  // Build project rows for invoice table
  const invoiceRows = invoiceData.map((d,i) => {
    const bg = i % 2 === 0 ? '#1a1a2e' : '#141428';
    const pct = d.val > 0 ? Math.round((d.invoiced/d.val)*100) : 0;
    const barW = Math.min(100, pct);
    return `<tr style="background:${bg}">
      <td style="padding:8px 10px;border-bottom:1px solid #2a2a4a;color:#fff;font-size:12px;font-weight:500">${d.p.clientName}</td>
      <td style="padding:8px 10px;border-bottom:1px solid #2a2a4a;color:#aaa;font-size:11px">$${Math.round(d.val).toLocaleString('en-AU')}</td>
      <td style="padding:8px 10px;border-bottom:1px solid #2a2a4a;font-size:11px">
        <div style="">
          <div style="flex:1;background:#2a2a4a;border-radius:3px;height:8px;min-width:60px">
            <div style="width:${barW}%;background:#00e8bb;height:8px;border-radius:3px"></div>
          </div>
          <span style="color:${d.invoiced>0?'#00e8bb':'#555'};min-width:60px;text-align:right">$${Math.round(d.invoiced).toLocaleString('en-AU')}</span>
        </div>
      </td>
      <td style="padding:8px 10px;border-bottom:1px solid #2a2a4a;color:${d.outstanding>0?'#ffd93d':'#00e8bb'};font-size:11px;font-weight:${d.outstanding>0?'600':'400'}">$${Math.round(d.outstanding).toLocaleString('en-AU')}</td>
      <td style="padding:8px 10px;border-bottom:1px solid #2a2a4a;color:#aaa;font-size:11px">${d.dueMonth}</td>
      <td style="padding:8px 10px;border-bottom:1px solid #2a2a4a;color:${d.paid>0?'#00e8bb':'#555'};font-size:11px">$${Math.round(d.paid).toLocaleString('en-AU')}</td>
    </tr>`;
  }).join('');

  // Monthly outstanding breakdown
  const monthlyRows = Object.entries(byMonth)
    .sort(([a],[b]) => new Date('1 '+a) - new Date('1 '+b))
    .map(([month, amt]) => `
      <div style="display:flex;justify-content:space-between;align-items:center;padding:6px 0;border-bottom:1px solid #2a2a4a22">
        <span style="color:#ccc;font-size:12px">${month}</span>
        <span style="color:#ffd93d;font-size:13px;font-weight:600">$${Math.round(amt).toLocaleString('en-AU')}</span>
      </div>`).join('');

  let invoicingSectionHtml = `
  <!-- Invoicing Summary -->
  <div style="background:#1a1a2e;border:1px solid #2a2a4a;border-radius:8px;padding:0;margin-bottom:20px;overflow:hidden">
    <div style="padding:14px 16px;border-bottom:1px solid #2a2a4a;display:flex;gap:24px;align-items:center">
      <div style="font-size:10px;color:#6aa3ff;text-transform:uppercase;letter-spacing:1px;flex:1">Invoicing Overview</div>
      <div style="text-align:center"><div style="font-size:16px;font-weight:700;color:#00e8bb">$${Math.round(totalInvoiced).toLocaleString('en-AU')}</div><div style="font-size:9px;color:#888;text-transform:uppercase">Invoiced</div></div>
      <div style="text-align:center"><div style="font-size:16px;font-weight:700;color:#6aa3ff">$${Math.round(totalPaid).toLocaleString('en-AU')}</div><div style="font-size:9px;color:#888;text-transform:uppercase">Paid</div></div>
      <div style="text-align:center"><div style="font-size:16px;font-weight:700;color:#ffd93d">$${Math.round(totalOutstanding).toLocaleString('en-AU')}</div><div style="font-size:9px;color:#888;text-transform:uppercase">Outstanding</div></div>
      <div style="text-align:center"><div style="font-size:16px;font-weight:700;color:#fff">$${Math.round(totalPortfolio).toLocaleString('en-AU')}</div><div style="font-size:9px;color:#888;text-transform:uppercase">Portfolio</div></div>
    </div>
    <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse">
      <thead><tr style="background:#141428">
        <th style="padding:8px 10px;text-align:left;font-size:10px;color:#666;font-weight:500;text-transform:uppercase;border-bottom:1px solid #2a2a4a">Client</th>
        <th style="padding:8px 10px;text-align:left;font-size:10px;color:#666;font-weight:500;text-transform:uppercase;border-bottom:1px solid #2a2a4a">Contract value</th>
        <th style="padding:8px 10px;text-align:left;font-size:10px;color:#666;font-weight:500;text-transform:uppercase;border-bottom:1px solid #2a2a4a">Invoiced to date</th>
        <th style="padding:8px 10px;text-align:left;font-size:10px;color:#666;font-weight:500;text-transform:uppercase;border-bottom:1px solid #2a2a4a">Outstanding</th>
        <th style="padding:8px 10px;text-align:left;font-size:10px;color:#666;font-weight:500;text-transform:uppercase;border-bottom:1px solid #2a2a4a">Due month</th>
        <th style="padding:8px 10px;text-align:left;font-size:10px;color:#666;font-weight:500;text-transform:uppercase;border-bottom:1px solid #2a2a4a">Paid</th>
      </tr></thead>
      <tbody>${invoiceRows}</tbody>
    </table>
    ${Object.keys(byMonth).length > 0 ? `
    <div style="padding:14px 16px;border-top:1px solid #2a2a4a">
      <div style="font-size:10px;color:#6aa3ff;text-transform:uppercase;letter-spacing:1px;margin-bottom:10px">Outstanding by month</div>
      ${monthlyRows}
    </div>` : ''}
  </div>`;


  // ── Build PM checklist summary for report ─────────────────────────────────
  const checklistRows = await Promise.all(standard.map(async p => {
    try {
      const state = await readPMChecklist(p.id);
      const weekKey = getWeekKey();
      const weekData = state.weeks?.[weekKey] || {};
      const phase = p.phase || 0;
      const relevantQs = PM_CHECKLIST_QUESTIONS; // All questions for every project
      const checkedCount = relevantQs.filter(q => weekData[q.id]?.checked).length;
      const total = relevantQs.length;
      const pct = total > 0 ? Math.round((checkedCount / total) * 100) : 0;
      return { p, checkedCount, total, pct };
    } catch { return { p, checkedCount: 0, total: 11, pct: 0 }; }
  }));

  const checklistHtml = `
  <div style="background:#1a1a2e;border:1px solid #2a2a4a;border-radius:8px;overflow:hidden;margin-bottom:20px">
    <div style="padding:12px 16px;border-bottom:1px solid #2a2a4a;display:flex;justify-content:space-between;align-items:center">
      <div style="font-size:10px;color:#6aa3ff;text-transform:uppercase;letter-spacing:1px">Weekly PM Checklist — Diane's completion this week</div>
      <div style="font-size:10px;color:#888">${getWeekKey()}</div>
    </div>
    <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse">
      <thead><tr style="background:#141428">
        <th style="padding:8px 10px;text-align:left;font-size:10px;color:#666;font-weight:500;text-transform:uppercase;border-bottom:1px solid #2a2a4a">Project</th>
        <th style="padding:8px 10px;text-align:left;font-size:10px;color:#666;font-weight:500;text-transform:uppercase;border-bottom:1px solid #2a2a4a">This week</th>
        <th style="padding:8px 10px;text-align:left;font-size:10px;color:#666;font-weight:500;text-transform:uppercase;border-bottom:1px solid #2a2a4a">Status</th>
      </tr></thead>
      <tbody>
        ${checklistRows.map((r, i) => {
          const bg = i % 2 === 0 ? '#1a1a2e' : '#141428';
          const statusColor = r.pct === 100 ? '#00e8bb' : r.pct >= 50 ? '#ffd93d' : '#ff608a';
          const statusText = r.pct === 100 ? '✓ Complete' : r.checkedCount === 0 ? 'Not started' : `${r.checkedCount}/${r.total} done`;
          return `<tr style="background:${bg}">
            <td style="padding:8px 10px;border-bottom:1px solid #2a2a4a;color:#fff;font-size:12px">${r.p.clientName}<br><span style="font-size:10px;color:#888">${r.p.projectName||''}</span></td>
            <td style="padding:8px 10px;border-bottom:1px solid #2a2a4a;font-size:11px">
              <div style="background:#2a2a4a;border-radius:3px;height:6px;width:100px">
                <div style="width:${r.pct}%;background:${statusColor};height:6px;border-radius:3px"></div>
              </div>
            </td>
            <td style="padding:8px 10px;border-bottom:1px solid #2a2a4a">
              <span style="font-size:11px;color:${statusColor}">${statusText}</span>
            </td>
          </tr>`;
        }).join('')}
      </tbody>
    </table>
  </div>`;
  const projectRows = standard.map((p, i) => {
    const phase = PHASES[p.phase || 0];
    const sb = statusBadge(p);
    const val = parseVal(p.value);
    const valStr = val ? '$' + val.toLocaleString('en-AU') : 'TBC';
    const briefed = p.consultant ? 'Yes' : 'No';
    const welcomed = p.clientEmail ? 'Pending review' : 'TBC';
    const dueStr = p.dueDate || 'TBC';
    const bgColor = i % 2 === 0 ? '#1a1a2e' : '#141428';

    return `<tr style="background:${bgColor}">
      <td style="padding:8px 10px;border-bottom:1px solid #2a2a4a;color:#e0e0e0;font-size:12px;vertical-align:top">
        <div style="font-weight:600;color:#fff">${p.clientName}</div>
        <div style="color:#aaa;font-size:11px">${p.projectName || ''}</div>
      </td>
      <td style="padding:8px 10px;border-bottom:1px solid #2a2a4a;vertical-align:top">
        <span style="background:#2a2a4a;color:#6aa3ff;padding:2px 8px;border-radius:10px;font-size:11px;white-space:nowrap">${phase}</span>
      </td>
      <td style="padding:8px 10px;border-bottom:1px solid #2a2a4a;vertical-align:top">
        <span style="background:${sb.color}22;color:${sb.color};padding:2px 8px;border-radius:10px;font-size:11px;border:1px solid ${sb.color}44;white-space:nowrap">${sb.label}</span>
      </td>
      <td style="padding:8px 10px;border-bottom:1px solid #2a2a4a;color:#aaa;font-size:11px;vertical-align:top">${valStr}</td>
      <td style="padding:8px 10px;border-bottom:1px solid #2a2a4a;color:#aaa;font-size:11px;vertical-align:top">${dueStr}</td>
      <td style="padding:8px 10px;border-bottom:1px solid #2a2a4a;vertical-align:top">
        <div style="color:#ccc;font-size:11px">${p.consultant || '<span style="color:#ff608a">Not assigned</span>'}</div>
        <div style="font-size:10px;color:${briefed==='Yes'?'#00e8bb':'#ff608a'}">${briefed === 'Yes' ? '✓ Briefed' : '✗ Not briefed'}</div>
      </td>
      <td style="padding:8px 10px;border-bottom:1px solid #2a2a4a;color:#aaa;font-size:11px;vertical-align:top">
        ${p.clientContact || 'TBC'}${p.clientEmail ? '<br><span style="font-size:10px;color:#6aa3ff">' + p.clientEmail + '</span>' : ''}
      </td>
    </tr>`;
  }).join('');

  // ── Summary stats ─────────────────────────────────────────────────────────
  const totalValue   = standard.reduce((s, p) => s + parseVal(p.value), 0);
  const atRisk       = standard.filter(p => { const d = p.dueDate ? Math.round((new Date(p.dueDate)-now)/(1000*60*60*24)) : 99; return d <= 14 && d >= 0; }).length;
  const noConsultant = standard.filter(p => !p.consultant).length;
  const onHold       = standard.filter(p => p.status === 'On Hold').length;

  // ── Generate AI narrative for the week ───────────────────────────────────
  let narrative = '';
  try {
    const contextSummary = standard.map(p => {
      const days = p.dueDate ? Math.round((new Date(p.dueDate)-now)/(1000*60*60*24)) : null;
      return `${p.clientName} (${PHASES[p.phase||0]}, ${days !== null ? days + ' days until due' : 'no due date'}, consultant: ${p.consultant || 'unassigned'})`;
    }).join('; ');

    narrative = await aurora('status_report',
      `Write a concise executive summary paragraph (4-6 sentences) for the R2S weekly project status report dated ${dateStr}.
      
Active projects: ${standard.length}
Total portfolio value: $${Math.round(totalValue).toLocaleString('en-AU')}
Projects at risk or due soon: ${atRisk}
Projects without consultant: ${noConsultant}
On hold: ${onHold}

Project summaries: ${contextSummary}

Write as if briefing the CEO, COO and PM. Highlight what needs attention this week. Plain text, no asterisks, no markdown, no bullet points in this paragraph — just clear professional prose.`,
      null
    );
  } catch(e) { narrative = `Weekly project status report for the week of ${dateStr}. ${standard.length} active projects with a total portfolio value of $${Math.round(totalValue).toLocaleString('en-AU')}.`; }

  // ── Build full HTML email ─────────────────────────────────────────────────
  const html = `<!DOCTYPE html>
<html xmlns:o="urn:schemas-microsoft-com:office:office">
<head><meta charset="utf-8"/><!--[if mso]><xml><o:OfficeDocumentSettings><o:AllowPNG/></o:OfficeDocumentSettings></xml><![endif]--></head>
<body style="margin:0;padding:0;background:#0f0f1a;font-family:Arial,sans-serif">
<div style="max-width:700px;margin:0 auto;background:#0f0f1a;padding:24px">

  <!-- Header -->
  <div style="background:#1a1a3e;border-radius:12px;padding:24px;margin-bottom:20px;border:1px solid #2a2a5a">
    <div style="font-size:11px;color:#6aa3ff;text-transform:uppercase;letter-spacing:2px;margin-bottom:6px">Risk 2 Solution Group</div>
    <div style="font-size:22px;font-weight:700;color:#fff;margin-bottom:4px">Weekly Project Status Report</div>
    <div style="font-size:13px;color:#aaa">${dateStr} · Generated by Aurora</div>
  </div>

  <!-- Stats row -->
  <div style="margin-bottom:20px">
    ${[
      { label:'Active projects', value: standard.length, color:'#6aa3ff' },
      { label:'Portfolio value', value: '$'+Math.round(totalValue/1000)+'k', color:'#00e8bb' },
      { label:'At risk / due soon', value: atRisk, color: atRisk > 0 ? '#ffd93d' : '#00e8bb' },
      { label:'Needs attention', value: noConsultant + onHold, color: (noConsultant+onHold) > 0 ? '#ff608a' : '#00e8bb' },
    ].map(s => `<div style="background:#1a1a2e;border:1px solid #2a2a4a;border-radius:8px;padding:12px;text-align:center">
      <div style="font-size:22px;font-weight:700;color:${s.color};font-family:Georgia,serif">${s.value}</div>
      <div style="font-size:10px;color:#888;margin-top:3px;text-transform:uppercase;letter-spacing:.5px">${s.label}</div>
    </div>`).join('')}
  </div>

  <!-- Narrative -->
  <div style="background:#1a1a2e;border:1px solid #2a2a4a;border-radius:8px;padding:16px;margin-bottom:20px;color:#ccc;font-size:13px;line-height:1.7">
    <div style="font-size:10px;color:#6aa3ff;text-transform:uppercase;letter-spacing:1px;margin-bottom:8px">Executive Summary</div>
    ${narrative}
  </div>

  ${invoicingSectionHtml}

  ${checklistHtml}

  <!-- Project table -->
  <div style="background:#1a1a2e;border:1px solid #2a2a4a;border-radius:8px;overflow:hidden;margin-bottom:20px">
    <div style="padding:12px 16px;border-bottom:1px solid #2a2a4a">
      <div style="font-size:10px;color:#6aa3ff;text-transform:uppercase;letter-spacing:1px">Active Projects</div>
    </div>
    <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse">
      <thead>
        <tr style="background:#141428">
          <th style="padding:8px 10px;text-align:left;font-size:10px;color:#666;font-weight:500;text-transform:uppercase;letter-spacing:.5px;border-bottom:1px solid #2a2a4a">Client / Project</th>
          <th style="padding:8px 10px;text-align:left;font-size:10px;color:#666;font-weight:500;text-transform:uppercase;letter-spacing:.5px;border-bottom:1px solid #2a2a4a">Phase</th>
          <th style="padding:8px 10px;text-align:left;font-size:10px;color:#666;font-weight:500;text-transform:uppercase;letter-spacing:.5px;border-bottom:1px solid #2a2a4a">Status</th>
          <th style="padding:8px 10px;text-align:left;font-size:10px;color:#666;font-weight:500;text-transform:uppercase;letter-spacing:.5px;border-bottom:1px solid #2a2a4a">Value</th>
          <th style="padding:8px 10px;text-align:left;font-size:10px;color:#666;font-weight:500;text-transform:uppercase;letter-spacing:.5px;border-bottom:1px solid #2a2a4a">Due date</th>
          <th style="padding:8px 10px;text-align:left;font-size:10px;color:#666;font-weight:500;text-transform:uppercase;letter-spacing:.5px;border-bottom:1px solid #2a2a4a">Consultant</th>
          <th style="padding:8px 10px;text-align:left;font-size:10px;color:#666;font-weight:500;text-transform:uppercase;letter-spacing:.5px;border-bottom:1px solid #2a2a4a">Client contact</th>
        </tr>
      </thead>
      <tbody>${projectRows}</tbody>
    </table>
  </div>

  <!-- This week section -->
  <div style="background:#1a1a2e;border:1px solid #2a2a4a;border-radius:8px;padding:16px;margin-bottom:20px">
    <div style="font-size:10px;color:#6aa3ff;text-transform:uppercase;letter-spacing:1px;margin-bottom:10px">Actions Required This Week</div>
    ${standard.filter(p => {
      const days = p.dueDate ? Math.round((new Date(p.dueDate)-now)/(1000*60*60*24)) : 99;
      return !p.consultant || days <= 14 || p.status === 'On Hold';
    }).map(p => {
      const items = [];
      if (!p.consultant) items.push(`Assign consultant/trainer to <strong>${p.clientName}</strong>`);
      const days = p.dueDate ? Math.round((new Date(p.dueDate)-now)/(1000*60*60*24)) : 99;
      if (days <= 7 && days >= 0) items.push(`<strong>${p.clientName}</strong> due in ${days} day${days!==1?'s':''} — review deliverables`);
      else if (days <= 14 && days >= 0) items.push(`<strong>${p.clientName}</strong> due in ${days} days — prepare for close-out`);
      if (p.status === 'On Hold') items.push(`<strong>${p.clientName}</strong> is On Hold — review with team`);
      return items;
    }).flat().map(item => `<div style="padding:5px 0;color:#ccc;font-size:12px;border-bottom:1px solid #2a2a4a22">→ ${item}</div>`).join('') || '<div style="color:#00e8bb;font-size:12px">No urgent actions — all projects on track</div>'}
  </div>

  <!-- Footer -->
  <div style="text-align:center;color:#444;font-size:11px;padding-top:10px">
    Aurora · R2S Project Management Intelligence · Confidential — internal use only<br>
    ${process.env.FRONTEND_URL ? '<a href="' + process.env.FRONTEND_URL + '" style="color:#6aa3ff">Open Aurora portal</a>' : ''}
  </div>

</div>
</body>
</html>`;

  // Send to all recipients
  for (const recipient of RECIPIENTS) {
    try {
      await sendEmail(recipient, `R2S Weekly Project Status — ${now.toLocaleDateString('en-AU', { day:'numeric', month:'short' })}`, html, false, [], true);
      console.log(`[WeeklyReport] Sent to ${recipient}`);
    } catch(e) {
      console.error(`[WeeklyReport] Failed to send to ${recipient}:`, e.message);
    }
  }
  console.log('[WeeklyReport] Weekly executive report complete');
}

// ── Daily batch (6am AEST = 8pm UTC) ─────────────────────────────────────────
async function runBatch() {
  const now = new Date();
  // Use AEST (UTC+10) for day-of-week checks — cron fires at 8pm UTC = 6am AEST next day
  const aestNow = new Date(now.getTime() + 10 * 60 * 60 * 1000);
  const dayOfWeek = aestNow.getDay(); // 0=Sun, 1=Mon ... 5=Fri, 6=Sat (in AEST)
  const isWeekday = dayOfWeek >= 1 && dayOfWeek <= 5;
  const isMonday  = dayOfWeek === 1;
  console.log(`[Batch] AEST day: ${['Sun','Mon','Tue','Wed','Thu','Fri','Sat'][dayOfWeek]} isMonday:${isMonday}`);

  console.log('\n[Batch] ═══ Aurora daily batch starting ═══');
  const projects = await db.getProjects();
  const standard = projects.filter(p => p.type === 'standard');
  const ongoing  = projects.filter(p => p.type === 'ongoing');
  console.log(`[Batch] ${standard.length} standard | ${ongoing.length} ongoing (skipped)`);

  // ── 0. Generate Aurora suggestions ──────────────────────────────────────────
  try {
    const newSuggestions = await generateSuggestions();
    const pending = newSuggestions.filter(s => s.status === 'pending');
    if (pending.length > 0) {
      console.log(`[Suggestions] ${pending.length} suggestion(s) generated`);
    }
  } catch (err) { console.error('[Suggestions] Error:', err.message); }

  // ── 1. Daily portal prompt to Diane (weekdays only) ───────────────────────
  if (isWeekday) {
    const activeCount = standard.filter(p => !['Completed','Terminated'].includes(p.status)).length;
    const draftCount  = (await db.getDrafts()).length;
    await sendEmail('diane.k@risk2solution.com',
      '[Aurora] Good morning — daily project check',
      `Good morning Diane,

Aurora here with your daily project summary.

Active projects: ${activeCount}
Drafts awaiting your review: ${draftCount}

Please log into Aurora to review any pending drafts and check project status.

${process.env.FRONTEND_URL ? `Aurora portal: ${process.env.FRONTEND_URL}` : ''}

Aurora
R2S Project Management Intelligence`,
      true
    );
  }

  // ── 2. Due date reminders ─────────────────────────────────────────────────
  await checkDueDateReminders();

  // Email polling runs on its own hourly schedule — not in the daily batch

  // ── 3. Check for stuck phases + monitor risks ───────────────────────────────
  try { await monitorRisks(standard); } catch(e) { console.error('[Risks]', e.message); }
  // ── 3. Check for stuck phases ─────────────────────────────────────────────
  await checkStuckPhases(standard);

  // ── 4. SOP-TRN-001: Materials submission & session report reminders ────────
  await checkMaterialsSubmissionReminders(standard);

  // ── 5. Weekly actions (Mondays only) ─────────────────────────────────────
  if (isMonday) {
    // 5b. Weekly client status email drafts
    for (const p of standard) {
      if (['Completed','Terminated','On Hold'].includes(p.status)) continue;
      try {
        const context = buildContext(p);
        const text = await aurora('status_email',
          `Draft a short weekly status update email to ${p.clientContact || 'the client'} at ${p.clientName} for the ${p.projectName || p.clientName} project. Current phase: ${PHASES[p.phase||0]}. 3-4 sentences: what happened this week, what is next, anything needed from the client. Professional and concise.`,
          context
        );
        const draft = {
          id: `d_${Date.now()}_${Math.random().toString(36).slice(2,6)}`,
          projectId: p.id, clientName: p.clientName, projectName: p.projectName,
          type: 'status_email', urgency: 'routine',
          toName: p.clientContact, toEmail: p.clientEmail,
          subject: `${p.projectName || p.clientName} — Weekly update`,
          body: text, source: 'batch',
        };
        await db.saveDraft(draft);
        await saveDraftEmail(draft);
      } catch (err) {
        if (err.message === 'MONTHLY_CAP_REACHED') break;
        console.error(`[Batch] Status email error on ${p.clientName}:`, err.message);
      }
    }

    // 5b. Weekly consultant check-in drafts
    const checkinCount = await generateWeeklyConsultantCheckins(standard);

    // 5c. Monday reminder to Diane to review and send check-in drafts
    await sendEmail('diane.k@risk2solution.com',
      '[Aurora] Weekly action — please review and send check-in emails',
      `Good morning Diane,

Aurora has prepared your weekly emails for review. Please log into Aurora to review and send:

• ${standard.filter(p => !['Completed','Terminated'].includes(p.status)).length} client status update emails
• ${checkinCount} consultant/trainer check-in emails

All drafts are in the info@risk2solution.com Outlook shared mailbox Drafts folder and visible in the Aurora portal under Comms Drafts.

Please review each one for accuracy before sending.

${process.env.FRONTEND_URL ? `Aurora portal: ${process.env.FRONTEND_URL}` : ''}

Aurora
R2S Project Management Intelligence`,
      true
    );

    // 5d. Read and analyse consultant reply emails
    await readConsultantReplies();
  }

  // ── 6. Weekly executive report (Mondays only at 6am AEST) ────────────────
  if (isMonday) {
    try {
      console.log('[Batch] Sending weekly executive report (Monday)...');
      await sendWeeklyExecutiveReport(standard);
    } catch(e) { console.error('[WeeklyReport] Error:', e.message); }
  }

  // ── 7. Internal programme reminders (daily) ───────────────────────────────
  try { await checkInternalProjectReminders(); } catch(e) { console.error('[Internal]', e.message); }

  // ── 8. Internal programme weekly report (Mondays) ────────────────────────
  if (isMonday) {
    try { await sendInternalProjectStatusReport(); } catch(e) { console.error('[InternalReport]', e.message); }
    // Operational weekly update to delivery team (Cherry, Janita, Diane, CC Dave)
    try { await sendInternalWeeklyOpsUpdate(); } catch(e) { console.error('[InternalOps]', e.message); }
  }

  // ── 9. Wednesday PM checklist (Wednesday = dayOfWeek 3 in AEST) ───────────
  const isWednesday = dayOfWeek === 3;
  if (isWednesday) {
    try { await sendWeeklyPMChecklist(); } catch(e) { console.error('[PMChecklist]', e.message); }
  }

  // ── 10. At-risk project summary ───────────────────────────────────────────
  const atRisk = standard.filter(p => {
    if (!p.dueDate) return false;
    const days = Math.round((new Date(p.dueDate) - new Date()) / (1000 * 60 * 60 * 24));
    return days <= 14 && days >= 0;
  });

  if (atRisk.length > 0) {
    const summary = atRisk.map(p => {
      const days = Math.round((new Date(p.dueDate) - new Date()) / (1000 * 60 * 60 * 24));
      return `• ${p.clientName} (${p.projectName || 'project'}) — due in ${days} days — Phase: ${PHASES[p.phase||0]}`;
    }).join('\n');
    await sendInternalEmail(
      '[Aurora] Projects due within 14 days',
      `Good morning,\n\nAurora has identified ${atRisk.length} project(s) due within the next 14 days:\n\n${summary}\n\nPlease log into Aurora to review and action.\n\nAurora\nR2S Project Management Intelligence`
    );
  }

  console.log('[Batch] ═══ Complete ═══\n');
}

// ── Read consultant calendar availability ────────────────────────────────────
async function getConsultantAvailability() {
  const token = await getOutlookToken();
  if (!token) return null;
  const calendarName = 'R2S Training & Education';
  const mailbox = process.env.OUTLOOK_SHARED_MAILBOX || 'info@risk2solution.com';

  try {
    // Get events in the next 30 days from the training calendar
    const start = new Date();
    const end   = new Date(Date.now() + 30 * 24 * 60 * 60 * 1000);

    // First find the calendar ID
    const calsRes = await axios.get(
      `https://graph.microsoft.com/v1.0/users/${mailbox}/calendars`,
      { headers: { Authorization: `Bearer ${token}` }, timeout: 10000 }
    );
    const cal = calsRes.data?.value?.find(c =>
      c.name?.toLowerCase().includes('training') ||
      c.name?.toLowerCase().includes('education') ||
      c.name === calendarName
    );

    if (!cal) return null;

    // Get events from that calendar
    const eventsRes = await axios.get(
      `https://graph.microsoft.com/v1.0/users/${mailbox}/calendars/${cal.id}/events?$filter=start/dateTime ge '${start.toISOString()}' and start/dateTime le '${end.toISOString()}'&$select=subject,start,end&$top=50`,
      { headers: { Authorization: `Bearer ${token}` }, timeout: 10000 }
    );

    const busySlots = (eventsRes.data?.value || []).map(e => ({
      subject: e.subject,
      start: new Date(e.start.dateTime),
      end: new Date(e.end.dateTime),
    }));

    // Find 3 available weekday morning slots in next 30 days
    const available = [];
    const checkDate = new Date();
    checkDate.setDate(checkDate.getDate() + 3); // start 3 days from now

    while (available.length < 3 && checkDate < end) {
      const dow = checkDate.getDay();
      if (dow >= 1 && dow <= 5) { // weekdays only
        // Check 9am and 10am slots
        for (const hour of [9, 10, 14]) {
          const slotStart = new Date(checkDate);
          slotStart.setHours(hour, 0, 0, 0);
          const slotEnd = new Date(slotStart.getTime() + 60 * 60 * 1000);

          const isBusy = busySlots.some(b =>
            slotStart < b.end && slotEnd > b.start
          );

          if (!isBusy && available.length < 3) {
            available.push(slotStart);
          }
        }
      }
      checkDate.setDate(checkDate.getDate() + 1);
    }

    return available;
  } catch (err) {
    console.error('[Calendar] Availability check failed:', err.message);
    return null;
  }
}

// ── Kick-off meeting agenda generation ────────────────────────────────────────
async function generateKickoffAgenda(project) {
  const context = buildContext(project);
  const agenda = await aurora('consultant_briefing',
    `Write a professional kick-off meeting agenda for the ${project.projectName || project.clientName} project.

FORMAT RULES:
- Plain text only. No asterisks, no bold markdown, no long dashes, no lines.
- Use numbered sections and bullet points (use the bullet character •).
- R2S branding: professional, clear, human-centred tone.
- This is a formal agenda document.

Write the agenda with these sections:

RISK 2 SOLUTION GROUP
Kick-off Meeting Agenda
${project.clientName} — ${project.projectName || 'Project Engagement'}
Date: [To be confirmed]
Location: [To be confirmed — virtual or on-site]
Attendees: ${project.clientContact || '[Client contact]'} (${project.clientName}), ${project.consultant || '[R2S Consultant]'} (R2S), Diane Kruger (R2S)
Duration: 60-90 minutes

1. Welcome and Introductions (10 minutes)
• Purpose and format of meeting
• Attendee introductions

2. Project Overview (15 minutes)
• Scope of engagement as per proposal
• Objectives and expected outcomes
• What success looks like for ${project.clientName}

3. Deliverables and Timeline (20 minutes)
• Confirmed deliverables (list each one from the project scope)
• Proposed timeline and key milestones
• Phasing and scheduling

4. Roles and Responsibilities (10 minutes)
• R2S team responsibilities
• ${project.clientName} team responsibilities
• Key contacts and escalation points

5. Communication and Reporting (10 minutes)
• Reporting frequency and format
• Primary communication channels
• How updates will be shared

6. Risk and Issue Management (10 minutes)
• How risks will be identified and managed
• Escalation process if issues arise
• Change request process

7. Next Steps and Close (10 minutes)
• Confirm immediate next steps
• Confirm dates for next check-in
• Any questions

Prepared by Aurora, R2S Project Management Intelligence
For review by Diane Kruger before distribution`,
    context
  );
  return agenda;
}

// ── Full kick-off system ───────────────────────────────────────────────────────
async function sendKickoffPrompt(project) {
  const DIANE = 'diane.k@risk2solution.com';

  // 1. Get consultant availability from training calendar
  let availabilityText = 'Aurora was unable to read the training calendar. Please check availability manually.';
  let availableDates = null;
  try {
    availableDates = await getConsultantAvailability();
    if (availableDates && availableDates.length > 0) {
      const opts = availableDates.map((d, i) => {
        const dateStr = d.toLocaleDateString('en-AU', { weekday: 'long', day: 'numeric', month: 'long' });
        const timeStr = d.toLocaleTimeString('en-AU', { hour: 'numeric', minute: '2-digit', hour12: true, timeZone: 'Australia/Brisbane' });
        return `Option ${i+1}: ${dateStr} at ${timeStr} AEST`;
      }).join('\n');
      availabilityText = 'Based on the R2S Training and Education calendar, the following slots appear available:\n\n' + opts + '\n\nPlease confirm with ' + (project.consultant || 'the consultant') + ' that these work before sending to the client.';
    }
  } catch (err) {
    console.error('[Kickoff] Calendar check failed:', err.message);
  }

  // 2. Generate kick-off agenda
  let agendaText = '';
  try {
    agendaText = await generateKickoffAgenda(project);
  } catch (err) {
    console.error('[Kickoff] Agenda generation failed:', err.message);
  }

  // 3. Save agenda as Outlook draft for Diane to review
  if (agendaText) {
    const agendaDraft = {
      id: `d_${Date.now()}_agenda`,
      projectId: project.id, clientName: project.clientName,
      projectName: project.projectName, type: 'kickoff_agenda',
      urgency: 'routine', toName: project.clientContact,
      toEmail: project.clientEmail,
      ccEmail: 'diane.k@risk2solution.com',
      subject: `Kick-off Meeting Agenda — ${project.clientName} — ${project.projectName || 'Project'}`,
      body: agendaText, source: 'auto',
    };
    await db.saveDraft(agendaDraft);
    await saveDraftEmail(agendaDraft);
  }

  // 4. Draft client onboarding email
  await draftClientOnboarding(project);

  // 5. Email Diane with availability options and next steps
  await sendEmail(DIANE,
    `[Aurora] Schedule kick-off meeting: ${project.clientName}`,
    `Hi Diane,

The consultant briefing for ${project.projectName || project.clientName} has been sent to ${project.consultant || 'the consultant'}.

The next step is to schedule the kick-off meeting with the client and consultant.

${availabilityText}

To schedule:
• Confirm the date with ${project.consultant || 'the consultant'}
• Send the meeting invite to ${project.clientContact || 'the client'} (${project.clientEmail || ''})
• The kick-off agenda has been saved to the info@risk2solution.com Outlook drafts for your review

Client contact: ${project.clientContact || 'See project record'}
Client email: ${project.clientEmail || 'See project record'}
Consultant: ${project.consultant || 'See project record'}

Once the kick-off meeting is confirmed, update the project phase to Deployment in Aurora.

${process.env.FRONTEND_URL ? 'Aurora portal: ' + process.env.FRONTEND_URL : ''}

Aurora
R2S Project Management Intelligence`,
    true
  );

  console.log('[Kickoff] Prompt sent to Diane with availability and agenda');
}

// ── Client onboarding email ───────────────────────────────────────────────────
async function draftClientOnboarding(project) {
  try {
    const context = buildContext(project);
    const onboardingBody = await aurora('consultant_briefing',
      `Write a professional client welcome and onboarding email from R2S to ${project.clientContact || 'the client'} at ${project.clientName}.

FORMAT RULES:
- Plain text only. No asterisks, no bold markdown, no long dashes.
- Professional, warm, and human-centred tone.
- Concise — this is a welcome email, not a report.

Write the email starting with "Hi ${(project.clientContact || 'there').split(' ')[0]}," and covering:

1. A warm welcome to the R2S engagement (1-2 sentences)

2. Brief confirmation of what R2S will be delivering (1-2 sentences referencing the project scope)

3. Your key R2S contacts:
   Project managed by: Diane Kruger, Corporate Operations Lead
   Email: diane.k@risk2solution.com | Phone: 1300 459 970
   ${project.consultant ? 'Assigned consultant/trainer: ' + project.consultant : ''}

4. What happens next (3 bullet points covering: kick-off meeting to be scheduled, scope and timeline to be confirmed, regular updates throughout)

5. A note that Diane is available for any questions

6. Then add this exact section as the final part of the email before the sign-off:

To help us set up your project correctly, please complete and return the following onboarding form:

ONBOARDING FORM — ${project.clientName}

Primary contact person:
Phone number:
Email address:
Position / Title:

Update report frequency (please select one):
[ ] Weekly
[ ] Fortnightly
[ ] Monthly

Is a Purchase Order (PO) required for invoicing?
[ ] Yes — PO number: ________________________________
[ ] No

Accounts payable contact name:
Accounts payable email:
Accounts payable phone:

Preferred dates for kick-off meeting (please provide 3 options):
Option 1:
Option 2:
Option 3:

Any other information or requirements we should be aware of:


Please return this form by reply email to diane.k@risk2solution.com at your earliest convenience so we can get your project started.

Sign off as:
Kind regards,

Diane Kruger
Corporate Operations Lead
Risk 2 Solution Group
P: 1300 459 970 | M: +61 415 748 747
E: diane.k@risk2solution.com
W: www.risk2solution.com
Queensland, Australia`,
      context
    );

    const draft = {
      id: `d_${Date.now()}_onboard`,
      projectId: project.id, clientName: project.clientName,
      projectName: project.projectName, type: 'client_onboarding',
      urgency: 'routine', toName: project.clientContact,
      toEmail: project.clientEmail,
      subject: `Welcome to R2S — ${project.projectName || project.clientName}`,
      body: onboardingBody, source: 'auto',
    };
    await db.saveDraft(draft);
    await saveDraftEmail(draft);
    console.log(`[Onboarding] Client welcome email drafted for ${project.clientName}`);
  } catch (err) {
    if (err.message === 'MONTHLY_CAP_REACHED') throw err;
    console.error('[Onboarding] Draft failed:', err.message);
  }
}

// ── Risk escalation drafter ───────────────────────────────────────────────────
async function draftRiskEscalation(project, riskDescription) {
  const DIANE = 'diane.k@risk2solution.com';
  try {
    const context = buildContext(project);

    // Draft escalation email to client (for Diane to review — NOT auto-sent)
    const escalationBody = await aurora('escalation_email',
      `Write a professional risk escalation email from R2S to ${project.clientContact || 'the client'} at ${project.clientName}.

FORMAT RULES:
- Plain text only. No asterisks, no bold markdown, no long dashes.
- Professional and measured tone — serious but not alarming.
- Keep it factual and solution-focused.

Risk description: ${riskDescription}
Project: ${project.projectName || project.clientName}
Current phase: ${PHASES[project.phase || 0]}

Write starting with "Hi ${(project.clientContact || 'there').split(' ')[0]},"

Cover:
1. Reason for the escalation (clear and factual, 2-3 sentences)
2. Current impact or risk to the project
3. Proposed actions or next steps (bullet points)
4. Request for a brief call or response to agree on the path forward

Sign off as Diane Kruger with full signature.`,
      context
    );

    const draft = {
      id: `d_${Date.now()}_escalation`,
      projectId: project.id, clientName: project.clientName,
      projectName: project.projectName, type: 'escalation_email',
      urgency: 'urgent', toName: project.clientContact,
      toEmail: project.clientEmail,
      subject: `Project Update — ${project.projectName || project.clientName} — Action Required`,
      body: escalationBody, source: 'auto',
    };
    await db.saveDraft(draft);
    await saveDraftEmail(draft);

    // Also notify Diane internally
    await sendEmail(DIANE,
      `[Aurora] Risk escalation draft ready: ${project.clientName}`,
      `Aurora has identified a risk on the ${project.clientName} project and drafted an escalation email for your review.

Risk: ${riskDescription}
Project: ${project.projectName || project.clientName}
Phase: ${PHASES[project.phase || 0]}
Client contact: ${project.clientContact || 'See project record'}

The draft escalation email is waiting in the Comms Drafts section of Aurora and in the info@risk2solution.com Outlook shared mailbox drafts folder.

Please review before sending.

${process.env.FRONTEND_URL ? 'Aurora portal: ' + process.env.FRONTEND_URL : ''}

Aurora
R2S Project Management Intelligence`,
      true
    );
    console.log(`[Escalation] Draft created for ${project.clientName}`);
  } catch (err) {
    if (err.message === 'MONTHLY_CAP_REACHED') throw err;
    console.error('[Escalation] Draft failed:', err.message);
  }
}

// ── Client satisfaction / feedback email (triggered on close-out) ─────────────
async function draftClientFeedback(project) {
  try {
    const context = buildContext(project);
    const feedbackBody = await aurora('status_email',
      `Write a professional post-project feedback request email from R2S to ${project.clientContact || 'the client'} at ${project.clientName}.

FORMAT RULES:
- Plain text only. No asterisks, no markdown, no long dashes.
- Warm, genuine, and brief.

Write starting with "Hi ${(project.clientContact || 'there').split(' ')[0]},"

Cover:
1. Thank the client for the engagement (1-2 sentences)
2. Brief note on what was delivered
3. A genuine request for feedback — how did R2S perform, what could be improved, would they recommend R2S
4. Optional: mention R2S would welcome the opportunity to continue supporting them
5. Offer to arrange a brief debrief call if they would find it useful

Sign off as Diane Kruger with full signature.`,
      context
    );

    const draft = {
      id: `d_${Date.now()}_feedback`,
      projectId: project.id, clientName: project.clientName,
      projectName: project.projectName, type: 'feedback_request',
      urgency: 'routine', toName: project.clientContact,
      toEmail: project.clientEmail,
      subject: `Thank you — ${project.projectName || project.clientName} — Your feedback`,
      body: feedbackBody, source: 'auto',
    };
    await db.saveDraft(draft);
    await saveDraftEmail(draft);
    console.log(`[Feedback] Draft created for ${project.clientName}`);
  } catch (err) {
    console.error('[Feedback] Draft failed:', err.message);
  }
}

// ── Cron schedules ───────────────────────────────────────────────────────────

// Daily batch: 6am AEST (UTC+10) = 8pm UTC previous day
cron.schedule('0 20 * * *', () => runBatch().catch(console.error), { timezone: 'UTC' });

// Wednesday afternoon follow-up: 4pm AEST = 6am UTC Thursday (UTC+10 offset)
// Wednesday 4pm AEST = Wednesday 6am UTC
cron.schedule('0 6 * * 3', () => sendPMChecklistFollowUp().catch(console.error), { timezone: 'UTC' });

// Email polling: every hour 7am-7pm AEST weekdays
// AEST UTC+10: 7am=9pm UTC, 7pm=9am UTC
cron.schedule('0 21-23,0-9 * * 1-5', async () => {
  console.log('[Poll] Scheduled hourly check...');
  try { await readConsultantReplies(); } catch(e) { console.error('[Poll] Error:', e.message); }
}, { timezone: 'UTC' });

// Keepalive ping every 5 minutes — prevents Azure from sleeping the app
cron.schedule('*/5 * * * *', async () => {
  try {
    const PORT = process.env.PORT || 8080;
    await axios.get(`http://localhost:${PORT}/api/health`, { timeout: 5000 }).catch(() => {});
  } catch(e) {}
}, { timezone: 'UTC' });

// ── API Routes ────────────────────────────────────────────────────────────────
app.get('/api/health', (req, res) => res.json({ status: 'ok', service: 'Aurora R2S v3', db: !!process.env.DATABASE_URL ? 'postgres' : 'json' }));

// Projects
app.get('/api/projects', async (req, res) => {
  try { res.json({ projects: await db.getProjects() }); }
  catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/projects', express.json(), async (req, res) => {
  try {
    const p = req.body;
    if (!p.clientName) return res.status(400).json({ error: 'clientName required' });
    p.id = p.id || `p_${Date.now()}`;
    p.type = p.type || 'standard';
    p.phase = p.phase || 0;
    await db.upsertProject(p);
    res.json({ project: p });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.put('/api/projects/:id', express.json(), async (req, res) => {
  try {
    const before = await db.getProject(req.params.id);
    const updated = await db.updateProjectField(req.params.id, req.body);
    // Log manual updates
    if (req.body.phase !== undefined && before && req.body.phase !== before.phase) {
      await db.logActivity(req.params.id, { type: 'phase_change', summary: `Phase updated to ${PHASES[req.body.phase]} by Diane` });
    }
    if (req.body.status !== undefined && before && req.body.status !== before.status) {
      await db.logActivity(req.params.id, { type: 'status_change', summary: `Status updated to ${req.body.status} by Diane` });
    }
    if (req.body.consultant !== undefined && before && req.body.consultant !== before.consultant) {
      await db.logActivity(req.params.id, { type: 'consultant_assigned', summary: `Consultant/trainer updated to: ${req.body.consultant}` });
    }
    if (req.body.dueDate !== undefined && before && req.body.dueDate !== before.dueDate) {
      await db.logActivity(req.params.id, { type: 'manual_note', summary: `Due date updated to: ${req.body.dueDate}` });
    }
    // If phase changed to Close-out (4), set 1yr and 2yr follow-up reminders
    if (req.body.phase === 4 || req.body.status === 'Completed') {
      const p = await db.getProject(req.params.id);
      if (p) {
        const yr1 = new Date(); yr1.setFullYear(yr1.getFullYear() + 1);
        const yr2 = new Date(); yr2.setFullYear(yr2.getFullYear() + 2);
        await createCalendarReminder(
          `1-year follow-up: ${p.clientName}`,
          `Check in with ${p.clientName} (${p.clientContact || ''}) — explore new service needs, pain points, and opportunities for R2S.`,
          yr1.toISOString().slice(0,10)
        );
        await createCalendarReminder(
          `2-year follow-up: ${p.clientName}`,
          `2-year relationship check-in with ${p.clientName}. Review their current situation and how R2S can help.`,
          yr2.toISOString().slice(0,10)
        );
        await sendInternalEmail(
          `[Aurora] Project closed: ${p.clientName}`,
          `The ${p.projectName || p.clientName} project has been marked complete.\n\nAurora has set 1-year and 2-year follow-up reminders in the Outlook calendar.\n\nAurora\nR2S Project Management Intelligence`
        );
        // Draft client feedback email
        try { await draftClientFeedback(p); } catch(e) { console.error('[Feedback]', e.message); }
      }
    }
    res.json({ project: updated });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.delete('/api/projects/:id', async (req, res) => {
  try {
    const id = req.params.id;
    const before = await db.getProjects();
    const exists = before.find(p => p.id === id);
    if (!exists) return res.status(404).json({ error: 'Project not found' });
    console.log(`[Delete] Removing: ${exists.clientName} (${id}), total before: ${before.length}`);
    await db.deleteProject(id);
    // Small delay to ensure Blob write completes before verification
    await new Promise(r => setTimeout(r, 500));
    const after = await db.getProjects();
    console.log(`[Delete] After: ${after.length} projects, deleted ID still present: ${!!after.find(p => p.id === id)}`);
    // Return the updated list so client doesn't need to re-fetch
    res.json({ success: true, projects: after });
  } catch (e) {
    console.error('[Delete] Error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// Trigger consultant briefing for a specific consultant on a project
app.post('/api/projects/:id/briefing', express.json(), async (req, res) => {
  try {
    const project = await db.getProject(req.params.id);
    if (!project) return res.status(404).json({ error: 'Project not found' });

    // Allow specifying a single consultant name (for multi-consultant support)
    const consultantName  = req.body?.consultantName || project.consultant;
    if (!consultantName) return res.status(400).json({ error: 'No consultant assigned to this project' });

    const CONSULTANT_EMAILS = {
      'Mick Harran':'info@risk2solution.com','Paul Johnston':'info@risk2solution.com',
      'Dave Cohen':'dave.c@risk2solution.com','Ross Mackenzie':'info@risk2solution.com',
      'Lawrence Phillips':'info@risk2solution.com','Marina Toailoa':'info@risk2solution.com',
      'Gavriel Schneider':'info@risk2solution.com','Pierre Andipatin':'info@risk2solution.com',
      'Daniel Du Plessis':'info@risk2solution.com','Gavriel Guriel':'info@risk2solution.com',
    };

    const docs    = await db.getDocuments(req.params.id);
    const context = buildContext(project, docs);

    // Build a single-consultant version of the project for the briefing
    const briefProject = { ...project, consultant: consultantName, consultantEmail: CONSULTANT_EMAILS[consultantName] || project.consultantEmail || 'info@risk2solution.com' };
    const extracted = {
      consultant: consultantName,
      consultantEmail: briefProject.consultantEmail,
      flightsRequired: project.flightsRequired,
      accommodationRequired: project.accommodationRequired,
    };

    const draft = await sendConsultantBriefing(briefProject, extracted, context);
    res.json({ success: true, draft });
  } catch (err) {
    if (err.message === 'MONTHLY_CAP_REACHED') return res.status(429).json({ error: 'Monthly cap reached' });
    console.error('[Briefing endpoint]', err.message);
    res.status(500).json({ error: err.message });
  }
});

// Monday sync (backup)
app.post('/api/projects/sync', async (req, res) => {
  try { res.json({ projects: await syncMonday() }); }
  catch (e) { res.status(500).json({ error: e.message }); }
});

// Contract upload — auto-creates project
app.post('/api/contracts/upload', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
    const spend = await db.getSpend();
    if (spend.total >= CAP_USD) return res.status(429).json({ error: 'Monthly cap reached' });

    console.log(`[Contract] Reading ${req.file.originalname}...`);
    const rawText = await extractTextFromFile(req.file.path, req.file.mimetype);
    if (!rawText || rawText.trim().length < 100) {
      fs.unlinkSync(req.file.path);
      return res.status(400).json({ error: 'Could not extract text from this file. Try a different format.' });
    }

    // ── Duplicate detection — do a quick name scan before full extraction ──────
    // Extract just client/project name cheaply using Haiku before full Sonnet extraction
    if (!req.body.forceCreate) {
      try {
        const quickScan = await aurora('status_email',
          `Read the first 2000 characters of this document and extract ONLY:
1. The client or organisation name
2. The project or engagement name or title

Return as JSON only: {"clientName": "...", "projectName": "..."}
If not found use empty string.

Document start:
${rawText.slice(0, 2000)}`,
          null
        );
        const quickData = JSON.parse(quickScan.replace(/\`\`\`json|\`\`\`/g,'').trim());
        const existingProjects = await db.getProjects();

        // Check for similar existing projects
        const duplicates = existingProjects.filter(p => {
          const clientMatch = quickData.clientName &&
            p.clientName?.toLowerCase().includes(quickData.clientName.toLowerCase().slice(0,8)) ||
            quickData.clientName?.toLowerCase().includes((p.clientName||'').toLowerCase().slice(0,8));
          const projectMatch = quickData.projectName &&
            p.projectName?.toLowerCase().includes(quickData.projectName.toLowerCase().slice(0,10)) ||
            quickData.projectName?.toLowerCase().includes((p.projectName||'').toLowerCase().slice(0,10));
          return clientMatch || projectMatch;
        });

        if (duplicates.length > 0) {
          // Don't delete the file yet — return duplicate warning so Diane can decide
          return res.status(409).json({
            duplicate: true,
            message: `A similar project may already exist in Aurora.`,
            existingProjects: duplicates.map(p => ({
              id: p.id,
              clientName: p.clientName,
              projectName: p.projectName,
              status: p.status,
              phase: p.phase,
            })),
            detectedClient: quickData.clientName,
            detectedProject: quickData.projectName,
            fileStillUploaded: true,
          });
        }
      } catch (scanErr) {
        // If quick scan fails, continue with full extraction
        console.log('[Contract] Quick scan failed, proceeding with full extraction:', scanErr.message);
      }
    }

    const extracted = await analyseContract(rawText, req.file.originalname);
    console.log(`[Contract] Extracted: ${extracted.clientName} — ${extracted.projectName}`);

    // Create project from extracted data
    const projectId = req.body.projectId || `p_${Date.now()}`;
    const project = {
      id:                   projectId,
      clientName:           extracted.organisationName || extracted.clientName || req.body.clientName || 'Unknown client',
      projectName:          extracted.projectName || req.file.originalname,
      clientContact:        extracted.clientContact || '',
      clientEmail:          extracted.clientEmail || '',
      clientPhone:          extracted.clientPhone || '',
      value:                extracted.value || '',
      contractStart:        extracted.contractStart || '',
      dueDate:              extracted.dueDate || '',
      summary:              extracted.summary || '',
      deliverables:         extracted.deliverables || '',
      milestones:           extracted.milestones || '',
      timeline:             extracted.timeline || '',
      invoicingNotes:       extracted.invoicingNotes || '',
      consultant:           extracted.consultant || '',
      consultantEmail:      extracted.consultantEmail || '',
      flightsRequired:      extracted.flightsRequired || '',
      accommodationRequired:extracted.accommodationRequired || '',
      notes:                extracted.notes || '',
      status:               'In Progress',
      phase:                0,
      type:                 'standard',
    };

    if (!req.body.projectId) {
      await db.upsertProject(project);
    } else {
      await db.updateProjectField(projectId, {
        summary: extracted.summary, deliverables: extracted.deliverables,
        milestones: extracted.milestones, timeline: extracted.timeline,
        value: extracted.value, dueDate: extracted.dueDate,
        contractStart: extracted.contractStart, invoicingNotes: extracted.invoicingNotes,
        consultant: extracted.consultant, consultantEmail: extracted.consultantEmail,
        clientContact: extracted.clientContact, clientEmail: extracted.clientEmail,
        clientPhone: extracted.clientPhone,
        flightsRequired: extracted.flightsRequired,
        accommodationRequired: extracted.accommodationRequired,
      });
    }

    // Save document extract (text only — delete the original file to save space)
    await db.saveDocument({
      id: `doc_${Date.now()}`,
      projectId, name: req.file.originalname,
      extract: rawText.slice(0, 8000),
      type: 'contract_extract',
    });

    // Delete the uploaded file immediately — we only need the extracted text
    try {
      fs.unlinkSync(req.file.path);
      console.log(`[Upload] Contract file deleted after extraction: ${req.file.originalname}`);
    } catch (unlinkErr) {
      console.error('[Upload] Could not delete file:', unlinkErr.message);
    }

    // Don't auto-send briefing on upload — consultant is pre-filled but Diane confirms via dropdown
    // The briefingPrepared flag tells the UI which consultants were found so Diane can confirm
    // Draft client onboarding email when project is first created
    if (!req.body.projectId) {
      try { await draftClientOnboarding(project); } catch(e) { console.error('[Onboarding]', e.message); }
      // Log project creation
      await db.logActivity(projectId, { type: 'project_created', summary: `Project created from contract upload — ${project.clientName} ${project.projectName || ''}` });
      if (project.consultant) await db.logActivity(projectId, { type: 'consultant_assigned', summary: `${project.consultant} identified as consultant/trainer from proposal` });
      if (project.value) await db.logActivity(projectId, { type: 'contract', summary: `Contract value extracted: ${project.value}` });
    }

    res.json({ project, extracted, briefingPrepared: false, suggestedConsultants: (() => {
      // Only suggest consultants actually named in the proposal
      const KNOWN = ['Mick Harran','Paul Johnston','Dave Cohen','Ross Mackenzie','Lawrence Phillips','Marina Toailoa','Gavriel Schneider','Pierre Andipatin','Daniel Du Plessis','Gavriel Guriel'];
      const raw = (extracted.consultant || '') + ' ' + (extracted.consultantEmail || '');
      const found = KNOWN.filter(name => {
        const [first, last] = name.toLowerCase().split(' ');
        return raw.toLowerCase().includes(first) || (last && raw.toLowerCase().includes(last));
      });
      // If name from proposal doesn't match our list, include it directly
      if (found.length === 0 && extracted.consultant) found.push(extracted.consultant.trim());
      return found;
    })() });
  } catch (e) {
    if (e.message === 'MONTHLY_CAP_REACHED') return res.status(429).json({ error: 'Monthly cap reached' });
    console.error('[Contract] Error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// Drafts
app.get('/api/drafts', async (req, res) => {
  try { res.json({ drafts: await db.getDrafts() }); }
  catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/drafts/generate', express.json(), async (req, res) => {
  try {
    const { projectId, taskType, prompt } = req.body;
    const p = await db.getProject(projectId);
    if (!p) return res.status(404).json({ error: 'Project not found' });
    if (p.type === 'ongoing') return res.status(400).json({ error: 'Ongoing projects do not use Aurora automation' });

    const docs = await db.getDocuments(projectId);
    const context = buildContext(p, docs);
    const text = await aurora(taskType || 'status_email', prompt || `Draft a ${(taskType||'status email').replace(/_/g,' ')} for ${p.projectName||p.clientName} at ${p.clientName}.`, context);

    const draft = {
      id: `d_${Date.now()}_${Math.random().toString(36).slice(2,6)}`,
      projectId, clientName: p.clientName, projectName: p.projectName,
      type: taskType || 'status_email', urgency: taskType?.includes('escalat') ? 'urgent' : 'routine',
      toName: p.clientContact, toEmail: p.clientEmail,
      subject: `${p.projectName || p.clientName}`, body: text, source: 'manual',
    };
    await db.saveDraft(draft);
    await saveDraftEmail(draft);
    res.json({ draft });
  } catch (e) {
    if (e.message === 'MONTHLY_CAP_REACHED') return res.status(429).json({ error: 'Monthly cap reached' });
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/drafts/:id/approve', async (req, res) => {
  try {
    const drafts = await db.getDrafts();
    const draft  = drafts.find(d => d.id === req.params.id);
    if (!draft) return res.status(404).json({ error: 'Draft not found' });

    // Save to Outlook Drafts folder ONLY — Diane sends manually from Outlook
    await saveDraftEmail(draft);
    console.log(`[Drafts] Saved to Outlook Drafts: ${draft.subject} → ${draft.toEmail}`);

    // Mark as approved in Aurora
    await db.updateDraft(req.params.id, { approved: true, approvedAt: new Date().toISOString() });

    // Log activity
    if (draft.projectId) {
      const typeLabels = {
        consultant_briefing: 'Consultant briefing saved to Outlook Drafts',
        client_onboarding: 'Client onboarding email saved to Outlook Drafts',
        status_email: 'Client status update saved to Outlook Drafts',
        kickoff_agenda: 'Kick-off agenda saved to Outlook Drafts',
        invoice_reminder: 'Invoice reminder saved to Outlook Drafts',
        escalation_email: 'Escalation email saved to Outlook Drafts',
        feedback_request: 'Client feedback email saved to Outlook Drafts',
        consultant_checkin: 'Consultant check-in saved to Outlook Drafts',
      };
      const label = typeLabels[draft.type] || `${(draft.type||'').replace(/_/g,' ')} saved to Outlook Drafts`;
      await db.logActivity(draft.projectId, { type: 'draft_sent', summary: `${label} — ready for Diane to review and send to ${draft.toName || draft.toEmail}` });
    }

    // If consultant briefing approved, prompt Diane about kick-off scheduling
    if (draft.type === 'consultant_briefing' && draft.projectId) {
      const project = await db.getProject(draft.projectId);
      if (project) {
        await sendKickoffPrompt(project);
        console.log(`[Drafts] Kick-off prompt sent for ${project.clientName}`);
      }
    }

    res.json({ success: true, message: 'Saved to Outlook Drafts — open Outlook to send' });
  } catch (e) {
    console.error('[Drafts] Approve error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/drafts/:id/reject', async (req, res) => {
  try { await db.updateDraft(req.params.id, { rejected: true }); res.json({ success: true }); }
  catch (e) { res.status(500).json({ error: e.message }); }
});

app.put('/api/drafts/:id', express.json(), async (req, res) => {
  try { await db.updateDraft(req.params.id, req.body); res.json({ success: true }); }
  catch (e) { res.status(500).json({ error: e.message }); }
});

// Chat
app.post('/api/chat', express.json(), async (req, res) => {
  try {
    const { message, projectId, history } = req.body;
    if (!message) return res.status(400).json({ error: 'message required' });
    let context = null;
    if (projectId) {
      const p    = await db.getProject(projectId);
      const docs = await db.getDocuments(projectId);
      if (p && p.type !== 'ongoing') context = buildContext(p, docs);
    }
    const turns = (history || []).slice(-6);
    const fullMessage = turns.length
      ? turns.map(t => `${t.role === 'user' ? 'User' : 'Aurora'}: ${t.content}`).join('\n') + `\nUser: ${message}`
      : message;
    const reply = await aurora('chat', fullMessage, context);
    res.json({ reply });
  } catch (e) {
    if (e.message === 'MONTHLY_CAP_REACHED') return res.status(429).json({ error: 'Monthly cap reached' });
    res.status(500).json({ error: e.message });
  }
});

// Documents
app.get('/api/documents', async (req, res) => {
  try { res.json({ documents: await db.getDocuments(req.query.projectId) }); }
  catch (e) { res.status(500).json({ error: e.message }); }
});

// Reports — generate and stream as downloadable files
app.post('/api/reports', express.json(), async (req, res) => {
  try {
    const { reportType, projectId } = req.body;
    const allProjects = await db.getProjects();
    const targets = projectId
      ? allProjects.filter(p => p.id === projectId && p.type !== 'ongoing')
      : allProjects.filter(p => p.type !== 'ongoing');
    if (!targets.length) return res.status(404).json({ error: 'No projects found' });

    const contextBlock = (await Promise.all(targets.map(async p => {
      const docs = await db.getDocuments(p.id);
      return buildContext(p, docs);
    }))).filter(Boolean).join('\n\n---\n\n');

    // Determine format
    const excelTypes = ['milestones','risks','invoices'];
    const isExcel = excelTypes.includes(reportType);

    const prompts = {
      status:    `Write a clear project status report. Use bullet points and numbered lists. No asterisks, no long dashes, no markdown symbols.
For each project write:
Project name and client
Current phase (one of: Kick-off, Deployment, Monitoring & Review, Reporting, Close-out)
Status: On Track / At Risk / On Hold
Delivered so far (bullet points)
Outstanding (bullet points)
Risks or blockers (bullet points if any)
Next steps (bullet points)`,

      milestones:`Generate a deliverables tracker. Return as tab-separated values with these exact columns:
Client\tProject\tDeliverable\tPhase\tDue Date\tStatus\tNotes
One row per deliverable. Status must be one of: Complete, In Progress, Outstanding, Overdue.
No extra text, no asterisks, no dashes. Just the data rows after the header.`,

      risks:     `Generate a risk register. Return as tab-separated values with these exact columns:
Client\tProject\tRisk Description\tLikelihood\tImpact\tRisk Level\tMitigation Action\tOwner
Likelihood and Impact: High/Medium/Low. Risk Level: High/Medium/Low.
No extra text, no asterisks, no dashes. Just data rows after the header.`,

      closeout:  `Write a project close-out report. Use bullet points and numbered lists. No asterisks, no markdown.
Include: Project summary, Deliverables completed (numbered list), Key outcomes, Invoicing summary, Lessons learned (bullet points), Recommendations for the client (numbered list).`,

      invoices:  `Generate an invoice summary. Return as tab-separated values with these exact columns:
Client\tProject\tContract Value\tInvoicing Terms\tInvoiced To Date\tOutstanding\tNext Invoice Due\tNotes
No extra text, no asterisks, no dashes. Just data rows after the header.`,

      portfolio: `Write a portfolio overview for R2S leadership. Use bullet points and numbered lists. No asterisks, no markdown.
Lead with anything needing immediate attention. Then list each active project: client, project name, phase, status, contract value, key date.
End with a 3-month revenue forecast based on contract values.`,
    };

    const taskType = ['portfolio','closeout'].includes(reportType) ? 'portfolio_report' : 'status_report';
    const reportContent = await aurora(taskType, prompts[reportType] || prompts.status, contextBlock);

    const now = new Date().toISOString().slice(0,10);
    const safeName = reportType.replace(/[^a-z0-9]/gi,'_');

    // If this was a risk or deliverable report for a specific project, save it
    if (projectId) {
      try {
        if (reportType === 'risks') {
          const risks = parseRiskRegister(reportContent, projectId);
          if (risks.length > 0) await db.saveRiskRegister(projectId, risks);
        } else if (reportType === 'milestones') {
          const deliverables = parseDeliverables(reportContent, projectId);
          if (deliverables.length > 0) await db.saveDeliverables(projectId, deliverables);
        }
      } catch (saveErr) {
        console.error('[Reports] Save to project failed:', saveErr.message);
      }
    }

    if (isExcel) {
      // Return as CSV (opens in Excel when saved as .csv)
      const lines = reportContent.split('\n').filter(l => l.trim());
      const csv = lines.map(line =>
        line.split('\t').map(cell => {
          const c = (cell || '').replace(/"/g, '""').trim();
          return c.includes(',') || c.includes('"') || c.includes('\n') ? `"${c}"` : c;
        }).join(',')
      ).join('\r\n');

      res.setHeader('Content-Type', 'text/csv; charset=utf-8');
      res.setHeader('Content-Disposition', `attachment; filename="R2S_${safeName}_${now}.csv"`);
      res.send('\uFEFF' + csv); // UTF-8 BOM so Excel opens correctly
    } else {
      // Return as plain text formatted for Word (.txt that Word can open)
      const wordContent = reportContent
        .replace(/\*\*?/g, '')       // remove asterisks
        .replace(/^-{3,}$/gm, '')    // remove horizontal rules
        .replace(/\u2014|\u2013/g, '-') // replace em/en dashes with hyphen
        .replace(/_{2,}/g, '')        // remove underscores used as dividers
        .trim();

      const header = `R2S PROJECT MANAGEMENT
${reportType.toUpperCase().replace(/_/g,' ')} REPORT
Generated: ${new Date().toLocaleDateString('en-AU')}
${targets.length > 1 ? 'Portfolio — All projects' : targets[0]?.clientName || ''}

${'='.repeat(60)}

`;
      res.setHeader('Content-Type', 'application/msword');
      res.setHeader('Content-Disposition', `attachment; filename="R2S_${safeName}_${now}.doc"`);
      res.send(header + wordContent);
    }

  } catch (e) {
    if (e.message === 'MONTHLY_CAP_REACHED') return res.status(429).json({ error: 'Monthly cap reached' });
    res.status(500).json({ error: e.message });
  }
});

// Suggestions
app.get('/api/suggestions', async (req, res) => {
  try {
    const suggestions = await db.getSuggestions();
    res.json({ suggestions });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/suggestions/generate', async (req, res) => {
  try {
    await generateSuggestions();
    const suggestions = await db.getSuggestions();
    res.json({ suggestions });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/suggestions/:id/approve', express.json(), async (req, res) => {
  try {
    const id = req.params.id;
    // Get suggestion from Blob (db.getSuggestions reads from Blob)
    const suggestions = await db.getSuggestions();
    const suggestion = suggestions.find(s => s.id === id);
    if (!suggestion) {
      console.error(`[Suggestions] Approve: suggestion ${id} not found. Available: ${suggestions.map(s=>s.id).join(',')}`);
      return res.status(404).json({ error: 'Suggestion not found' });
    }
    console.log(`[Suggestions] Approving: ${suggestion.title} for ${suggestion.clientName}`);
    // Apply the action FIRST
    await applySuggestion(suggestion);
    // Then mark as approved so it doesn't reappear
    await db.updateSuggestion(id, 'approved');
    console.log(`[Suggestions] ✓ Approved and applied: ${suggestion.title}`);
    res.json({ success: true });
  } catch (e) {
    console.error('[Suggestions] Approve error:', e.message);
    if (e.message === 'MONTHLY_CAP_REACHED') return res.status(429).json({ error: 'Monthly cap reached' });
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/suggestions/:id/dismiss', async (req, res) => {
  try {
    const id = req.params.id;
    console.log(`[Suggestions] Dismissing: ${id}`);
    await db.updateSuggestion(id, 'dismissed');
    console.log(`[Suggestions] ✓ Dismissed: ${id}`);
    res.json({ success: true });
  } catch (e) {
    console.error('[Suggestions] Dismiss error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ── Risk Register routes ─────────────────────────────────────────────────────
app.get('/api/projects/:id/risks', async (req, res) => {
  try { res.json({ risks: await db.getRiskRegister(req.params.id) }); }
  catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/projects/:id/risks/generate', async (req, res) => {
  try {
    const project = await db.getProject(req.params.id);
    if (!project) return res.status(404).json({ error: 'Project not found' });
    const docs = await db.getDocuments(req.params.id);
    const context = buildContext(project, docs);

    const tsvContent = await aurora('status_report',
      `Generate a risk register for this project. Return as tab-separated values with these exact columns:
Client\tProject\tRisk Description\tLikelihood\tImpact\tRisk Level\tMitigation Action\tOwner
Likelihood and Impact: High/Medium/Low. Risk Level: High/Medium/Low.
Owner should be Diane Kruger unless a consultant is clearly responsible.
Include 5-8 realistic risks for this type of engagement.
No extra text, no asterisks, no dashes. Just data rows after the header.`,
      context
    );

    const risks = parseRiskRegister(tsvContent, req.params.id);
    await db.saveRiskRegister(req.params.id, risks);
    res.json({ risks });
  } catch (e) {
    if (e.message === 'MONTHLY_CAP_REACHED') return res.status(429).json({ error: 'Monthly cap reached' });
    res.status(500).json({ error: e.message });
  }
});

app.put('/api/projects/:id/risks/:riskId', express.json(), async (req, res) => {
  try {
    const risk = await db.updateRisk(req.params.id, req.params.riskId, req.body);
    res.json({ risk });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// Also generate and save risk register when report is downloaded
// (handled by hooking into the reports endpoint — see reports endpoint update below)

// ── Invoice tracking routes ──────────────────────────────────────────────────
app.get('/api/projects/:id/invoices', async (req, res) => {
  try {
    const invoices = await readInvoices(req.params.id);
    res.json({ invoices });
  } catch (e) { res.json({ invoices: [] }); }
});

app.post('/api/projects/:id/invoices', express.json(), async (req, res) => {
  try {
    const key = `invoices_${req.params.id}.json`;
    const existing = await readInvoices(req.params.id);
    const invoice = { id: `inv_${Date.now()}`, ...req.body, createdAt: new Date().toISOString(), paid: false };
    existing.push(invoice);
    await writeInvoices(req.params.id, existing);
    // Log activity
    await db.logActivity(req.params.id, { type: 'contract', summary: `Invoice ${invoice.invoiceNumber || ''} sent to client — $${invoice.amount || '0'}` });
    res.json({ invoice });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.put('/api/projects/:id/invoices/:invId', express.json(), async (req, res) => {
  try {
    const key = `invoices_${req.params.id}.json`;
    const invoices = await readInvoices(req.params.id);
    const idx = invoices.findIndex(i => i.id === req.params.invId);
    if (idx >= 0) {
      Object.assign(invoices[idx], req.body);
      if (req.body.paid) {
        invoices[idx].paidAt = new Date().toISOString();
        await db.logActivity(req.params.id, { type: 'contract', summary: `Invoice ${invoices[idx].invoiceNumber || ''} marked as paid — $${invoices[idx].amount || '0'}` });
      }
      await writeInvoices(req.params.id, invoices);
    }
    res.json({ invoice: invoices[idx] });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.delete('/api/projects/:id/invoices/:invId', async (req, res) => {
  try {
    const invoices = await readInvoices(req.params.id);
    const filtered = invoices.filter(i => i.id !== req.params.invId);
    await writeInvoices(req.params.id, filtered);
    await db.logActivity(req.params.id, { type: 'contract', summary: `Invoice deleted` });
    res.json({ success: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── Deliverables Tracker routes ───────────────────────────────────────────────
app.get('/api/projects/:id/deliverables', async (req, res) => {
  try { res.json({ deliverables: await db.getDeliverables(req.params.id) }); }
  catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/projects/:id/deliverables/generate', async (req, res) => {
  try {
    const project = await db.getProject(req.params.id);
    if (!project) return res.status(404).json({ error: 'Project not found' });
    const docs = await db.getDocuments(req.params.id);
    const context = buildContext(project, docs);

    const tsvContent = await aurora('status_report',
      `Generate a deliverables tracker for this project. Return as tab-separated values with these exact columns:
Client\tProject\tDeliverable\tPhase\tDue Date\tStatus\tAssigned To\tNotes
Status must be one of: Complete, In Progress, Outstanding, Overdue.
Extract specific deliverables from the project scope and contract details.
Assigned To should be the consultant/trainer if known, otherwise Diane Kruger.
No extra text, no asterisks, no dashes. Just data rows after the header.`,
      context
    );

    const deliverables = parseDeliverables(tsvContent, req.params.id);
    await db.saveDeliverables(req.params.id, deliverables);
    res.json({ deliverables });
  } catch (e) {
    if (e.message === 'MONTHLY_CAP_REACHED') return res.status(429).json({ error: 'Monthly cap reached' });
    res.status(500).json({ error: e.message });
  }
});

app.put('/api/projects/:id/deliverables/:delId', express.json(), async (req, res) => {
  try {
    const del = await db.updateDeliverable(req.params.id, req.params.delId, req.body);
    res.json({ deliverable: del });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.delete('/api/projects/:id/deliverables/:delId', async (req, res) => {
  try {
    const deliverables = await db.getDeliverables(req.params.id);
    const filtered = deliverables.filter(d => d.id !== req.params.delId);
    await db.saveDeliverables(req.params.id, filtered);
    await db.logActivity(req.params.id, { type: 'manual_note', summary: `Deliverable deleted` });
    res.json({ success: true, deliverables: filtered });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── Activity log ─────────────────────────────────────────────────────────────
app.get('/api/projects/:id/activity', async (req, res) => {
  try {
    let activity = await db.getActivityLog(req.params.id);
    console.log(`[Activity] Project ${req.params.id}: ${activity.length} entries found`);

    // If empty, auto-generate historical log from project data
    if (!activity.length) {
      console.log(`[Activity] Auto-generating history for project ${req.params.id}`);
      const project = await db.getProject(req.params.id);
      if (project) {
        const base = project.createdAt || project.updatedAt || new Date().toISOString();
        const entries = [
          { type: 'project_created', summary: `Project created — ${project.clientName}${project.projectName ? ': ' + project.projectName : ''}`, timestamp: base },
          ...(project.consultant ? [{ type: 'consultant_assigned', summary: `${project.consultant} assigned as consultant/trainer`, timestamp: base }] : []),
          ...(project.value ? [{ type: 'contract', summary: `Contract value: ${project.value}`, timestamp: base }] : []),
          ...(project.dueDate ? [{ type: 'contract', summary: `Project due date: ${project.dueDate}`, timestamp: base }] : []),
          { type: 'phase_change', summary: `Current phase: ${PHASES[project.phase || 0]}`, timestamp: project.updatedAt || base },
          ...(project.flightsRequired === 'yes' ? [{ type: 'manual_note', summary: 'Flights required for this engagement', timestamp: base }] : []),
          ...(project.accommodationRequired === 'yes' ? [{ type: 'manual_note', summary: 'Accommodation required for this engagement', timestamp: base }] : []),
        ];
        for (const e of entries) {
          await db.logActivity(req.params.id, e);
        }
        activity = await db.getActivityLog(req.params.id);
        console.log(`[Activity] Generated ${activity.length} historical entries`);
      }
    }

    res.json({ activity });
  } catch (e) {
    console.error('[Activity] Error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/projects/:id/activity', express.json(), async (req, res) => {
  try {
    const entry = await db.logActivity(req.params.id, req.body);
    res.json({ entry });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// ── Debug email reader — shows what Aurora sees without processing ────────────
app.get('/api/debug/emails', async (req, res) => {
  try {
    const token = await getOutlookToken();
    if (!token) return res.json({ error: 'No Outlook token — check Azure app credentials' });
    const mailbox = process.env.OUTLOOK_SHARED_MAILBOX || 'info@risk2solution.com';
    const since = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000).toISOString();
    const url = `https://graph.microsoft.com/v1.0/users/${mailbox}/mailFolders/inbox/messages?$select=id,subject,from,receivedDateTime,isRead,categories&$top=20&$orderby=receivedDateTime desc`;
    const res2 = await axios.get(url, { headers: { Authorization: `Bearer ${token}` }, timeout: 10000 });
    const msgs = res2.data?.value || [];
    const sinceDate = new Date(since);
    res.json({
      mailbox,
      totalFound: msgs.length,
      since: since,
      messages: msgs.map(m => ({
        subject: m.subject,
        from: m.from?.emailAddress?.address,
        received: m.receivedDateTime,
        isRead: m.isRead,
        categories: m.categories || [],
        withinWindow: new Date(m.receivedDateTime) >= sinceDate,
        alreadyTagged: (m.categories || []).includes('Aurora Processed'),
      }))
    });
  } catch(e) {
    res.json({ error: e.response?.data?.error?.message || e.message });
  }
});

// ── Data backup / export ─────────────────────────────────────────────────────
app.get('/api/backup', async (req, res) => {
  try {
    const projects     = await db.getProjects();
    const drafts       = await db.getDrafts();
    const documents    = await db.getDocuments();
    const suggestions  = await db.getSuggestions();
    const spend        = await db.getSpend();

    const backup = {
      exportedAt: new Date().toISOString(),
      version: '3.0',
      projects,
      drafts,
      documents: documents.map(d => ({ ...d, extract: d.extract?.slice(0, 500) })), // trim extracts
      suggestions,
      spend,
    };

    res.setHeader('Content-Type', 'application/json');
    res.setHeader('Content-Disposition', `attachment; filename="aurora-backup-${new Date().toISOString().slice(0,10)}.json"`);
    res.json(backup);
    console.log(`[Backup] Exported ${projects.length} projects, ${drafts.length} drafts`);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── Data restore from backup ──────────────────────────────────────────────────
app.post('/api/restore', express.json({ limit: '50mb' }), async (req, res) => {
  try {
    const { projects, drafts, documents } = req.body;
    let restored = 0;
    if (projects?.length) {
      for (const p of projects) { await db.upsertProject(p); restored++; }
    }
    if (drafts?.length) {
      for (const d of drafts) { try { await db.saveDraft(d); } catch(e) {} }
    }
    if (documents?.length) {
      for (const d of documents) { try { await db.saveDocument(d); } catch(e) {} }
    }
    console.log(`[Restore] Restored ${restored} projects`);
    res.json({ success: true, projectsRestored: restored });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Cost
app.get('/api/cost', async (req, res) => {
  try {
    const spend = await db.getSpend();
    res.json({
      month: spend.month, totalUSD: spend.total.toFixed(4),
      totalAUD: (spend.total * 1.55).toFixed(2),
      calls: spend.calls, capUSD: CAP_USD,
      percentUsed: ((spend.total / CAP_USD) * 100).toFixed(1),
    });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// Manual batch + reminder triggers
app.post('/api/batch', async (req, res) => {
  try { await runBatch(); res.json({ success: true }); }
  catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/report/send', async (req, res) => {
  try {
    const projects = await db.getProjects();
    const standard = projects.filter(p => p.type === 'standard');
    await sendWeeklyExecutiveReport(standard);
    res.json({ success: true, message: 'Report sent to Dave, Kandia and Diane' });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// Manual calendar booking from deliverable tile
app.post('/api/projects/:id/deliverables/:delId/book', express.json(), async (req, res) => {
  try {
    const project = await db.getProject(req.params.id);
    if (!project) return res.status(404).json({ error: 'Project not found' });

    const { date, time, timeDisplay, durationMinutes, title, location, online } = req.body;
    if (!date || !time) return res.status(400).json({ error: 'Date and time are required' });

    const deliverables = await db.getDeliverables(req.params.id);
    const del = deliverables.find(d => d.id === req.params.delId);

    const eventTitle = title || (del ? del.name : 'Meeting') + ` — ${project.clientName}`;
    const displayTime = timeDisplay || time;
    const dateFormatted = new Date(date).toLocaleDateString('en-AU', { weekday:'long', day:'numeric', month:'long', year:'numeric' });

    const booking = {
      title: eventTitle,
      description: `Deliverable: ${del?.name || title}\nProject: ${project.projectName || project.clientName}\nClient: ${project.clientName}`,
      startDateTime: `${date}T${time}:00`,
      durationMinutes: parseInt(durationMinutes) || 60,
      location: location || project.clientName,
      online: online || false,
      clientName: project.clientName,
      clientEmail: project.clientEmail,
      consultantName: project.consultant,
      consultantEmail: project.consultantEmail || 'info@risk2solution.com',
      projectId: project.id,
    };

    // Create calendar event (tentative — not confirmed until invite is sent)
    // Continue even if calendar booking fails — still create the draft invite
    const eventId = await createCalendarBooking(booking, true);
    if (!eventId) {
      console.error(`[Booking] Calendar event creation failed for ${eventTitle} — continuing to create draft invite`);
    } else {
      console.log(`[Booking] Calendar event created: ${eventTitle} on ${date} at ${displayTime}`);
    }

    // Build attendee list for the invite email
    const attendees = [
      project.consultant ? `• ${project.consultant} (Consultant/Trainer)` : null,
      project.clientContact ? `• ${project.clientContact} — ${project.clientName} (Client)` : null,
      '• Diane Kruger — Risk 2 Solution (Project Manager)',
    ].filter(Boolean).join('\n');

    // Draft the meeting invite email — saved to Outlook Drafts, NOT sent
    const inviteBody = `Dear ${project.clientContact || 'Team'},

I hope this message finds you well. I am writing to confirm the following meeting scheduled in relation to the ${project.projectName || project.clientName} engagement.

Meeting Details:
Event: ${eventTitle}
Date: ${dateFormatted}
Time: ${displayTime} AEST
Duration: ${durationMinutes || 60} minutes
Location: ${location || (online ? 'Online — link to follow' : 'TBC')}
${online ? 'Format: Online meeting\n' : ''}
Attendees:
${attendees}

Please confirm your attendance by replying to this email. If you have any questions or need to reschedule, please do not hesitate to contact us.

Kind regards,

Diane Kruger
Corporate Operations Lead
Risk 2 Solution Group
P: 1300 459 970 | M: +61 415 748 747
E: diane.k@risk2solution.com
W: www.risk2solution.com`;

    const inviteDraft = {
      id: `d_${Date.now()}_invite`,
      projectId: project.id,
      clientName: project.clientName,
      projectName: project.projectName,
      type: 'meeting_invite',
      urgency: 'routine',
      toName: project.clientContact,
      toEmail: project.clientEmail || 'info@risk2solution.com',
      ccEmail: 'diane.k@risk2solution.com',
      subject: `Meeting Confirmation — ${eventTitle} — ${dateFormatted}`,
      body: inviteBody,
      source: 'booking',
      createdAt: new Date().toISOString(),
    };

    await db.saveDraft(inviteDraft);
    await saveDraftEmail(inviteDraft);

    // Update deliverable with booking details
    if (del) {
      await db.updateDeliverable(req.params.id, req.params.delId, {
        status: 'In Progress',
        calendarEvent: eventTitle,
        calendarDate: date,
      });
    }

    // Log activity
    await db.logActivity(project.id, {
      type: 'calendar_booking',
      summary: `Meeting ${eventId ? 'booked in calendar' : 'draft invite created (calendar unavailable)'}: ${eventTitle} on ${dateFormatted} at ${displayTime}`,
    });

    // Alert Diane
    await sendEmail('diane.k@risk2solution.com',
      `[Aurora] ${eventId ? 'Meeting booked' : 'Draft invite ready'}: ${eventTitle}`,
      `Aurora has ${eventId ? 'created a calendar event and ' : ''}saved a draft meeting invite to your Outlook Drafts folder.\n\nEvent: ${eventTitle}\nDate: ${dateFormatted}\nTime: ${displayTime} AEST\nDuration: ${durationMinutes || 60} minutes\nLocation: ${location || 'TBC'}\n${!eventId ? '\nNote: The calendar event could not be created automatically — please add it manually in Outlook.\n' : ''}
Open Outlook → Drafts → review the invite → hit Send when ready.\n\nAurora\nR2S Project Management Intelligence`,
      true
    );

    res.json({ success: true, eventId: eventId || null, calendarCreated: !!eventId, booking });
  } catch (e) {
    console.error('[Booking]', e.message);
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/emails/read', async (req, res) => {
  try {
    await readConsultantReplies();
    res.json({ success: true, message: 'Inbox checked and processed' });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// Delete a document from the library
app.delete('/api/documents/:id', async (req, res) => {
  try {
    await db.deleteDocument(req.params.id);
    res.json({ success: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/reminders/check', async (req, res) => {
  try { await checkDueDateReminders(); res.json({ success: true }); }
  catch (e) { res.status(500).json({ error: e.message }); }
});

// ── Start ─────────────────────────────────────────────────────────────────────
async function start() {
  await db.initDB();
  await db.ensureSpendConstraint();
  app.listen(PORT, async () => {
    console.log(`\n╔══════════════════════════════════════╗`);
    console.log(`║  Aurora R2S v3 — port ${PORT}          ║`);
    console.log(`╚══════════════════════════════════════╝`);
    console.log(`DB: ${process.env.AZURE_STORAGE_CONNECTION_STRING ? 'Azure Blob Storage' : 'JSON files'}`);
    console.log(`Spend cap: $${CAP_USD} USD/month`);
    console.log(`Internal emails: ${INTERNAL_EMAILS.join(', ')}`);
    console.log(`Reminders: 14, 7, 3 days before due date`);

    // Test Outlook connection on startup
    const outlookToken = await getOutlookToken();
    if (outlookToken) {
      console.log(`[Outlook] ✓ Connected — emails send from ${process.env.OUTLOOK_SHARED_MAILBOX || 'info@risk2solution.com'}\n`);
    } else {
      console.error(`[Outlook] ✗ NOT CONNECTED — emails will be logged only. Check Azure env vars: OUTLOOK_TENANT_ID, OUTLOOK_CLIENT_ID, OUTLOOK_CLIENT_SECRET\n`);
    }
  });
}

start().catch(console.error);

// ══════════════════════════════════════════════════════════════════════════════
// INTERNAL PROJECTS — IoP Grad Cert / Grad Dip Programme Management
// ══════════════════════════════════════════════════════════════════════════════

// Staff email directory for internal projects
const INTERNAL_STAFF = {
  'Dave Cohen':       process.env.DAVE_EMAIL        || 'dave.c@risk2solution.com',
  'Diane Kruger':     process.env.DIANE_EMAIL        || 'diane.k@risk2solution.com',
  'Cherry Abadeza':   process.env.CHERRY_EMAIL       || 'cherry.a@risk2solution.com',
  'Janita Zhang':     process.env.JANITA_EMAIL       || 'janita.z@risk2solution.com',
  'Dr Paul Johnston': process.env.PAUL_EMAIL         || 'paul.j@risk2solution.com',
  'Kandia':           process.env.KANDIA_EMAIL       || 'kandia@risk2solution.com',
  'Trainer':          process.env.TRAINER_EMAIL      || 'info@risk2solution.com',
};

// The master checklist — 22 steps across 5 phases
const IOP_CHECKLIST_TEMPLATE = [
  // PHASE 1 — Program Setup and Cohort Launch
  { phase: 1, phaseName: 'Phase 1 — Program Setup & Cohort Launch', step: 1, group: 'Set the delivery dates',
    task: 'Confirm the start date and module schedule for the next Grad Cert / Grad Dip intake with academic lead.',
    owner: 'Dave Cohen', trigger: 'program_start', daysOffset: -90 },
  { phase: 1, phaseName: 'Phase 1 — Program Setup & Cohort Launch', step: 1, group: 'Set the delivery dates',
    task: 'Capture final dates in the Master Communication & Cohort Tracker.',
    owner: 'Diane Kruger', trigger: 'program_start', daysOffset: -90 },
  { phase: 1, phaseName: 'Phase 1 — Program Setup & Cohort Launch', step: 2, group: 'Update public-facing information',
    task: 'Update the Institute of Presilience website with the new intake dates and any pricing changes.',
    owner: 'Janita Zhang', trigger: 'program_start', daysOffset: -85 },
  { phase: 1, phaseName: 'Phase 1 — Program Setup & Cohort Launch', step: 2, group: 'Update public-facing information',
    task: 'Ensure enquiry/enrolment forms, landing pages, and brochure PDFs reflect the correct dates.',
    owner: 'Janita Zhang', trigger: 'program_start', daysOffset: -85 },
  { phase: 1, phaseName: 'Phase 1 — Program Setup & Cohort Launch', step: 3, group: 'Lock in delivery team',
    task: 'Book trainers and guest speakers for each module.',
    owner: 'Cherry Abadeza', trigger: 'program_start', daysOffset: -80 },
  { phase: 1, phaseName: 'Phase 1 — Program Setup & Cohort Launch', step: 3, group: 'Lock in delivery team',
    task: 'Confirm availability, topic, time slot, delivery mode (live, pre-recorded, panel).',
    owner: 'Cherry Abadeza', trigger: 'program_start', daysOffset: -80 },
  { phase: 1, phaseName: 'Phase 1 — Program Setup & Cohort Launch', step: 3, group: 'Lock in delivery team',
    task: 'Record confirmed trainers and speakers in the tracker.',
    owner: 'Diane Kruger', trigger: 'program_start', daysOffset: -78 },
  { phase: 1, phaseName: 'Phase 1 — Program Setup & Cohort Launch', step: 4, group: 'Create the cohort Microsoft Teams Classroom',
    task: 'Create a new Teams Class Team for this intake (Grad Cert or Grad Dip) — name consistently e.g. Graduate Certificate (11056NAT) - [Month Year].',
    owner: 'Janita Zhang', trigger: 'program_start', daysOffset: -75 },
  { phase: 1, phaseName: 'Phase 1 — Program Setup & Cohort Launch', step: 4, group: 'Create the cohort Microsoft Teams Classroom',
    task: 'Set up branding, homepage, and channels by module/topic in Teams Classroom.',
    owner: 'Janita Zhang', trigger: 'program_start', daysOffset: -74 },
  { phase: 1, phaseName: 'Phase 1 — Program Setup & Cohort Launch', step: 4, group: 'Add core staff/admin as owners and create M365 Distribution Group.',
    task: 'Add core teaching staff and admin as owners/members. Create M365 Distribution Group.',
    owner: 'Janita Zhang', trigger: 'program_start', daysOffset: -71 },
  { phase: 2, phaseName: 'Phase 2 — Two Weeks Before Each Module', step: 4, group: 'Upload module resources to Teams Classroom',
    task: 'Upload module-specific resources to Teams: slides, EUO, required reading, and any pre-reading for this module.',
    owner: 'Diane Kruger', trigger: 'module', daysOffset: -14 },
  { phase: 2, phaseName: 'Phase 2 — Two Weeks Before Each Module', step: 4, group: 'Upload module resources to Teams Classroom',
    task: 'Create or confirm assignments for this module are visible and accurate in Teams Classroom.',
    owner: 'Cherry Abadeza', trigger: 'module', daysOffset: -14 },
  { phase: 1, phaseName: 'Phase 1 — Program Setup & Cohort Launch', step: 5, group: 'Build the marketing & awareness activity',
    task: 'Create social media campaign (LinkedIn, internal networks, partner channels) — align messaging with program theme and confirmed dates. Schedule posts.',
    owner: 'Janita Zhang', trigger: 'program_start', daysOffset: -70 },
  { phase: 1, phaseName: 'Phase 1 — Program Setup & Cohort Launch', step: 6, group: 'Student recruitment and confirmation',
    task: 'Confirm which students are attending this intake (verbals, EOI, or carried over from previous intake).',
    owner: 'Cherry Abadeza', trigger: 'program_start', daysOffset: -60 },
  { phase: 1, phaseName: 'Phase 1 — Program Setup & Cohort Launch', step: 6, group: 'Student recruitment and confirmation',
    task: 'Send enrolment email to each confirmed student: offer/welcome, course dates, fee information, next steps for enrolment/payment.',
    owner: 'Cherry Abadeza', trigger: 'program_start', daysOffset: -58 },
  { phase: 1, phaseName: 'Phase 1 — Program Setup & Cohort Launch', step: 6, group: 'Student recruitment and confirmation',
    task: 'Update the Communication & Cohort Tracker spreadsheet after each student outreach.',
    owner: 'Cherry Abadeza', trigger: 'program_start', daysOffset: -57 },
  { phase: 1, phaseName: 'Phase 1 — Program Setup & Cohort Launch', step: 7, group: 'Enrolment administration',
    task: 'When student returns enrolment form: Add student to Axcelerate (official registration).',
    owner: 'Cherry Abadeza', trigger: 'program_start', daysOffset: -55 },
  { phase: 1, phaseName: 'Phase 1 — Program Setup & Cohort Launch', step: 7, group: 'Enrolment administration',
    task: 'Confirm the fee amount for each student with Dave (standard, scholarship, partner, or corporate rate). Notify Diane of agreed pricing so she can raise the invoice.',
    owner: 'Cherry Abadeza', trigger: 'program_start', daysOffset: -55 },
  // PHASE 2 — Two weeks before program start
  { phase: 2, phaseName: 'Phase 2 — Two Weeks Before Each Module', step: 8, group: 'Set up student access to systems',
    task: 'Create an Institute of Presilience (IoP) email address for each enrolled student (first module only — confirm access is active for subsequent modules).',
    owner: 'Janita Zhang', trigger: 'module', daysOffset: -14 },
  { phase: 2, phaseName: 'Phase 2 — Two Weeks Before Each Module', step: 8, group: 'Set up student access to systems',
    task: 'Confirm each student can access the cohort Teams Classroom using their IoP email. Test access (sign in, view channels, see this module files).',
    owner: 'Janita Zhang', trigger: 'module', daysOffset: -13 },
  { phase: 2, phaseName: 'Phase 2 — Two Weeks Before Each Module', step: 9, group: 'Send module reminder email to students (T-2 weeks)',
    task: 'Send each student a 2-week module reminder email — confirm module dates/times, Teams access, and any preparation required for this module.',
    owner: 'Cherry Abadeza', trigger: 'module', daysOffset: -14 },
  { phase: 2, phaseName: 'Phase 2 — Two Weeks Before Each Module', step: 9, group: 'Send module reminder email to students (T-2 weeks)',
    task: 'Upload the EUO and any updated materials for this module into the Teams Classroom. Notify students they are available.',
    owner: 'Diane Kruger', trigger: 'module', daysOffset: -13 },
  { phase: 2, phaseName: 'Phase 2 — Two Weeks Before Each Module', step: 10, group: 'Readiness check',
    task: 'Ask each student to confirm they can access this module\'s materials in Teams and see the module assignments. Follow up with anyone who hasn\'t confirmed. Update tracker.',
    owner: 'Cherry Abadeza', trigger: 'module', daysOffset: -10 },
  // PHASE 3 — One week before program start
  { phase: 3, phaseName: 'Phase 3 — One Week Before Each Module', step: 11, group: 'Finalise learning materials',
    task: 'Confirm training slides and facilitator decks for this module are final. Upload slides and any pre-reading to the correct channel in Teams Classroom.',
    owner: 'Diane Kruger', trigger: 'module', daysOffset: -7 },
  { phase: 3, phaseName: 'Phase 3 — One Week Before Each Module', step: 12, group: 'One-week reminder email to students',
    task: 'Send group email: "Module delivery in one week" — include session dates and times, Teams link, expectations for attendance/camera/mic/engagement, and any preparation required.',
    owner: 'Cherry Abadeza', trigger: 'module', daysOffset: -7 },
  { phase: 3, phaseName: 'Phase 3 — One Week Before Each Module', step: 13, group: 'Guest speaker / trainer coordination',
    task: 'Confirm all guest speakers and trainers for this module. Send them their session date/time, Teams link, and cohort context. Confirm special requirements (breakout rooms, polls, recording, chat access). Note in tracker.',
    owner: 'Cherry Abadeza', trigger: 'module', daysOffset: -6 },
  // PHASE 4 — During delivery (per module)
  { phase: 4, phaseName: 'Phase 4 — During Delivery', step: 14, group: 'Module delivery rhythm',
    task: 'Confirm slides and materials are uploaded to Teams before each module session.',
    owner: 'Diane Kruger', trigger: 'module', daysOffset: -2 },
  { phase: 4, phaseName: 'Phase 4 — During Delivery', step: 14, group: 'Module delivery rhythm',
    task: 'Confirm all assignment due dates related to the module are visible and accurate in Teams.',
    owner: 'Cherry Abadeza', trigger: 'module', daysOffset: -2 },
  { phase: 4, phaseName: 'Phase 4 — During Delivery', step: 14, group: 'Module delivery rhythm',
    task: 'Confirm trainer(s) and guest speaker(s) are still available and briefed.',
    owner: 'Cherry Abadeza', trigger: 'module', daysOffset: -7 },
  { phase: 4, phaseName: 'Phase 4 — During Delivery', step: 14, group: 'Module delivery rhythm',
    task: 'Send 1-week-before reminder email to students: delivery dates, session link, assessment submission reminders.',
    owner: 'Cherry Abadeza', trigger: 'module', daysOffset: -7 },
  { phase: 4, phaseName: 'Phase 4 — During Delivery', step: 14, group: 'Module delivery rhythm',
    task: 'Send day-before reminder: reconfirm start time, trainer, and any materials to have ready.',
    owner: 'Cherry Abadeza', trigger: 'module', daysOffset: -1 },
  { phase: 4, phaseName: 'Phase 4 — During Delivery', step: 14, group: 'Module delivery rhythm',
    task: 'Run and record the live online session in the Teams Classroom.',
    owner: 'Trainer', trigger: 'module', daysOffset: 0 },
  { phase: 4, phaseName: 'Phase 4 — During Delivery', step: 14, group: 'Module delivery rhythm',
    task: 'Track student attendance and engagement in the tracker (Attended / Missed / Follow up). Follow up with any absent students via email.',
    owner: 'Cherry Abadeza', trigger: 'module', daysOffset: 1 },
  { phase: 4, phaseName: 'Phase 4 — During Delivery', step: 14, group: 'Module delivery rhythm',
    task: 'Send post-module follow-up email: where to find recordings, assessment submission deadlines, encourage discussion.',
    owner: 'Cherry Abadeza', trigger: 'module', daysOffset: 1 },
  { phase: 4, phaseName: 'Phase 4 — During Delivery', step: 15, group: 'Student support and admin during delivery',
    task: 'Monitor Teams Classroom chat for questions, assessment clarifications, and pastoral/support needs.',
    owner: 'Dr Paul Johnston', trigger: 'module', daysOffset: 0 },
  { phase: 4, phaseName: 'Phase 4 — During Delivery', step: 15, group: 'Student support and admin during delivery',
    task: 'Follow up on any missed invoices or partial payments during delivery.',
    owner: 'Diane Kruger', trigger: 'module', daysOffset: 1 },
  { phase: 4, phaseName: 'Phase 4 — During Delivery', step: 15, group: 'Student support and admin during delivery',
    task: 'Keep Axcelerate updated with attendance, progress, and assessment submissions.',
    owner: 'Cherry Abadeza', trigger: 'module', daysOffset: 1 },
  // PHASE 5 — Completion / Wrap-up
  { phase: 5, phaseName: 'Phase 5 — Program Completion', step: 16, group: 'Financial completion',
    task: 'Confirm all students have paid in full. Chase any outstanding invoices before issuing awards. Mark in tracker.',
    owner: 'Diane Kruger', trigger: 'program_end', daysOffset: 0 },
  { phase: 5, phaseName: 'Phase 5 — Program Completion', step: 17, group: 'Academic completion',
    task: 'Confirm all required assessments are submitted and marked according to competency/criteria. Confirm which students have met requirements for award.',
    owner: 'Dr Paul Johnston', trigger: 'program_end', daysOffset: 0 },
  { phase: 5, phaseName: 'Phase 5 — Program Completion', step: 18, group: 'Records and documentation',
    task: 'Export academic records from Teams Classroom. Prepare formal transcripts for each eligible student (official Institute of Presilience format).',
    owner: 'Janita Zhang', trigger: 'program_end', daysOffset: 7 },
  { phase: 5, phaseName: 'Phase 5 — Program Completion', step: 19, group: 'Certification and graduation',
    task: 'Issue Certificates (Grad Cert / Grad Dip) and official Transcripts to eligible students.',
    owner: 'Diane Kruger', trigger: 'program_end', daysOffset: 14 },
  { phase: 5, phaseName: 'Phase 5 — Program Completion', step: 19, group: 'Certification and graduation',
    task: 'Set the date for Graduation / Recognition ceremony (virtual or in-person). Schedule as Microsoft Teams Webinar. Send details to students. Prepare slide deck. Deliver and produce ceremony. Edit and upload recording to IoP YouTube.',
    owner: 'Janita Zhang', trigger: 'program_end', daysOffset: 21 },
  { phase: 5, phaseName: 'Phase 5 — Program Completion', step: 19, group: 'Certification and graduation',
    task: 'Send thank-you email with recording and alumni welcome message to graduates.',
    owner: 'Cherry Abadeza', trigger: 'program_end', daysOffset: 28 },
  { phase: 5, phaseName: 'Phase 5 — Program Completion', step: 20, group: 'Post-nominal approval',
    task: 'Prepare and submit Post Nominal application for each graduate so they are approved to use post-nominals. Notify graduates once cleared.',
    owner: 'Cherry Abadeza', trigger: 'program_end', daysOffset: 30 },
  { phase: 5, phaseName: 'Phase 5 — Program Completion', step: 21, group: 'Graduate engagement, close-out and handover',
    task: 'Add graduating students to: Alumni database/mailing list, relevant professional communities, IoP alumni association (with welcome email and next-touch timeline). Update final status in Communication & Cohort Tracker.',
    owner: 'Cherry Abadeza', trigger: 'program_end', daysOffset: 35 },
  { phase: 5, phaseName: 'Phase 5 — Program Completion', step: 21, group: 'Graduate engagement, close-out and handover',
    task: 'Archive the Teams Classroom: lock posting permissions or set to read-only for students. Ensure all learning materials, chat logs, and recordings are stored per retention requirements.',
    owner: 'Janita Zhang', trigger: 'program_end', daysOffset: 35 },
  { phase: 5, phaseName: 'Phase 5 — Program Completion', step: 21, group: 'Graduate engagement, close-out and handover',
    task: 'Capture lessons learned: trainer feedback, student feedback, operational issues to fix for next intake.',
    owner: 'Dave Cohen', trigger: 'program_end', daysOffset: 40 },
  { phase: 5, phaseName: 'Phase 5 — Program Completion', step: 22, group: 'Graduate Journal preparation & publication',
    task: 'Review assignments and select submissions for inclusion in the Graduate Journal.',
    owner: 'Dr Paul Johnston', trigger: 'program_end', daysOffset: 45 },
  { phase: 5, phaseName: 'Phase 5 — Program Completion', step: 22, group: 'Graduate Journal preparation & publication',
    task: 'Contact selected students individually to request written consent to publish their work. Collect headshots and LinkedIn profiles. Compile and PDF all confirmed submissions and send to Janita.',
    owner: 'Cherry Abadeza', trigger: 'program_end', daysOffset: 50 },
  { phase: 5, phaseName: 'Phase 5 — Program Completion', step: 22, group: 'Graduate Journal preparation & publication',
    task: 'Create and design Graduate Journal using pre-designed template. Conduct final proof review before publication.',
    owner: 'Janita Zhang', trigger: 'program_end', daysOffset: 55 },
  { phase: 5, phaseName: 'Phase 5 — Program Completion', step: 22, group: 'Graduate Journal preparation & publication',
    task: 'Register ISBN, generate barcode, update Graduate Journal file. Upload to Adobe, add download form to webpage, connect to HubSpot.',
    owner: 'Janita Zhang', trigger: 'program_end', daysOffset: 60 },
  { phase: 5, phaseName: 'Phase 5 — Program Completion', step: 22, group: 'Graduate Journal preparation & publication',
    task: 'Send launch email to all featured graduates. Schedule social media announcement. Track downloads and engagement via HubSpot.',
    owner: 'Cherry Abadeza', trigger: 'program_end', daysOffset: 65 },
];

// ── Internal project CRUD ─────────────────────────────────────────────────────
async function readInternalProjects() {
  try {
    if (db.read) return await db.read('internal_projects.json', []);
    const f = require('path').join(db.DATA, 'internal_projects.json');
    return require('fs').existsSync(f) ? JSON.parse(require('fs').readFileSync(f, 'utf8')) : [];
  } catch { return []; }
}

async function writeInternalProjects(projects) {
  if (db.write) return await db.write('internal_projects.json', projects);
  require('fs').writeFileSync(
    require('path').join(db.DATA, 'internal_projects.json'),
    JSON.stringify(projects, null, 2)
  );
}

async function readInternalChecklist(projectId) {
  try {
    if (db.read) return await db.read(`internal_checklist_${projectId}.json`, []);
    const f = require('path').join(db.DATA, `internal_checklist_${projectId}.json`);
    return require('fs').existsSync(f) ? JSON.parse(require('fs').readFileSync(f, 'utf8')) : [];
  } catch { return []; }
}

async function writeInternalChecklist(projectId, items) {
  if (db.write) return await db.write(`internal_checklist_${projectId}.json`, items);
  require('fs').writeFileSync(
    require('path').join(db.DATA, `internal_checklist_${projectId}.json`),
    JSON.stringify(items, null, 2)
  );
}

// Instantiate the checklist from the template for a new cohort
// moduleDates: array of 'YYYY-MM-DD' strings for each module delivery weekend
function buildChecklist(projectId, programStart, programEnd, moduleDates = []) {
  const start = programStart ? new Date(programStart) : null;
  const end   = programEnd   ? new Date(programEnd)   : null;
  const items = [];
  let idx = 0;

  for (const t of IOP_CHECKLIST_TEMPLATE) {
    if (t.trigger === 'module') {
      // Expand once per module date
      const dates = moduleDates.length > 0 ? moduleDates : (start ? [programStart] : []);
      dates.forEach((modDate, modIdx) => {
        const d = new Date(modDate);
        d.setDate(d.getDate() + t.daysOffset);
        items.push({
          id: `chk_${projectId}_${idx++}`,
          projectId,
          phase: t.phase,
          phaseName: t.phaseName,
          step: t.step,
          group: `${t.group} — Module ${modIdx + 1} (${new Date(modDate).toLocaleDateString('en-AU', { day: 'numeric', month: 'short', year: 'numeric' })})`,
          task: t.task,
          owner: t.owner,
          ownerEmail: INTERNAL_STAFF[t.owner] || 'info@risk2solution.com',
          trigger: t.trigger,
          moduleDate: modDate,
          moduleNumber: modIdx + 1,
          dueDate: d.toISOString().slice(0, 10),
          status: 'Not Started',
          reminderSent: false,
          completedAt: null,
          createdAt: new Date().toISOString(),
        });
      });
    } else {
      let dueDate = null;
      if (t.trigger === 'program_start' && start) {
        const d = new Date(start);
        d.setDate(d.getDate() + t.daysOffset);
        dueDate = d.toISOString().slice(0, 10);
      } else if (t.trigger === 'program_end' && end) {
        const d = new Date(end);
        d.setDate(d.getDate() + t.daysOffset);
        dueDate = d.toISOString().slice(0, 10);
      }
      items.push({
        id: `chk_${projectId}_${idx++}`,
        projectId,
        phase: t.phase,
        phaseName: t.phaseName,
        step: t.step,
        group: t.group,
        task: t.task,
        owner: t.owner,
        ownerEmail: INTERNAL_STAFF[t.owner] || 'info@risk2solution.com',
        trigger: t.trigger,
        moduleDate: null,
        moduleNumber: null,
        dueDate,
        status: 'Not Started',
        reminderSent: false,
        completedAt: null,
        createdAt: new Date().toISOString(),
      });
    }
  }
  return items;
}

// ── Send task reminder to staff member ───────────────────────────────────────
async function sendInternalTaskReminder(project, item) {
  const ownerEmail = item.ownerEmail || INTERNAL_STAFF[item.owner] || 'info@risk2solution.com';
  const programName = project.programType === 'grad_cert'
    ? 'Graduate Certificate in Organisational Resilience, Risk and High Reliability (11056NAT)'
    : 'Graduate Diploma of Organisational Presilience®, Risk and High Performance (11066NAT)';

  const subject = `[Aurora] Action required — ${project.cohortName}: ${item.group}`;
  const body = `Hi ${item.owner.split(' ')[0]},

Aurora is writing to remind you of an upcoming task for the ${programName} — ${project.cohortName} cohort.

TASK: ${item.task}
PHASE: ${item.phaseName}
DUE: ${item.dueDate || 'As soon as possible'}
ASSIGNED TO: ${item.owner}

Please action this task and reply to this email (to info@risk2solution.com) once it is complete. Aurora will automatically mark it as done when your reply is received.

Aurora
R2S Project Management Intelligence

---
Task ID: ${item.id} (include this in your reply to auto-confirm)`;

  // Only CC Diane if she is NOT the primary recipient (i.e. she's not the assigned owner)
  const cc = item.owner === 'Diane Kruger' ? [] : [];
  await sendEmail(ownerEmail, subject, body, false, cc);
  console.log(`[Internal] Reminder sent to ${item.owner} (${ownerEmail}): ${item.task.slice(0, 60)}`);
}

// ── Check internal project checklists daily ───────────────────────────────────
async function checkInternalProjectReminders() {
  const projects = await readInternalProjects();
  if (!projects.length) return;
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  for (const project of projects) {
    if (project.status === 'Completed' || project.status === 'On Hold' || project.status === 'Paused') continue;
    const checklist = await readInternalChecklist(project.id);
    if (!checklist.length) continue;

    for (const item of checklist) {
      if (item.status === 'Completed' || item.reminderSent) continue;
      if (!item.dueDate) continue;

      const due = new Date(item.dueDate);
      due.setHours(0, 0, 0, 0);
      const daysUntil = Math.round((due - today) / (1000 * 60 * 60 * 24));

      // Send reminder at 7 days and 2 days before due date, and on the day
      if ([7, 2, 0].includes(daysUntil) || (daysUntil < 0 && daysUntil >= -3)) {
        try {
          await sendInternalTaskReminder(project, item);
          item.reminderSent = true;
          item.lastReminderAt = new Date().toISOString();
        } catch(e) { console.error('[Internal] Reminder error:', e.message); }
      }
    }
    await writeInternalChecklist(project.id, checklist);
  }
}

// ── Detect staff confirmation reply emails ────────────────────────────────────
async function checkInternalTaskConfirmations(emailBody, emailSubject, fromEmail) {
  // Task IDs look like chk_ip_1234567890_0 — may be buried in quoted reply text
  const fullText = emailBody + ' ' + emailSubject;
  const taskIdMatch = fullText.match(/chk_ip_\d+_[a-z0-9_]*/i) ||
                      fullText.match(/chk_[a-z0-9_]{10,}/i);
  if (!taskIdMatch) return false;

  const taskId = taskIdMatch[0].trim();
  console.log(`[Internal] Possible task confirmation — ID: ${taskId} from ${fromEmail}`);

  const projects = await readInternalProjects();
  for (const project of projects) {
    const checklist = await readInternalChecklist(project.id);
    const item = checklist.find(i => i.id === taskId);
    if (item && item.status === 'Completed') {
      console.log(`[Internal] Task already completed: ${taskId}`);
      return true;
    }
    if (item && item.status !== 'Completed') {
      item.status = 'Completed';
      item.completedAt = new Date().toISOString();
      item.completedBy = fromEmail;
      await writeInternalChecklist(project.id, checklist);
      console.log(`[Internal] ✓ Task auto-confirmed by ${fromEmail}: ${item.task.slice(0, 60)}`);
      await sendEmail('diane.k@risk2solution.com',
        `[Aurora] Task confirmed: ${project.cohortName} — ${item.group}`,
        `${fromEmail} has confirmed completion of the following task:\n\nProgramme: ${project.cohortName}\nPhase: ${item.phaseName}\nTask: ${item.task}\nCompleted by: ${fromEmail}\n\nAurora has marked this item as complete in the checklist.\n\n${process.env.FRONTEND_URL || ''}\n\nAurora\nR2S Project Management Intelligence`,
        true
      );
      return true;
    }
  }
  console.log(`[Internal] Task ID found but no matching item: ${taskId}`);
  return false;
}

// ── Weekly internal project status report ─────────────────────────────────────
async function sendInternalProjectStatusReport() {
  const projects = await readInternalProjects();
  if (!projects.length) return;

  const RECIPIENTS = ['dave.c@risk2solution.com', 'kandia@risk2solution.com', 'diane.k@risk2solution.com'];
  const now = new Date();
  const now0 = new Date(now); now0.setHours(0,0,0,0);
  const dateStr = now.toLocaleDateString('en-AU', { weekday:'long', day:'numeric', month:'long', year:'numeric' });

  const ownerColors = {
    'Dave Cohen':'#6aa3ff','Diane Kruger':'#ff608a',
    'Cherry Abadeza':'#00e8bb','Janita Zhang':'#ffd93d',
    'Dr Paul Johnston':'#a78bfa','Trainer':'#888',
  };
  const fmt = d => new Date(d).toLocaleDateString('en-AU',{day:'numeric',month:'short'});

  let reportSections = '';

  for (const project of projects) {
    const checklist   = await readInternalChecklist(project.id);
    const programName = project.programType === 'grad_cert' ? 'Graduate Certificate (11056NAT)' : 'Graduate Diploma (11066NAT)';
    const moduleDates = (project.moduleDates || []).map(d => { const dt=new Date(d); dt.setHours(0,0,0,0); return dt; }).filter(d=>!isNaN(d)).sort((a,b)=>a-b);
    const nextModule  = moduleDates.find(d => d >= now0);
    const daysToNext  = nextModule ? Math.round((nextModule-now0)/(1000*60*60*24)) : null;

    const totalDone   = checklist.filter(i => i.status==='Completed').length;
    const totalTasks  = checklist.length;
    const pct         = totalTasks>0 ? Math.round(totalDone/totalTasks*100) : 0;
    const pctColor    = pct>=80?'#00e8bb':pct>=50?'#ffd93d':'#ff608a';

    const overdue     = checklist.filter(i => i.status!=='Completed' && i.dueDate && new Date(i.dueDate)<now0);
    const nextWeek    = new Date(now0.getTime()+7*24*60*60*1000);
    const threeWeeks  = new Date(now0.getTime()+21*24*60*60*1000);
    const dueThisWeek = checklist.filter(i => i.status!=='Completed' && i.dueDate && new Date(i.dueDate)>=now0 && new Date(i.dueDate)<=nextWeek);
    const dueSoon     = checklist.filter(i => i.status!=='Completed' && i.dueDate && new Date(i.dueDate)>nextWeek && new Date(i.dueDate)<=threeWeeks);
    const completedThisWeek = checklist.filter(i => i.completedAt && (now0-new Date(i.completedAt))<=7*24*60*60*1000);

    // Per-phase stats
    const phaseMap = {};
    checklist.forEach(i => {
      const k = i.phaseName||'General';
      if (!phaseMap[k]) phaseMap[k] = { total:0, done:0, overdue:0 };
      phaseMap[k].total++;
      if (i.status==='Completed') phaseMap[k].done++;
      if (i.status!=='Completed' && i.dueDate && new Date(i.dueDate)<now0) phaseMap[k].overdue++;
    });

    // Module timeline chips
    const moduleTimeline = moduleDates.map((d,idx) => {
      const isPast=d<now0, isNext=nextModule&&d.getTime()===nextModule.getTime();
      const days=Math.round((d-now0)/(1000*60*60*24));
      return {num:idx+1, date:d.toLocaleDateString('en-AU',{day:'numeric',month:'short',year:'numeric'}), isPast, isNext, days};
    });

    const statusBadge = daysToNext===null ? {text:'Programme active',color:'#00e8bb',bg:'#00e8bb15'}
      : daysToNext<=0   ? {text:'⚡ Module delivery TODAY',color:'#ff608a',bg:'#ff608a20'}
      : daysToNext<=7   ? {text:`🔴 Module in ${daysToNext} days — CRITICAL`,color:'#ff608a',bg:'#ff608a15'}
      : daysToNext<=14  ? {text:`🟡 Module in ${daysToNext} days — preparation underway`,color:'#ffd93d',bg:'#ffd93d15'}
      : daysToNext<=21  ? {text:`🟢 Module in ${daysToNext} days`,color:'#00e8bb',bg:'#00e8bb15'}
      : {text:`🔵 Module in ${daysToNext} days`,color:'#6aa3ff',bg:'#6aa3ff15'};

    reportSections += `
    <!-- Cohort card -->
    <div style="background:#1a1a2e;border:1px solid #2a2a4a;border-radius:12px;overflow:hidden;margin-bottom:20px">

      <!-- Card header -->
      <div style="background:#1a1a3e;padding:20px 24px">
        <div style="margin-bottom:12px">
          <div>
            <div style="font-size:16px;font-weight:700;color:#fff;margin-bottom:3px">${project.cohortName}</div>
            <div style="font-size:11px;color:#8899bb">${programName} · ${project.status}</div>
          </div>
          <div style="background:${statusBadge.bg};border:1px solid ${statusBadge.color}44;border-radius:8px;padding:6px 12px;text-align:center">
            <span style="color:${statusBadge.color};font-size:11px;font-weight:600">${statusBadge.text}</span>
          </div>
        </div>
        <!-- Overall progress -->
        <div style="margin-bottom:6px">
          <span style="font-size:11px;color:#8899bb">Overall progress — ${totalDone}/${totalTasks} tasks</span>
          <span style="font-size:16px;font-weight:700;color:${pctColor}">${pct}%</span>
        </div>
        <div style="background:#2a2a4a;border-radius:6px;height:12px">
          <div style="width:${pct}%;background:#00e8bb;height:12px;border-radius:6px"></div>
        </div>
      </div>

      <!-- Stats row -->
      <div style="border-bottom:1px solid #2a2a4a">
        ${[
          {label:'Completed',val:`${totalDone}`,sub:`of ${totalTasks}`,color:'#00a878'},
          {label:'Due this week',val:`${dueThisWeek.length}`,sub:'tasks',color:dueThisWeek.length>0?'#cc8800':'#00a878'},
          {label:'Overdue',val:`${overdue.length}`,sub:'tasks',color:overdue.length>0?'#cc2222':'#00a878'},
          {label:'Coming up',val:`${dueSoon.length}`,sub:'next 3 weeks',color:'#3366cc'},
        ].map((s,i) => `<div style="padding:12px;text-align:center;${i<3?'border-right:1px solid #eee':''}">
          <div style="font-size:20px;font-weight:700;color:${s.color}">${s.val}</div>
          <div style="font-size:10px;color:#666;margin-top:1px">${s.label}</div>
          <div style="font-size:9px;color:#aaa">${s.sub}</div>
        </div>`).join('')}
      </div>

      <div style="padding:16px 20px">

        <!-- Module timeline -->
        ${moduleTimeline.length>0 ? `<div style="margin-bottom:14px">
          <div style="font-size:10px;font-weight:700;color:#333;text-transform:uppercase;letter-spacing:1px;margin-bottom:8px">Module Timeline</div>
          <div style="">
            ${moduleTimeline.map(m => `<div style="min-width:90px;text-align:center;padding:8px 6px;border-radius:7px;border:2px solid ${m.isNext?'#6aa3ff':m.isPast?'#00e8bb':'#2a2a4a'};background:${m.isNext?'#1a1a3e':m.isPast?'#0d1f1a':'#141428'}">
              <div style="font-size:9px;color:#888;margin-bottom:2px">MOD ${m.num}</div>
              <div style="font-size:11px;font-weight:600;color:${m.isNext?'#6aa3ff':m.isPast?'#00e8bb':'#cccccc'}">${m.date}</div>
              <div style="font-size:9px;margin-top:2px;color:${m.isPast?'#00a878':m.isNext?'#3366cc':'#aaa'}">${m.isPast?'✓ Delivered':m.isNext?`${m.days}d away`:'Upcoming'}</div>
            </div>`).join('')}
          </div>
        </div>` : ''}

        <!-- Phase breakdown -->
        <div style="margin-bottom:14px">
          <div style="font-size:10px;font-weight:700;color:#333;text-transform:uppercase;letter-spacing:1px;margin-bottom:8px">Progress by Phase</div>
          ${Object.entries(phaseMap).map(([phase,stats]) => {
            const p2=stats.total>0?Math.round(stats.done/stats.total*100):0;
            const c2=p2===100?'#00e8bb':stats.overdue>0?'#ff608a':p2>=50?'#ffd93d':'#6aa3ff';
            return `<div style="margin-bottom:8px">
              <div style="margin-bottom:3px">
                <span style="font-size:11px;color:#cccccc">${phase.replace('Phase ','Ph ')}</span>
                <div style="">
                  ${stats.overdue>0?`<span style="font-size:9px;background:#ff608a15;color:#cc2222;border:1px solid #ff608a44;padding:1px 5px;border-radius:6px">⚠ ${stats.overdue} overdue</span>`:''}
                  <span style="font-size:11px;font-weight:600;color:${c2}">${stats.done}/${stats.total}</span>
                </div>
              </div>
              <div style="background:#2a2a4a;border-radius:3px;height:8px">
                <div style="width:${p2}%;background:${c2};height:8px;border-radius:3px"></div>
              </div>
            </div>`;
          }).join('')}
        </div>

        ${overdue.length>0 ? `<!-- Overdue -->
        <div style="background:#1a1a2e;border:1px solid #ff608a44;border-radius:8px;overflow:hidden;margin-bottom:12px;border-left:4px solid #ff3860">
          <div style="padding:8px 12px;background:#ff608a15;border-bottom:1px solid #ff608a33">
            <span style="font-size:10px;font-weight:700;color:#cc0033;text-transform:uppercase;letter-spacing:1px">⚠ Overdue — Immediate Attention Required (${overdue.length})</span>
          </div>
          ${overdue.map((i,idx) => `<div style="padding:9px 12px;border-bottom:1px solid #2a2a4a;background:${idx%2===0?'#1a1a2e':'#141428'}">
            <div style="flex:1;font-size:11px;color:#cccccc">${i.task.slice(0,90)}${i.task.length>90?'…':''}</div>
            <span style="font-size:10px;padding:2px 7px;border-radius:8px;background:${ownerColors[i.owner]||'#888'}22;color:${ownerColors[i.owner]||'#888'};font-weight:600;flex-shrink:0">${i.owner.split(' ')[0]}</span>
            <span style="font-size:10px;color:#cc0033;font-weight:600;flex-shrink:0;min-width:50px;text-align:right">${fmt(i.dueDate)}</span>
          </div>`).join('')}
        </div>` : `<div style="background:#00e8bb15;border:1px solid #00e8bb44;border-radius:8px;padding:10px 12px;margin-bottom:12px;text-align:center">
          <span style="color:#00e8bb;font-size:12px;font-weight:600">✓ No overdue tasks</span>
        </div>`}

        ${dueThisWeek.length>0 ? `<!-- Due this week -->
        <div style="background:#1a1a2e;border:1px solid #ffd93d44;border-radius:8px;overflow:hidden;margin-bottom:12px;border-left:4px solid #ffd93d">
          <div style="padding:8px 12px;background:#ffd93d15;border-bottom:1px solid #ffd93d33">
            <span style="font-size:10px;font-weight:700;color:#997700;text-transform:uppercase;letter-spacing:1px">📋 Due This Week (${dueThisWeek.length})</span>
          </div>
          ${dueThisWeek.map((i,idx) => `<div style="padding:9px 12px;border-bottom:1px solid #2a2a4a;background:${idx%2===0?'#1a1a2e':'#141428'}">
            <div style="flex:1;font-size:11px;color:#cccccc">${i.task.slice(0,90)}${i.task.length>90?'…':''}</div>
            <span style="font-size:10px;padding:2px 7px;border-radius:8px;background:${ownerColors[i.owner]||'#888'}22;color:${ownerColors[i.owner]||'#888'};font-weight:600;flex-shrink:0">${i.owner.split(' ')[0]}</span>
            <span style="font-size:10px;color:#997700;font-weight:600;flex-shrink:0;min-width:50px;text-align:right">${fmt(i.dueDate)}</span>
          </div>`).join('')}
        </div>` : ''}

        ${dueSoon.length>0 ? `<!-- Coming up -->
        <div style="background:#1a1a2e;border:1px solid #6aa3ff44;border-radius:8px;overflow:hidden;margin-bottom:12px;border-left:4px solid #6aa3ff">
          <div style="padding:8px 12px;background:#6aa3ff15;border-bottom:1px solid #6aa3ff33">
            <span style="font-size:10px;font-weight:700;color:#3366cc;text-transform:uppercase;letter-spacing:1px">📅 Coming Up — Next 3 Weeks (${dueSoon.length})</span>
          </div>
          ${dueSoon.slice(0,8).map((i,idx) => `<div style="padding:9px 12px;border-bottom:1px solid #2a2a4a;background:${idx%2===0?'#1a1a2e':'#141428'}">
            <div style="flex:1;font-size:11px;color:#cccccc">${i.task.slice(0,90)}${i.task.length>90?'…':''}</div>
            <span style="font-size:10px;padding:2px 7px;border-radius:8px;background:${ownerColors[i.owner]||'#888'}22;color:${ownerColors[i.owner]||'#888'};font-weight:600;flex-shrink:0">${i.owner.split(' ')[0]}</span>
            <span style="font-size:10px;color:#3366cc;font-weight:600;flex-shrink:0;min-width:50px;text-align:right">${fmt(i.dueDate)}</span>
          </div>`).join('')}
          ${dueSoon.length>8?`<div style="padding:8px 12px;text-align:center;font-size:10px;color:#888">+ ${dueSoon.length-8} more tasks</div>`:''}
        </div>` : ''}

        ${completedThisWeek.length>0 ? `<!-- Completed this week -->
        <div style="background:#1a1a2e;border:1px solid #00e8bb44;border-radius:8px;overflow:hidden">
          <div style="padding:8px 12px;background:#00e8bb15;border-bottom:1px solid #00e8bb33">
            <span style="font-size:10px;font-weight:700;color:#00a878;text-transform:uppercase;letter-spacing:1px">✓ Completed This Week (${completedThisWeek.length})</span>
          </div>
          ${completedThisWeek.map((i,idx) => `<div style="padding:8px 12px;border-bottom:1px solid #2a2a4a;background:${idx%2===0?'#1a1a2e':'#141428'}">
            <span style="color:#00a878;font-size:13px;flex-shrink:0">✓</span>
            <div style="flex:1;font-size:11px;color:#aaaaaa">${i.task.slice(0,90)}${i.task.length>90?'…':''}</div>
            <span style="font-size:10px;padding:2px 7px;border-radius:8px;background:${ownerColors[i.owner]||'#888'}22;color:${ownerColors[i.owner]||'#888'};font-weight:600;flex-shrink:0">${i.owner.split(' ')[0]}</span>
          </div>`).join('')}
        </div>` : ''}

      </div>
    </div>`;
  }

  const html = `<!DOCTYPE html><html xmlns:o="urn:schemas-microsoft-com:office:office"><head><meta charset="utf-8"/><!--[if mso]><xml><o:OfficeDocumentSettings><o:AllowPNG/></o:OfficeDocumentSettings></xml><![endif]--></head>
<body style="margin:0;padding:0;background:#0f0f1a;font-family:Arial,sans-serif">
<table width="720" cellpadding="20" cellspacing="0" border="0" align="center" style="max-width:720px;"><tr><td>

  <!-- Header -->
  <div style="background:#1a1a3e;border-radius:12px;padding:28px;margin-bottom:20px;border:1px solid #2a2a4a">
    <div style="font-size:10px;color:#6aa3ff;text-transform:uppercase;letter-spacing:2px;margin-bottom:8px">Institute of Presilience · Leadership Report</div>
    <div style="font-size:24px;font-weight:700;color:#fff;margin-bottom:4px">Internal Programmes — Weekly Status</div>
    <div style="font-size:12px;color:#8899bb">${dateStr} · Generated by Aurora</div>
  </div>

  ${reportSections}

  <div style="text-align:center;color:#555577;font-size:10px;padding:12px 0">
    Aurora · R2S Project Management Intelligence · Confidential
    ${process.env.FRONTEND_URL ? '<br><a href="' + process.env.FRONTEND_URL + '" style="color:#6aa3ff">Open Aurora portal</a>' : ''}
  </div>

</div></body></html>`;

  const subject = `IoP Internal Programmes — Weekly Leadership Report (${now.toLocaleDateString('en-AU',{day:'numeric',month:'short'})})`;
  for (const r of RECIPIENTS) {
    await sendEmail(r, subject, html, false, [], true);
  }
  console.log('[Internal] Weekly leadership report sent');
}

// ── Determine if a cohort should send its weekly ops update this week ─────────
// Rules:
// - Always send from cohort creation until 3 weeks before first module
// - Resume 3 weeks before each module, stop after that module passes
// - Stop after final module is delivered (unless wrap-up tasks remain)
function shouldSendWeeklyOpsUpdate(project) {
  // Send every Monday for any Active or Scheduled cohort
  // Only stop when the cohort is Completed, On Hold, or Paused
  // The per-module task sections naturally show what's relevant each week
  return !['Completed', 'On Hold', 'Paused'].includes(project.status);
}

// ── Weekly operational checklist update to delivery team ─────────────────────
async function buildChecklistExcel(project, checklist) {
  // Build Excel checklist using ExcelJS (or fall back to CSV if not available)
  try {
    const ExcelJS = require('exceljs');
    const wb = new ExcelJS.Workbook();
    wb.creator = 'Aurora — R2S Project Management Intelligence';

    const programName = project.programType === 'grad_cert'
      ? 'Graduate Certificate (11056NAT)'
      : 'Graduate Diploma (11066NAT)';

    // Group by phase
    const phases = {};
    checklist.forEach(i => {
      const key = i.phaseName || 'General';
      if (!phases[key]) phases[key] = [];
      phases[key].push(i);
    });

    const ownerColors = {
      'Dave Cohen':       'FF6AA3FF',
      'Diane Kruger':     'FFFF608A',
      'Cherry Abadeza':   'FF00E8BB',
      'Janita Zhang':     'FFFFD93D',
      'Dr Paul Johnston': 'FFA78BFA',
      'Trainer':          'FF888888',
    };

    // Summary sheet
    const summary = wb.addWorksheet('Summary');
    summary.columns = [
      { header: '', key: 'label', width: 30 },
      { header: '', key: 'value', width: 40 },
    ];
    const titleRow = summary.addRow([`${project.cohortName} — ${programName}`, '']);
    titleRow.font = { bold: true, size: 14, color: { argb: 'FF1A1A3E' } };
    summary.addRow(['Generated', new Date().toLocaleDateString('en-AU', { weekday:'long', day:'numeric', month:'long', year:'numeric' })]);
    summary.addRow(['Status', project.status]);
    summary.addRow(['Programme start', project.programStart || 'TBC']);
    summary.addRow(['Programme end', project.programEnd || 'TBC']);
    summary.addRow(['Module dates', (project.moduleDates || []).join(', ') || 'Not set']);
    summary.addRow([]);
    const total = checklist.length;
    const done  = checklist.filter(i => i.status === 'Completed').length;
    const overdue = checklist.filter(i => i.status !== 'Completed' && i.dueDate && new Date(i.dueDate) < new Date()).length;
    summary.addRow(['Total tasks', total]);
    summary.addRow(['Completed', done]);
    summary.addRow(['Outstanding', total - done]);
    summary.addRow(['Overdue', overdue]);
    summary.addRow(['% Complete', `${Math.round(done/total*100)}%`]);

    // One sheet per phase
    for (const [phaseName, items] of Object.entries(phases)) {
      const sheetName = phaseName.replace(/[\/\\*\[\]:?]/g, '').slice(0, 31);
      const ws = wb.addWorksheet(sheetName);
      ws.columns = [
        { header: 'Status',    key: 'status',    width: 14 },
        { header: 'Group',     key: 'group',     width: 35 },
        { header: 'Task',      key: 'task',      width: 60 },
        { header: 'Owner',     key: 'owner',     width: 20 },
        { header: 'Due Date',  key: 'dueDate',   width: 14 },
        { header: 'Completed', key: 'completedAt',width: 20 },
        { header: 'Completed By', key: 'completedBy', width: 25 },
      ];

      // Header row styling
      ws.getRow(1).eachCell(cell => {
        cell.fill   = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A1A3E' } };
        cell.font   = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
        cell.border = { bottom: { style: 'medium', color: { argb: 'FF6AA3FF' } } };
        cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      });
      ws.getRow(1).height = 24;

      items.forEach((item, idx) => {
        const row = ws.addRow({
          status:      item.status === 'Completed' ? '✓ Complete' : item.dueDate && new Date(item.dueDate) < new Date() ? '⚠ Overdue' : 'Outstanding',
          group:       item.group.replace(/ — Module \d+ \(.*?\)/,'').trim(),
          task:        item.task,
          owner:       item.owner,
          dueDate:     item.dueDate || '',
          completedAt: item.completedAt ? new Date(item.completedAt).toLocaleDateString('en-AU') : '',
          completedBy: item.completedBy || '',
        });
        row.height = 30;

        const isCompleted = item.status === 'Completed';
        const isOverdue   = !isCompleted && item.dueDate && new Date(item.dueDate) < new Date();
        const bg = isCompleted ? 'FFE6FFF8' : isOverdue ? 'FFFFF0F0' : idx % 2 === 0 ? 'FFFFFFFF' : 'FFF8F9FF';

        row.eachCell((cell, colNum) => {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
          cell.alignment = { vertical: 'middle', wrapText: true };
          cell.border = { bottom: { style: 'thin', color: { argb: 'FFE0E0E0' } } };
          if (colNum === 1) {
            cell.font = { bold: true, color: { argb: isCompleted ? 'FF00A878' : isOverdue ? 'FFCC0000' : 'FF666666' } };
          }
          if (colNum === 4) {
            const oColor = ownerColors[item.owner] || 'FF888888';
            cell.font = { color: { argb: oColor }, bold: true };
          }
        });
      });

      ws.views = [{ state: 'frozen', ySplit: 1 }];
    }

    // Write to buffer
    const buffer = await wb.xlsx.writeBuffer();
    return buffer;
  } catch(xlsxErr) {
    console.error('[Excel] ExcelJS not available:', xlsxErr.message);
    // Fallback to CSV
    const lines = ['Status,Group,Task,Owner,Due Date,Completed At,Completed By'];
    checklist.forEach(i => {
      const status = i.status === 'Completed' ? 'Complete' : i.dueDate && new Date(i.dueDate) < new Date() ? 'OVERDUE' : 'Outstanding';
      const esc = s => `"${(s||'').replace(/"/g,'""')}"`;
      lines.push([status, esc(i.group), esc(i.task), esc(i.owner), i.dueDate||'', i.completedAt||'', i.completedBy||''].join(','));
    });
    return Buffer.from(lines.join('\n'), 'utf8');
  }
}

async function sendInternalWeeklyOpsUpdate(forceAll = false) {
  const projects = await readInternalProjects();
  const activeProjects = projects.filter(p =>
    p.status !== 'Completed' && p.status !== 'On Hold' && p.status !== 'Paused'
  );
  if (!activeProjects.length) return;

  const now = new Date();
  const now0 = new Date(now); now0.setHours(0,0,0,0);
  const dateStr = now.toLocaleDateString('en-AU', { weekday:'long', day:'numeric', month:'long', year:'numeric' });

  const TO = global._testMode ? [global._testEmail] : [
    process.env.CHERRY_EMAIL || 'cherry.a@risk2solution.com',
    process.env.JANITA_EMAIL || 'janita.z@risk2solution.com',
    'diane.k@risk2solution.com',
  ];
  const CC = global._testMode ? [] : ['dave.c@risk2solution.com', 'kandia@risk2solution.com'];

  const ownerColors = {
    'Dave Cohen': '#6aa3ff', 'Diane Kruger': '#ff608a',
    'Cherry Abadeza': '#00e8bb', 'Janita Zhang': '#ffd93d',
    'Dr Paul Johnston': '#a78bfa', 'Trainer': '#888',
  };

  const fmt = d => { try { return new Date(d).toLocaleDateString('en-AU',{day:'numeric',month:'short'}); } catch { return '—'; } }

  let anySent = false;

  for (const project of activeProjects) {
    try {
    if (!forceAll && !shouldSendWeeklyOpsUpdate(project)) {
      console.log(`[Internal Ops] Skipping ${project.cohortName} — outside active window`);
      continue;
    }

    const checklist = await readInternalChecklist(project.id);
    const moduleDates = (project.moduleDates || []).map(d => { const dt = new Date(d); dt.setHours(0,0,0,0); return dt; }).filter(d=>!isNaN(d)).sort((a,b)=>a-b);
    const nextModule  = moduleDates.find(d => d >= now0);
    const pastModules = moduleDates.filter(d => d < now0);
    const daysToNext  = nextModule ? Math.round((nextModule - now0)/(1000*60*60*24)) : null;
    const programName = project.programType === 'grad_cert' ? 'Graduate Certificate (11056NAT)' : 'Graduate Diploma (11066NAT)';

    // Segment tasks
    const overdue      = checklist.filter(i => i.status !== 'Completed' && i.dueDate && new Date(i.dueDate) < now0);
    const nextWeek     = new Date(now0.getTime() + 7*24*60*60*1000);
    const threeWeeks   = new Date(now0.getTime() + 21*24*60*60*1000);
    const dueThisWeek  = checklist.filter(i => i.status !== 'Completed' && i.dueDate && new Date(i.dueDate) >= now0 && new Date(i.dueDate) <= nextWeek);
    const dueSoon      = checklist.filter(i => i.status !== 'Completed' && i.dueDate && new Date(i.dueDate) > nextWeek && new Date(i.dueDate) <= threeWeeks);
    const completedThisWeek = checklist.filter(i => i.completedAt && (now0 - new Date(i.completedAt)) <= 7*24*60*60*1000);
    const totalDone    = checklist.filter(i => i.status === 'Completed').length;
    const totalTasks   = checklist.length;
    const pct          = totalTasks > 0 ? Math.round(totalDone/totalTasks*100) : 0;

    // Per-phase stats
    const phaseMap = {};
    checklist.forEach(i => {
      const k = i.phaseName || 'General';
      if (!phaseMap[k]) phaseMap[k] = { total:0, done:0, overdue:0 };
      phaseMap[k].total++;
      if (i.status === 'Completed') phaseMap[k].done++;
      if (i.status !== 'Completed' && i.dueDate && new Date(i.dueDate) < now0) phaseMap[k].overdue++;
    });

    // Module timeline
    const moduleTimeline = moduleDates.map((d, idx) => {
      const isPast = d < now0;
      const isNext = nextModule && d.getTime() === nextModule.getTime();
      const days = Math.round((d - now0)/(1000*60*60*24));
      return { num: idx+1, date: d.toLocaleDateString('en-AU',{day:'numeric',month:'short',year:'numeric'}), isPast, isNext, days };
    });

    // Status badge
    const statusBadge = daysToNext === null ? { text:'Programme active', color:'#00e8bb', bg:'#00e8bb15' }
      : daysToNext <= 0   ? { text:'⚡ Module delivery TODAY', color:'#ff608a', bg:'#ff608a20' }
      : daysToNext <= 7   ? { text:`🔴 Next module in ${daysToNext} days — CRITICAL WEEK`, color:'#ff608a', bg:'#ff608a15' }
      : daysToNext <= 14  ? { text:`🟡 Next module in ${daysToNext} days — preparation underway`, color:'#ffd93d', bg:'#ffd93d15' }
      : daysToNext <= 21  ? { text:`🟢 Next module in ${daysToNext} days`, color:'#00e8bb', bg:'#00e8bb15' }
      : { text:`🔵 Next module in ${daysToNext} days — setup phase`, color:'#6aa3ff', bg:'#6aa3ff15' };

    const fmt = d => new Date(d).toLocaleDateString('en-AU',{day:'numeric',month:'short'});

    // Dark theme colour variables — must be before the html template
    const BG    = '#0f0f1a';
    const CARD  = '#1a1a2e';
    const CARD2 = '#141428';
    const BORDER= '#2a2a4a';
    const TEXT  = '#e0e0e0';
    const TEXT2 = '#aaaaaa';
    const TEXT3 = '#666680';
    const pctColor = pct>=80?'#00e8bb':pct>=50?'#ffd93d':'#ff608a';

    // Build Excel and get token before html (both referenced in template footer)
    let excelBuffer = null;
    let excelFilename = `${project.cohortName.replace(/[^a-zA-Z0-9 ]/g,'').trim()} — Checklist ${now.toISOString().slice(0,10)}.xlsx`;
    try {
      excelBuffer = await buildChecklistExcel(project, checklist);
      console.log(`[Internal Ops] Excel built: ${excelFilename}`);
    } catch(e) {
      console.error('[Excel] Build failed (will send without attachment):', e.message);
    }
    const token = await getOutlookToken();
    const fromMailbox = process.env.OUTLOOK_SHARED_MAILBOX || 'info@risk2solution.com';

    console.log(`[Internal Ops] Sending to: ${JSON.stringify(TO)} CC: ${JSON.stringify(CC)}`);

    const html = `<!DOCTYPE html>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:v="urn:schemas-microsoft-com:vml">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width"/>
<!--[if mso]><xml><o:OfficeDocumentSettings><o:AllowPNG/><o:PixelsPerInch>96</o:PixelsPerInch></o:OfficeDocumentSettings></xml><![endif]-->
<style>
  body { margin:0; padding:0; background:#0f0f1a; -ms-text-size-adjust:100%; -webkit-text-size-adjust:100%; }
  table { border-collapse:collapse; mso-table-lspace:0pt; mso-table-rspace:0pt; }
  td { border-collapse:collapse; }
  img { border:0; outline:none; text-decoration:none; }
</style>
</head>
<body style="margin:0;padding:0;background:#0f0f1a;">
<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#0f0f1a;">
<tr><td align="center" style="padding:16px 8px;">
<table width="680" cellpadding="0" cellspacing="0" border="0" style="max-width:680px;width:100%;">

  <!-- HEADER -->
  <tr><td style="padding-bottom:12px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#1a1a3e;border-radius:10px;border:1px solid #2a2a4a;">
    <tr><td style="padding:24px 28px;">
      <div style="font-size:10px;color:#6aa3ff;text-transform:uppercase;letter-spacing:2px;margin-bottom:8px;font-family:Arial,sans-serif;">${global._testMode ? '&#x1F9EA; TEST &middot; ' : ''}Institute of Presilience &middot; Weekly Programme Dashboard</div>
      <div style="font-size:22px;font-weight:700;color:#ffffff;margin-bottom:4px;font-family:Arial,sans-serif;">${project.cohortName}</div>
      <div style="font-size:12px;color:#8899bb;margin-bottom:16px;font-family:Arial,sans-serif;">${programName} &middot; ${dateStr}</div>
      <table cellpadding="0" cellspacing="0" border="0"><tr><td style="background:${statusBadge.bg};border:1px solid ${statusBadge.color};border-radius:6px;padding:8px 14px;">
        <span style="color:${statusBadge.color};font-size:13px;font-weight:700;font-family:Arial,sans-serif;">${statusBadge.text}</span>
      </td></tr></table>
    </td></tr>
    </table>
  </td></tr>

  <!-- STAT TILES -->
  <tr><td style="padding-bottom:12px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0">
    <tr>
      ${[
        { label:'Total tasks',   val:String(totalTasks),        color:'#6aa3ff' },
        { label:'Completed',     val:`${totalDone} (${pct}%)`,  color:'#00e8bb' },
        { label:'Due this week', val:String(dueThisWeek.length), color:dueThisWeek.length>0?'#ffd93d':'#00e8bb' },
        { label:'Overdue',       val:String(overdue.length),    color:overdue.length>0?'#ff608a':'#00e8bb' },
      ].map((s,i) => `<td width="25%" style="padding:0 ${i<3?'6px':'0'} 0 0;vertical-align:top;">
        <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#1a1a2e;border:1px solid #2a2a4a;border-radius:8px;">
        <tr><td align="center" style="padding:14px 8px;">
          <div style="font-size:20px;font-weight:700;color:${s.color};font-family:Arial,sans-serif;">${s.val}</div>
          <div style="font-size:10px;color:#888;text-transform:uppercase;letter-spacing:0.5px;margin-top:4px;font-family:Arial,sans-serif;">${s.label}</div>
        </td></tr>
        </table>
      </td>`).join('')}
    </tr>
    </table>
  </td></tr>

  <!-- OVERALL PROGRESS -->
  <tr><td style="padding-bottom:12px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#1a1a2e;border:1px solid #2a2a4a;border-radius:8px;">
    <tr><td style="padding:16px 18px;">
      <table width="100%" cellpadding="0" cellspacing="0" border="0">
      <tr>
        <td style="font-size:12px;font-weight:700;color:#e0e0e0;font-family:Arial,sans-serif;">Overall Programme Progress</td>
        <td align="right" style="font-size:16px;font-weight:700;color:${pctColor};font-family:Arial,sans-serif;">${pct}%</td>
      </tr>
      </table>
      <div style="margin-top:8px;background:#2a2a4a;border-radius:4px;height:12px;">
        <!--[if mso]><v:rect xmlns:v="urn:schemas-microsoft-com:vml" style="width:${Math.round(pct*5.6)}pt;height:12pt;" fillcolor="${pctColor}" stroke="f"><v:fill type="solid"/></v:rect><![endif]-->
        <!--[if !mso]><!-->
        <div style="width:${pct}%;background:${pctColor};height:12px;border-radius:4px;min-width:${pct>0?'4px':'0'};"></div>
        <!--<![endif]-->
      </div>
    </td></tr>
    </table>
  </td></tr>

  <!-- MODULE TIMELINE -->
  ${moduleTimeline.length > 0 ? `<tr><td style="padding-bottom:12px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#1a1a2e;border:1px solid #2a2a4a;border-radius:8px;">
    <tr><td style="padding:16px 18px;">
      <div style="font-size:10px;font-weight:700;color:#e0e0e0;text-transform:uppercase;letter-spacing:1px;margin-bottom:12px;font-family:Arial,sans-serif;">Module Delivery Timeline</div>
      <table width="100%" cellpadding="0" cellspacing="0" border="0"><tr>
        ${moduleTimeline.map((m,mi) => `<td width="${Math.floor(100/moduleTimeline.length)}%" style="padding:0 ${mi<moduleTimeline.length-1?'6px':'0'} 0 0;vertical-align:top;">
          <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border:2px solid ${m.isNext?'#6aa3ff':m.isPast?'#00e8bb':'#2a2a4a'};border-radius:6px;background:${m.isNext?'#1a1a3e':m.isPast?'#0d1f1a':'#141428'};">
          <tr><td align="center" style="padding:10px 6px;">
            <div style="font-size:9px;color:#666688;text-transform:uppercase;margin-bottom:3px;font-family:Arial,sans-serif;">MODULE ${m.num}</div>
            <div style="font-size:11px;font-weight:700;color:${m.isNext?'#6aa3ff':m.isPast?'#00e8bb':'#cccccc'};font-family:Arial,sans-serif;">${m.date}</div>
            <div style="font-size:9px;color:${m.isPast?'#00e8bb':m.isNext?'#6aa3ff':'#666688'};margin-top:3px;font-family:Arial,sans-serif;">${m.isPast?'&#10003; Delivered':m.isNext?`${m.days}d away`:'Upcoming'}</div>
          </td></tr>
          </table>
        </td>`).join('')}
      </tr></table>
    </td></tr>
    </table>
  </td></tr>` : ''}

  <!-- PHASE PROGRESS -->
  <tr><td style="padding-bottom:12px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#1a1a2e;border:1px solid #2a2a4a;border-radius:8px;">
    <tr><td style="padding:16px 18px;">
      <div style="font-size:10px;font-weight:700;color:#e0e0e0;text-transform:uppercase;letter-spacing:1px;margin-bottom:14px;font-family:Arial,sans-serif;">Progress by Phase</div>
      ${Object.entries(phaseMap).map(([phase, stats]) => {
        const p2 = stats.total>0 ? Math.round(stats.done/stats.total*100) : 0;
        const c2 = p2===100?'#00e8bb':stats.overdue>0?'#ff608a':p2>=50?'#ffd93d':'#6aa3ff';
        return `<table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom:10px;">
        <tr>
          <td style="font-size:11px;color:#aaaaaa;font-family:Arial,sans-serif;">${phase.replace('Phase ','Ph ')}</td>
          <td align="right">
            ${stats.overdue>0?`<span style="font-size:9px;color:#ff608a;background:#ff608a22;padding:1px 6px;border-radius:4px;font-family:Arial,sans-serif;">&#9888; ${stats.overdue} overdue</span> `:''}
            <span style="font-size:11px;font-weight:700;color:${c2};font-family:Arial,sans-serif;">${stats.done}/${stats.total}</span>
          </td>
        </tr>
        <tr><td colspan="2" style="padding-top:4px;">
          <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#2a2a4a;border-radius:3px;height:8px;">
          <tr><td style="height:8px;">
            <!--[if !mso]><!-->
            <div style="width:${p2}%;background:${c2};height:8px;border-radius:3px;min-width:${p2>0?'4px':'0'};"></div>
            <!--<![endif]-->
          </td></tr>
          </table>
        </td></tr>
        </table>`;
      }).join('')}
    </td></tr>
    </table>
  </td></tr>

  <!-- OVERDUE -->
  ${overdue.length > 0 ? `<tr><td style="padding-bottom:12px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#1a1a2e;border:1px solid #ff386044;border-left:4px solid #ff3860;border-radius:8px;">
    <tr><td style="background:#ff386015;padding:10px 16px;border-radius:8px 8px 0 0;">
      <span style="font-size:10px;font-weight:700;color:#ff608a;text-transform:uppercase;letter-spacing:1px;font-family:Arial,sans-serif;">&#9888; Overdue &mdash; Requires Immediate Attention (${overdue.length})</span>
    </td></tr>
    ${overdue.map((item,ri) => `<tr><td style="padding:10px 16px;border-top:1px solid #2a2a4a;background:${ri%2===0?'#1a1a2e':'#141428'};">
      <table width="100%" cellpadding="0" cellspacing="0" border="0"><tr>
        <td style="font-size:12px;color:#e0e0e0;font-family:Arial,sans-serif;padding-right:8px;">${item.task.slice(0,90)}${item.task.length>90?'&hellip;':''}</td>
        <td nowrap style="padding-right:8px;"><span style="font-size:10px;color:${ownerColors[item.owner]||'#aaa'};background:${ownerColors[item.owner]||'#888'}22;padding:2px 8px;border-radius:10px;white-space:nowrap;font-family:Arial,sans-serif;font-weight:700;">${item.owner.split(' ')[0]}</span></td>
        <td nowrap style="font-size:11px;color:#ff608a;font-weight:700;font-family:Arial,sans-serif;white-space:nowrap;">${fmt(item.dueDate)}</td>
      </tr></table>
    </td></tr>`).join('')}
    </table>
  </td></tr>` : `<tr><td style="padding-bottom:12px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#00e8bb15;border:1px solid #00e8bb44;border-radius:8px;">
    <tr><td align="center" style="padding:14px;font-size:13px;font-weight:700;color:#00e8bb;font-family:Arial,sans-serif;">&#10003; No overdue tasks &mdash; great work!</td></tr>
    </table>
  </td></tr>`}

  <!-- DUE THIS WEEK -->
  ${dueThisWeek.length > 0 ? `<tr><td style="padding-bottom:12px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#1a1a2e;border:1px solid #ffd93d44;border-left:4px solid #ffd93d;border-radius:8px;">
    <tr><td style="background:#ffd93d15;padding:10px 16px;border-radius:8px 8px 0 0;">
      <span style="font-size:10px;font-weight:700;color:#ffd93d;text-transform:uppercase;letter-spacing:1px;font-family:Arial,sans-serif;">&#128203; Due This Week (${dueThisWeek.length})</span>
    </td></tr>
    ${dueThisWeek.map((item,ri) => `<tr><td style="padding:10px 16px;border-top:1px solid #2a2a4a;background:${ri%2===0?'#1a1a2e':'#141428'};">
      <table width="100%" cellpadding="0" cellspacing="0" border="0"><tr>
        <td style="font-size:12px;color:#e0e0e0;font-family:Arial,sans-serif;padding-right:8px;">${item.task.slice(0,90)}${item.task.length>90?'&hellip;':''}</td>
        <td nowrap style="padding-right:8px;"><span style="font-size:10px;color:${ownerColors[item.owner]||'#aaa'};background:${ownerColors[item.owner]||'#888'}22;padding:2px 8px;border-radius:10px;white-space:nowrap;font-family:Arial,sans-serif;font-weight:700;">${item.owner.split(' ')[0]}</span></td>
        <td nowrap style="font-size:11px;color:#ffd93d;font-weight:700;font-family:Arial,sans-serif;white-space:nowrap;">${fmt(item.dueDate)}</td>
      </tr></table>
    </td></tr>`).join('')}
    </table>
  </td></tr>` : ''}

  <!-- COMING UP -->
  ${dueSoon.length > 0 ? `<tr><td style="padding-bottom:12px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#1a1a2e;border:1px solid #6aa3ff44;border-left:4px solid #6aa3ff;border-radius:8px;">
    <tr><td style="background:#6aa3ff15;padding:10px 16px;border-radius:8px 8px 0 0;">
      <span style="font-size:10px;font-weight:700;color:#6aa3ff;text-transform:uppercase;letter-spacing:1px;font-family:Arial,sans-serif;">&#128197; Coming Up &mdash; Next 3 Weeks (${dueSoon.length})</span>
    </td></tr>
    ${dueSoon.slice(0,8).map((item,ri) => `<tr><td style="padding:10px 16px;border-top:1px solid #2a2a4a;background:${ri%2===0?'#1a1a2e':'#141428'};">
      <table width="100%" cellpadding="0" cellspacing="0" border="0"><tr>
        <td style="font-size:12px;color:#e0e0e0;font-family:Arial,sans-serif;padding-right:8px;">${item.task.slice(0,90)}${item.task.length>90?'&hellip;':''}</td>
        <td nowrap style="padding-right:8px;"><span style="font-size:10px;color:${ownerColors[item.owner]||'#aaa'};background:${ownerColors[item.owner]||'#888'}22;padding:2px 8px;border-radius:10px;white-space:nowrap;font-family:Arial,sans-serif;font-weight:700;">${item.owner.split(' ')[0]}</span></td>
        <td nowrap style="font-size:11px;color:#6aa3ff;font-weight:700;font-family:Arial,sans-serif;white-space:nowrap;">${fmt(item.dueDate)}</td>
      </tr></table>
    </td></tr>`).join('')}
    ${dueSoon.length>8?`<tr><td align="center" style="padding:8px;font-size:10px;color:#666688;font-family:Arial,sans-serif;">+ ${dueSoon.length-8} more tasks</td></tr>`:''}
    </table>
  </td></tr>` : ''}

  <!-- COMPLETED THIS WEEK -->
  ${completedThisWeek.length > 0 ? `<tr><td style="padding-bottom:12px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#1a1a2e;border:1px solid #00e8bb44;border-left:4px solid #00e8bb;border-radius:8px;">
    <tr><td style="background:#00e8bb15;padding:10px 16px;border-radius:8px 8px 0 0;">
      <span style="font-size:10px;font-weight:700;color:#00e8bb;text-transform:uppercase;letter-spacing:1px;font-family:Arial,sans-serif;">&#10003; Completed This Week (${completedThisWeek.length})</span>
    </td></tr>
    ${completedThisWeek.map((item,ri) => `<tr><td style="padding:9px 16px;border-top:1px solid #2a2a4a;background:${ri%2===0?'#1a1a2e':'#141428'};">
      <table width="100%" cellpadding="0" cellspacing="0" border="0"><tr>
        <td width="20" style="color:#00e8bb;font-size:14px;font-family:Arial,sans-serif;padding-right:8px;">&#10003;</td>
        <td style="font-size:12px;color:#aaaaaa;font-family:Arial,sans-serif;">${item.task.slice(0,90)}${item.task.length>90?'&hellip;':''}</td>
        <td nowrap><span style="font-size:10px;color:${ownerColors[item.owner]||'#aaa'};background:${ownerColors[item.owner]||'#888'}22;padding:2px 8px;border-radius:10px;white-space:nowrap;font-family:Arial,sans-serif;font-weight:700;">${item.owner.split(' ')[0]}</span></td>
      </tr></table>
    </td></tr>`).join('')}
    </table>
  </td></tr>` : ''}

  <!-- FOOTER -->
  <tr><td align="center" style="padding-top:8px;padding-bottom:4px;font-size:10px;color:#555577;font-family:Arial,sans-serif;line-height:1.6;">
    Aurora &middot; R2S Project Management Intelligence &middot; Confidential<br/>
    Reply to this email to confirm task completion &mdash; Aurora will automatically update the checklist.<br/>
    ${global._testMode ? '&#x1F9EA; TEST SEND &mdash; not sent to full recipient list' : 'The full updated checklist is attached as an Excel file.'}
  </td></tr>

</table>
</td></tr></table>
</body></html>`;

    const subject = `${global._testMode?'[TEST] ':''}[Aurora] ${project.cohortName} — Weekly Programme Dashboard${overdue.length > 0 ? ` ⚠️ ${overdue.length} overdue` : daysToNext !== null && daysToNext <= 7 ? ' 🔴 Module this week' : ''}`;

    // Send with attachment if we have one, otherwise plain HTML
    if (excelBuffer && token) {
      try {
        const toArray = Array.isArray(TO) ? TO : [TO];
        const ccArray = Array.isArray(CC) ? CC : (CC ? [CC] : []);
        const message = {
          subject,
          body: { contentType: 'HTML', content: html },
          toRecipients: toArray.map(addr => ({ emailAddress: { address: addr } })),
          ccRecipients: ccArray.map(addr => ({ emailAddress: { address: addr } })),
          attachments: [{
            '@odata.type': '#microsoft.graph.fileAttachment',
            name: excelFilename,
            contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            contentBytes: excelBuffer.toString('base64'),
          }],
        };
        await axios.post(
          `https://graph.microsoft.com/v1.0/users/${fromMailbox}/sendMail`,
          { message },
          { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }, timeout: 30000 }
        );
        console.log(`[Internal Ops] ✓ Dashboard + Excel sent for ${project.cohortName}`);
      } catch(sendErr) {
        console.error('[Internal Ops] Graph send failed, falling back:', sendErr.message);
        await sendEmail(TO, subject, html, false, CC, true);
      }
    } else {
      await sendEmail(TO, subject, html, false, CC, true);
      console.log(`[Internal Ops] ✓ Dashboard sent (no Excel) for ${project.cohortName}`);
    }
    anySent = true;
    } catch(projectErr) {
      console.error(`[Internal Ops] Error processing ${project.cohortName}:`, projectErr.message);
      global._testMode = false;
      throw projectErr; // re-throw so the endpoint catches it
    }
  }

  if (!anySent) console.log('[Internal Ops] No cohorts in active window this week — no ops emails sent');
}



// ── API routes for internal projects ─────────────────────────────────────────
app.get('/api/internal/projects', async (req, res) => {
  try { res.json({ projects: await readInternalProjects() }); }
  catch(e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/internal/projects', express.json(), async (req, res) => {
  try {
    const { cohortName, programType, programStart, programEnd, studentCount, notes, status, moduleDates } = req.body;
    if (!cohortName || !programType) return res.status(400).json({ error: 'cohortName and programType required' });
    const id = `ip_${Date.now()}`;
    const project = {
      id, cohortName, programType, programStart, programEnd,
      studentCount: studentCount || 0, notes: notes || '',
      status: status || 'Scheduled',
      moduleDates: moduleDates || [],
      createdAt: new Date().toISOString(), updatedAt: new Date().toISOString()
    };
    const checklist = buildChecklist(id, programStart, programEnd, moduleDates || []);
    const projects = await readInternalProjects();
    projects.push(project);
    await writeInternalProjects(projects);
    await writeInternalChecklist(id, checklist);
    console.log(`[Internal] Created: ${cohortName} (${checklist.length} checklist items, ${(moduleDates||[]).length} modules)`);

    // Auto-send reminders for any items due within 14 days or already overdue
    // (catches late additions like this cohort)
    const today = new Date(); today.setHours(0,0,0,0);
    let autoSent = 0;
    for (const item of checklist) {
      if (!item.dueDate) continue;
      const due = new Date(item.dueDate); due.setHours(0,0,0,0);
      const daysUntil = Math.round((due - today) / (1000*60*60*24));
      // Send immediately for anything due within 14 days or overdue (not yet completed)
      if (daysUntil <= 14 && item.status !== 'Completed') {
        try {
          await sendInternalTaskReminder(project, item);
          item.reminderSent = true;
          item.lastReminderAt = new Date().toISOString();
          autoSent++;
          // Small delay to avoid rate limiting
          await new Promise(r => setTimeout(r, 300));
        } catch(e) { console.error('[Internal] Auto-reminder error:', e.message); }
      }
    }
    // Save updated reminder flags
    await writeInternalChecklist(id, checklist);
    console.log(`[Internal] Auto-sent ${autoSent} immediate reminders for upcoming/overdue tasks`);

    res.json({ project, checklist, autoSent });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.put('/api/internal/projects/:id', express.json(), async (req, res) => {
  try {
    const projects = await readInternalProjects();
    const idx = projects.findIndex(p => p.id === req.params.id);
    if (idx < 0) return res.status(404).json({ error: 'Not found' });
    Object.assign(projects[idx], req.body, { updatedAt: new Date().toISOString() });
    await writeInternalProjects(projects);

    // If module dates were updated, regenerate the module-specific checklist items
    if (req.body.moduleDates) {
      const p = projects[idx];
      const existing = await readInternalChecklist(req.params.id);
      // Keep non-module items as-is (preserve their completion status)
      const nonModuleItems = existing.filter(i => i.trigger !== 'module');
      // Rebuild module items from new dates
      const newModuleItems = [];
      let idx2 = nonModuleItems.length;
      const moduleTasks = IOP_CHECKLIST_TEMPLATE.filter(t => t.trigger === 'module');
      (req.body.moduleDates || []).forEach((modDate, modIdx) => {
        moduleTasks.forEach(t => {
          const d = new Date(modDate);
          d.setDate(d.getDate() + t.daysOffset);
          // Check if this item already exists (same module number + task)
          const existingItem = existing.find(i =>
            i.trigger === 'module' && i.moduleNumber === modIdx + 1 &&
            i.task === t.task
          );
          newModuleItems.push(existingItem ? {
            ...existingItem,
            moduleDate: modDate,
            dueDate: d.toISOString().slice(0, 10),
            group: `${t.group} — Module ${modIdx + 1} (${new Date(modDate).toLocaleDateString('en-AU', { day:'numeric',month:'short',year:'numeric' })})`,
          } : {
            id: `chk_${req.params.id}_m${idx2++}`,
            projectId: req.params.id,
            phase: t.phase, phaseName: t.phaseName, step: t.step,
            group: `${t.group} — Module ${modIdx + 1} (${new Date(modDate).toLocaleDateString('en-AU', { day:'numeric',month:'short',year:'numeric' })})`,
            task: t.task, owner: t.owner,
            ownerEmail: INTERNAL_STAFF[t.owner] || 'info@risk2solution.com',
            trigger: 'module', moduleDate: modDate, moduleNumber: modIdx + 1,
            dueDate: d.toISOString().slice(0, 10),
            status: 'Not Started', reminderSent: false, completedAt: null,
            createdAt: new Date().toISOString(),
          });
        });
      });
      await writeInternalChecklist(req.params.id, [...nonModuleItems, ...newModuleItems]);
      console.log(`[Internal] Checklist regenerated: ${nonModuleItems.length} base + ${newModuleItems.length} module items`);
    }

    res.json({ project: projects[idx] });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.delete('/api/internal/projects/:id', async (req, res) => {
  try {
    const projects = await readInternalProjects();
    const filtered = projects.filter(p => p.id !== req.params.id);
    await writeInternalProjects(filtered);
    res.json({ success: true, projects: filtered });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/internal/projects/:id/checklist', async (req, res) => {
  try { res.json({ checklist: await readInternalChecklist(req.params.id) }); }
  catch(e) { res.status(500).json({ error: e.message }); }
});

app.put('/api/internal/projects/:id/checklist/:itemId', express.json(), async (req, res) => {
  try {
    const checklist = await readInternalChecklist(req.params.id);
    const idx = checklist.findIndex(i => i.id === req.params.itemId);
    if (idx < 0) return res.status(404).json({ error: 'Item not found' });
    Object.assign(checklist[idx], req.body);
    if (req.body.status === 'Completed' && !checklist[idx].completedAt) {
      checklist[idx].completedAt = new Date().toISOString();
    }
    await writeInternalChecklist(req.params.id, checklist);
    res.json({ item: checklist[idx] });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/internal/projects/:id/checklist/:itemId/remind', async (req, res) => {
  try {
    const projects = await readInternalProjects();
    const project = projects.find(p => p.id === req.params.id);
    if (!project) return res.status(404).json({ error: 'Project not found' });
    const checklist = await readInternalChecklist(req.params.id);
    const item = checklist.find(i => i.id === req.params.itemId);
    if (!item) return res.status(404).json({ error: 'Item not found' });
    await sendInternalTaskReminder(project, item);
    res.json({ success: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/internal/report', async (req, res) => {
  try { await sendInternalProjectStatusReport(); res.json({ success: true }); }
  catch(e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/internal/ops-update', async (req, res) => {
  try { await sendInternalWeeklyOpsUpdate(); res.json({ success: true }); }
  catch(e) { res.status(500).json({ error: e.message }); }
});

// ── Test endpoints — send all reports to Kandia only ──────────────────────────
const TEST_EMAIL = 'kandia@risk2solution.com';

app.post('/api/test/report/weekly', async (req, res) => {
  try {
    const projects = await db.getProjects();
    const standard = projects.filter(p => p.type === 'standard');
    // Temporarily override sendEmail to only send to Kandia
    const origSend = sendEmail;
    const testSend = async (to, subject, body, isInternal, cc, isHtml) => {
      return origSend(TEST_EMAIL, `[TEST] ${subject}`, body, false, [], isHtml);
    };
    // Call with overridden recipients
    const html = await buildWeeklyReportHtml(standard);
    await sendEmail(TEST_EMAIL, `[TEST] R2S Weekly Project Status`, html, false, [], true);
    res.json({ success: true, sentTo: TEST_EMAIL });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/test/report/internal-leadership', async (req, res) => {
  try {
    const projects = await readInternalProjects();
    if (!projects.length) return res.status(400).json({ error: 'No internal projects found' });
    const now = new Date();
    const now0 = new Date(now); now0.setHours(0,0,0,0);
    const dateStr = now.toLocaleDateString('en-AU', { weekday:'long', day:'numeric', month:'long', year:'numeric' });
    const ownerColors = { 'Dave Cohen':'#6aa3ff','Diane Kruger':'#ff608a','Cherry Abadeza':'#00e8bb','Janita Zhang':'#ffd93d','Dr Paul Johnston':'#a78bfa','Trainer':'#888' };
    const fmt = d => new Date(d).toLocaleDateString('en-AU',{day:'numeric',month:'short'});
    let reportSections = '';
    for (const project of projects) {
      const checklist = await readInternalChecklist(project.id);
      const programName = project.programType === 'grad_cert' ? 'Graduate Certificate (11056NAT)' : 'Graduate Diploma (11066NAT)';
      const moduleDates = (project.moduleDates||[]).map(d=>{const dt=new Date(d);dt.setHours(0,0,0,0);return dt;}).filter(d=>!isNaN(d)).sort((a,b)=>a-b);
      const nextModule = moduleDates.find(d=>d>=now0);
      const daysToNext = nextModule?Math.round((nextModule-now0)/(1000*60*60*24)):null;
      const totalDone = checklist.filter(i=>i.status==='Completed').length;
      const totalTasks = checklist.length;
      const pct = totalTasks>0?Math.round(totalDone/totalTasks*100):0;
      const pctColor = pct>=80?'#00e8bb':pct>=50?'#ffd93d':'#ff608a';
      const overdue = checklist.filter(i=>i.status!=='Completed'&&i.dueDate&&new Date(i.dueDate)<now0);
      const nextWeek = new Date(now0.getTime()+7*24*60*60*1000);
      const dueThisWeek = checklist.filter(i=>i.status!=='Completed'&&i.dueDate&&new Date(i.dueDate)>=now0&&new Date(i.dueDate)<=nextWeek);
      const statusBadge = daysToNext===null?{text:'Programme active',color:'#00e8bb'} : daysToNext<=7?{text:`🔴 Module in ${daysToNext} days`,color:'#ff608a'} : {text:`🟢 Module in ${daysToNext} days`,color:'#00e8bb'};
      reportSections += `<div style="background:#1a1a2e;border:1px solid #2a2a4a;border-radius:12px;padding:20px;margin-bottom:16px">
        <div style="font-size:15px;font-weight:700;color:#ffffff;margin-bottom:4px">${project.cohortName}</div>
        <div style="font-size:11px;color:#6aa3ff;margin-bottom:12px">${programName} · <span style="color:${statusBadge.color}">${statusBadge.text}</span></div>
        <div style="margin-bottom:4px"><span style="font-size:11px;color:#aaaaaa">${totalDone}/${totalTasks} complete</span><span style="font-weight:700;color:${pctColor}">${pct}%</span></div>
        <div style="background:#2a2a4a;border-radius:4px;height:10px;margin-bottom:12px"><div style="width:${pct}%;background:${pctColor};height:10px;border-radius:4px"></div></div>
        ${overdue.length>0?`<div style="background:#ff608a15;border:1px solid #ff608a44;border-radius:6px;padding:10px;margin-bottom:8px"><div style="font-size:10px;font-weight:700;color:#ff608a;margin-bottom:6px">⚠ ${overdue.length} OVERDUE</div>${overdue.map(i=>`<div style="font-size:11px;color:#cccccc;margin-bottom:3px">• <b style="color:${ownerColors[i.owner]||'#aaa'}">${i.owner.split(' ')[0]}</b>: ${i.task.slice(0,70)} — <span style="color:#ff608a">${fmt(i.dueDate)}</span></div>`).join('')}</div>`:''}
        ${dueThisWeek.length>0?`<div style="background:#ffd93d15;border:1px solid #ffd93d44;border-radius:6px;padding:10px"><div style="font-size:10px;font-weight:700;color:#ffd93d;margin-bottom:6px">📋 DUE THIS WEEK (${dueThisWeek.length})</div>${dueThisWeek.map(i=>`<div style="font-size:11px;color:#cccccc;margin-bottom:3px">• <b style="color:${ownerColors[i.owner]||'#aaa'}">${i.owner.split(' ')[0]}</b>: ${i.task.slice(0,70)}</div>`).join('')}</div>`:'<div style="font-size:11px;color:#00e8bb">✓ Nothing due this week</div>'}
      </div>`;
    }
    const html = `<!DOCTYPE html><html xmlns:o="urn:schemas-microsoft-com:office:office"><body style="margin:0;padding:0;background:#0f0f1a;font-family:Arial,sans-serif"><table width="700" cellpadding="20" cellspacing="0" border="0" align="center" style="max-width:700px;"><tr><td>
      <div style="background:#1a1a3e;border-radius:12px;padding:24px;margin-bottom:16px;border:1px solid #2a2a4a"><div style="font-size:10px;color:#6aa3ff;text-transform:uppercase;letter-spacing:2px;margin-bottom:6px">TEST SEND · Institute of Presilience</div><div style="font-size:22px;font-weight:700;color:#fff;margin-bottom:4px">Internal Programmes — Leadership Report</div><div style="font-size:12px;color:#8899bb">${dateStr} · Aurora</div></div>
      ${reportSections}
      <div style="text-align:center;color:#555577;font-size:10px;padding:12px">🧪 TEST SEND — Aurora · R2S Project Management Intelligence</div>
    </div></body></html>`;
    await sendEmail(TEST_EMAIL, `[TEST] IoP Internal Programmes — Leadership Report`, html, false, [], true);
    res.json({ success: true, sentTo: TEST_EMAIL });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/pmchecklist/followup', async (req, res) => {
  try { await sendPMChecklistFollowUp(); res.json({ success: true }); }
  catch(e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/test/report/weekly-pm', async (req, res) => {
  try {
    console.log('[Test] Sending PM checklist follow-up to Kandia...');
    await sendPMChecklistFollowUp();
    res.json({ success: true, sentTo: 'kandia@risk2solution.com' });
  } catch(e) {
    console.error('[Test] PM follow-up error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/test/report/internal-ops', async (req, res) => {
  try {
    global._testMode = true;
    global._testEmail = TEST_EMAIL;
    console.log('[Test Ops] Starting ops dashboard test send...');
    await sendInternalWeeklyOpsUpdate(true);
    global._testMode = false;
    console.log('[Test Ops] Done');
    res.json({ success: true, sentTo: TEST_EMAIL });
  } catch(e) {
    global._testMode = false;
    console.error('[Test Ops] ERROR:', e.message, e.stack?.slice(0,300));
    res.status(500).json({ error: e.message });
  }
});

// ══════════════════════════════════════════════════════════════════════════════
// WEEKLY PM CHECKLIST — Wednesday accountability check-in for Diane
// ══════════════════════════════════════════════════════════════════════════════

// The 11 checklist questions — with metadata about whether they are one-time or weekly
const PM_CHECKLIST_QUESTIONS = [
  { id: 'q1',  text: 'Have you sent the notification email to the consultant regarding the project?',           weekly: false, phase: [0,1,2,3,4] },
  { id: 'q2',  text: 'Have you sent the welcome email to the client with the onboarding information we request?', weekly: false, phase: [0,1,2,3,4] },
  { id: 'q3',  text: 'Have you created and shared the SP (SharePoint) link with the consultant/s on the project?', weekly: false, phase: [0,1,2,3,4] },
  { id: 'q4',  text: 'Have you scheduled the kick-off meeting with the client, yourself, and the consultant/s?',  weekly: false, phase: [0,1,2,3,4] },
  { id: 'q5',  text: 'Have you followed up and touched base with the consultant/s to ensure the project is running correctly and on time?', weekly: true,  phase: [1,2,3] },
  { id: 'q6',  text: 'Have you ensured the consultant/s are saving their work as they go in the correct client folder this week?',         weekly: true,  phase: [1,2,3] },
  { id: 'q7',  text: 'Have you checked on the deliverables to ensure they are tracking on time?',               weekly: true,  phase: [1,2,3] },
  { id: 'q8',  text: 'Have you checked in with the consultant/s this week?',                                    weekly: true,  phase: [1,2,3] },
  { id: 'q9',  text: 'Have you ensured the invoicing is accurate and up to date?',                              weekly: true,  phase: [1,2,3,4] },
  { id: 'q10', text: 'Has the final invoice been sent to the client?',                                          weekly: false, phase: [4] },
  { id: 'q11', text: 'Has the project been closed out and all deliverables met?',                               weekly: false, phase: [4,5] },
];

// ── Read/write PM checklist state per project ─────────────────────────────────
async function readPMChecklist(projectId) {
  try {
    if (db.read) return await db.read(`pm_checklist_${projectId}.json`, {});
    const f = require('path').join(db.DATA, `pm_checklist_${projectId}.json`);
    return require('fs').existsSync(f) ? JSON.parse(require('fs').readFileSync(f, 'utf8')) : {};
  } catch { return {}; }
}

async function writePMChecklist(projectId, data) {
  if (db.write) return await db.write(`pm_checklist_${projectId}.json`, data);
  require('fs').writeFileSync(
    require('path').join(db.DATA, `pm_checklist_${projectId}.json`),
    JSON.stringify(data, null, 2)
  );
}

// Build the checklist state for a project — returns current week's state + history
async function getPMChecklistState(project) {
  const state = await readPMChecklist(project.id);
  const weekKey = getWeekKey();

  // Initialise this week's entry if needed
  if (!state.weeks) state.weeks = {};
  if (!state.weeks[weekKey]) state.weeks[weekKey] = {};

  // ALL 11 questions shown for every project regardless of phase
  const questions = PM_CHECKLIST_QUESTIONS;

  return questions.map(q => {
    const thisWeek = state.weeks[weekKey][q.id] || { checked: false, checkedAt: null };

    // Weekly items: always start fresh each week — never carry over
    // One-time items: check if ever ticked in ANY previous week
    let previouslyCompleted = false;
    let previouslyCompletedAt = null;
    if (!q.weekly) {
      const pastWeeks = Object.entries(state.weeks || {})
        .filter(([k]) => k !== weekKey)
        .sort(([a], [b]) => b.localeCompare(a));
      for (const [, wkData] of pastWeeks) {
        if (wkData[q.id]?.checked) {
          previouslyCompleted = true;
          previouslyCompletedAt = wkData[q.id].checkedAt;
          break;
        }
      }
    }
    // Weekly items are NEVER carried over — always unchecked at start of new week
    const effectivelyChecked = q.weekly
      ? thisWeek.checked  // only this week's tick counts
      : (thisWeek.checked || previouslyCompleted);

    return {
      ...q,
      checked: effectivelyChecked,
      checkedAt: thisWeek.checkedAt || (!q.weekly ? previouslyCompletedAt : null),
      previouslyCompleted: !q.weekly && previouslyCompleted,
      previouslyCompletedAt,
      isHistoric: !q.weekly && !thisWeek.checked && previouslyCompleted,
    };
  });
}

function getWeekKey() {
  // ISO week key: YYYY-Www
  const now = new Date();
  const d = new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()));
  d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  const weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  return `${d.getUTCFullYear()}-W${String(weekNo).padStart(2, '0')}`;
}

// ── Send the weekly PM checklist email to Diane ───────────────────────────────
async function sendWeeklyPMChecklist() {
  const projects = await db.getProjects();
  const active = projects.filter(p =>
    p.type === 'standard' &&
    !['Completed', 'Terminated', 'Closed'].includes(p.status) &&
    (p.phase || 0) < 5
  );

  if (!active.length) {
    console.log('[PMChecklist] No active projects — skipping');
    return;
  }

  const weekKey = getWeekKey();
  const dateStr = new Date().toLocaleDateString('en-AU', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });

  // Build HTML checklist for each project
  let projectSections = '';
  for (const p of active) {
    const questions = await getPMChecklistState(p);
    const done = questions.filter(q => q.checked).length;
    const phaseLabel = ['Kick-off', 'Deployment', 'Monitoring & Review', 'Reporting', 'Close-out', 'Completed'][p.phase || 0];

    projectSections += `
      <div style="background:#1e1e3a;border:1px solid #2a2a5a;border-radius:8px;padding:18px;margin-bottom:18px">
        <div style="margin-bottom:14px">
          <div>
            <div style="font-size:15px;font-weight:700;color:#fff">${p.clientName}</div>
            <div style="font-size:11px;color:#6aa3ff;margin-top:2px">${p.projectName || ''} · Phase: ${phaseLabel}</div>
          </div>
          <div style="text-align:right">
            <div style="font-size:13px;font-weight:600;color:${done===questions.length?'#00e8bb':'#ffd93d'}">${done}/${questions.length} complete</div>
          </div>
        </div>
        ${questions.map((q, i) => `
        <div style="padding:8px 0;border-bottom:1px solid #2a2a4a22">
          <div style="width:20px;height:20px;border-radius:4px;border:2px solid ${q.checked ? '#00e8bb' : '#3a3a6a'};background:${q.checked ? '#00e8bb' : 'transparent'};margin-top:1px;justify-content:center">
            ${q.checked ? '<span style="color:#0f0f1a;font-weight:bold;font-size:13px">✓</span>' : ''}
          </div>
          <div style="flex:1">
            <div style="font-size:12px;color:${q.checked ? '#aaa' : '#e0e0e0'};${q.checked ? 'text-decoration:line-through;' : ''}">${i + 1}. ${q.text}</div>
            ${q.isHistoric ? `<div style="font-size:10px;color:#00e8bb;margin-top:2px">✓ Previously completed ${q.previouslyCompletedAt ? new Date(q.previouslyCompletedAt).toLocaleDateString('en-AU', {day:'numeric',month:'short',year:'numeric'}) : ''} — one-time task</div>` : ''}
            ${q.checked && q.checkedAt && !q.isHistoric ? `<div style="font-size:10px;color:#888;margin-top:2px">Checked ${new Date(q.checkedAt).toLocaleDateString('en-AU', {day:'numeric',month:'short'})}</div>` : ''}
            ${q.weekly && !q.checked ? `<div style="font-size:10px;color:#ffd93d;margin-top:2px">↻ Weekly task — please complete and tick off this week</div>` : ''}
          </div>
        </div>`).join('')}
      </div>`;
  }

  const html = `<!DOCTYPE html><html xmlns:o="urn:schemas-microsoft-com:office:office"><body style="margin:0;padding:0;background:#0f0f1a;font-family:Arial,sans-serif">
  <table width="680" cellpadding="24" cellspacing="0" border="0" align="center" style="max-width:680px;"><tr><td>
    <div style="background:#1a1a3e;border-radius:12px;padding:24px;margin-bottom:20px;border:1px solid #2a2a5a">
      <div style="font-size:11px;color:#6aa3ff;text-transform:uppercase;letter-spacing:2px;margin-bottom:6px">Aurora · Weekly PM Checklist</div>
      <div style="font-size:22px;font-weight:700;color:#fff;margin-bottom:6px">Wednesday Check-in</div>
      <div style="font-size:13px;color:#aaa">${dateStr}</div>
    </div>

    <div style="background:#1a1a2e;border:1px solid #2a2a4a;border-radius:8px;padding:16px;margin-bottom:20px;color:#ccc;font-size:13px;line-height:1.7">
      Hi Diane,<br><br>
      For each project currently running, please complete the following check-in to confirm everything is on track and up to date.
      Items marked <span style="color:#00e8bb">✓ Previously completed</span> are <strong>one-time tasks</strong> — already done, no need to re-check.<br/>
      Items marked <span style="color:#ffd93d">↻ Weekly</span> must be completed and ticked off <strong>fresh every week</strong> — they have been reset for this week.
      <br><br>
      Once you have reviewed each section, please reply to this email confirming you have completed the checklist. Aurora will log your responses automatically in the portal.
    </div>

    ${projectSections}

    <div style="background:#1a1a2e;border:1px solid #2a2a4a;border-radius:8px;padding:14px;margin-top:8px;text-align:center">
      <div style="font-size:12px;color:#888;margin-bottom:8px">You can also complete and review checklists directly in Aurora under each project → Checklist tab</div>
      ${process.env.FRONTEND_URL ? `<a href="${process.env.FRONTEND_URL}" style="color:#6aa3ff;font-size:12px">Open Aurora portal →</a>` : ''}
    </div>

    <div style="text-align:center;color:#444;font-size:11px;padding-top:16px">
      Aurora · R2S Project Management Intelligence · Confidential
    </div>
  </div></body></html>`;

  await sendEmail(
    'diane.k@risk2solution.com',
    `[Aurora] Wednesday PM Check-in — ${active.length} active project${active.length !== 1 ? 's' : ''}`,
    html, false, ['kandia@risk2solution.com'], true
  );
  console.log(`[PMChecklist] Wednesday checklist sent to Diane (CC: Kandia) — ${active.length} projects`);
}

// ── Wednesday follow-up: check completion and notify Kandia ──────────────────
async function sendPMChecklistFollowUp() {
  const projects = await db.getProjects();
  const active = projects.filter(p =>
    p.type === 'standard' &&
    !['Completed', 'Terminated', 'Closed'].includes(p.status) &&
    (p.phase || 0) < 5
  );
  if (!active.length) return;

  const now = new Date();
  const dateStr = now.toLocaleDateString('en-AU', { weekday:'long', day:'numeric', month:'long', year:'numeric' });
  const weekKey = getWeekKey();

  // Build per-project data
  const projectData = [];
  let totalDone = 0, totalItems = 0, totalOutstanding = 0;

  for (const p of active) {
    const questions = await getPMChecklistState(p);
    const done      = questions.filter(q => q.checked || q.isHistoric).length;
    const total     = questions.length;
    const outstanding = questions.filter(q => !q.checked && !q.isHistoric);
    const completedThisWeek = questions.filter(q => q.checked && !q.isHistoric);
    totalDone += done;
    totalItems += total;
    totalOutstanding += outstanding.length;
    projectData.push({ p, questions, done, total, outstanding, completedThisWeek });
  }

  const allComplete = totalOutstanding === 0;
  const pct = totalItems > 0 ? Math.round(totalDone / totalItems * 100) : 0;
  const pctColor = allComplete ? '#00e8bb' : pct >= 70 ? '#ffd93d' : '#ff608a';
  const phaseLabel = p => ['Kick-off','Deployment','Monitoring & Review','Reporting','Close-out','Completed'][p.phase||0];

  // Build project rows
  let projectRows = '';
  for (const { p, done, total, outstanding, completedThisWeek } of projectData) {
    const rowPct = total > 0 ? Math.round(done/total*100) : 0;
    const rowColor = rowPct === 100 ? '#00e8bb' : rowPct >= 70 ? '#ffd93d' : '#ff608a';

    projectRows += `
    <!-- Project row -->
    <tr><td style="padding:0 0 10px 0;">
      <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#1a1a2e;border:1px solid ${outstanding.length>0?'#ff386044':'#2a2a4a'};border-left:4px solid ${outstanding.length>0?'#ff3860':rowPct===100?'#00e8bb':'#ffd93d'};border-radius:8px;">

        <!-- Project header -->
        <tr><td style="padding:12px 16px;border-bottom:1px solid #2a2a4a;">
          <table width="100%" cellpadding="0" cellspacing="0" border="0"><tr>
            <td>
              <div style="font-size:13px;font-weight:700;color:#ffffff;font-family:Arial,sans-serif;">${p.clientName}</div>
              <div style="font-size:10px;color:#8899bb;margin-top:2px;font-family:Arial,sans-serif;">${p.projectName||''} &middot; ${phaseLabel(p)}</div>
            </td>
            <td align="right" style="padding-left:12px;white-space:nowrap;">
              <div style="font-size:18px;font-weight:700;color:${rowColor};font-family:Arial,sans-serif;">${rowPct}%</div>
              <div style="font-size:10px;color:#8899bb;font-family:Arial,sans-serif;">${done}/${total} complete</div>
            </td>
          </tr></table>
          <!-- Mini progress bar -->
          <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:8px;background:#2a2a4a;border-radius:3px;height:6px;">
          <tr><td style="height:6px;">
            <div style="width:${rowPct}%;background:${rowColor};height:6px;border-radius:3px;min-width:${rowPct>0?'4px':'0'};"></div>
          </td></tr></table>
        </td></tr>

        ${outstanding.length > 0 ? `
        <!-- Outstanding items -->
        <tr><td style="padding:10px 16px;${completedThisWeek.length>0?'border-bottom:1px solid #2a2a4a;':''}">
          <div style="font-size:9px;font-weight:700;color:#ff608a;text-transform:uppercase;letter-spacing:1px;margin-bottom:6px;font-family:Arial,sans-serif;">&#9888; Outstanding (${outstanding.length})</div>
          ${outstanding.map(q => `
          <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom:4px;"><tr>
            <td width="14" style="vertical-align:top;padding-top:1px;">
              <div style="width:12px;height:12px;border:2px solid #3a3a6a;border-radius:3px;background:#141428;"></div>
            </td>
            <td style="padding-left:8px;font-size:11px;color:#cccccc;font-family:Arial,sans-serif;line-height:1.4;">
              ${q.text}
              ${q.weekly ? '<span style="font-size:9px;color:#ffd93d;background:#ffd93d22;padding:1px 5px;border-radius:4px;margin-left:4px;font-family:Arial,sans-serif;">Weekly</span>' : ''}
            </td>
          </tr></table>`).join('')}
        </td></tr>` : ''}

        ${completedThisWeek.length > 0 ? `
        <!-- Completed this week -->
        <tr><td style="padding:10px 16px;">
          <div style="font-size:9px;font-weight:700;color:#00e8bb;text-transform:uppercase;letter-spacing:1px;margin-bottom:6px;font-family:Arial,sans-serif;">&#10003; Completed this week (${completedThisWeek.length})</div>
          ${completedThisWeek.map(q => `
          <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom:4px;"><tr>
            <td width="14" style="vertical-align:top;padding-top:1px;color:#00e8bb;font-size:12px;font-family:Arial,sans-serif;">&#10003;</td>
            <td style="padding-left:8px;font-size:11px;color:#888;font-family:Arial,sans-serif;text-decoration:line-through;line-height:1.4;">${q.text}</td>
          </tr></table>`).join('')}
        </td></tr>` : ''}

      </table>
    </td></tr>`;
  }

  const html = `<!DOCTYPE html>
<html xmlns:o="urn:schemas-microsoft-com:office:office">
<head><meta charset="utf-8"/><!--[if mso]><xml><o:OfficeDocumentSettings><o:AllowPNG/></o:OfficeDocumentSettings></xml><![endif]--></head>
<body style="margin:0;padding:0;background:#0f0f1a;">
<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#0f0f1a;">
<tr><td align="center" style="padding:16px 8px;">
<table width="680" cellpadding="0" cellspacing="0" border="0" style="max-width:680px;width:100%;">

  <!-- HEADER -->
  <tr><td style="padding-bottom:12px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#1a1a3e;border-radius:10px;border:1px solid #2a2a4a;">
    <tr><td style="padding:24px 28px;">
      <div style="font-size:10px;color:#6aa3ff;text-transform:uppercase;letter-spacing:2px;margin-bottom:8px;font-family:Arial,sans-serif;">Aurora &middot; Wednesday PM Checklist Review</div>
      <div style="font-size:22px;font-weight:700;color:#ffffff;margin-bottom:4px;font-family:Arial,sans-serif;">Diane's Weekly Check-in Status</div>
      <div style="font-size:12px;color:#8899bb;margin-bottom:16px;font-family:Arial,sans-serif;">${dateStr} &middot; Week ${weekKey}</div>
      <!-- Status badge -->
      <table cellpadding="0" cellspacing="0" border="0"><tr><td style="background:${allComplete?'#00e8bb15':'#ff608a15'};border:1px solid ${allComplete?'#00e8bb':'#ff608a'};border-radius:6px;padding:8px 16px;">
        <span style="color:${allComplete?'#00e8bb':'#ff608a'};font-size:13px;font-weight:700;font-family:Arial,sans-serif;">${allComplete ? '&#10003; All items complete this week' : `&#9888; ${totalOutstanding} item${totalOutstanding!==1?'s':''} outstanding across ${projectData.filter(d=>d.outstanding.length>0).length} project${projectData.filter(d=>d.outstanding.length>0).length!==1?'s':''}`}</span>
      </td></tr></table>
    </td></tr>
    </table>
  </td></tr>

  <!-- SUMMARY TILES -->
  <tr><td style="padding-bottom:12px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0">
    <tr>
      ${[
        { label:'Active projects', val:String(active.length),      color:'#6aa3ff' },
        { label:'Items complete',  val:`${totalDone}/${totalItems}`, color:'#00e8bb' },
        { label:'Outstanding',     val:String(totalOutstanding),   color:totalOutstanding>0?'#ff608a':'#00e8bb' },
        { label:'Overall',         val:`${pct}%`,                  color:pctColor },
      ].map((s,i) => `<td width="25%" style="padding:0 ${i<3?'6px':'0'} 0 0;vertical-align:top;">
        <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#1a1a2e;border:1px solid #2a2a4a;border-radius:8px;">
        <tr><td align="center" style="padding:14px 8px;">
          <div style="font-size:20px;font-weight:700;color:${s.color};font-family:Arial,sans-serif;">${s.val}</div>
          <div style="font-size:10px;color:#888;text-transform:uppercase;letter-spacing:0.5px;margin-top:4px;font-family:Arial,sans-serif;">${s.label}</div>
        </td></tr>
        </table>
      </td>`).join('')}
    </tr>
    </table>
  </td></tr>

  <!-- OVERALL PROGRESS BAR -->
  <tr><td style="padding-bottom:16px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#1a1a2e;border:1px solid #2a2a4a;border-radius:8px;">
    <tr><td style="padding:14px 18px;">
      <table width="100%" cellpadding="0" cellspacing="0" border="0">
      <tr>
        <td style="font-size:12px;font-weight:700;color:#e0e0e0;font-family:Arial,sans-serif;">Overall PM Checklist Completion</td>
        <td align="right" style="font-size:16px;font-weight:700;color:${pctColor};font-family:Arial,sans-serif;">${pct}%</td>
      </tr></table>
      <div style="margin-top:8px;background:#2a2a4a;border-radius:4px;height:10px;">
        <div style="width:${pct}%;background:${pctColor};height:10px;border-radius:4px;min-width:${pct>0?'4px':'0'};"></div>
      </div>
    </td></tr>
    </table>
  </td></tr>

  <!-- PROJECT CARDS -->
  <table width="100%" cellpadding="0" cellspacing="0" border="0">
  ${projectRows}
  </table>

  <!-- FOOTER -->
  <tr><td align="center" style="padding-top:8px;font-size:10px;color:#555577;font-family:Arial,sans-serif;line-height:1.6;">
    Aurora &middot; R2S Project Management Intelligence &middot; Confidential<br/>
    This report is for your review only. Diane has not been copied on this email.
  </td></tr>

</table>
</td></tr></table>
</body></html>`;

  const subject = `[Aurora] PM Checklist — ${allComplete ? 'All complete ✓' : `${totalOutstanding} outstanding across ${projectData.filter(d=>d.outstanding.length>0).length} project${projectData.filter(d=>d.outstanding.length>0).length!==1?'s':''}`} · ${now.toLocaleDateString('en-AU',{day:'numeric',month:'short'})}`;

  // Send to Kandia ONLY — no CC to Diane
  await sendEmail('kandia@risk2solution.com', subject, html, false, [], true);
  console.log(`[PMChecklist] Follow-up dashboard sent to Kandia only — ${allComplete ? 'all complete' : `${totalOutstanding} outstanding`}`);
}

// ── API routes for PM checklist ───────────────────────────────────────────────
app.get('/api/projects/:id/pmchecklist', async (req, res) => {
  try {
    const project = await db.getProject(req.params.id);
    if (!project) return res.status(404).json({ error: 'Project not found' });
    const questions = await getPMChecklistState(project);
    const weekKey = getWeekKey();
    const state = await readPMChecklist(req.params.id);
    // Return full history too
    const history = Object.entries(state.weeks || {})
      .filter(([k]) => k !== weekKey)
      .sort(([a], [b]) => b.localeCompare(a))
      .slice(0, 8) // last 8 weeks
      .map(([week, data]) => ({ week, items: data }));
    res.json({ questions, weekKey, history });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.put('/api/projects/:id/pmchecklist/:questionId', express.json(), async (req, res) => {
  try {
    const project = await db.getProject(req.params.id);
    if (!project) return res.status(404).json({ error: 'Project not found' });
    const state = await readPMChecklist(req.params.id);
    const weekKey = getWeekKey();
    if (!state.weeks) state.weeks = {};
    if (!state.weeks[weekKey]) state.weeks[weekKey] = {};
    state.weeks[weekKey][req.params.questionId] = {
      checked: req.body.checked,
      checkedAt: req.body.checked ? new Date().toISOString() : null,
      checkedBy: 'Diane',
    };
    await writePMChecklist(req.params.id, state);
    // Log activity
    const q = PM_CHECKLIST_QUESTIONS.find(x => x.id === req.params.questionId);
    if (q) {
      await db.logActivity(req.params.id, {
        type: 'checklist',
        summary: `PM Checklist: "${q.text.slice(0, 80)}" — ${req.body.checked ? '✓ Checked by Diane' : 'Unchecked'}`,
      });
    }
    res.json({ success: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/pmchecklist/send', async (req, res) => {
  try { await sendWeeklyPMChecklist(); res.json({ success: true }); }
  catch(e) { res.status(500).json({ error: e.message }); }
});

// ══════════════════════════════════════════════════════════════════════════════
// GENERIC INTERNAL PROJECT — upload/extract/manual creation
// ══════════════════════════════════════════════════════════════════════════════

// Extract project requirements from text using Claude
async function extractInternalProjectRequirements(text, filename) {
  const prompt = `You are an AI assistant for Risk 2 Solution (R2S), an Australian risk and resilience consulting firm.

A staff member has uploaded a document or email. Extract all project/task requirements from it.

Document content:
---
${text.slice(0, 8000)}
---

Return a JSON object (no markdown, no backticks) with these fields:
{
  "projectName": "short descriptive project name",
  "description": "2-3 sentence summary of what this project/task requires",
  "deadline": "YYYY-MM-DD if found, else null",
  "priority": "High|Medium|Low",
  "tasks": [
    {
      "task": "specific action item",
      "owner": "most relevant staff member from: Dave Cohen, Diane Kruger, Cherry Abadeza, Janita Zhang, Dr Paul Johnston, Kandia Du Bruyn, or null if unclear",
      "dueDate": "YYYY-MM-DD if found, else null",
      "notes": "any relevant context"
    }
  ],
  "deliverables": ["list", "of", "outputs", "required"],
  "stakeholders": ["names or organisations mentioned"],
  "notes": "any other important information"
}`;

  try {
    const result = await aurora('internal_extract', prompt);
    // Strip any markdown formatting
    const clean = result.replace(/```json\s*/gi, '').replace(/```\s*/g, '').trim();
    // Try to find JSON object in the response
    const jsonMatch = clean.match(/\{[\s\S]*\}/);
    if (!jsonMatch) throw new Error('No JSON found in response');
    return JSON.parse(jsonMatch[0]);
  } catch(e) {
    console.error('[Extract] Failed:', e.message);
    return null;
  }
}

// Read internal projects (generic type)
async function readGenericInternalProjects() {
  try {
    if (db.read) return await db.read('generic_internal_projects.json', []);
    const f = require('path').join(db.DATA, 'generic_internal_projects.json');
    return require('fs').existsSync(f) ? JSON.parse(require('fs').readFileSync(f, 'utf8')) : [];
  } catch { return []; }
}

async function writeGenericInternalProjects(projects) {
  if (db.write) return await db.write('generic_internal_projects.json', projects);
  require('fs').writeFileSync(require('path').join(db.DATA, 'generic_internal_projects.json'), JSON.stringify(projects, null, 2));
}

async function readGenericProjectTasks(projectId) {
  try {
    if (db.read) return await db.read(`generic_tasks_${projectId}.json`, []);
    const f = require('path').join(db.DATA, `generic_tasks_${projectId}.json`);
    return require('fs').existsSync(f) ? JSON.parse(require('fs').readFileSync(f, 'utf8')) : [];
  } catch { return []; }
}

async function writeGenericProjectTasks(projectId, tasks) {
  if (db.write) return await db.write(`generic_tasks_${projectId}.json`, tasks);
  require('fs').writeFileSync(require('path').join(db.DATA, `generic_tasks_${projectId}.json`), JSON.stringify(tasks, null, 2));
}

// Extract from uploaded file
const genericUpload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024 } });

app.post('/api/generic-internal/extract', (req, res, next) => {
  res.setHeader('Content-Type', 'application/json');
  genericUpload.fields([{ name: 'file', maxCount: 1 }, { name: 'text', maxCount: 1 }])(req, res, (err) => {
    if (err) { return res.status(400).json({ error: 'Upload error: ' + err.message }); }
    next();
  });
}, async (req, res) => {
  res.setHeader('Content-Type', 'application/json');
  try {
    let text = '';
    const uploadedFile = req.files?.file?.[0];

    if (uploadedFile) {
      const buf = uploadedFile.buffer;
      const filename = (uploadedFile.originalname || '').toLowerCase();
      console.log(`[Extract] File: ${uploadedFile.originalname} (${buf.length} bytes)`);

      if (filename.endsWith('.pdf')) {
        try { const pp = require('pdf-parse'); const parsed = await pp(buf); text = parsed.text; }
        catch(e) { text = buf.toString('utf8').replace(/[^\x20-\x7E\n]/g, ' '); }
      } else if (filename.endsWith('.docx') || filename.endsWith('.doc')) {
        try { const m = require('mammoth'); const r = await m.extractRawText({ buffer: buf }); text = r.value; }
        catch(e) { text = buf.toString('utf8').replace(/[^\x20-\x7E\n]/g, ' '); }
      } else {
        let raw = buf.toString('utf8');
        raw = raw.replace(/=\r?\n/g, '').replace(/=([0-9A-F]{2})/gi, (_, h) => {
          try { return String.fromCharCode(parseInt(h, 16)); } catch { return ''; }
        });
        if (raw.includes('<html') || raw.includes('<body') || raw.includes('<div')) {
          raw = raw.replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
                   .replace(/<[^>]+>/g, ' ')
                   .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&')
                   .replace(/&lt;/g, '<').replace(/&gt;/g, '>');
        }
        text = raw;
      }
    } else {
      // Text field from FormData
      const textField = req.files?.text?.[0];
      if (textField) {
        text = textField.buffer.toString('utf8');
      } else {
        text = req.body?.text || '';
      }
    }

    text = (text || '').replace(/\s+/g, ' ').trim().slice(0, 12000);
    console.log(`[Extract] Text: ${text.length} chars`);

    if (text.length < 20) {
      return res.status(400).json({ error: 'Could not extract readable text. Please paste the content directly.' });
    }

    const extracted = await extractInternalProjectRequirements(text, uploadedFile?.originalname || 'pasted text');
    if (!extracted) return res.status(500).json({ error: 'AI extraction failed — please try again.' });

    console.log(`[Extract] Success: "${extracted.projectName}" with ${extracted.tasks?.length || 0} tasks`);
    res.json({ extracted, textLength: text.length });
  } catch(e) {
    console.error('[Extract] Error:', e.message);
    res.status(500).json({ error: e.message });
  }
});


app.get('/api/generic-internal/projects', async (req, res) => {
  try { res.json({ projects: await readGenericInternalProjects() }); }
  catch(e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/generic-internal/projects', express.json(), async (req, res) => {
  try {
    const { projectName, description, deadline, priority, status, notes, tags } = req.body;
    if (!projectName) return res.status(400).json({ error: 'projectName required' });
    const id = `gip_${Date.now()}`;
    const project = { id, projectName, description, deadline, priority: priority||'Medium', status: status||'Active', notes, tags: tags||[], createdAt: new Date().toISOString(), updatedAt: new Date().toISOString() };
    const projects = await readGenericInternalProjects();
    projects.push(project);
    await writeGenericInternalProjects(projects);
    // Save tasks if provided
    if (req.body.tasks && req.body.tasks.length) {
      const tasks = req.body.tasks.map((t, i) => ({
        id: `gt_${id}_${i}`, projectId: id,
        task: t.task, owner: t.owner||null, ownerEmail: INTERNAL_STAFF[t.owner]||null,
        dueDate: t.dueDate||null, status: 'Not Started', notes: t.notes||'',
        createdAt: new Date().toISOString(),
      }));
      await writeGenericProjectTasks(id, tasks);
      // Send task reminders
      for (const task of tasks.filter(t => t.ownerEmail)) {
        try {
          await sendEmail(task.ownerEmail,
            `[Aurora] New task assigned: ${projectName}`,
            `Hi ${task.owner?.split(' ')[0]||''},\n\nYou have been assigned a task for the internal project: ${projectName}\n\nTask: ${task.task}\nDue: ${task.dueDate || 'TBC'}\nPriority: ${priority||'Medium'}\n${task.notes ? 'Notes: ' + task.notes : ''}\n\nPlease reply to this email once complete — Aurora will mark it as done automatically.\n\nTask ID: ${task.id}\n\nAurora\nR2S Project Management Intelligence`,
            false, []
          );
        } catch(e) { console.error('[GenericInternal] Task email error:', e.message); }
      }
    }
    console.log(`[GenericInternal] Created: ${projectName}`);
    res.json({ project });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.put('/api/generic-internal/projects/:id', express.json(), async (req, res) => {
  try {
    const projects = await readGenericInternalProjects();
    const idx = projects.findIndex(p => p.id === req.params.id);
    if (idx < 0) return res.status(404).json({ error: 'Not found' });
    Object.assign(projects[idx], req.body, { updatedAt: new Date().toISOString() });
    await writeGenericInternalProjects(projects);
    res.json({ project: projects[idx] });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.delete('/api/generic-internal/projects/:id', async (req, res) => {
  try {
    const projects = await readGenericInternalProjects();
    const filtered = projects.filter(p => p.id !== req.params.id);
    await writeGenericInternalProjects(filtered);
    res.json({ success: true, projects: filtered });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/generic-internal/projects/:id/tasks', async (req, res) => {
  try { res.json({ tasks: await readGenericProjectTasks(req.params.id) }); }
  catch(e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/generic-internal/projects/:id/tasks', express.json(), async (req, res) => {
  try {
    const tasks = await readGenericProjectTasks(req.params.id);
    const task = { id: `gt_${req.params.id}_${Date.now()}`, projectId: req.params.id, ...req.body, status: req.body.status||'Not Started', createdAt: new Date().toISOString() };
    task.ownerEmail = INTERNAL_STAFF[task.owner] || null;
    tasks.push(task);
    await writeGenericProjectTasks(req.params.id, tasks);
    // Send assignment email
    if (task.ownerEmail) {
      const projects = await readGenericInternalProjects();
      const project = projects.find(p => p.id === req.params.id);
      await sendEmail(task.ownerEmail,
        `[Aurora] New task assigned: ${project?.projectName || 'Internal project'}`,
        `Hi ${task.owner?.split(' ')[0]||''},\n\nYou have been assigned a new task.\n\nProject: ${project?.projectName||''}\nTask: ${task.task}\nDue: ${task.dueDate||'TBC'}\n${task.notes ? 'Notes: ' + task.notes : ''}\n\nReply to this email once complete.\n\nTask ID: ${task.id}\n\nAurora\nR2S Project Management Intelligence`,
        false, []
      );
    }
    res.json({ task });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.put('/api/generic-internal/projects/:id/tasks/:taskId', express.json(), async (req, res) => {
  try {
    const tasks = await readGenericProjectTasks(req.params.id);
    const idx = tasks.findIndex(t => t.id === req.params.taskId);
    if (idx < 0) return res.status(404).json({ error: 'Task not found' });
    Object.assign(tasks[idx], req.body);
    if (req.body.status === 'Complete' && !tasks[idx].completedAt) tasks[idx].completedAt = new Date().toISOString();
    await writeGenericProjectTasks(req.params.id, tasks);
    res.json({ task: tasks[idx] });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.delete('/api/generic-internal/projects/:id/tasks/:taskId', async (req, res) => {
  try {
    const tasks = await readGenericProjectTasks(req.params.id);
    await writeGenericProjectTasks(req.params.id, tasks.filter(t => t.id !== req.params.taskId));
    res.json({ success: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});
