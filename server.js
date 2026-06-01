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
  if (req.path === '/auth/login') return next();
  const authHeader = req.headers['authorization'] || '';
  const token = authHeader.replace('Bearer ', '');
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
  chat:              'claude-sonnet-4-6',
  document_analysis: 'claude-sonnet-4-6',
  contract_extract:  'claude-sonnet-4-6',
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

// ── Email (Outlook via nodemailer / Graph API) ────────────────────────────────
async function getOutlookToken() {
  const tenantId     = process.env.OUTLOOK_TENANT_ID;
  const clientId     = process.env.OUTLOOK_CLIENT_ID;
  const clientSecret = process.env.OUTLOOK_CLIENT_SECRET;
  if (!tenantId || !clientId || !clientSecret) return null;
  try {
    const res = await axios.post(
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
      new URLSearchParams({ client_id: clientId, client_secret: clientSecret, scope: 'https://graph.microsoft.com/.default', grant_type: 'client_credentials' }),
      { headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, timeout: 10000 }
    );
    return res.data.access_token;
  } catch (err) {
    console.error('[Email] Token failed:', err.message);
    return null;
  }
}

async function sendEmail(to, subject, body, isInternal = false, cc = []) {
  const token = await getOutlookToken();
  const fromMailbox = process.env.OUTLOOK_SHARED_MAILBOX || 'info@risk2solution.com';

  if (token) {
    try {
      const toArray = Array.isArray(to) ? to : [to];
      const ccArray = Array.isArray(cc) ? cc : (cc ? [cc] : []);
      const message = {
        subject,
        body: { contentType: 'Text', content: body },
        toRecipients: toArray.map(addr => ({ emailAddress: { address: addr } })),
        saveToSentItems: true,
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
  // Save draft to Outlook shared mailbox drafts folder
  const token = await getOutlookToken();
  const fromMailbox = process.env.OUTLOOK_SHARED_MAILBOX || 'info@risk2solution.com';
  if (token) {
    try {
      await axios.post(
        `https://graph.microsoft.com/v1.0/users/${fromMailbox}/messages`,
        {
          subject: draft.subject,
          body: { contentType: 'Text', content: draft.body },
          toRecipients: [{ emailAddress: { address: draft.toEmail || fromMailbox } }],
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
  const truncated = rawText.slice(0, 15000);
  const text = await aurora(
    'contract_extract',
    `You are reading a client contract or proposal for Risk 2 Solution (R2S). Extract ALL of the following information and return it as a valid JSON object with exactly these keys. Be thorough — read the entire document carefully before responding.

{
  "organisationName": "full legal organisation/company name of the client",
  "clientName": "organisation name (same as above, used for display)",
  "projectName": "project title or name of the engagement as written in the document",
  "clientContact": "primary client contact person full name",
  "clientEmail": "primary client contact email address",
  "clientPhone": "primary client contact phone number",
  "value": "total contract value as written e.g. $25,000 — include any per-session rates if relevant",
  "contractStart": "contract start date or engagement commencement date",
  "dueDate": "project completion date, contract end date, or due date",
  "summary": "full description of services R2S is providing — extract the key paragraphs word for word describing what R2S will do for the client",
  "deliverables": "all specific deliverables listed — e.g. reports, training sessions (with numbers and dates if given), assessments, presentations, workshops",
  "milestones": "any key milestones, phases, or stages mentioned with dates or conditions",
  "timeline": "overall project timeline description — start to finish with any phasing or scheduling mentioned",
  "invoicingNotes": "full payment terms, invoicing schedule, milestone payment triggers, and invoicing frequency",
  "consultant": "name(s) of any R2S consultant, trainer, or staff member assigned or mentioned",
  "consultantEmail": "email address of the assigned consultant or trainer if mentioned",
  "flightsRequired": "yes or no — are flights required for this engagement",
  "accommodationRequired": "yes or no — is accommodation required for this engagement",
  "notes": "any special conditions, exclusions, important requirements, or things the team should be aware of"
}

Return ONLY the JSON object. No markdown, no explanation, no other text. If a field is not found in the document, use an empty string "".

Document: ${filename}
---
${truncated}`,
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
async function sendConsultantBriefing(project, extracted) {
  const DIANE = 'diane.k@risk2solution.com';

  // Consultant email — use extracted email if available, otherwise fall back to info@ for now
  const consultantEmail = extracted.consultantEmail || process.env.CONSULTANT_DEFAULT_EMAIL || 'info@risk2solution.com';
  const consultantName  = extracted.consultant || project.consultant || 'Team';

  const context = buildContext(project);

  const briefingBody = await aurora(
    'status_email',
    `Draft a professional project briefing email from R2S to ${consultantName}, who has been assigned as the consultant/trainer on this project.

The email should brief them on:
- The client and project
- What we are delivering (scope and deliverables)
- Key dates (start date, due/completion date, any milestones)
- The timeline and any phasing
- Invoicing terms (so they know how we are getting paid)
- Whether flights or accommodation are required: ${extracted.flightsRequired || 'not specified'} / ${extracted.accommodationRequired || 'not specified'}
- Any special requirements, conditions, or important notes
- What is expected of them and next steps

Write it as if from the R2S project management team. Professional, clear, and to the point. Not too long. End by asking them to confirm receipt and that they are clear on the requirements.

This email will be CC'd to diane.k@risk2solution.com for our records.`,
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
const PHASES = ['Kick-off', 'Deployment', 'Monitoring & Review', 'Reporting', 'Close-out'];

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

// ── Daily batch (6am AEST = 8pm UTC) ─────────────────────────────────────────
async function runBatch() {
  console.log('\n[Batch] ═══ Aurora daily batch starting ═══');
  const projects = await db.getProjects();
  const standard = projects.filter(p => p.type === 'standard');
  const ongoing  = projects.filter(p => p.type === 'ongoing');
  console.log(`[Batch] ${standard.length} standard | ${ongoing.length} ongoing (skipped)`);

  // 1. Check due date reminders
  await checkDueDateReminders();

  // 2. Monday status emails (Mondays only)
  if (new Date().getDay() === 1) {
    for (const p of standard) {
      if (['Completed','Terminated','On Hold'].includes(p.status)) continue;
      try {
        const context = buildContext(p);
        const text = await aurora('status_email',
          `Draft a short weekly status update email to ${p.clientContact || 'the client'} at ${p.clientName} for the ${p.projectName || p.clientName} project. Current phase: ${PHASES[p.phase||0]}. Keep it to 3-4 sentences covering what happened this week, what's next, and anything needed from the client.`,
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
        if (err.message === 'MONTHLY_CAP_REACHED') { console.error('[Batch] Cap reached'); break; }
        console.error(`[Batch] Error on ${p.clientName}:`, err.message);
      }
    }
  }

  // 3. Send internal summary to team
  const atRisk = standard.filter(p => p.dueDate && (() => {
    const days = Math.round((new Date(p.dueDate) - new Date()) / (1000 * 60 * 60 * 24));
    return days <= 14 && days >= 0;
  })());

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

// Schedule: 6am AEST (UTC+10) = 8pm UTC
cron.schedule('0 20 * * *', () => runBatch().catch(console.error), { timezone: 'UTC' });

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
    const updated = await db.updateProjectField(req.params.id, req.body);
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
      }
    }
    res.json({ project: updated });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.delete('/api/projects/:id', async (req, res) => {
  try { await db.deleteProject(req.params.id); res.json({ success: true }); }
  catch (e) { res.status(500).json({ error: e.message }); }
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
      return res.status(400).json({ error: 'Could not extract text from this file. Try a different format.' });
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

    // Save document extract
    await db.saveDocument({
      id: `doc_${Date.now()}`,
      projectId, name: req.file.originalname,
      extract: rawText.slice(0, 8000),
    });

    // Send consultant briefing if a consultant/trainer is named
    let briefingDraft = null;
    if (extracted.consultant && extracted.consultant.trim()) {
      try {
        briefingDraft = await sendConsultantBriefing(project, extracted);
        console.log(`[Contract] ✓ Consultant briefing prepared for ${extracted.consultant}`);
      } catch (briefErr) {
        console.error('[Contract] Briefing failed:', briefErr.message);
      }
    }

    res.json({ project, extracted, briefingPrepared: !!briefingDraft });
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
    const drafts = db.readJSON('drafts.json');
    const draft  = drafts.find(d => d.id === req.params.id);
    if (!draft) return res.status(404).json({ error: 'Not found' });
    // Send via Outlook
    await sendEmail(draft.toEmail || INTERNAL_EMAILS[1], draft.subject, draft.body);
    await db.updateDraft(req.params.id, { approved: true });
    res.json({ success: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
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

// Reports
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

    const prompts = {
      status:    'Write a concise project status report. For each project: current phase, status, what has been delivered so far, what is outstanding, any risks or blockers. Keep it factual and easy to read.',
      milestones:'List deliverables and milestones for each project. Format clearly with project name, deliverable, and status (complete/in progress/outstanding).',
      risks:     'Summarise risks across all projects. Format: project | risk | level | recommended action. Keep it brief and factual.',
      closeout:  'Write a project close-out report. What was delivered, outcomes achieved, and 2-3 forward recommendations for the client.',
      invoices:  'Summarise invoice and payment status. Format: client | project | contract value | invoicing terms | status.',
      portfolio: 'Write a one-page portfolio overview for R2S leadership. Status per project, anything needing immediate attention at the top, key dates ahead.',
    };

    const taskType = reportType === 'portfolio' ? 'portfolio_report' : reportType === 'closeout' ? 'closeout_report' : 'status_report';
    const content  = await aurora(taskType, prompts[reportType] || prompts.status, contextBlock);
    res.json({ content, reportType, generatedAt: new Date().toISOString() });
  } catch (e) {
    if (e.message === 'MONTHLY_CAP_REACHED') return res.status(429).json({ error: 'Monthly cap reached' });
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

app.post('/api/reminders/check', async (req, res) => {
  try { await checkDueDateReminders(); res.json({ success: true }); }
  catch (e) { res.status(500).json({ error: e.message }); }
});

// ── Start ─────────────────────────────────────────────────────────────────────
async function start() {
  await db.initDB();
  await db.ensureSpendConstraint();
  app.listen(PORT, () => {
    console.log(`\n╔══════════════════════════════════════╗`);
    console.log(`║  Aurora R2S v3 — port ${PORT}          ║`);
    console.log(`╚══════════════════════════════════════╝`);
    console.log(`DB: ${process.env.DATABASE_URL ? 'PostgreSQL' : 'JSON files'}`);
    console.log(`Spend cap: $${CAP_USD} USD/month`);
    console.log(`Internal emails: ${INTERNAL_EMAILS.join(', ')}`);
    console.log(`Reminders: 14, 7, 3 days before due date\n`);
  });
}

start().catch(console.error);
