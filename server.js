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
  chat:              'claude-haiku-4-5-20251001', // Haiku for chat to save credits
  document_analysis: 'claude-haiku-4-5-20251001',
  contract_extract:  'claude-sonnet-4-6', // Sonnet for accuracy with tables and financials
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

// ── Read consultant reply emails from info@ inbox ─────────────────────────────
async function readConsultantReplies() {
  const token = await getOutlookToken();
  if (!token) { console.log('[Poll] No Outlook token — skipping'); return; }
  const mailbox = process.env.OUTLOOK_SHARED_MAILBOX || 'info@risk2solution.com';

  try {
    // Get unread emails from last 24 hours (polling catches them quickly)
    const since = new Date(Date.now() - 24 * 60 * 60 * 1000).toISOString();

    // Check both inbox and sent items
    const folders = ['inbox', 'sentitems'];
    let allMessages = [];

    for (const folder of folders) {
      try {
        const res = await axios.get(
          `https://graph.microsoft.com/v1.0/users/${mailbox}/mailFolders/${folder}/messages?$filter=receivedDateTime ge '${since}' and isRead eq false&$select=id,subject,from,toRecipients,body,receivedDateTime&$top=25`,
          { headers: { Authorization: `Bearer ${token}` }, timeout: 10000 }
        );
        const msgs = (res.data?.value || []).map(m => ({ ...m, folder }));
        allMessages = allMessages.concat(msgs);
      } catch(folderErr) {
        console.error(`[Poll] Error reading ${folder}:`, folderErr.message);
      }
    }

    if (!allMessages.length) { console.log('[Poll] No new emails'); return; }
    console.log(`[Poll] Processing ${allMessages.length} new email(s)`);

    const projects = await db.getProjects();
    // Accept emails from any R2S staff, consultants, or forwarded from R2S addresses
    const isFromConsultant = (email) => email && (
      email.toLowerCase().endsWith('@risk2solution.com') ||
      email.toLowerCase().endsWith('@presilience.com')
    );

    for (const msg of allMessages) {
      const fromEmail = msg.from?.emailAddress?.address || '';
      if (!isFromConsultant(fromEmail)) continue;

      const subject = msg.subject || '';
      const bodyText = msg.body?.content?.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').slice(0, 3000) || '';

      // Find matching project by subject or body content
      const bodyLower = bodyText.toLowerCase();
      const subjectLower = subject.toLowerCase();
      const matchedProject = projects.find(p =>
        p.type === 'standard' && p.clientName && (
          subjectLower.includes(p.clientName.toLowerCase()) ||
          subjectLower.includes((p.projectName || '').toLowerCase().slice(0, 12)) ||
          bodyLower.includes(p.clientName.toLowerCase()) ||
          bodyLower.includes((p.projectName || '').toLowerCase().slice(0, 12))
        )
      );

      // If no specific project match, send general update to Diane
      if (!matchedProject) {
        // Still notify Diane of any R2S staff email with project content
        if (bodyText.length > 100) {
          await sendEmail('diane.k@risk2solution.com',
            `[Aurora] R2S staff update received: ${subject}`,
            `Aurora received an email from ${fromEmail} that may contain project updates.

Subject: ${subject}

Summary:
${bodyText.slice(0, 800)}

Please review and update Aurora project records as needed.

Aurora
R2S Project Management Intelligence`,
            true
          );
        }
        // Mark as read and continue
        try {
          await axios.patch(
            `https://graph.microsoft.com/v1.0/users/${mailbox}/messages/${msg.id}`,
            { isRead: true },
            { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }, timeout: 5000 }
          );
        } catch(e) {}
        continue;
      }

      // Ask Aurora to analyse the reply and determine if phase should change
      try {
        const analysis = await aurora('status_email',
          `You are analysing a reply email from a consultant/trainer to determine if it indicates a project phase change or status update.

Project: ${matchedProject.projectName || matchedProject.clientName}
Current phase: ${PHASES[matchedProject.phase || 0]}
Current status: ${matchedProject.status}

Email from: ${fromEmail}
Subject: ${subject}
Content: ${bodyText.slice(0, 1500)}

Respond in JSON only with exactly this structure:
{
  "phaseChange": true or false,
  "newPhase": 0-4 (only if phaseChange is true, use: 0=Kick-off, 1=Deployment, 2=Monitoring & Review, 3=Reporting, 4=Close-out),
  "statusUpdate": "brief summary of what the consultant reported (1-2 sentences)",
  "requiresAttention": true or false,
  "attentionReason": "reason if requiresAttention is true, else empty string"
}`,
          buildContext(matchedProject)
        );

        let parsed;
        try { parsed = JSON.parse(analysis.replace(/\`\`\`json|\`\`\`/g, '').trim()); } catch { continue; }

        // Auto-update phase if clearly indicated
        if (parsed.phaseChange && typeof parsed.newPhase === 'number' && parsed.newPhase !== matchedProject.phase) {
          await db.updateProjectField(matchedProject.id, { phase: parsed.newPhase });
          console.log(`[Replies] Phase updated for ${matchedProject.clientName}: ${PHASES[matchedProject.phase]} → ${PHASES[parsed.newPhase]}`);
        }

        // Always notify Diane of the reply and any updates
        await sendEmail('diane.k@risk2solution.com',
          `[Aurora] Consultant reply received: ${matchedProject.clientName}`,
          `Aurora has received a reply from ${fromEmail} regarding the ${matchedProject.projectName || matchedProject.clientName} project.

Status update: ${parsed.statusUpdate}

${parsed.phaseChange ? `Phase automatically updated to: ${PHASES[parsed.newPhase]}

` : ''}${parsed.requiresAttention ? `ACTION REQUIRED: ${parsed.attentionReason}

` : ''}Please log into Aurora to review and update the project record if needed.

Aurora
R2S Project Management Intelligence`,
          true
        );

        // Mark email as read
        await axios.patch(
          `https://graph.microsoft.com/v1.0/users/${mailbox}/messages/${msg.id}`,
          { isRead: true },
          { headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' }, timeout: 5000 }
        );

      } catch (err) {
        if (err.message === 'MONTHLY_CAP_REACHED') break;
        console.error('[Replies] Analysis failed:', err.message);
      }
    }
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
  const standard = projects.filter(p =>
    p.type === 'standard' && !['Completed','Terminated','Closed'].includes(p.status)
  );
  const suggestions = [];

  for (const p of standard) {
    const phase = p.phase || 0;
    const days  = p.dueDate ? Math.round((new Date(p.dueDate) - new Date()) / (1000*60*60*24)) : null;
    const updatedDaysAgo = p.updatedAt ? Math.round((new Date() - new Date(p.updatedAt)) / (1000*60*60*24)) : 999;

    // ── Phase 0 (Kick-off) ────────────────────────────────────────────────────
    // If kick-off has been going on for 5+ days with a consultant assigned, suggest moving to Deployment
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
    // If deployment has been going for 45+ days, suggest moving to Monitoring
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
        action: { phase: 4, status: 'Completed' },
        confirmLabel: 'Mark as Complete & Close out',
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
  if (!project) return;

  if (suggestion.action?.phase !== undefined) {
    await db.updateProjectField(suggestion.projectId, {
      phase: suggestion.action.phase,
      ...(suggestion.action.status ? { status: suggestion.action.status } : {}),
    });

    // If moving to Close-out, trigger close-out actions
    if (suggestion.action.phase === 4 || suggestion.action.status === 'Completed') {
      // Draft client feedback email
      try { await draftClientFeedback(project); } catch(e) { console.error('[Feedback]', e.message); }
      const yr1 = new Date(); yr1.setFullYear(yr1.getFullYear() + 1);
      const yr2 = new Date(); yr2.setFullYear(yr2.getFullYear() + 2);
      await createCalendarReminder(
        `1-year follow-up: ${project.clientName}`,
        `Check in with ${project.clientName} — explore new opportunities for R2S.`,
        yr1.toISOString().slice(0,10)
      );
      await createCalendarReminder(
        `2-year follow-up: ${project.clientName}`,
        `2-year relationship check-in with ${project.clientName}.`,
        yr2.toISOString().slice(0,10)
      );
      await sendInternalEmail(
        `[Aurora] Project closed: ${project.clientName}`,
        `The ${project.projectName||project.clientName} project has been marked complete.

Aurora has set 1-year and 2-year follow-up reminders in the Outlook calendar.

Aurora
R2S Project Management Intelligence`
      );
    }

    console.log(`[Suggestions] Phase updated for ${project.clientName}: ${PHASES[project.phase||0]} → ${PHASES[suggestion.action.phase]}`);
  }

  // If suggestion was to draft an escalation
  if (suggestion.action?.draftEscalation) {
    await draftRiskEscalation(project, suggestion.action.riskDescription || 'Risk identified by Aurora');
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

// ── Daily batch (6am AEST = 8pm UTC) ─────────────────────────────────────────
async function runBatch() {
  const now = new Date();
  const dayOfWeek = now.getDay(); // 0=Sun, 1=Mon ... 5=Fri, 6=Sat
  const isWeekday = dayOfWeek >= 1 && dayOfWeek <= 5;
  const isMonday  = dayOfWeek === 1;

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
    // 5a. Weekly client status email drafts
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

  // ── 6. At-risk project summary ────────────────────────────────────────────
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

// Email polling: every hour from 7am to 7pm AEST (9pm to 9am UTC)
// AEST is UTC+10, so 7am AEST = 9pm UTC previous day, 7pm AEST = 9am UTC
// Run every hour: 21,22,23,0,1,2,3,4,5,6,7,8,9 UTC = 7am-7pm AEST
cron.schedule('0 21-23,0-9 * * *', async () => {
  const now = new Date();
  const dayOfWeek = now.getDay();
  const isWeekday = dayOfWeek >= 1 && dayOfWeek <= 5;
  if (!isWeekday) return; // weekdays only
  console.log('[Poll] Checking inbox...');
  try { await readConsultantReplies(); } catch(e) { console.error('[Poll] Error:', e.message); }
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
        // Draft client feedback email
        try { await draftClientFeedback(p); } catch(e) { console.error('[Feedback]', e.message); }
      }
    }
    res.json({ project: updated });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.delete('/api/projects/:id', async (req, res) => {
  try { await db.deleteProject(req.params.id); res.json({ success: true }); }
  catch (e) { res.status(500).json({ error: e.message }); }
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
    }

    res.json({ project, extracted, briefingPrepared: false, suggestedConsultants: extracted.consultant ? extracted.consultant.split(/[,;&]+/).map(s => s.trim()).filter(Boolean) : [] });
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
    const cc = draft.ccEmail ? [draft.ccEmail] : [];
    await sendEmail(draft.toEmail || INTERNAL_EMAILS[1], draft.subject, draft.body, false, cc);
    await db.updateDraft(req.params.id, { approved: true });

    // If this was a consultant briefing, send kick-off meeting prompt to Diane
    if (draft.type === 'consultant_briefing' && draft.projectId) {
      const project = await db.getProject(draft.projectId);
      if (project) {
        await sendKickoffPrompt(project);
        console.log(`[Approve] Kick-off prompt sent for ${project.clientName}`);
      }
    }

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
    const allSuggestions = await db.getSuggestions();
    const allWithDismissed = await (async () => { try { return JSON.parse(require('fs').readFileSync(require('path').join(db.DATA,'suggestions.json'),'utf8')); } catch { return []; } })();
    const suggestions = allWithDismissed;
    const suggestion = suggestions.find(s => s.id === req.params.id);
    if (!suggestion) return res.status(404).json({ error: 'Not found' });
    await applySuggestion(suggestion);
    await db.updateSuggestion(req.params.id, 'approved');
    res.json({ success: true });
  } catch (e) {
    if (e.message === 'MONTHLY_CAP_REACHED') return res.status(429).json({ error: 'Monthly cap reached' });
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/suggestions/:id/dismiss', async (req, res) => {
  try {
    await db.updateSuggestion(req.params.id, 'dismissed');
    res.json({ success: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
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

app.post('/api/emails/read', async (req, res) => {
  try {
    await readConsultantReplies();
    res.json({ success: true, message: 'Inbox checked and processed' });
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
