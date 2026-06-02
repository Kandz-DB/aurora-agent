const express = require("express");
const session = require("express-session");
const Anthropic = require("@anthropic-ai/sdk");
const path = require("path");
const fs = require("fs");
const https = require("https");
const multer = require("multer");
const { BlobServiceClient } = require("@azure/storage-blob");
const { generateProposal } = require("./generateProposal");

const app = express();
const PORT = process.env.PORT || 3000;
const PASSWORD = process.env.APP_PASSWORD || "Ariel1!";
const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY;
const TEMPLATE_URL = process.env.TEMPLATE_URL || "https://www.dropbox.com/scl/fi/xs9x7fuhxbby0wjmyly79/Presentation1_Ariel_Template.pptx?rlkey=y316ja6bp7ri1nbfyqhf1wdfm&st=46qije4i&dl=1";
const TEMPLATE_PATH = "/tmp/r2s_template.pptx";
const AZURE_CLIENT_ID = process.env.AZURE_CLIENT_ID;
const AZURE_TENANT_ID = process.env.AZURE_TENANT_ID;
const AZURE_CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;
const SHARED_MAILBOX = process.env.SHARED_MAILBOX || "eps@risk2solution.com";
const GARIMA_EMAIL = process.env.GARIMA_EMAIL || "garima.a@risk2solution.com";
const REINETTE_NAME = process.env.REINETTE_NAME || "Reinette Kruger";

// ── BLOB STORAGE ───────────────────────────────────────────────────────────────
const STORAGE_CONNECTION_STRING = process.env.AZURE_STORAGE_CONNECTION_STRING;
const BLOB_CONTAINER = "ariel-data";
const ACTIVITY_BLOB = "activityLog.json";
const QUEUE_BLOB = "reviewQueue.json";

let containerClient = null;
if (STORAGE_CONNECTION_STRING) {
  const blobServiceClient = BlobServiceClient.fromConnectionString(STORAGE_CONNECTION_STRING);
  containerClient = blobServiceClient.getContainerClient(BLOB_CONTAINER);
  console.log("Blob storage connected.");
} else {
  console.log("No blob storage configured — using in-memory only.");
}

async function loadFromBlob(blobName, fallback) {
  if (!containerClient) return fallback;
  try {
    const blob = containerClient.getBlobClient(blobName);
    const exists = await blob.exists();
    if (!exists) return fallback;
    const buf = await blob.downloadToBuffer();
    return JSON.parse(buf.toString());
  } catch (err) {
    console.error("Blob load error (" + blobName + "):", err.message);
    return fallback;
  }
}

async function saveToBlob(blobName, data) {
  if (!containerClient) return;
  try {
    const blob = containerClient.getBlockBlobClient(blobName);
    const content = JSON.stringify(data);
    await blob.upload(content, Buffer.byteLength(content), {
      blobHTTPHeaders: { blobContentType: "application/json" }
    });
  } catch (err) {
    console.error("Blob save error (" + blobName + "):", err.message);
  }
}

// ── OPTIMISATION: prompt caching beta header on all calls ──
const client = new Anthropic({
  apiKey: ANTHROPIC_API_KEY,
  defaultHeaders: { "anthropic-beta": "prompt-caching-2024-07-31" }
});

let activityLog = [];
let reviewQueue = [];
let graphToken = null;
let graphTokenExpiry = null;

// ── OPTIMISATION: conversation history trim ──
const MAX_CHAT_MESSAGES = 20;

const upload = multer({
  dest: "/tmp/uploads/",
  fileFilter: (req, file, cb) => {
    const allowed = [".xlsx", ".xls", ".csv"];
    cb(null, allowed.includes(path.extname(file.originalname).toLowerCase()));
  },
  limits: { fileSize: 10 * 1024 * 1024 }
});

app.use(express.json({ limit: "50mb" }));
app.use(express.urlencoded({ extended: true }));
app.set("trust proxy", 1);
app.use(session({
  secret: process.env.SESSION_SECRET || "ariel-r2s-secret-2026",
  resave: false, saveUninitialized: false,
  cookie: { secure: process.env.NODE_ENV === "production", maxAge: 24 * 60 * 60 * 1000 }
}));
app.use(express.static("public"));

// ── TEMPLATE ──────────────────────────────────────────────────────────────────
function downloadTemplate() {
  return new Promise((resolve, reject) => {
    if (fs.existsSync(TEMPLATE_PATH)) { console.log("Template cached."); return resolve(); }
    console.log("Downloading template...");
    const file = fs.createWriteStream(TEMPLATE_PATH);
    const tryDownload = (url, n = 0) => {
      if (n > 10) return reject(new Error("Too many redirects"));
      const p = new URL(url);
      https.get({ hostname: p.hostname, path: p.pathname + p.search, headers: { "User-Agent": "ariel-r2s", "Accept": "application/octet-stream" } }, (res) => {
        if ([301,302,307,308].includes(res.statusCode)) return tryDownload(res.headers.location, n+1);
        if (res.statusCode !== 200) return reject(new Error("Failed: " + res.statusCode));
        res.pipe(file);
        file.on("finish", () => { file.close(); console.log("Template downloaded."); resolve(); });
        file.on("error", (err) => { fs.unlink(TEMPLATE_PATH, () => {}); reject(err); });
      }).on("error", (err) => { fs.unlink(TEMPLATE_PATH, () => {}); reject(err); });
    };
    tryDownload(TEMPLATE_URL);
  });
}

// ── GRAPH API ─────────────────────────────────────────────────────────────────
async function getGraphToken() {
  if (graphToken && graphTokenExpiry && Date.now() < graphTokenExpiry - 60000) return graphToken;
  if (!AZURE_CLIENT_ID) throw new Error("Azure not configured");
  const res = await fetch("https://login.microsoftonline.com/" + AZURE_TENANT_ID + "/oauth2/v2.0/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ client_id: AZURE_CLIENT_ID, client_secret: AZURE_CLIENT_SECRET, scope: "https://graph.microsoft.com/.default", grant_type: "client_credentials" }).toString()
  });
  const data = await res.json();
  if (!data.access_token) throw new Error("Token failed: " + JSON.stringify(data));
  graphToken = data.access_token;
  graphTokenExpiry = Date.now() + data.expires_in * 1000;
  console.log("Graph token acquired.");
  return graphToken;
}

async function graphRequest(method, endpoint, body, token) {
  const t = token || await getGraphToken();
  const url = endpoint.startsWith("https") ? endpoint : "https://graph.microsoft.com/v1.0" + endpoint;
  const opts = { method, headers: { "Authorization": "Bearer " + t, "Content-Type": "application/json" } };
  if (body) opts.body = JSON.stringify(body);
  const res = await fetch(url, opts);
  if (res.status === 204) return null;
  const text = await res.text();
  try { return JSON.parse(text); } catch { return text; }
}

async function attachLargeFile(token, messageId, fileName, filePath) {
  const buf = fs.readFileSync(filePath);
  const size = buf.length;
  console.log("Attaching " + fileName + " (" + Math.round(size/1024) + "KB)...");
  if (size < 3 * 1024 * 1024) {
    await graphRequest("POST", "/users/" + SHARED_MAILBOX + "/messages/" + messageId + "/attachments",
      { "@odata.type": "#microsoft.graph.fileAttachment", name: fileName, contentBytes: buf.toString("base64") }, token);
  } else {
    const sess = await graphRequest("POST", "/users/" + SHARED_MAILBOX + "/messages/" + messageId + "/attachments/createUploadSession",
      { AttachmentItem: { attachmentType: "file", name: fileName, size } }, token);
    if (!sess || !sess.uploadUrl) throw new Error("No upload session");
    const chunk = 4 * 1024 * 1024;
    for (let offset = 0; offset < size; offset += chunk) {
      const end = Math.min(offset + chunk - 1, size - 1);
      const slice = buf.slice(offset, end + 1);
      await fetch(sess.uploadUrl, { method: "PUT", headers: { "Content-Range": "bytes " + offset + "-" + end + "/" + size, "Content-Length": slice.length.toString(), "Content-Type": "application/octet-stream" }, body: slice });
    }
  }
  console.log("Attachment complete.");
}

// ── OPTIMISATION: business hours check (Brisbane UTC+10) ──
function isBrisbaneBusinessHours() {
  const hour = parseInt(
    new Date().toLocaleString("en-AU", {
      timeZone: "Australia/Brisbane",
      hour: "numeric",
      hour12: false
    })
  );
  return hour >= 7 && hour <= 18;
}

// ── GARIMA EXCEL ATTACHMENT PROCESSING ───────────────────────────────────────
async function fetchExcelAttachment(emailId, token) {
  try {
    const result = await graphRequest("GET",
      "/users/" + SHARED_MAILBOX + "/messages/" + emailId + "/attachments",
      null, token);
    if (!result || !result.value || !result.value.length) return null;
    const excelExts = [".xlsx", ".xls", ".csv"];
    const attachment = result.value.find(a => {
      const name = (a.name || "").toLowerCase();
      return excelExts.some(ext => name.endsWith(ext));
    });
    if (!attachment) return null;
    console.log("Found Excel attachment: " + attachment.name);
    // Fetch full attachment with content
    const full = await graphRequest("GET",
      "/users/" + SHARED_MAILBOX + "/messages/" + emailId + "/attachments/" + attachment.id,
      null, token);
    if (!full || !full.contentBytes) return null;
    const buf = Buffer.from(full.contentBytes, "base64");
    const tmpPath = "/tmp/garima_attach_" + Date.now() + "_" + attachment.name;
    fs.writeFileSync(tmpPath, buf);
    return { path: tmpPath, name: attachment.name };
  } catch (err) {
    console.error("Attachment fetch error:", err.message);
    return null;
  }
}

async function parseExcelBuffer(filePath, fileName) {
  try {
    const XLSX = require("xlsx");
    let textContent = "";
    if (fileName.toLowerCase().endsWith(".csv")) {
      textContent = fs.readFileSync(filePath, "utf8");
    } else {
      const wb = XLSX.readFile(filePath);
      wb.SheetNames.forEach(n => {
        const csv = XLSX.utils.sheet_to_csv(wb.Sheets[n]);
        if (csv.trim().replace(/,/g, "").trim()) textContent += "Sheet: " + n + "\n" + csv + "\n\n";
      });
    }
    return textContent.substring(0, 8000);
  } catch (err) {
    console.error("Excel parse error:", err.message);
    return null;
  }
}

async function extractClientFromEmailBody(bodyText) {
  const res = await client.messages.create({
    model: "claude-haiku-4-5-20251001",
    max_tokens: 600,
    messages: [{
      role: "user",
      content: "Extract the CLIENT contact details from this forwarded email thread. The client is the EXTERNAL person (not from risk2solution.com) who originally sent the enquiry.\n\nReturn ONLY valid JSON:\n{\"client_name\":\"\",\"contact_name\":\"\",\"contact_position\":\"\",\"contact_email\":\"\",\"summary\":\"one sentence about what they need\"}\n\nEmail thread:\n" + bodyText.substring(0, 4000)
    }]
  });
  let raw = res.content[0].text.trim().replace(/^```json\s*/i,"").replace(/^```\s*/i,"").replace(/```\s*$/g,"").trim();
  try { return JSON.parse(raw); } catch { return { client_name: "", contact_name: "", contact_position: "", contact_email: "", summary: "" }; }
}

async function processGarimaAttachment(email, token) {
  console.log("Garima email detected — checking for Excel attachment...");
  const attachment = await fetchExcelAttachment(email.id, token);
  if (!attachment) { console.log("No Excel attachment found — falling through to normal processing."); return false; }

  try {
    const bodyText = email.body && email.body.content
      ? email.body.content.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim()
      : email.bodyPreview || "";

    // Extract client details from the email thread
    const clientInfo = await extractClientFromEmailBody(bodyText);
    console.log("Extracted client:", JSON.stringify(clientInfo));

    // Parse the Excel attachment
    const excelText = await parseExcelBuffer(attachment.path, attachment.name);
    fs.unlink(attachment.path, () => {});
    if (!excelText || !excelText.trim()) { console.log("Excel appears empty — falling through."); return false; }

    // Extract drawing hours from Excel
    const xlRes = await client.messages.create({
      model: "claude-haiku-4-5-20251001",
      max_tokens: 600,
      messages: [{
        role: "user",
        content: "Extract drawing hours from this R2S quote spreadsheet. Return ONLY valid JSON:\n{\"drawing_hours\":0,\"notes\":\"\"}\nSpreadsheet:\n" + excelText
      }]
    });
    let xlRaw = xlRes.content[0].text.trim().replace(/^```json\s*/i,"").replace(/^```\s*/i,"").replace(/```\s*$/g,"").trim();
    let xlData;
    try { xlData = JSON.parse(xlRaw); } catch { console.log("Could not parse Excel data — falling through."); return false; }

    if (!xlData.drawing_hours || xlData.drawing_hours <= 0) {
      // Try to extract hours from email body if Excel didn't have them
      const hoursMatch = bodyText.match(/(\d+(?:\.\d+)?)\s*hours?\s*(?:of\s*work|drawing|quoted)?/i);
      if (hoursMatch) { xlData.drawing_hours = parseFloat(hoursMatch[1]); console.log("Hours from email body: " + xlData.drawing_hours); }
    }

    if (!xlData.drawing_hours || xlData.drawing_hours <= 0) { console.log("No drawing hours found — falling through."); return false; }

    // Apply R2S pricing rules: add 2 admin hours, rate based on quoted hours
    const quoted = xlData.drawing_hours;
    const rate = quoted <= 20 ? 180 : 150;
    const totalHours = quoted + 2;
    const totalCost = parseFloat((totalHours * rate).toFixed(2));
    const diagramsPrice = "$" + totalCost.toLocaleString("en-AU", { minimumFractionDigits: 2 });
    const diagramsPriceGst = "$" + (totalCost * 1.1).toLocaleString("en-AU", { minimumFractionDigits: 2 });
    console.log("Diagrams pricing: " + quoted + " quoted + 2 admin = " + totalHours + "hrs @ $" + rate + "/hr = " + diagramsPrice);

    const today = new Date().toLocaleDateString("en-AU", { day: "numeric", month: "long", year: "numeric" });
    const proposalData = {
      client_name: clientInfo.client_name || "Client",
      contact_name: clientInfo.contact_name || "",
      contact_position: clientInfo.contact_position || "",
      contact_email: clientInfo.contact_email || "",
      proposal_title: "Emergency Management Services",
      contract_type: "standard",
      date: today,
      selected_service_ids: ["diagrams"],
      diagrams_price: diagramsPrice,
      online_warden_learners: 0,
      online_warden_price: "",
      equipment: [],
      equipment_subtotal: 0,
      equipment_total_with_shipping: 0,
      total_standard: diagramsPrice,
      total_three_year: diagramsPrice,
      total_five_year: diagramsPrice,
      total_inc_gst_standard: diagramsPriceGst,
      total_inc_gst_three_year: diagramsPriceGst,
      total_inc_gst_five_year: diagramsPriceGst,
      email_body: [
        "Dear " + (clientInfo.contact_name || "there") + ",",
        "",
        "Thank you for your enquiry. Please find attached our proposal for the provision of evacuation diagrams.",
        "",
        "We look forward to working with you.",
        "",
        "Kind regards,",
        "",
        "Reinette Kruger",
        "Administration Officer",
        "Risk 2 Solution",
        "eps@risk2solution.com"
      ].join("\n")
    };

    if (!fs.existsSync(TEMPLATE_PATH)) await downloadTemplate();
    const outputPath = "/tmp/proposal_" + Date.now() + ".pptx";
    const pythonCmd = (() => { try { require("child_process").execSync("python3 --version"); return "python3"; } catch { return "python"; } })();
    const spawnResult = require("child_process").spawnSync(pythonCmd, ["generate_proposal.py", JSON.stringify(proposalData), TEMPLATE_PATH, outputPath], { timeout: 60000, cwd: __dirname, encoding: "utf8" });
    if (spawnResult.status !== 0) throw new Error(spawnResult.stderr || "Python script failed");
    if (!fs.existsSync(outputPath)) throw new Error("PPTX not created");

    const clientSlug = (proposalData.client_name || "Proposal").replace(/\s+/g, "");
    const dateSlug = new Date().toLocaleDateString("en-AU", { month: "short", year: "numeric" }).replace(" ", "");
    const fileName = clientSlug + "_Proposal_" + dateSlug + ".pptx";

    // Create draft reply to the CLIENT (not Garima)
    const clientEmail = clientInfo.contact_email;
    let draftCreated = false;
    if (clientEmail) {
      try {
        const replyDraft = await graphRequest("POST", "/users/" + SHARED_MAILBOX + "/messages", {
          subject: "Proposal — " + (clientInfo.client_name || "Evacuation Diagrams"),
          toRecipients: [{ emailAddress: { address: clientEmail, name: clientInfo.contact_name || "" } }],
          body: { contentType: "HTML", content: "<p>" + proposalData.email_body.replace(/\n/g, "<br>") + "</p>" },
          isDraft: true
        }, token);
        if (replyDraft && replyDraft.id) {
          await attachLargeFile(token, replyDraft.id, fileName, outputPath);
          const queueEntry = {
            id: Date.now(), timestamp: new Date().toISOString(), email_id: email.id,
            subject: email.subject,
            from_name: email.from && email.from.emailAddress ? email.from.emailAddress.name : "",
            from_email: email.from && email.from.emailAddress ? email.from.emailAddress.address : "",
            body_preview: email.bodyPreview || "", web_link: email.webLink || "",
            received_at: email.receivedDateTime || null,
            classification: ["diagrams"], summary: "Garima diagrams quote — " + quoted + " hrs → " + diagramsPrice + " excl. GST. " + (clientInfo.summary || ""),
            client_name: proposalData.client_name, contact_name: proposalData.contact_name,
            contact_position: proposalData.contact_position, contact_email: proposalData.contact_email,
            services_mentioned: ["diagrams"], status: "pending",
            proposal_file: outputPath, proposal_filename: fileName, proposal_data: proposalData,
            drafts: [{ type: "client_reply", draft_id: replyDraft.id, to: clientEmail, subject: "Proposal — " + (clientInfo.client_name || "Evacuation Diagrams") }]
          };
          reviewQueue.push(queueEntry);
          await saveToBlob(QUEUE_BLOB, reviewQueue);
          activityLog.push({
            id: Date.now(), timestamp: new Date().toISOString(),
            client_name: proposalData.client_name, contact_name: proposalData.contact_name,
            contact_position: proposalData.contact_position, contact_email: proposalData.contact_email,
            contract_type: "standard", proposal_title: proposalData.proposal_title,
            services: ["diagrams"], total_standard: diagramsPrice,
            total_three_year: diagramsPrice, total_five_year: diagramsPrice,
            total_inc_gst: diagramsPriceGst, email_body: proposalData.email_body,
            filename: fileName, file_path: outputPath, source: "email_garima"
          });
          await saveToBlob(ACTIVITY_BLOB, activityLog);
          draftCreated = true;
          console.log("Garima attachment: proposal generated and draft created for " + clientEmail);
        }
      } catch (err) { console.error("Draft creation error:", err.message); }
    }

    if (!draftCreated) {
      // No client email found — add to queue without draft so Reinette can review
      const queueEntry = {
        id: Date.now(), timestamp: new Date().toISOString(), email_id: email.id,
        subject: email.subject,
        from_name: email.from && email.from.emailAddress ? email.from.emailAddress.name : "",
        from_email: email.from && email.from.emailAddress ? email.from.emailAddress.address : "",
        body_preview: email.bodyPreview || "", web_link: email.webLink || "",
        received_at: email.receivedDateTime || null,
        classification: ["diagrams"], summary: "Garima diagrams quote — " + quoted + " hrs → " + diagramsPrice + " excl. GST. Could not find client email — please send manually.",
        client_name: proposalData.client_name, contact_name: proposalData.contact_name,
        contact_position: proposalData.contact_position, contact_email: "",
        services_mentioned: ["diagrams"], status: "pending",
        proposal_file: outputPath, proposal_filename: fileName, proposal_data: proposalData,
        drafts: []
      };
      reviewQueue.push(queueEntry);
      await saveToBlob(QUEUE_BLOB, reviewQueue);
      console.log("Garima attachment: proposal generated but no client email found — added to queue.");
    }

    return true; // handled — skip normal processing
  } catch (err) {
    console.error("Garima attachment processing error:", err.message);
    fs.unlink(attachment.path, () => {});
    return false; // fall through to normal processing
  }
}

// ── EMAIL POLLING ─────────────────────────────────────────────────────────────
async function pollInbox() {
  if (!AZURE_CLIENT_ID) { console.log("Graph not configured."); return; }
  console.log("Polling inbox...");
  try {
    const token = await getGraphToken();
    const result = await graphRequest("GET",
      "/users/" + SHARED_MAILBOX + "/mailFolders/inbox/messages?$top=20&$orderby=receivedDateTime desc&$select=id,subject,from,bodyPreview,body,receivedDateTime,categories,webLink",
      null, token);
    if (result && result.error) { console.log("Graph API error:", JSON.stringify(result.error)); return; }
    if (!result || !result.value) { console.log("No result from Graph API."); return; }
    console.log("Inbox has " + result.value.length + " email(s).");
    const unprocessed = result.value.filter(e => !e.categories || !e.categories.includes("Ariel Processed"));
    console.log("Found " + unprocessed.length + " unprocessed email(s).");
    for (const email of unprocessed) {
      try {
        await processEmail(email, token);
        await graphRequest("PATCH", "/users/" + SHARED_MAILBOX + "/messages/" + email.id,
          { categories: [...(email.categories || []), "Ariel Processed"] }, token);
        console.log("Tagged: " + email.subject);
      } catch (err) { console.error("Error processing:", email.subject, err.message); }
    }
  } catch (err) { console.error("Poll error:", err.message); }
}

async function processEmail(email, token) {
  console.log("Processing: \"" + email.subject + "\" from " + (email.from && email.from.emailAddress ? email.from.emailAddress.address : "unknown"));

  const senderEmail = (email.from && email.from.emailAddress ? email.from.emailAddress.address : "").toLowerCase();
  const senderName = (email.from && email.from.emailAddress ? email.from.emailAddress.name : "").toLowerCase();
  if (senderEmail.includes("monday.com") || senderName.includes("monday.com")) { console.log("Skipping monday.com."); return; }

  const odataType = (email["@odata.type"] || "").toLowerCase();
  if (odataType.includes("eventmessage") || odataType.includes("meeting")) { console.log("Skipping calendar invite."); return; }

  const subjectLower = (email.subject || "").toLowerCase();
  if (["[in-person]","[online]","[teams meeting]","[teams]","[zoom]"].some(p => subjectLower.includes(p))) { console.log("Skipping calendar pattern."); return; }

  // ── GARIMA ATTACHMENT SHORTCUT ─────────────────────────────────────────────
  const isFromGarima = senderEmail === GARIMA_EMAIL.toLowerCase();
  if (isFromGarima) {
    const handled = await processGarimaAttachment(email, token);
    if (handled) return; // proposal generated and queued — skip normal flow
  }

  const bodyText = email.body && email.body.content
    ? email.body.content.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim()
    : email.bodyPreview || "";

  const emailText = "From: " + (email.from && email.from.emailAddress ? email.from.emailAddress.name + " <" + email.from.emailAddress.address + ">" : "unknown") +
    "\nSubject: " + (email.subject || "") +
    "\nBody: " + bodyText.substring(0, 3000);

  const bodyLower = (bodyText + " " + (email.bodyPreview || "")).toLowerCase();
  const explicitPhrases = ["provide a quote","please quote","send a quote","get a quote","can i get a quote","i get a quote","quote for","quote on","provide a proposal","send a proposal","request a quote","how much","what do you charge","what is the cost","what is your pricing","pricing for","cost for","rates for","fee for","interested in your services","looking for costs","looking for pricing","could you please quote","can you please quote","could you provide a quote","can you provide a quote","request for quote","request for proposal","rfq"];
  const hasExplicitPhrase = explicitPhrases.some(p => bodyLower.includes(p));

  const classifyRes = await client.messages.create({
    model: "claude-haiku-4-5-20251001",
    max_tokens: 800,
    system: [
      {
        type: "text",
        text: [
          "You are Ariel, email classifier for Risk 2 Solution (R2S).",
          "Return ONLY valid JSON: {\"is_quote_request\":true/false,\"request_types\":[],\"client_name\":\"\",\"contact_name\":\"\",\"contact_email\":\"\",\"services_mentioned\":[],\"summary\":\"\"}",
          "Set is_quote_request TRUE if asking for price/quote/proposal for any R2S service.",
          "request_types: training (warden/fire/exercise/chief warden/staff awareness/bushfire/school/armed offender/online warden), diagrams (ONLY physical evacuation diagrams/floor plans/base plans/warden maps to be drawn — NOT EMP/EP/CIMP which are document services), equipment (vests/hats/bags/backpacks/whistles/megaphones/radios/first aid kits).",
          "Set FALSE for: monday.com notifications, booking confirmations for already-agreed training, scheduling/delivery logistics, reports being completed, signed proposals, rescheduling, invoices, calendar invites.",
          "KEY TEST: Is this person asking what something will COST or for a PROPOSAL? Yes=TRUE, No=FALSE."
        ].join("\n"),
        cache_control: { type: "ephemeral" }
      }
    ],
    messages: [{ role: "user", content: "Classify this email:\n" + emailText }]
  });

  let rawText = classifyRes.content[0].text.trim().replace(/^```json\s*/i,"").replace(/^```\s*/i,"").replace(/```\s*$/g,"").trim();
  let classification;
  try { classification = JSON.parse(rawText); }
  catch (e) {
    classification = { is_quote_request: hasExplicitPhrase, request_types: hasExplicitPhrase ? ["training"] : [], summary: "Could not classify.", client_name: "", contact_name: "", contact_email: "" };
  }

  if (hasExplicitPhrase && !classification.is_quote_request) {
    console.log("Content override: explicit phrase found");
    classification.is_quote_request = true;
    if (!classification.request_types || !classification.request_types.length) {
      const hasDiagrams = bodyLower.includes("diagram") || bodyLower.includes("floor plan") || bodyLower.includes("base plan");
      classification.request_types = hasDiagrams ? ["diagrams"] : ["training"];
    }
  }

  console.log("Classification: is_quote=" + classification.is_quote_request + ", types=" + JSON.stringify(classification.request_types));
  console.log("Body available: " + (bodyText.length > 0) + ", explicit phrase: " + hasExplicitPhrase);

  if (!classification.is_quote_request) {
    reviewQueue.push({ id: Date.now(), timestamp: new Date().toISOString(), email_id: email.id, subject: email.subject, from_name: email.from && email.from.emailAddress ? email.from.emailAddress.name : "", from_email: email.from && email.from.emailAddress ? email.from.emailAddress.address : "", body_preview: email.bodyPreview || "", web_link: email.webLink || "", classification: "not_quote", summary: classification.summary || "Not a quote request", status: "pending", drafts: [] });
    await saveToBlob(QUEUE_BLOB, reviewQueue);
    return;
  }

  const queueEntry = {
    id: Date.now(), timestamp: new Date().toISOString(), email_id: email.id, subject: email.subject,
    from_name: email.from && email.from.emailAddress ? email.from.emailAddress.name : "",
    from_email: email.from && email.from.emailAddress ? email.from.emailAddress.address : "",
    body_preview: email.bodyPreview || "", web_link: email.webLink || "",
    classification: classification.request_types || [], summary: classification.summary || "",
    client_name: classification.client_name || "", contact_name: classification.contact_name || "",
    contact_email: classification.contact_email || (email.from && email.from.emailAddress ? email.from.emailAddress.address : ""),
    services_mentioned: classification.services_mentioned || [], status: "pending", drafts: [], proposal_file: null
  };

  const mentionsOnlineWarden = bodyLower.includes("online warden") ||
    bodyLower.includes("elearning") || bodyLower.includes("e-learning") ||
    (classification.services_mentioned && classification.services_mentioned.some(s => s.toLowerCase().includes("online")));

  const learnerMatch = bodyText.match(/\b(\d+)\s*(learners?|students?|participants?|people|staff|employees?|users?)\b/i);
  const learnerCount = learnerMatch ? parseInt(learnerMatch[1]) : 0;

  if (mentionsOnlineWarden && !learnerCount) {
    try {
      const replyDraft = await graphRequest("POST", "/users/" + SHARED_MAILBOX + "/messages/" + email.id + "/createReply", {}, token);
      if (replyDraft && replyDraft.id) {
        const replyBody = [
          "Dear " + (queueEntry.contact_name || "there") + ",",
          "",
          "Thank you for your enquiry regarding Online Warden Training (eLearning).",
          "",
          "Could you please advise how many learners will be completing the training? Once we have this, we will prepare your proposal promptly.",
          "",
          "Kind regards,",
          "",
          "Reinette Kruger",
          "Administration Officer",
          "Risk 2 Solution",
          "eps@risk2solution.com"
        ].join("\n");
        await graphRequest("PATCH", "/users/" + SHARED_MAILBOX + "/messages/" + replyDraft.id, {
          subject: "Re: " + email.subject,
          toRecipients: [{ emailAddress: { address: queueEntry.contact_email, name: queueEntry.contact_name } }],
          body: { contentType: "HTML", content: "<p>" + replyBody.replace(/\n/g, "<br>") + "</p>" }
        }, token);
        queueEntry.drafts.push({ type: "more_info_request", draft_id: replyDraft.id, to: queueEntry.contact_email, subject: "Re: " + email.subject });
        console.log("More info request draft created for online warden.");
      }
    } catch (err) { console.error("More info draft error:", err.message); }
    queueEntry.status = "more_info_required";
    queueEntry.more_info_note = "Learner count required to quote Online Warden Training. A draft email asking the client has been prepared and is ready to send from Outlook Drafts. Once their reply arrives with the learner count, Ariel will generate the proposal automatically.";
    reviewQueue.push(queueEntry);
    await saveToBlob(QUEUE_BLOB, reviewQueue);
    return;
  }

  if (mentionsOnlineWarden && learnerCount) {
    classification.online_warden_learners = learnerCount;
    queueEntry.online_warden_learners = learnerCount;
    console.log("Online warden learner count detected: " + learnerCount);
  }

  const isTrainingOrEquipment = classification.request_types && (classification.request_types.includes("training") || classification.request_types.includes("equipment"));
  if (isTrainingOrEquipment) {
    try {
      const result = await generateProposalFromEmail(classification, email);
      if (result) {
        queueEntry.proposal_file = result.filePath;
        queueEntry.proposal_filename = result.fileName;
        queueEntry.proposal_data = result.proposalData;
        const replyDraft = await graphRequest("POST", "/users/" + SHARED_MAILBOX + "/messages/" + email.id + "/createReply", {}, token);
        if (replyDraft && replyDraft.id) {
          await attachLargeFile(token, replyDraft.id, result.fileName, result.filePath);
          await graphRequest("PATCH", "/users/" + SHARED_MAILBOX + "/messages/" + replyDraft.id, {
            subject: "Re: " + email.subject,
            toRecipients: [{ emailAddress: { address: queueEntry.contact_email, name: queueEntry.contact_name } }],
            body: { contentType: "HTML", content: "<p>" + result.emailBody.replace(/\n/g, "<br>") + "</p>" }
          }, token);
          queueEntry.drafts.push({ type: "client_reply", draft_id: replyDraft.id, to: queueEntry.contact_email, subject: "Re: " + email.subject });
          console.log("Training/equipment draft created.");
        }
      }
    } catch (err) { console.error("Error creating draft:", err.message); queueEntry.drafts_error = err.message; }
  }

  if (classification.request_types && classification.request_types.includes("diagrams")) {
    try {
      const fwd = await graphRequest("POST", "/users/" + SHARED_MAILBOX + "/messages/" + email.id + "/createForward", {}, token);
      if (fwd && fwd.id) {
        const body = "Hi Garima,\n\nThe client below is requesting a quote for evacuation diagrams. Could you please provide an Excel quote for the drawing hours?\n\nClient: " + (queueEntry.client_name || queueEntry.from_name) + "\nContact: " + queueEntry.contact_name + " (" + queueEntry.contact_email + ")\n\nThanks,\n" + REINETTE_NAME + "\nRisk 2 Solution";
        await graphRequest("PATCH", "/users/" + SHARED_MAILBOX + "/messages/" + fwd.id, {
          toRecipients: [{ emailAddress: { address: GARIMA_EMAIL, name: "Garima" } }],
          subject: "FWD: Diagrams Quote Required — " + (queueEntry.client_name || queueEntry.from_name),
          body: { contentType: "HTML", content: "<p>" + body.replace(/\n/g, "<br>") + "</p>" }
        }, token);
        queueEntry.drafts.push({ type: "garima_forward", draft_id: fwd.id, to: GARIMA_EMAIL, subject: "FWD: Diagrams Quote Required — " + (queueEntry.client_name || queueEntry.from_name) });
        console.log("Diagrams forward created.");
      }
    } catch (err) { console.error("Diagrams draft error:", err.message); }
  }

  reviewQueue.push(queueEntry);
  await saveToBlob(QUEUE_BLOB, reviewQueue);
  console.log("Queue entry added. Drafts: " + queueEntry.drafts.length);
}

async function generateProposalFromEmail(classification, email) {
  if (!fs.existsSync(TEMPLATE_PATH)) await downloadTemplate();
  const emailText = email.body && email.body.content
    ? email.body.content.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim().substring(0, 3000)
    : email.bodyPreview || "";

  const proposalSystemPrompt = [
    "You are Ariel, the R2S proposal agent. Generate proposal data from an email enquiry.",
    "",
    "SERVICES AND PRICING:",
    "- Onsite Emergency Exercise (exercise): Standard $1,200/yr | 3yr $1,140/yr | 5yr $1,116/yr",
    "- Warden Training 1-2 hours (warden): Standard $1,500/yr | 3yr $1,425/yr | 5yr $1,395/yr",
    "- Chief Warden Training 1-2 hours (chief): Standard $1,500/yr | 3yr $1,425/yr | 5yr $1,395/yr",
    "- Fire Extinguisher Training 1-2 hours (fire): Standard $1,200/yr | 3yr $1,140/yr | 5yr $1,116/yr",
    "- School Critical Incident Management Training (school): Standard $1,800/yr | 3yr $1,710/yr | 5yr $1,674/yr",
    "- Active Armed Offender Training (armed): Standard $1,500/yr | 3yr $1,425/yr | 5yr $1,395/yr",
    "- Staff Awareness Training 1 hour (staff): Standard $750/yr | 3yr $712/yr | 5yr $697/yr",
    "- Bushfire Preparedness Training (bushfire): Standard $1,500/yr | 3yr $1,425/yr | 5yr $1,395/yr",
    "- Annual Assessment with feedback (assessment): $600 standard only",
    "- Emergency Management Plan (emp): $1,600 standard only",
    "- Emergency Procedure Folder x2 (folder): $1,600 standard only",
    "- Critical Incident Management Plan (cimp): $1,600 standard only",
    "- Online Warden Training eLearning (online_warden): tiered — 1-20 learners $35pp, 21-50 $25pp, 51-75 $20pp, 76-120 $15pp, 120+ $10pp. Show total only.",
    "",
    "DRAWING/DIAGRAMS PRICING RULES:",
    "- Always add 2 hours admin time to drawing hours from any quote",
    "- Rate based on QUOTED hours only (before admin): quoted 1-20hrs = $180/hr, quoted over 20hrs = $150/hr",
    "- Apply that rate to ALL hours including the 2 admin hours",
    "- Example: 16 quoted + 2 admin = 18hrs x $180 = $3,240",
    "- Example: 20 quoted + 2 admin = 22hrs x $180 = $3,960",
    "- Example: 21 quoted + 2 admin = 23hrs x $150 = $3,450",
    "",
    "EQUIPMENT CATALOGUE — extract items and quantities from email:",
    "- Warden: Vest $18.35, Drawstring Bag $19.31, Hard Hat $21.71, Backpack $32.04, Pealess Whistle $8.03",
    "- Chief Warden: Vest $18.35, Drawstring Bag $19.31, Hard Hat $21.71, Backpack $32.04, Pealess Whistle $8.03",
    "- Comms Officer: Vest $18.35, Drawstring Bag $19.31, Hard Hat $21.71, Backpack $32.04, Megaphone $58.66, Uniden 2-Way Radio Set $155.71",
    "- First Aid Responder: Vest $18.35, Drawstring Bag $19.31, Hard Hat $21.71, Backpack $32.04, Compact First Aid Kit $28.69",
    "- Shipping: $24.50 per order — include in equipment_total_with_shipping, never show separately",
    "- If email mentions equipment items, populate equipment array with {name, qty, unit_price (as number not string), total (as number)} for each item",
    "- Set equipment_subtotal (number) and equipment_total_with_shipping = subtotal + 24.50",
    "- Add equipment_total_with_shipping to total_standard, total_three_year and total_five_year",
    "",
    "IMPORTANT TOTALS RULE:",
    "Always calculate ALL THREE totals. For documentation services include their cost in all three totals.",
    "Always show dollar amounts — never use dash for totals.",
    "",
    "Return ONLY valid JSON, no markdown:",
    "{\"client_name\":\"\",\"contact_name\":\"\",\"contact_position\":\"\",\"contact_email\":\"\",\"proposal_title\":\"Emergency Management Services\",\"contract_type\":\"standard\",\"selected_service_ids\":[],\"online_warden_learners\":0,\"online_warden_price\":\"\",\"equipment\":[],\"equipment_subtotal\":0,\"equipment_total_with_shipping\":0,\"total_standard\":\"$X,XXX\",\"total_three_year\":\"$X,XXX/yr\",\"total_five_year\":\"$X,XXX/yr\",\"total_inc_gst_standard\":\"$X,XXX\",\"total_inc_gst_three_year\":\"$X,XXX/yr\",\"total_inc_gst_five_year\":\"$X,XXX/yr\",\"email_body\":\"\"}",
    "",
    "The email_body must end with: Kind regards,\\n\\nReinette Kruger\\nAdministration Officer\\nRisk 2 Solution\\neps@risk2solution.com"
  ].join("\n");

  const proposalRes = await client.messages.create({
    model: "claude-sonnet-4-5",
    max_tokens: 3000,
    system: [{ type: "text", text: proposalSystemPrompt, cache_control: { type: "ephemeral" } }],
    messages: [{
      role: "user",
      content: "Generate proposal data from this email:\n\nFrom: " + (classification.contact_name || "") + " <" + (classification.contact_email || "") + ">\nOrganisation: " + (classification.client_name || "") + "\nServices mentioned: " + (classification.services_mentioned || []).join(", ") + "\nDate: " + new Date().toLocaleDateString("en-AU", { day: "numeric", month: "long", year: "numeric" }) + "\n\nEmail content: " + emailText
    }]
  });

  let rawText = proposalRes.content[0].text.trim().replace(/^```json\s*/i,"").replace(/^```\s*/i,"").replace(/```\s*$/g,"").trim();
  const proposalData = JSON.parse(rawText);

  console.log("Proposal data from email:", JSON.stringify({ client: proposalData.client_name, services: proposalData.selected_service_ids, equipment: (proposalData.equipment||[]).length + " items", total_std: proposalData.total_standard }));

  if ((!proposalData.selected_service_ids || !proposalData.selected_service_ids.length) && (!proposalData.equipment || !proposalData.equipment.length)) {
    console.log("No services or equipment identified — skipping PPTX.");
    return null;
  }

  if (proposalData.equipment) {
    proposalData.equipment = proposalData.equipment.map(item => ({
      ...item,
      unit_price: parseFloat(String(item.unit_price).replace('$','').replace(',','')),
      total: parseFloat(String(item.total).replace('$','').replace(',',''))
    }));
  }

  if (!proposalData.date) proposalData.date = new Date().toLocaleDateString("en-AU", { day: "numeric", month: "long", year: "numeric" });

  const outputPath = "/tmp/proposal_" + Date.now() + ".pptx";
  await generateProposal(TEMPLATE_PATH, proposalData, outputPath);
  if (!fs.existsSync(outputPath)) throw new Error("PPTX not created");

  const clientName = (proposalData.client_name || "Proposal").replace(/\s+/g, "");
  const date = new Date().toLocaleDateString("en-AU", { month: "short", year: "numeric" }).replace(" ", "");
  const fileName = clientName + "_Proposal_" + date + ".pptx";

  activityLog.push({
    id: Date.now(), timestamp: new Date().toISOString(),
    client_name: proposalData.client_name || "", contact_name: proposalData.contact_name || classification.contact_name || "",
    contact_position: proposalData.contact_position || "", contact_email: proposalData.contact_email || classification.contact_email || "",
    contract_type: proposalData.contract_type || "standard", proposal_title: proposalData.proposal_title || "",
    services: proposalData.selected_service_ids || [], total_standard: proposalData.total_standard || "",
    total_three_year: proposalData.total_three_year || "", total_five_year: proposalData.total_five_year || "",
    total_inc_gst: proposalData.total_inc_gst_standard || "", email_body: proposalData.email_body || "",
    filename: fileName, file_path: outputPath, source: "email"
  });
  await saveToBlob(ACTIVITY_BLOB, activityLog);

  return { filePath: outputPath, fileName, proposalData, emailBody: proposalData.email_body || "" };
}

// ── AUTH ──────────────────────────────────────────────────────────────────────
function requireAuth(req, res, next) {
  if (req.session && req.session.authenticated) return next();
  res.status(401).json({ error: "Unauthorised" });
}
app.post("/api/login", (req, res) => {
  if (req.body.password === PASSWORD) { req.session.authenticated = true; res.json({ success: true }); }
  else res.status(401).json({ error: "Incorrect password" });
});
app.post("/api/logout", (req, res) => { req.session.destroy(); res.json({ success: true }); });
app.get("/api/check-auth", (req, res) => { res.json({ authenticated: !!(req.session && req.session.authenticated) }); });

// ── ACTIVITY ──────────────────────────────────────────────────────────────────
app.get("/api/activity", requireAuth, (req, res) => {
  const today = new Date().toDateString();
  res.json(activityLog.filter(e => new Date(e.timestamp).toDateString() === today));
});

// ── REVIEW QUEUE ──────────────────────────────────────────────────────────────
app.get("/api/review-queue", requireAuth, (req, res) => {
  const pending = reviewQueue.filter(e => e.status === "pending").reverse();
  pending.forEach(e => console.log("Queue item " + e.id + ": proposal_file=" + e.proposal_file + ", drafts=" + e.drafts.length));
  res.json(pending);
});
app.post("/api/poll-now", requireAuth, (req, res) => {
  res.json({ success: true, message: "Polling started" });
  pollInbox().catch(err => console.error("Poll error:", err.message));
});
app.get("/api/review-queue/count", requireAuth, (req, res) => {
  res.json({ count: reviewQueue.filter(e => e.status === "pending").length });
});
app.post("/api/review-queue/:id/dismiss", requireAuth, async (req, res) => {
  const e = reviewQueue.find(e => e.id === parseInt(req.params.id));
  if (e) {
    e.status = "dismissed";
    await saveToBlob(QUEUE_BLOB, reviewQueue);
  }
  res.json({ success: true });
});
app.get("/api/review-queue/:id/download", requireAuth, async (req, res) => {
  const e = reviewQueue.find(e => e.id === parseInt(req.params.id));
  if (!e || !e.proposal_file) return res.status(404).json({ error: "No proposal file" });

  if (fs.existsSync(e.proposal_file)) {
    return res.download(e.proposal_file, e.proposal_filename);
  }

  if (!e.proposal_data) return res.status(404).json({ error: "File expired. Please re-process the email by removing the Ariel Processed tag in Outlook and clicking Poll Now." });

  try {
    console.log("Regenerating expired proposal for queue item " + e.id);
    if (!fs.existsSync(TEMPLATE_PATH)) await downloadTemplate();
    const outputPath = "/tmp/proposal_regen_" + Date.now() + ".pptx";
    await generateProposal(TEMPLATE_PATH, e.proposal_data, outputPath);
    if (!fs.existsSync(outputPath)) throw new Error("PPTX not created");
    e.proposal_file = outputPath;
    console.log("Proposal regenerated successfully.");
    res.download(outputPath, e.proposal_filename);
  } catch (err) {
    console.error("Regeneration error:", err.message);
    res.status(500).json({ error: "Could not regenerate proposal: " + err.message });
  }
});

// ── CHAT ──────────────────────────────────────────────────────────────────────
const ARIEL_SYSTEM_PROMPT_TEXT = [
  "You are Ariel, the proposal agent for Risk 2 Solution (R2S).",
  "",
  "IDENTITY: Reinette Kruger, Administration Officer, eps@risk2solution.com",
  "",
  "PRICING — STANDARD RATE ONLY (no multi-year):",
  "- Base Plan / Evacuation Diagrams / Summary Charts / Warden Mapping: quote from Excel",
  "- Annual Assessment with feedback: $600",
  "- Emergency Management Plan: $1,600",
  "- Emergency Procedure Folder x2: $1,600",
  "- Critical Incident Management Plan: $1,600",
  "- Online Warden Training eLearning: tiered per learner",
  "",
  "PRICING — MULTI-YEAR ELIGIBLE:",
  "- Onsite Emergency Exercise: Standard $1,200/yr | 3yr $1,140/yr | 5yr $1,116/yr",
  "- Warden Training (1-2 hours): Standard $1,500/yr | 3yr $1,425/yr | 5yr $1,395/yr",
  "- Chief Warden Training (1-2 hours): Standard $1,500/yr | 3yr $1,425/yr | 5yr $1,395/yr",
  "- Fire Extinguisher Training (1-2 hours): Standard $1,200/yr | 3yr $1,140/yr | 5yr $1,116/yr",
  "- School Critical Incident Management Training: Standard $1,800/yr | 3yr $1,710/yr | 5yr $1,674/yr",
  "- Active Armed Offender Training: Standard $1,500/yr | 3yr $1,425/yr | 5yr $1,395/yr",
  "- Staff Awareness Training (1 hour): Standard $750/yr | 3yr $712/yr | 5yr $697/yr",
  "- Bushfire Preparedness Training: Standard $1,500/yr | 3yr $1,425/yr | 5yr $1,395/yr",
  "",
  "ONLINE WARDEN TRAINING PRICING (tiered per learner, total only):",
  "- 1-20 learners: $35/person | 21-50: $25/person | 51-75: $20/person | 76-120: $15/person | 120+: $10/person",
  "",
  "DRAWING/DIAGRAMS PRICING:",
  "- Add 2 admin hours to quoted hours. Rate based on QUOTED hours only.",
  "- Quoted 1-20hrs: $180/hr on all hours. Quoted over 20hrs: $150/hr on all hours.",
  "",
  "EQUIPMENT PRICING (per unit, shipping $24.50 per order included in total):",
  "- Vest (any role): $18.35 | Drawstring Bag: $19.31 | Hard Hat: $21.71 | Backpack: $32.04",
  "- Pealess Whistle: $8.03 | Megaphone: $58.66 | Uniden 2-Way Radio Set: $155.71 | First Aid Kit: $28.69",
  "",
  "TOTALS RULE — MANDATORY:",
  "Always calculate and populate ALL THREE totals (standard, 3-year, 5-year). Never leave them blank.",
  "- total_standard: sum of all services at standard rate",
  "- total_three_year: sum of all multi-year eligible services at 3yr rate + any standard-only services at full price",
  "- total_five_year: sum of all multi-year eligible services at 5yr rate + any standard-only services at full price",
  "- Always calculate total_inc_gst versions by multiplying each total by 1.1",
  "- Never use a dash or empty string for any total field — always show a dollar amount.",
  "",
  "SERVICE IDs: exercise, warden, chief, fire, school, armed, staff, bushfire, assessment, emp, folder, cimp, diagrams, online_warden",
  "",
  "Respond with ONLY valid JSON:",
  "{\"message\":\"conversational response\",\"ready_to_generate\":false,\"proposal_data\":{\"client_name\":\"\",\"contact_name\":\"\",\"contact_position\":\"\",\"contact_email\":\"\",\"proposal_title\":\"Emergency Management Services\",\"contract_type\":\"standard\",\"date\":\"\",\"selected_service_ids\":[],\"diagrams_price\":\"\",\"online_warden_learners\":0,\"online_warden_price\":\"\",\"equipment\":[],\"equipment_subtotal\":0,\"equipment_total_with_shipping\":0,\"total_standard\":\"\",\"total_three_year\":\"\",\"total_five_year\":\"\",\"total_inc_gst_standard\":\"\",\"total_inc_gst_three_year\":\"\",\"total_inc_gst_five_year\":\"\",\"email_body\":\"\"}}",
  "Set ready_to_generate true when you have client name, contact name, position, email and at least one service.",
  "Email body must end: Kind regards,\\n\\nReinette Kruger\\nAdministration Officer\\nRisk 2 Solution\\neps@risk2solution.com",
  "",
  "CRITICAL: You MUST ALWAYS respond with a raw JSON object and nothing else. No plain text. No markdown. Every response must start with { and end with }. Put any questions inside the message field."
].join("\n");

app.post("/api/chat", requireAuth, async (req, res) => {
  try {
    let messages = req.body.messages || [];
    if (messages.length > MAX_CHAT_MESSAGES) {
      const first = messages[0];
      messages = messages.slice(-MAX_CHAT_MESSAGES);
      if (first && !messages.includes(first)) messages.unshift(first);
      console.log("[TRIMMED] Chat history trimmed to " + messages.length + " messages");
    }
    const response = await client.messages.create({
      model: "claude-sonnet-4-5",
      max_tokens: 2000,
      temperature: 0,
      system: [{ type: "text", text: ARIEL_SYSTEM_PROMPT_TEXT, cache_control: { type: "ephemeral" } }],
      messages
    });
    let raw = response.content[0].text.trim().replace(/^```json\s*/i,"").replace(/^```\s*/i,"").replace(/```\s*$/g,"").trim();
    let parsed;
    try { parsed = JSON.parse(raw); }
    catch { const m = raw.match(/\{[\s\S]*\}/); try { parsed = m ? JSON.parse(m[0]) : { message: raw, ready_to_generate: false, proposal_data: null }; } catch { parsed = { message: raw, ready_to_generate: false, proposal_data: null }; } }
    res.json(parsed);
  } catch (err) { console.error("Chat error:", err); res.status(500).json({ error: "Something went wrong." }); }
});

// ── EXCEL PARSE ───────────────────────────────────────────────────────────────
app.post("/api/parse-excel", requireAuth, upload.single("excel"), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: "No file uploaded" });
  try {
    const XLSX = require("xlsx");
    let textContent = "";
    if (path.extname(req.file.originalname).toLowerCase() === ".csv") {
      textContent = fs.readFileSync(req.file.path, "utf8");
    } else {
      const wb = XLSX.readFile(req.file.path);
      wb.SheetNames.forEach(n => { const csv = XLSX.utils.sheet_to_csv(wb.Sheets[n]); if (csv.trim().replace(/,/g,"").trim()) textContent += "Sheet: " + n + "\n" + csv + "\n\n"; });
    }
    if (!textContent.trim()) return res.json({ client_name:"", contact_name:"", drawing_hours:0, summary:"File appears empty." });
    const response = await client.messages.create({
      model: "claude-haiku-4-5-20251001",
      max_tokens: 1000,
      messages: [{ role: "user", content: "Extract details from this R2S quote spreadsheet and return ONLY valid JSON:\n{\"client_name\":\"\",\"contact_name\":\"\",\"contact_position\":\"\",\"contact_email\":\"\",\"drawing_hours\":0,\"hourly_rate\":0,\"total_drawing_cost\":0,\"notes\":\"\",\"summary\":\"one sentence\"}\nSpreadsheet:\n" + textContent.substring(0, 8000) }]
    });
    let raw = response.content[0].text.trim().replace(/^```json\s*/i,"").replace(/^```\s*/i,"").replace(/```\s*$/g,"").trim();
    const parsed = JSON.parse(raw);
    if (parsed.drawing_hours > 0) {
      const quoted = parsed.drawing_hours;
      const rate = quoted <= 20 ? 180 : 150;
      const total = quoted + 2;
      parsed.drawing_hours = total;
      parsed.hourly_rate = rate;
      parsed.total_drawing_cost = parseFloat((total * rate).toFixed(2));
      parsed.summary = (parsed.summary || "") + " " + quoted + " quoted hrs + 2 admin = " + total + "hrs @ $" + rate + "/hr = $" + parsed.total_drawing_cost;
    }
    fs.unlink(req.file.path, () => {});
    res.json(parsed);
  } catch (err) { if (req.file) fs.unlink(req.file.path, () => {}); res.status(500).json({ error: "Could not read file." }); }
});

// ── GENERATE PPTX ─────────────────────────────────────────────────────────────
app.post("/api/generate-pptx", requireAuth, async (req, res) => {
  const { proposal_data } = req.body;
  if (!proposal_data) return res.status(400).json({ error: "No proposal data" });
  if (!proposal_data.date) proposal_data.date = new Date().toLocaleDateString("en-AU", { day: "numeric", month: "long", year: "numeric" });
  try {
    if (!fs.existsSync(TEMPLATE_PATH)) await downloadTemplate();
    const outputPath = "/tmp/proposal_" + Date.now() + ".pptx";
    await generateProposal(TEMPLATE_PATH, proposal_data, outputPath);
    if (!fs.existsSync(outputPath)) throw new Error("PPTX not created");
    const clientName = (proposal_data.client_name || "Proposal").replace(/\s+/g, "");
    const date = new Date().toLocaleDateString("en-AU", { month: "short", year: "numeric" }).replace(" ", "");
    const fileName = clientName + "_Proposal_" + date + ".pptx";
    activityLog.push({
      id: Date.now(), timestamp: new Date().toISOString(),
      client_name: proposal_data.client_name, contact_name: proposal_data.contact_name,
      contact_position: proposal_data.contact_position || "", contact_email: proposal_data.contact_email || "",
      contract_type: proposal_data.contract_type, proposal_title: proposal_data.proposal_title || "",
      services: proposal_data.selected_service_ids || [], total_standard: proposal_data.total_standard || "",
      total_three_year: proposal_data.total_three_year || "", total_five_year: proposal_data.total_five_year || "",
      total_inc_gst: proposal_data.total_inc_gst_standard || "", email_body: proposal_data.email_body || "",
      filename: fileName, file_path: outputPath
    });
    await saveToBlob(ACTIVITY_BLOB, activityLog);
    res.download(outputPath, fileName, (err) => {
      if (err) console.error("Download error:", err);
      setTimeout(() => fs.unlink(outputPath, () => {}), 60000);
    });
  } catch (err) { console.error("PPTX error:", err.message); res.status(500).json({ error: err.message || "Failed." }); }
});

app.get("/api/redownload/:id", requireAuth, (req, res) => {
  const e = activityLog.find(e => e.id === parseInt(req.params.id));
  if (!e) return res.status(404).json({ error: "Not found" });
  if (!fs.existsSync(e.file_path)) return res.status(404).json({ error: "File expired." });
  res.download(e.file_path, e.filename);
});

// ── START ─────────────────────────────────────────────────────────────────────
async function start() {
  activityLog.push(...(await loadFromBlob(ACTIVITY_BLOB, [])));
  reviewQueue.push(...(await loadFromBlob(QUEUE_BLOB, [])));
  console.log("Loaded " + activityLog.length + " activity entries and " + reviewQueue.length + " queue entries from blob storage.");

  try { await downloadTemplate(); } catch (err) { console.error("Template download failed:", err.message); }

  app.listen(PORT, () => {
    console.log("Ariel is running on port " + PORT);
    if (AZURE_CLIENT_ID) {
      if (isBrisbaneBusinessHours()) { pollInbox(); console.log("Initial poll on startup."); }
      setInterval(() => {
        if (isBrisbaneBusinessHours()) { pollInbox(); } else { console.log("💤 Outside business hours — Ariel is resting."); }
      }, 20 * 60 * 1000);
      console.log("Email polling started — every 20 minutes during business hours.");
    }
  });
}

start().catch(err => console.error("Startup error:", err.message));
