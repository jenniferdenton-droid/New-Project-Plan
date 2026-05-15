// ════════════════════════════════════════════════════════════════
// /api/ai-assistant — unified Vercel serverless function
// Handles: generate-deck · process-transcript · send-status · quick-send
// ════════════════════════════════════════════════════════════════
//
// Required Vercel env vars (Hobby-tier, no deck gen):
//   ANTHROPIC_API_KEY        — Claude API key (console.anthropic.com)
//   SLACK_BOT_TOKEN          — Slack bot token (xoxb-...), preferred
//   SLACK_CHANNEL            — Default Slack channel name (#bookkeeping-launch)
//   SLACK_WEBHOOK_URL        — Alternative: incoming webhook (use if no bot token)
//   GOOGLE_CLIENT_ID         — OAuth Client ID (covers Gmail+Drive+Sheets+Calendar)
//   GOOGLE_CLIENT_SECRET     — OAuth Client Secret
//   GOOGLE_REFRESH_TOKEN     — OAuth refresh token (from OAuth Playground)
//   GMAIL_FROM_EMAIL         — Email address to send AS (e.g., jennifer.denton@joinmoxie.com)
//
// Optional (for deck gen on Vercel Pro):
//   GOOGLE_SERVICE_ACCOUNT_JSON, GDRIVE_FOLDER_ID
//
// ════════════════════════════════════════════════════════════════

const Anthropic = require('@anthropic-ai/sdk');
const { google } = require('googleapis');

// ── Init clients (cold start) ─────────────────────────────
const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

// Build an OAuth2 client using the unified GOOGLE_* credentials.
// One token works for Gmail, Drive, Sheets, and Calendar APIs.
function getGoogleAuth() {
  // Support both new GOOGLE_* names and legacy GMAIL_* names for back-compat
  const clientId     = process.env.GOOGLE_CLIENT_ID     || process.env.GMAIL_CLIENT_ID;
  const clientSecret = process.env.GOOGLE_CLIENT_SECRET || process.env.GMAIL_CLIENT_SECRET;
  const refreshToken = process.env.GOOGLE_REFRESH_TOKEN || process.env.GMAIL_REFRESH_TOKEN;
  if (!clientId || !clientSecret || !refreshToken) {
    throw new Error('Google OAuth credentials missing — set GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET, GOOGLE_REFRESH_TOKEN in Vercel env vars.');
  }
  const oAuth2 = new google.auth.OAuth2(
    clientId, clientSecret, 'https://developers.google.com/oauthplayground'
  );
  oAuth2.setCredentials({ refresh_token: refreshToken });
  return oAuth2;
}
function getGmailClient()    { return google.gmail({ version: 'v1', auth: getGoogleAuth() }); }
function getSheetsClient()   { return google.sheets({ version: 'v4', auth: getGoogleAuth() }); }
function getCalendarClient() { return google.calendar({ version: 'v3', auth: getGoogleAuth() }); }
function getDriveOAuthClient(){ return google.drive({ version: 'v3', auth: getGoogleAuth() }); }

function getDriveClient() {
  if (!process.env.GOOGLE_SERVICE_ACCOUNT_JSON) return null;
  const creds = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  const auth = new google.auth.JWT(
    creds.client_email, null, creds.private_key, ['https://www.googleapis.com/auth/drive']
  );
  return google.drive({ version: 'v3', auth });
}

// ════════════════════════════════════════════════════════════════
// CLAUDE CALLS
// ════════════════════════════════════════════════════════════════
const MODEL = 'claude-sonnet-4-6';

async function claudeJson(systemPrompt, userPrompt, maxTokens = 4096) {
  const res = await anthropic.messages.create({
    model: MODEL,
    max_tokens: maxTokens,
    system: systemPrompt,
    messages: [{ role: 'user', content: userPrompt }],
  });
  const text = res.content.map(c => c.text || '').join('');
  // Strip markdown code fences if present
  const cleaned = text.replace(/^```(?:json)?\s*/i, '').replace(/\s*```\s*$/, '').trim();
  try { return JSON.parse(cleaned); }
  catch (e) {
    throw new Error('Claude returned non-JSON: ' + cleaned.slice(0, 300));
  }
}

async function claudeText(systemPrompt, userPrompt, maxTokens = 2048) {
  const res = await anthropic.messages.create({
    model: MODEL,
    max_tokens: maxTokens,
    system: systemPrompt,
    messages: [{ role: 'user', content: userPrompt }],
  });
  return res.content.map(c => c.text || '').join('').trim();
}

// ════════════════════════════════════════════════════════════════
// ACTION 1: GENERATE DECK
// ════════════════════════════════════════════════════════════════
async function generateDeck({ title, focus, distribution, project, settings, fbProjectId }) {
  // 1) Ask Claude for slide content as JSON
  const systemPrompt = `You are a Principal Strategy Manager building executive decks at a SaaS company (Moxie — MedSpa software). Output exec-ready content: bottom line first, structured, no fluff.`;
  const userPrompt = `Build a working session deck (8 slides) for project "${project.name}".
Meeting title: ${title || 'Working Session — ' + new Date().toLocaleDateString()}
Focus areas: ${focus || 'open discussion'}

Project state (use this for facts, do not invent):
${JSON.stringify({
  name: project.name, lead: project.lead, dueDate: project.dueDate,
  description: project.description,
  milestones: project.milestones.map(m => ({ name: m.name, date: m.date, status: m.status })),
  taskCounts: countByStatus(project.tasks),
  actionItemCounts: countByStatus(project.actionItems),
  topRisks: (project.risks || []).slice(0, 5),
  ragStatus: project.ragStatus,
}, null, 2)}

Return ONLY a JSON object with this exact shape:
{
  "deckTitle": "string",
  "slides": [
    { "type": "title", "title": "string", "subtitle": "string", "kicker": "string" },
    { "type": "summary", "title": "string", "bullets": [{"label":"...","desc":"..."}] },
    { "type": "kpis", "title": "string", "kpis": [{"label":"...","value":"...","sub":"..."}] },
    { "type": "twocol", "title": "string", "kicker":"...", "leftHeader":"...", "leftItems":["..."], "rightHeader":"...", "rightItems":["..."] },
    { "type": "threecol", "title":"string", "kicker":"...", "cols":[{"title":"","items":["..."]}] },
    { "type": "table", "title":"string", "kicker":"...", "headers":["When","What","Who"], "rows":[["...","...","..."]] },
    { "type": "asks", "title":"string", "asks":[{"question":"...","why":"..."}] },
    { "type": "closing", "headline":"...", "subhead":"..." }
  ]
}
Slides MUST be relevant to the meeting focus. Be concise — slides have limited space.`;

  const deckSpec = await claudeJson(systemPrompt, userPrompt, 4096);

  // 2) Build pptx from spec
  const pptxBuffer = await buildPptxFromSpec(deckSpec);

  // 3) Distribute
  const result = { driveLink: null, notifications: {} };

  if (distribution.drive) {
    result.driveLink = await uploadToDrive(pptxBuffer, deckSpec.deckTitle || (title || 'Working Session Deck'));
  }

  // 4) Notifications
  const link = result.driveLink || '(deck generated, no Drive link)';
  if (distribution.slack) {
    await postToSlack({
      text: `:rocket: *${deckSpec.deckTitle || title || 'New deck ready'}*\nGenerated from current ${project.name} project state.\n${result.driveLink ? `:point_right: <${result.driveLink}|Open deck in Drive>` : ''}\nReply with feedback in this thread.`,
      settings,
    });
    result.notifications.slack = true;
  }
  if (distribution.email) {
    const recipients = collectRecipients(settings, project);
    if (recipients.length === 0) throw new Error('No recipients configured (Settings recipients OR Stakeholder POC emails).');
    const summary = await summarizeForEmail(project, deckSpec);
    await sendEmail({
      to: recipients,
      cc: settings.cc || [],
      subject: deckSpec.deckTitle || title || 'Working Session Deck',
      html: emailHtml({ project, deckSpec, summary, link: result.driveLink, fromName: settings.fromName }),
      attachments: result.driveLink ? [] : [{
        content: pptxBuffer.toString('base64'),
        filename: (deckSpec.deckTitle || 'deck').replace(/[^a-z0-9]+/gi, '_') + '.pptx',
        type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        disposition: 'attachment',
      }],
      fromName: settings.fromName,
      projectName: project.name,
    });
    result.notifications.email = recipients.length;
  }
  return result;
}

// ════════════════════════════════════════════════════════════════
// ACTION 2: PROCESS TRANSCRIPT
// ════════════════════════════════════════════════════════════════
async function processTranscript({ title, transcript, extract, project }) {
  const systemPrompt = `You extract structured data from meeting transcripts. Be precise. Owners must be real names mentioned. Dates in YYYY-MM-DD if explicit, else empty string.`;
  const userPrompt = `Transcript title: ${title}
Project context: ${project.name} (lead: ${project.lead}, due: ${project.dueDate})
Existing owners on this project: ${[...new Set([
    ...(project.actionItems || []).map(a => a.owner),
    ...(project.tasks || []).map(t => t.owner),
  ].filter(Boolean))].join(', ') || 'none yet'}

Transcript:
"""
${transcript.slice(0, 30000)}
"""

Extract these (return ONLY a JSON object with these exact keys, even if some are empty arrays):
${extract.actions   ? '- actionItems: [{item, owner, dueDate, priority (Critical|High|Medium|Low), notes}]' : ''}
${extract.decisions ? '- decisions:   [{decision, context, owner, date}]' : ''}
${extract.risks     ? '- risks:       [{risk, category, likelihood (High|Medium|Low), impact (High|Medium|Low), mitigation, owner}]' : ''}
${extract.summary   ? '- summary:     "2-4 sentence summary of the meeting"\n- attendees:   "comma-separated list of attendees mentioned"' : ''}

Return JSON only — no prose, no code fences.`;

  const extracted = await claudeJson(systemPrompt, userPrompt, 4096);
  return { extracted };
}

// ════════════════════════════════════════════════════════════════
// ACTION 3: SEND STATUS UPDATE
// ════════════════════════════════════════════════════════════════
async function sendStatus({ audience, cadence, channels, project, settings }) {
  const result = { notifications: {} };

  // 1) Generate copy with Claude
  const systemPrompt = `You write executive status updates for a Principal Strategy Manager at Moxie (SaaS for MedSpas). Bottom line first, exec-ready, no fluff. Use plain language. Match the requested audience level.`;
  const userPrompt = `Draft a ${cadence} status update for ${audience}.
Project: ${project.name} (lead: ${project.lead}, due: ${project.dueDate})
Current state:
${JSON.stringify({
  ragStatus: project.ragStatus,
  milestones: (project.milestones || []).map(m => ({ name: m.name, date: m.date, status: m.status })),
  taskCounts: countByStatus(project.tasks),
  actionItemCounts: countByStatus(project.actionItems),
  topOpenAIs: (project.actionItems || []).filter(a => a.status !== 'Done').slice(0, 8).map(a => ({ item: a.item, owner: a.owner, due: a.dueDate, priority: a.priority })),
  topRisks: (project.risks || []).slice(0, 4),
}, null, 2)}

Return JSON: { "subject": "...", "slackMrkdwn": "...", "emailHtml": "..." }
- slackMrkdwn: Slack mrkdwn format with section headers, bullets (•), bold (*text*)
- emailHtml: clean HTML, no inline styles needed`;

  const copy = await claudeJson(systemPrompt, userPrompt, 3000);

  if (channels.slack) {
    await postToSlack({ text: copy.slackMrkdwn, settings });
    result.notifications.slack = true;
  }
  if (channels.email) {
    const recipients = collectRecipients(settings, project);
    if (recipients.length === 0) throw new Error('No recipients configured in Settings AND no POC emails on Stakeholder Plan.');
    await sendEmail({
      to: recipients, cc: settings.cc || [],
      subject: copy.subject || `${cadence} update`,
      html: copy.emailHtml || `<pre>${copy.slackMrkdwn}</pre>`,
      fromName: settings.fromName,
      projectName: project.name,
    });
    result.notifications.email = recipients.length;
  }
  return result;
}

// Returns merged unique list of: Settings.recipients + POC emails from stakeholders.included
function collectRecipients(settings, project) {
  const list = new Set();
  (settings.recipients || []).forEach(r => { if (r && r.trim()) list.add(r.trim().toLowerCase()); });
  if (project && project.stakeholders) {
    Object.values(project.stakeholders).forEach(s => {
      if (s && s.included && s.contactEmail && s.contactEmail.trim()) {
        list.add(s.contactEmail.trim().toLowerCase());
      }
    });
  }
  return [...list];
}

// ════════════════════════════════════════════════════════════════
// ACTION 4: QUICK SEND
// ════════════════════════════════════════════════════════════════
async function quickSend({ channel, subject, body, polish, project, settings }) {
  let finalText = body;
  if (polish) {
    const systemPrompt = channel === 'slack'
      ? `Polish this into a clean, exec-ready Slack message (mrkdwn). Keep it punchy. Bottom line first.`
      : `Polish this into a clean, exec-ready email body (HTML allowed but minimal). Bottom line first.`;
    finalText = await claudeText(systemPrompt, `Project: ${project.name}\n\nDraft:\n${body}`, 1500);
  }
  if (channel === 'slack') {
    await postToSlack({ text: finalText, settings });
  } else {
    const recipients = collectRecipients(settings, project);
    if (recipients.length === 0) throw new Error('No recipients configured (Settings recipients OR Stakeholder POC emails).');
    await sendEmail({
      to: recipients, cc: settings.cc || [],
      subject, html: finalText.includes('<') ? finalText : `<p>${finalText.replace(/\n/g, '<br>')}</p>`,
      fromName: settings.fromName,
      projectName: project.name,
    });
  }
  return { ok: true, finalText };
}

// ════════════════════════════════════════════════════════════════
// HELPERS — Slack, Email, Drive, PPTX
// ════════════════════════════════════════════════════════════════
async function postToSlack({ text, settings }) {
  // Prefer Bot Token (chat.postMessage API) if available — works for any channel the bot is in.
  // Fall back to Webhook URL — channel-locked but no install required.
  const botToken = process.env.SLACK_BOT_TOKEN;
  const webhook  = process.env.SLACK_WEBHOOK_URL;

  if (botToken) {
    const channel = (settings && settings.slackChannel) || process.env.SLACK_CHANNEL;
    if (!channel) throw new Error('SLACK_BOT_TOKEN set but no channel — set SLACK_CHANNEL env var or fill Slack Channel in Settings.');
    const res = await fetch('https://slack.com/api/chat.postMessage', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json; charset=utf-8',
        'Authorization': `Bearer ${botToken}`,
      },
      body: JSON.stringify({ channel, text, mrkdwn: true }),
    });
    const json = await res.json().catch(() => ({}));
    if (!res.ok || !json.ok) {
      throw new Error('Slack chat.postMessage failed: ' + (json.error || res.statusText) +
        (json.error === 'not_in_channel' ? ' — invite the bot: /invite @<your-app-name> in the channel' : ''));
    }
    return;
  }

  if (webhook) {
    const payload = { text };
    if (settings && settings.slackChannel) payload.channel = settings.slackChannel;
    const res = await fetch(webhook, {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });
    if (!res.ok) throw new Error('Slack webhook post failed: ' + await res.text());
    return;
  }

  throw new Error('No Slack credentials — set either SLACK_BOT_TOKEN (+ SLACK_CHANNEL) or SLACK_WEBHOOK_URL in Vercel env vars.');
}

async function sendEmail({ to, cc, subject, html, attachments, fromName, projectName }) {
  const fromEmail = process.env.GMAIL_FROM_EMAIL;
  if (!fromEmail) throw new Error('GMAIL_FROM_EMAIL not set.');

  // Auto-prefix subject with [Project Name] if not already present
  if (projectName && subject && !subject.includes(`[${projectName}]`) && !subject.toLowerCase().includes(projectName.toLowerCase())) {
    subject = `[${projectName}] ${subject}`;
  }
  const gmail = getGmailClient();
  const fromHeader = fromName ? `"${fromName.replace(/"/g, '')}" <${fromEmail}>` : fromEmail;
  const toList = Array.isArray(to) ? to.join(', ') : to;
  const ccList = (cc && cc.length) ? (Array.isArray(cc) ? cc.join(', ') : cc) : '';

  // Build a MIME message. If attachments present, build multipart/mixed; else simple HTML.
  const boundary = 'mox_bk_ai_' + Date.now();
  let raw;
  if (attachments && attachments.length) {
    const parts = [];
    parts.push(`Content-Type: text/html; charset="UTF-8"\r\n`);
    parts.push(`MIME-Version: 1.0\r\n`);
    parts.push(`Content-Transfer-Encoding: 7bit\r\n\r\n`);
    parts.push(html);
    let mimeBody =
      `From: ${fromHeader}\r\n` +
      `To: ${toList}\r\n` +
      (ccList ? `Cc: ${ccList}\r\n` : '') +
      `Subject: ${subject}\r\n` +
      `MIME-Version: 1.0\r\n` +
      `Content-Type: multipart/mixed; boundary="${boundary}"\r\n\r\n` +
      `--${boundary}\r\n` +
      parts.join('') + `\r\n`;
    attachments.forEach(att => {
      mimeBody +=
        `--${boundary}\r\n` +
        `Content-Type: ${att.type}; name="${att.filename}"\r\n` +
        `MIME-Version: 1.0\r\n` +
        `Content-Disposition: attachment; filename="${att.filename}"\r\n` +
        `Content-Transfer-Encoding: base64\r\n\r\n` +
        att.content + `\r\n`;
    });
    mimeBody += `--${boundary}--`;
    raw = mimeBody;
  } else {
    raw =
      `From: ${fromHeader}\r\n` +
      `To: ${toList}\r\n` +
      (ccList ? `Cc: ${ccList}\r\n` : '') +
      `Subject: ${subject}\r\n` +
      `MIME-Version: 1.0\r\n` +
      `Content-Type: text/html; charset="UTF-8"\r\n` +
      `Content-Transfer-Encoding: 7bit\r\n\r\n` +
      html;
  }

  // Gmail expects URL-safe base64
  const encoded = Buffer.from(raw).toString('base64')
    .replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');

  await gmail.users.messages.send({
    userId: 'me',
    requestBody: { raw: encoded },
  });
}

async function uploadToDrive(buffer, name) {
  const drive = getDriveClient();
  if (!drive) throw new Error('GOOGLE_SERVICE_ACCOUNT_JSON not set.');
  if (!process.env.GDRIVE_FOLDER_ID) throw new Error('GDRIVE_FOLDER_ID not set.');
  const fileMetadata = {
    name: name.replace(/[^a-z0-9]+/gi, '_') + '_' + new Date().toISOString().slice(0,10) + '.pptx',
    parents: [process.env.GDRIVE_FOLDER_ID],
  };
  const media = {
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    body: Readable.from(buffer),
  };
  const res = await drive.files.create({
    requestBody: fileMetadata, media,
    fields: 'id, webViewLink',
  });
  // Make it viewable by anyone with the link
  try {
    await drive.permissions.create({
      fileId: res.data.id,
      requestBody: { role: 'reader', type: 'anyone' },
    });
  } catch(e) { /* may fail in restricted orgs — link still works for org members */ }
  return res.data.webViewLink;
}

function countByStatus(arr) {
  const out = {};
  (arr || []).forEach(x => { const s = x.status || 'Open'; out[s] = (out[s] || 0) + 1; });
  return out;
}

async function summarizeForEmail(project, deckSpec) {
  const sys = `You write 3-bullet exec summaries for emails. Bottom line first.`;
  const user = `Summarize this deck in 3 bullets max for an email body.\n\nDeck: ${JSON.stringify(deckSpec, null, 2).slice(0, 4000)}\n\nReturn plain HTML <ul><li>...</li></ul>`;
  return await claudeText(sys, user, 800);
}

function emailHtml({ project, deckSpec, summary, link, fromName }) {
  return `
<div style="font-family:-apple-system,Segoe UI,sans-serif;max-width:640px;color:#1A1A2E;">
  <h2 style="color:#1F1A47;margin-bottom:6px;">${escapeHtml(deckSpec.deckTitle || project.name)}</h2>
  <p style="color:#666;font-size:13px;margin-top:0;">Working session deck — generated ${new Date().toLocaleString()}</p>
  ${summary || ''}
  ${link ? `<p style="margin:18px 0;"><a href="${link}" style="background:#6B4EFF;color:#fff;padding:10px 18px;border-radius:6px;text-decoration:none;font-weight:600;">→ Open deck in Google Drive</a></p>` : ''}
  <hr style="border:none;border-top:1px solid #eee;margin:20px 0;">
  <p style="font-size:12px;color:#888;">Reply with feedback. — ${escapeHtml(fromName || 'Jen')}</p>
</div>`;
}

function escapeHtml(s) {
  return String(s || '').replace(/[&<>"']/g, c => ({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;' }[c]));
}

// ── PPTX builder (compact version of your full deck script) ──────
async function buildPptxFromSpec(spec) {
  const pres = new PptxGenJS();
  pres.layout = 'LAYOUT_16x9';
  pres.title  = spec.deckTitle || 'Working Session';
  const C = { ink:'1F1A47', purple:'6B4EFF', teal:'00BFA6', cream:'FBF9F4', paper:'FFFFFF', border:'E6E2D8', text:'1A1A2E', textMute:'6B6680' };

  for (const sl of (spec.slides || [])) {
    const s = pres.addSlide();
    if (sl.type === 'title') {
      s.background = { color: C.ink };
      if (sl.kicker) s.addText(sl.kicker, { x:0.6, y:0.7, w:8, h:0.3, fontSize:11, color:C.teal, bold:true, charSpacing:6, margin:0 });
      s.addText(sl.title || '', { x:0.6, y:1.5, w:9, h:0.9, fontSize:42, bold:true, color:'FFFFFF', margin:0 });
      if (sl.subtitle) s.addText(sl.subtitle, { x:0.6, y:2.5, w:9, h:0.5, fontSize:20, color:'CADCFC', italic:true, margin:0 });
    } else if (sl.type === 'closing') {
      s.background = { color: C.ink };
      s.addText('THANK YOU', { x:0.6, y:1.5, w:9, h:0.4, fontSize:13, color:C.teal, bold:true, charSpacing:5, margin:0 });
      s.addText(sl.headline || '', { x:0.6, y:2.0, w:9, h:0.9, fontSize:34, bold:true, color:'FFFFFF', margin:0 });
      if (sl.subhead) s.addText(sl.subhead, { x:0.6, y:3.0, w:9, h:0.5, fontSize:16, color:'CADCFC', italic:true, margin:0 });
    } else {
      // Content slide
      s.background = { color: C.cream };
      if (sl.kicker) s.addText(sl.kicker, { x:0.5, y:0.32, w:9, h:0.28, fontSize:11, color:C.purple, bold:true, charSpacing:4, margin:0 });
      s.addText(sl.title || '', { x:0.5, y:0.55, w:9, h:0.7, fontSize:26, bold:true, color:C.ink, margin:0 });
      s.addShape(pres.shapes.RECTANGLE, { x:0.5, y:1.28, w:0.6, h:0.04, fill:{color:C.teal}, line:{color:C.teal} });

      if (sl.type === 'summary' && sl.bullets) {
        sl.bullets.forEach((b, i) => {
          const y = 1.6 + i * 1.1;
          s.addShape(pres.shapes.OVAL, { x:0.5, y, w:0.6, h:0.6, fill:{color:C.purple}, line:{color:C.purple} });
          s.addText(String(i+1), { x:0.5, y, w:0.6, h:0.6, fontSize:22, bold:true, color:'FFFFFF', align:'center', valign:'middle', margin:0 });
          s.addShape(pres.shapes.RECTANGLE, { x:1.3, y:y-0.05, w:8.2, h:0.85, fill:{color:C.paper}, line:{color:C.border} });
          s.addText(b.label || '', { x:1.5, y:y-0.02, w:7.9, h:0.38, fontSize:14, bold:true, color:C.ink, valign:'middle', margin:0 });
          s.addText(b.desc || '',  { x:1.5, y:y+0.32, w:7.9, h:0.45, fontSize:11, color:C.textMute, valign:'top', margin:0 });
        });
      } else if (sl.type === 'kpis' && sl.kpis) {
        sl.kpis.forEach((k, i) => {
          const x = 0.5 + i * (9 / sl.kpis.length);
          const w = (9 / sl.kpis.length) - 0.2;
          s.addShape(pres.shapes.RECTANGLE, { x, y:1.7, w, h:1.6, fill:{color:C.paper}, line:{color:C.border} });
          s.addShape(pres.shapes.RECTANGLE, { x, y:1.7, w, h:0.08, fill:{color:C.teal}, line:{color:C.teal} });
          s.addText(k.label || '', { x:x+0.15, y:1.85, w:w-0.3, h:0.3, fontSize:10, bold:true, color:C.textMute, charSpacing:3, margin:0 });
          s.addText(k.value || '', { x:x+0.15, y:2.15, w:w-0.3, h:0.7, fontSize:30, bold:true, color:C.ink, margin:0 });
          s.addText(k.sub   || '', { x:x+0.15, y:2.85, w:w-0.3, h:0.3, fontSize:10, color:C.textMute, margin:0 });
        });
      } else if (sl.type === 'twocol') {
        ['left','right'].forEach((side, i) => {
          const x = 0.5 + i * 4.6;
          const items = side === 'left' ? sl.leftItems : sl.rightItems;
          const header = side === 'left' ? sl.leftHeader : sl.rightHeader;
          s.addShape(pres.shapes.RECTANGLE, { x, y:1.6, w:4.4, h:3.4, fill:{color:C.paper}, line:{color:C.border} });
          s.addShape(pres.shapes.RECTANGLE, { x, y:1.6, w:4.4, h:0.4, fill:{color:i?C.purple:C.teal}, line:{color:i?C.purple:C.teal} });
          s.addText(header || '', { x:x+0.2, y:1.6, w:4.0, h:0.4, fontSize:11, bold:true, color:'FFFFFF', charSpacing:3, valign:'middle', margin:0 });
          if (items && items.length) {
            s.addText(items.map((t,j) => ({ text:t, options:{ bullet:{code:'25AA'}, breakLine:j<items.length-1, paraSpaceAfter:5 } })),
              { x:x+0.2, y:2.15, w:4.0, h:2.7, fontSize:12, color:C.text, valign:'top', margin:0 });
          }
        });
      } else if (sl.type === 'threecol' && sl.cols) {
        sl.cols.slice(0,3).forEach((c, i) => {
          const x = 0.5 + i * 3.05;
          s.addShape(pres.shapes.RECTANGLE, { x, y:1.7, w:2.95, h:3.2, fill:{color:C.paper}, line:{color:C.border} });
          s.addShape(pres.shapes.RECTANGLE, { x, y:1.7, w:2.95, h:0.45, fill:{color:[C.purple,C.teal,C.ink][i]}, line:{color:[C.purple,C.teal,C.ink][i]} });
          s.addText(c.title || '', { x:x+0.2, y:1.7, w:2.6, h:0.45, fontSize:13, bold:true, color:'FFFFFF', valign:'middle', margin:0 });
          if (c.items && c.items.length) {
            s.addText(c.items.map((t,j) => ({ text:t, options:{ bullet:{code:'25AA'}, breakLine:j<c.items.length-1, paraSpaceAfter:5 } })),
              { x:x+0.2, y:2.3, w:2.6, h:2.5, fontSize:11, color:C.text, valign:'top', margin:0 });
          }
        });
      } else if (sl.type === 'table' && sl.headers && sl.rows) {
        s.addShape(pres.shapes.RECTANGLE, { x:0.5, y:1.6, w:9, h:0.35, fill:{color:C.ink}, line:{color:C.ink} });
        sl.headers.forEach((h, i) => {
          s.addText(h, { x: 0.65 + i * (8.7/sl.headers.length), y:1.6, w: 8.7/sl.headers.length, h:0.35, fontSize:10, bold:true, color:'FFFFFF', charSpacing:3, valign:'middle', margin:0 });
        });
        sl.rows.slice(0, 8).forEach((row, ri) => {
          const y = 2.0 + ri * 0.4;
          s.addShape(pres.shapes.RECTANGLE, { x:0.5, y, w:9, h:0.38, fill:{color: ri%2 ? 'F4F1EA' : C.paper}, line:{color:C.border} });
          row.slice(0, sl.headers.length).forEach((cell, ci) => {
            s.addText(String(cell), { x: 0.65 + ci * (8.7/sl.headers.length), y, w: 8.7/sl.headers.length, h:0.38, fontSize:11, color:C.text, valign:'middle', margin:0 });
          });
        });
      } else if (sl.type === 'asks' && sl.asks) {
        sl.asks.forEach((a, i) => {
          const y = 1.55 + i * 1.15;
          s.addShape(pres.shapes.RECTANGLE, { x:0.5, y, w:0.55, h:0.55, fill:{color:[C.purple,C.teal,C.ink][i%3]}, line:{color:[C.purple,C.teal,C.ink][i%3]} });
          s.addText(String(i+1), { x:0.5, y, w:0.55, h:0.55, fontSize:22, bold:true, color:'FFFFFF', align:'center', valign:'middle', margin:0 });
          s.addText(a.question || '', { x:1.2, y:y-0.05, w:8.3, h:0.45, fontSize:14, bold:true, color:C.ink, valign:'middle', margin:0 });
          s.addText([
            { text: 'Why  →  ', options: { color:C.purple, bold:true } },
            { text: a.why || '',  options: { color:C.textMute } },
          ], { x:1.2, y:y+0.4, w:8.3, h:0.5, fontSize:11, valign:'top', margin:0 });
        });
      }
    }
  }

  // Return as Buffer (Node)
  const data = await pres.write({ outputType: 'nodebuffer' });
  return data;
}

// ════════════════════════════════════════════════════════════════
// VERCEL HANDLER
// ════════════════════════════════════════════════════════════════
module.exports = async (req, res) => {
  // CORS for browser → API
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(204).end();
  if (req.method !== 'POST')    return res.status(405).json({ error: 'POST only' });

  try {
    const { action, payload, project, settings, fbProjectId } = req.body || {};
    if (!action || !project) return res.status(400).json({ error: 'Missing action or project' });

    let result;
    switch (action) {
      case 'generate-deck':
        result = await generateDeck({ ...payload, project, settings, fbProjectId });
        break;
      case 'process-transcript':
        result = await processTranscript({ ...payload, project });
        break;
      case 'send-status':
        result = await sendStatus({ ...payload, project, settings });
        break;
      case 'quick-send':
        result = await quickSend({ ...payload, project, settings });
        break;
      default:
        return res.status(400).json({ error: 'Unknown action: ' + action });
    }
    return res.status(200).json(result);
  } catch (e) {
    console.error(e);
    return res.status(500).json({ error: e.message || String(e) });
  }
};
