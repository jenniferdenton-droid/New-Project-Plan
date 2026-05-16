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
// ACTION 5: SCHEDULE MEETING (Google Calendar event + invites)
// ════════════════════════════════════════════════════════════════
async function scheduleMeetingAction({ title, date, startTime, duration, description, location, attendees, sendEmail: doEmail, postSlack, recurrence, recurrenceEnd, addMeet, project, settings }) {
  if (!title || !date || !startTime) throw new Error('title, date, and startTime are required.');
  if (!attendees || attendees.length === 0) throw new Error('No attendees.');

  // Build start + end in ISO format (assume local time, no TZ conversion server-side)
  const startISO = `${date}T${startTime}:00`;
  const startDt = new Date(startISO);
  const endDt   = new Date(startDt.getTime() + (duration || 30) * 60000);
  const endISO  = endDt.toISOString().slice(0, 19);

  const result = { eventLink: null, meetLink: null, notifications: {} };

  // Build RRULE for recurrence (Google Calendar uses iCal RRULE format)
  // recurrence values: DAILY, WEEKLY, BIWEEKLY, MONTHLY, WEEKDAYS
  let rrule = null;
  if (recurrence && recurrence !== 'none') {
    const map = {
      DAILY:    'RRULE:FREQ=DAILY',
      WEEKLY:   'RRULE:FREQ=WEEKLY',
      BIWEEKLY: 'RRULE:FREQ=WEEKLY;INTERVAL=2',
      MONTHLY:  'RRULE:FREQ=MONTHLY',
      WEEKDAYS: 'RRULE:FREQ=WEEKLY;BYDAY=MO,TU,WE,TH,FR',
    };
    rrule = map[recurrence] || null;
    if (rrule && recurrenceEnd) {
      // UNTIL must be UTC in YYYYMMDDTHHMMSSZ format. We use end of day.
      const untilDate = new Date(recurrenceEnd + 'T23:59:59');
      const until = untilDate.toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '');
      rrule += `;UNTIL=${until}`;
    }
  }

  // 1) Create the Google Calendar event with invites
  try {
    const calendar = getCalendarClient();
    const event = {
      summary: `[${project.name}] ${title}`,
      description: `${description || ''}\n\n— Project: ${project.name}\n— Lead: ${project.lead || 'TBD'}\n— Project tracker: https://moxie-ops-project-plans.vercel.app/`.trim(),
      location: location || '',
      start: { dateTime: startISO, timeZone: 'America/New_York' },
      end:   { dateTime: endISO,   timeZone: 'America/New_York' },
      attendees: attendees.map(a => ({ email: a.email, displayName: a.name || undefined })),
      reminders: { useDefault: true },
    };
    if (rrule) event.recurrence = [rrule];
    // Auto-generate Google Meet link
    if (addMeet) {
      event.conferenceData = {
        createRequest: {
          requestId: 'moxie-meet-' + Date.now() + '-' + Math.random().toString(36).slice(2, 8),
          conferenceSolutionKey: { type: 'hangoutsMeet' },
        },
      };
    }
    const created = await calendar.events.insert({
      calendarId: 'primary',
      requestBody: event,
      sendUpdates: 'all',
      conferenceDataVersion: addMeet ? 1 : 0,
    });
    result.eventLink = created.data.htmlLink;
    // Extract Meet link if generated
    if (created.data.conferenceData && created.data.conferenceData.entryPoints) {
      const videoEntry = created.data.conferenceData.entryPoints.find(ep => ep.entryPointType === 'video');
      if (videoEntry) result.meetLink = videoEntry.uri;
    }
  } catch (e) {
    throw new Error('Calendar create failed: ' + (e.message || String(e)));
  }

  // 2) Optional confirmation email (in addition to native Calendar invite)
  if (doEmail) {
    const recipients = attendees.map(a => a.email);
    const dateFmt = startDt.toLocaleString('en-US', { weekday:'short', month:'short', day:'numeric', hour:'numeric', minute:'2-digit' });
    const recurText = recurrence && recurrence !== 'none'
      ? ({ DAILY:'Daily', WEEKLY:'Weekly', BIWEEKLY:'Every 2 weeks', MONTHLY:'Monthly', WEEKDAYS:'Every weekday (Mon–Fri)' }[recurrence] || recurrence)
        + (recurrenceEnd ? ` until ${recurrenceEnd}` : '')
      : null;
    const html = `
      <div style="font-family:-apple-system,Segoe UI,sans-serif;max-width:600px;color:#1A1A2E;">
        <h2 style="color:#1F1A47;margin-bottom:6px;">${escapeHtml(title)}</h2>
        <p style="color:#666;font-size:13px;margin-top:0;">${escapeHtml(project.name)} · ${escapeHtml(dateFmt)} · ${duration} min${recurText ? ' · 🔁 ' + escapeHtml(recurText) : ''}</p>
        ${description ? `<p style="font-size:14px;line-height:1.5;">${escapeHtml(description)}</p>` : ''}
        ${result.meetLink ? `<p style="margin:14px 0;"><a href="${result.meetLink}" style="background:#34A853;color:#fff;padding:10px 18px;border-radius:6px;text-decoration:none;font-weight:600;display:inline-block;">📹 Join Google Meet</a></p>` : ''}
        ${location && location !== result.meetLink ? `<p style="font-size:14px;"><strong>Location:</strong> ${escapeHtml(location)}</p>` : ''}
        <p style="margin:18px 0;"><a href="${result.eventLink}" style="background:#1565C0;color:#fff;padding:10px 18px;border-radius:6px;text-decoration:none;font-weight:600;">→ Open in Google Calendar</a></p>
        <p style="font-size:12px;color:#888;">You'll also receive a separate calendar invite from Google. — ${escapeHtml(settings.fromName || 'Jen')}</p>
      </div>`;
    try {
      await sendEmail({
        to: recipients, cc: [],
        subject: title,
        html,
        fromName: settings.fromName,
        projectName: project.name,
      });
      result.notifications.email = recipients.length;
    } catch (e) {
      // Don't fail the whole schedule if email fails — calendar invite already went
      console.warn('Confirmation email failed:', e.message);
    }
  }

  // 3) Optional Slack post
  if (postSlack) {
    try {
      const dateFmt = startDt.toLocaleString('en-US', { weekday:'short', month:'short', day:'numeric', hour:'numeric', minute:'2-digit' });
      const slackText = `:calendar: *${title}* — ${project.name}\n${dateFmt} · ${duration} min\n${description ? description + '\n' : ''}${location ? '*Location:* ' + location + '\n' : ''}${result.eventLink ? '<' + result.eventLink + '|Open in Calendar>' : ''}\n*Attendees:* ${attendees.map(a => a.name || a.email).join(', ')}`;
      await postToSlack({ text: slackText, settings });
      result.notifications.slack = true;
    } catch (e) {
      console.warn('Slack post failed:', e.message);
    }
  }

  return result;
}

// ════════════════════════════════════════════════════════════════
// ACTION 6: LIST CALENDAR EVENTS (pull from Google Calendar)
// Used to import existing meetings (including recurring) into the project.
// ════════════════════════════════════════════════════════════════
async function listCalendarEvents({ daysPast, daysFuture, search, expandRecurring, project }) {
  const calendar = getCalendarClient();
  const now = new Date();
  const timeMin = new Date(now.getTime() - (daysPast || 30) * 86400000).toISOString();
  const timeMax = new Date(now.getTime() + (daysFuture || 60) * 86400000).toISOString();

  // Pass 1: get expanded instances (singleEvents=true) — this returns REAL future
  // occurrences of recurring events with their actual dates, plus single events.
  const instancesRes = await calendar.events.list({
    calendarId: 'primary',
    timeMin, timeMax,
    singleEvents: true,
    orderBy: 'startTime',
    showDeleted: false,
    maxResults: 500,
    q: search || undefined,
  });

  // Pass 2: get the masters (singleEvents=false) just to read each series' RRULE
  // so we can show a human-readable recurrence label.
  const mastersRes = await calendar.events.list({
    calendarId: 'primary',
    timeMin, timeMax,
    singleEvents: false,
    showDeleted: false,
    maxResults: 250,
    q: search || undefined,
  });
  const masterRrule = {};
  (mastersRes.data.items || []).forEach(m => {
    if (m.recurrence && m.recurrence.length > 0) masterRrule[m.id] = m.recurrence[0];
  });

  const describeRrule = (r) => {
    if (!r) return 'Recurring';
    if (/FREQ=DAILY/.test(r)) return 'Daily';
    if (/FREQ=WEEKLY;INTERVAL=2/.test(r)) return 'Every 2 weeks';
    if (/FREQ=WEEKLY;BYDAY=MO,TU,WE,TH,FR/.test(r)) return 'Weekdays';
    if (/FREQ=WEEKLY/.test(r)) return 'Weekly';
    if (/FREQ=MONTHLY/.test(r)) return 'Monthly';
    return 'Recurring';
  };

  // For recurring series: pick the NEXT UPCOMING instance (>= now), not the earliest.
  // Falls back to the most recent past instance only if no future one exists in the window.
  const nowMs = Date.now();
  const seriesInstances = {};
  const items = [];

  function mapEvent(ev, seriesId) {
    const startObj = ev.start || {};
    return {
      id: ev.id,
      title: ev.summary || '(no title)',
      description: ev.description || '',
      location: ev.location || '',
      startISO: startObj.dateTime || startObj.date || '',
      attendees: (ev.attendees || []).map(a => ({ email: a.email, name: a.displayName || '' })),
      isRecurring: !!seriesId,
      recurDesc: seriesId ? describeRrule(masterRrule[seriesId]) : '',
      meetLink: (ev.conferenceData && ev.conferenceData.entryPoints || [])
        .filter(e => e.entryPointType === 'video').map(e => e.uri)[0] || '',
      htmlLink: ev.htmlLink || '',
      organizer: (ev.organizer && (ev.organizer.email || ev.organizer.displayName)) || '',
    };
  }

  (instancesRes.data.items || []).forEach(ev => {
    const seriesId = ev.recurringEventId || null;
    if (seriesId && !expandRecurring) {
      // Default mode: group by series so we can pick the next-upcoming
      if (!seriesInstances[seriesId]) seriesInstances[seriesId] = [];
      seriesInstances[seriesId].push(ev);
    } else {
      // Single event OR expandRecurring mode: pass every instance through as its own item
      items.push(mapEvent(ev, seriesId));
    }
  });

  // If NOT expanding, dedupe each recurring series to its next-upcoming instance
  if (!expandRecurring) {
    Object.keys(seriesInstances).forEach(sid => {
      const arr = seriesInstances[sid];
      const future = arr.find(ev => {
        const start = ev.start && (ev.start.dateTime || ev.start.date);
        return start && new Date(start).getTime() >= nowMs;
      });
      const chosen = future || arr[arr.length - 1];
      items.push(mapEvent(chosen, sid));
    });
  }

  items.sort((a, b) => (a.startISO || '').localeCompare(b.startISO || ''));

  return {
    events: items,
    serverNow: new Date().toISOString(),       // for debugging: confirms fresh code
    futureLogicVersion: 'v2-pick-next-upcoming',// bumped when the dedupe logic changes
  };
}

// ════════════════════════════════════════════════════════════════
// ACTION 7: PIN PROJECT TO SLACK CHANNEL
// Posts a rich project info card to the channel. Attempts to pin and set
// channel topic if optional scopes are available — gracefully falls back.
// ════════════════════════════════════════════════════════════════
// ════════════════════════════════════════════════════════════════
// ASK THE DASHBOARD — natural-language Q&A grounded in project data
// ════════════════════════════════════════════════════════════════
async function askDashboard({ question, project }) {
  if (!question || !question.trim()) throw new Error('Question is required.');
  const t0 = Date.now();

  // Compress project to a focused snapshot Claude can reason over (token-budget-friendly).
  const snapshot = buildProjectSnapshot(project);

  const system = `You are a sharp project analyst assistant for a SaaS company called Moxie (MedSpa platform).
You answer questions about the user's project using ONLY the JSON snapshot they provide.
- Be direct and concise — exec-ready, no fluff.
- Use bullet points or short paragraphs as appropriate.
- Quote specific task names, owner names, and dates when supporting an answer.
- If the data doesn't contain the answer, say so plainly — don't speculate.
- When asked to summarize or draft, keep it under 6 sentences unless the user asks for more.
- Today's date is ${new Date().toISOString().split('T')[0]}.`;

  const user = `PROJECT SNAPSHOT:
\`\`\`json
${JSON.stringify(snapshot, null, 2)}
\`\`\`

QUESTION:
${question.trim()}`;

  const msg = await anthropic.messages.create({
    model: 'claude-sonnet-4-5',
    max_tokens: 1500,
    system,
    messages: [{ role: 'user', content: user }],
  });

  const answer = (msg.content && msg.content[0] && msg.content[0].text) || '';
  const tokensUsed = (msg.usage?.input_tokens || 0) + (msg.usage?.output_tokens || 0);
  return {
    answer,
    tokensUsed,
    elapsedMs: Date.now() - t0,
  };
}

// ════════════════════════════════════════════════════════════════
// AUTO-RAG — recommend Green/Amber/Red with reasoning
// ════════════════════════════════════════════════════════════════
async function suggestRag({ project }) {
  const t0 = Date.now();

  // Compute deterministic signals first — these give Claude something to ground on
  const today = new Date(); today.setHours(0,0,0,0);
  const todayMs = today.getTime();
  const dueMs = project.dueDate ? new Date(project.dueDate + 'T00:00:00').getTime() : null;
  const daysToDue = dueMs ? Math.round((dueMs - todayMs) / 86400000) : null;

  const tasks = project.tasks || [];
  const ais = project.actionItems || [];
  const risks = project.risks || [];
  const checklist = project.launchChecklist || [];

  const pastDueTasks = tasks.filter(t => t.dueDate && t.status !== 'Done' && new Date(t.dueDate+'T00:00:00').getTime() < todayMs);
  const pastDueAIs   = ais.filter(a => a.dueDate && a.status !== 'Done' && new Date(a.dueDate+'T00:00:00').getTime() < todayMs);
  const blockedTasks = tasks.filter(t => t.status === 'Blocked');
  const atRiskTasks  = tasks.filter(t => t.status === 'At Risk' || t.flaggedRisk);
  const criticalAIs  = ais.filter(a => a.status !== 'Done' && a.priority === 'Critical');
  const openRisks    = risks.filter(r => r.status === 'Open' || r.status === 'Monitoring' || !r.status);
  const highRisks    = openRisks.filter(r => (r.impact || '').toLowerCase() === 'high');
  const taskDone     = tasks.filter(t => t.status === 'Done').length;
  const checklistDone = checklist.filter(i => i.checked).length;
  const checklistTotal = checklist.filter(i => !i.na).length;
  const requiredOpen = checklist.filter(i => i.required && !i.checked && !i.na).length;

  const signals = {
    daysToDue,
    totalTasks: tasks.length,
    tasksDone: taskDone,
    pastDueTasks: pastDueTasks.length,
    pastDueAIs: pastDueAIs.length,
    blockedTasks: blockedTasks.length,
    atRiskTasks: atRiskTasks.length,
    criticalAIs: criticalAIs.length,
    openRisks: openRisks.length,
    highImpactRisks: highRisks.length,
    launchChecklistProgress: `${checklistDone}/${checklistTotal}`,
    requiredChecklistItemsOpen: requiredOpen,
  };

  const system = `You are a project health analyst. Based on the signals provided, recommend a RAG (Red/Amber/Green) health rating with a clear, defensible one-line reason and 3-5 key factors.

Use these rules:
- GREEN: 0 past-due tasks, 0 critical AIs, ≤1 high-impact risk, on-track for due date, no blockers.
- AMBER: Some past-due items (≤3), 1-2 critical AIs, 2-3 high-impact risks, OR launch checklist behind required gates, OR <30 days to launch with required items still open.
- RED: Multiple past-due (>3), 3+ critical AIs, blockers preventing progress, OR past the due date with required items still open, OR 4+ high-impact risks.

Respond with ONLY valid JSON in this exact shape:
{
  "status": "Green" | "Amber" | "Red",
  "reason": "one-line summary (max 100 chars)",
  "factors": ["factor 1", "factor 2", "factor 3"]
}`;

  const user = `PROJECT: ${project.name || 'Unnamed'}
DUE DATE: ${project.dueDate || 'TBD'} (${daysToDue !== null ? daysToDue + ' days from today' : 'no due date set'})

SIGNALS:
${JSON.stringify(signals, null, 2)}

Recommend a RAG status with reason and factors.`;

  const msg = await anthropic.messages.create({
    model: 'claude-sonnet-4-5',
    max_tokens: 600,
    system,
    messages: [{ role: 'user', content: user }],
  });

  const raw = (msg.content && msg.content[0] && msg.content[0].text) || '{}';
  // Strip ```json fences if Claude wrapped them
  const cleaned = raw.replace(/^```(?:json)?\s*/i, '').replace(/\s*```\s*$/i, '').trim();
  let parsed;
  try {
    parsed = JSON.parse(cleaned);
  } catch (e) {
    // Fallback: try to extract first JSON object
    const m = cleaned.match(/\{[\s\S]*\}/);
    parsed = m ? JSON.parse(m[0]) : { status: 'Amber', reason: 'Unable to parse AI response.', factors: [] };
  }
  return {
    status: parsed.status || 'Amber',
    reason: parsed.reason || '',
    factors: parsed.factors || [],
    signals,
    elapsedMs: Date.now() - t0,
  };
}

// ════════════════════════════════════════════════════════════════
// Shared Claude helper for the Tier 2 / Tier 3 features below
// ════════════════════════════════════════════════════════════════
async function callClaude(system, user, maxTokens = 2000) {
  const msg = await anthropic.messages.create({
    model: 'claude-sonnet-4-5',
    max_tokens: maxTokens,
    system,
    messages: [{ role: 'user', content: user }],
  });
  return (msg.content && msg.content[0] && msg.content[0].text) || '';
}
// Strip ```json fences if present, then JSON.parse defensively
function parseJsonSafe(raw, fallback) {
  const cleaned = String(raw || '').replace(/^```(?:json)?\s*/i, '').replace(/\s*```\s*$/i, '').trim();
  try { return JSON.parse(cleaned); }
  catch {
    const m = cleaned.match(/[\[{][\s\S]*[\]}]/);
    if (m) { try { return JSON.parse(m[0]); } catch {} }
    return fallback;
  }
}

// ════════════════════════════════════════════════════════════════
// MEETING PREP BRIEF (Tier 2)
// ════════════════════════════════════════════════════════════════
async function meetingPrepBrief({ meetingId, project }) {
  const t0 = Date.now();
  const meeting = (project.meetings || []).find(m => String(m.id) === String(meetingId));
  if (!meeting) throw new Error('Meeting not found in project state.');
  const snapshot = buildProjectSnapshot(project);

  const system = `You are an executive assistant preparing a 1-page meeting prep brief.
Today's date: ${new Date().toISOString().split('T')[0]}.
Output a clean, scannable plain-text brief (no markdown headers, just labels and bullets).
Keep it tight — designed to skim in 60 seconds.

Required sections:
1) MEETING — title, date, attendees
2) PURPOSE & SUGGESTED AGENDA (3-5 items, infer from project state if not stated)
3) WHAT'S CHANGED SINCE LAST UPDATE (3 bullets max — completed tasks, key decisions, status shifts)
4) ITEMS REQUIRING ATTENDEE INPUT (open AIs/blockers tied to attendees by team)
5) TOP RISKS TO FLAG (max 3)
6) DECISIONS NEEDED IN THIS MEETING (concrete, with proposed options)`;

  const user = `MEETING:
${JSON.stringify(meeting, null, 2)}

PROJECT SNAPSHOT:
\`\`\`json
${JSON.stringify(snapshot, null, 2)}
\`\`\``;

  const brief = await callClaude(system, user, 1800);
  return { brief, elapsedMs: Date.now() - t0 };
}

// ════════════════════════════════════════════════════════════════
// AUTO-CATEGORIZE TASKS (Tier 2)
// ════════════════════════════════════════════════════════════════
async function suggestTaskTeams({ taskIds, project }) {
  const t0 = Date.now();
  const teamLabels = ['PSM','Sales','Ops Leader','Onboarding','MD','Supplies','Billing','Marketing','Internal Comms','Product','Rev Ops','Biz Ops','External','Senior Leadership Owner'];
  const tasks = (project.tasks || []).filter(t => (taskIds || []).map(String).includes(String(t.id)));
  if (!tasks.length) return { suggestions: [], elapsedMs: Date.now() - t0 };

  const stakeholders = Object.entries(project.stakeholders || {})
    .filter(([_, s]) => s && s.included)
    .map(([k, s]) => ({ team: k, contact: s.contactName || '' }));

  const system = `You assign tasks to the right team. Available teams (use these labels EXACTLY):
${teamLabels.map(t => '- ' + t).join('\n')}

For each task, infer the best team based on the task description and notes. Use the stakeholder roster as a tiebreaker.

Respond with ONLY a JSON array, no prose:
[
  { "taskId": <number>, "suggestedTeam": "<one of the teams above>", "confidence": "High"|"Med"|"Low", "reason": "<one short sentence>" },
  ...
]`;

  const user = `STAKEHOLDER ROSTER:
${JSON.stringify(stakeholders, null, 2)}

TASKS TO CATEGORIZE:
${JSON.stringify(tasks.map(t => ({ id: t.id, task: t.task, notes: t.notes, currentOwner: t.owner })), null, 2)}

Return your team suggestions as JSON array now.`;

  const raw = await callClaude(system, user, 2000);
  const suggestions = parseJsonSafe(raw, []);
  return {
    suggestions: Array.isArray(suggestions) ? suggestions : [],
    elapsedMs: Date.now() - t0,
  };
}

// ════════════════════════════════════════════════════════════════
// SOP DRAFT (Tier 2)
// ════════════════════════════════════════════════════════════════
async function draftSop({ focus, project }) {
  const t0 = Date.now();
  const snapshot = buildProjectSnapshot(project);

  const system = `You write Standard Operating Procedures (SOPs) for SaaS operations teams.
Output a clean, well-structured SOP in plain text (NOT markdown). Use ALL-CAPS section headers and numbered steps.

Required sections (adapt to what the project state contains):
1) PURPOSE (2-3 sentences)
2) WHEN TO USE THIS SOP (triggers)
3) OWNER & ESCALATION PATH
4) STEP-BY-STEP PROCEDURE (numbered, concrete, action-verb-led)
5) INPUTS / TOOLS REQUIRED
6) OUTPUTS / DELIVERABLES
7) SUCCESS CRITERIA
8) FAQ / EDGE CASES (3-5 anticipated questions)
9) REVISION TRIGGERS

Keep it operational — actionable for someone who has never run this process before.`;

  const focusLine = focus ? `\nSOP FOCUS REQUESTED BY USER: ${focus}\nWrite the SOP narrowly for this focus area, not the whole project.\n` : '';

  const user = `${focusLine}
PROJECT SNAPSHOT (use as source material):
\`\`\`json
${JSON.stringify(snapshot, null, 2)}
\`\`\`

Now draft the SOP.`;

  const sop = await callClaude(system, user, 2500);
  return { sop, elapsedMs: Date.now() - t0 };
}

// ════════════════════════════════════════════════════════════════
// DAILY STANDUP DRAFT (Tier 3)
// ════════════════════════════════════════════════════════════════
async function draftStandup({ project }) {
  const t0 = Date.now();
  const snapshot = buildProjectSnapshot(project);

  const system = `You write daily standup updates. Today's date: ${new Date().toISOString().split('T')[0]}.

Output a standup grouped by OWNER (team or person). For each owner show:
*Owner Name*
:white_check_mark: Yesterday: <2-3 bullets of recent completions or in-progress momentum>
:dart: Today: <2-3 bullets of upcoming work — due today or this week>
:no_entry: Blockers: <bullets of blocked tasks / open critical AIs — or "None">

Use Slack mrkdwn formatting (*bold*, :emoji:, • bullets). Keep it scannable — total length under 600 words. Only include owners with relevant items.`;

  const user = `PROJECT SNAPSHOT:
\`\`\`json
${JSON.stringify(snapshot, null, 2)}
\`\`\`

Write today's standup.`;

  const standup = await callClaude(system, user, 1500);
  return { standup, elapsedMs: Date.now() - t0 };
}

// ════════════════════════════════════════════════════════════════
// DECISION EXTRACTION FROM THREAD (Tier 3)
// ════════════════════════════════════════════════════════════════
async function extractDecisions({ threadText, project }) {
  const t0 = Date.now();
  if (!threadText || !threadText.trim()) throw new Error('Thread text is required.');

  const system = `You extract DECISIONS from conversations (Slack threads, meeting transcripts, emails).

A decision is a concrete commitment or choice — NOT a question, NOT an opinion, NOT a task assignment.

Examples of decisions:
- "We're going with vendor X over vendor Y"
- "Launch is moving to June 15"
- "Pricing will be tier-based, $50/$100/$200"

NOT decisions (skip these):
- "We should think about pricing" (question/opinion)
- "Jenn will draft the doc" (task — goes elsewhere)
- "I'm worried about timeline" (concern)

Respond with ONLY a JSON array, no prose:
[
  { "decision": "<one-sentence decision statement>", "decidedBy": "<person or team>", "date": "YYYY-MM-DD (best guess from context)", "context": "<one-sentence context, optional>" },
  ...
]
If no decisions found, return [].`;

  const user = `CONVERSATION TO ANALYZE:
"""
${threadText}
"""

Extract decisions as JSON array now.`;

  const raw = await callClaude(system, user, 1500);
  const decisions = parseJsonSafe(raw, []);
  return {
    decisions: Array.isArray(decisions) ? decisions : [],
    elapsedMs: Date.now() - t0,
  };
}

// ════════════════════════════════════════════════════════════════
// ONBOARDING BRIEF (Tier 3)
// ════════════════════════════════════════════════════════════════
async function onboardingBrief({ role, project }) {
  const t0 = Date.now();
  const snapshot = buildProjectSnapshot(project);

  const system = `You write CATCH-UP BRIEFS for project team members joining mid-stream. Today: ${new Date().toISOString().split('T')[0]}.

The reader has zero context. Goal: get them productive in 15 minutes of reading.

Required sections (plain text, ALL-CAPS headers):
1) PROJECT IN ONE PARAGRAPH (what, why, target launch)
2) CURRENT HEALTH (RAG status + one-line reason)
3) WHO'S WHO (stakeholder teams + key contacts, 5-8 names max)
4) WHAT'S BEEN DECIDED (3-5 most important decisions, with brief context)
5) WHERE THE WORK STANDS (Phase / milestone breakdown — what's done, what's in flight, what's next)
6) OPEN RISKS & BLOCKERS (top 3-5, with owners)
7) WHAT THIS PERSON CAN PICK UP THIS WEEK (concrete suggestions based on their role)
8) RESOURCES (key links, channels, docs)`;

  const roleLine = role ? `READER'S ROLE: ${role}\nTailor "WHAT THIS PERSON CAN PICK UP" specifically for this role.\n` : '';

  const user = `${roleLine}
PROJECT SNAPSHOT:
\`\`\`json
${JSON.stringify(snapshot, null, 2)}
\`\`\`

Write the catch-up brief.`;

  const brief = await callClaude(system, user, 2500);
  return { brief, elapsedMs: Date.now() - t0 };
}

// ════════════════════════════════════════════════════════════════
// EMAIL RESPONSE DRAFTER (Tier 3)
// ════════════════════════════════════════════════════════════════
async function draftEmailReply({ incoming, project }) {
  const t0 = Date.now();
  if (!incoming || !incoming.trim()) throw new Error('Incoming email is required.');
  const snapshot = buildProjectSnapshot(project);

  const system = `You draft email replies on behalf of the project lead. Today: ${new Date().toISOString().split('T')[0]}.

Rules:
- Lead with the answer / bottom line — exec-ready, no fluff.
- Use specifics from the project state (real task names, dates, owners) when they answer the question.
- Match the tone of the incoming email (formal vs casual).
- Keep it tight — 3-6 sentences for most replies. Use bullets only if listing more than 3 items.
- If the question can't be answered from the data, say so honestly and propose a next step.
- End with a clear next-step or sign-off. Don't say "Best," — leave that to the sender.

Output ONLY the reply body. No subject line. No "Hi <name>," — start with the substance.`;

  const user = `INCOMING EMAIL:
"""
${incoming}
"""

PROJECT SNAPSHOT (use for facts):
\`\`\`json
${JSON.stringify(snapshot, null, 2)}
\`\`\`

Draft the reply.`;

  const reply = await callClaude(system, user, 1500);
  return { reply, elapsedMs: Date.now() - t0 };
}

// ════════════════════════════════════════════════════════════════
// PROPOSAL BUILDER — generates a proposal in Moxie's standard format
// (Modeled after the Bookkeeping Process Remediation Proposal template)
// ════════════════════════════════════════════════════════════════
async function buildProposal({ context, project }) {
  const t0 = Date.now();
  const snapshot = buildProjectSnapshot(project);

  const system = `You write project proposals for Moxie (MedSpa SaaS platform) in their established template format. Output a complete, paste-ready proposal in plain text. Today: ${new Date().toISOString().split('T')[0]}.

═══════ MOXIE PROPOSAL TEMPLATE — FOLLOW EXACTLY ═══════

Use these Roman-numeral sections IN THIS ORDER:

[Cover block]
MOXIE
PROJECT PROPOSAL — <PROJECT NAME>
Date: <today>  |  Prepared by: <project lead>

I. EXECUTIVE SUMMARY
   2-4 short paragraphs. Open with one-sentence thesis of the problem, then the proposed approach. Name the workstreams (2-3 max). List the team. ~120 words.

II. CURRENT STATE
   Categorized inventory of what is broken today, in ALL-CAPS labeled buckets like:
   DATA — what's broken about data today.
   PROCESS — what's broken about process today.
   BILLING — etc.
   COMPLIANCE — etc.
   Use 4-6 buckets that fit the project. Each 1-2 sentences. Plainspoken, no spin.

III. PROBLEM STATEMENT
   One-sentence root cause. Then 5-7 bullets listing downstream consequences. Direct, no hedging.

IV. GOALS / OBJECTIVES
   5-8 bulleted outcome statements, each starting with an action verb (Reduce, Eliminate, Establish, Replace, Document, etc.). Measurable where possible.

V. PROPOSED NEW PROCESS
   Three-column comparison rendered in plain text:
   AREA | CURRENT (Broken) | PROPOSED
   <repeated rows>
   Then a "NEW COMMUNICATION FRAMEWORK" subsection — 2-3 paragraphs on the new operating model.

VI. SCOPE OF WORK
   Two workstreams with priority tables:

   WORKSTREAM 1 — IMMEDIATE FIXES
   PRIORITY | ACTION | OWNER | DUE DATE
   <5-7 rows: CRITICAL / HIGH / MEDIUM>

   WORKSTREAM 2 — PROCESS REBUILD
   PRIORITY | ACTION | OWNER | DUE DATE
   <5-7 rows>

VII. TIMETABLE
   Phase table:
   PHASE | FOCUS | TARGET DATES
   <4-5 rows>

VIII. KEY PERSONNEL
   ROLE | NAME | RESPONSIBILITIES
   <one row per stakeholder team that's actively engaged>

IX. EVALUATION
   5-7 bulleted measurable success criteria with dates and percentages where possible.

X. RISKS
   RISK | SEVERITY (High/Med/Low) | MITIGATION
   <5-7 rows>

XI. NEXT STEPS
   Numbered immediate action list with owners and dates. 7-10 items.

XII. APPENDIX
   Bulleted list of supporting artifacts (links to relevant docs, dashboards, contracts).

[Signature block]
Project Lead: ___________________  Date: __________
Sponsor:       ___________________  Date: __________
Stakeholder:   ___________________  Date: __________

═══════ STYLE RULES ═══════
- Use Roman numerals (I, II, III...) for section headers.
- Use ALL-CAPS for category labels and table headers.
- Plain text only — NO markdown, NO bold syntax, NO asterisks.
- Em-dashes for clarifying lists.
- Every action item has a named owner (a team from the stakeholder list) and a concrete date.
- Tables: pipe-delimited (use ASCII | character), monospace-friendly.
- Moxie/MedSpa vocabulary: "providers" (not customers), "PSM" (Provider Success Manager), use real team names from stakeholder list.
- Formal, executive-facing, plainspoken. Acknowledge problems without spin.
- Section length: most sections ~100-250 words.
- Operational specificity (names, dates, percentages) under executive framing.

═══════ REPRESENTATIVE VOICE EXAMPLES ═══════
"The bookkeeping product launched and is live, but the operational layer beneath it was never fully built. What exists today is a fragile, manual process held together by Slack messages, individual tribal knowledge, and workarounds."
"The PSM team is absorbing manual work that should be automated — contract routing, kickoff scheduling, feature flag activation — at the expense of provider relationship quality."
"Replace Slack-only coordination with a formal operating model — defined escalation paths, documented decisions, and SLAs for issue resolution."

Now: read the project snapshot and additional context, then output the FULL proposal in the template above. Be concrete — pull real task names, real owners, real dates from the snapshot.`;

  const user = `${context ? 'ADDITIONAL CONTEXT FROM USER:\n' + context + '\n\n' : ''}PROJECT SNAPSHOT (source of truth):
\`\`\`json
${JSON.stringify(snapshot, null, 2)}
\`\`\`

Now write the complete proposal. Plain text only, no markdown.`;

  const proposal = await callClaude(system, user, 4000);
  return { proposal, elapsedMs: Date.now() - t0 };
}

// ════════════════════════════════════════════════════════════════
// READ SLACK CHANNEL — pull last N days of messages, extract candidate tasks/decisions/risks
// ════════════════════════════════════════════════════════════════
// Needed Slack scopes: channels:history (public), groups:history (private), channels:read
async function readSlackChannel({ channel, days, project }) {
  const t0 = Date.now();
  const botToken = process.env.SLACK_BOT_TOKEN;
  if (!botToken) throw new Error('Slack bot token not configured.');
  if (!channel) throw new Error('Channel is required.');
  const lookbackDays = Math.max(1, Math.min(30, parseInt(days, 10) || 7));

  // Resolve channel ID — accept #name, name, or raw ID
  const cleanName = String(channel).replace(/^#/, '').trim();
  let channelId = cleanName;
  if (!/^C[A-Z0-9]+$/i.test(cleanName)) {
    const listRes = await fetch('https://slack.com/api/conversations.list?types=public_channel,private_channel&limit=1000', {
      headers: { 'Authorization': `Bearer ${botToken}` },
    });
    const listJson = await listRes.json().catch(() => ({}));
    if (!listJson.ok) throw new Error('Slack conversations.list failed: ' + (listJson.error || 'unknown'));
    const found = (listJson.channels || []).find(c => c.name === cleanName);
    if (!found) throw new Error(`Channel #${cleanName} not found or bot isn't in it. Invite the bot first.`);
    channelId = found.id;
  }

  // Pull messages (Slack returns most recent first)
  const oldest = Math.floor((Date.now() - lookbackDays * 86400000) / 1000);
  const histRes = await fetch(`https://slack.com/api/conversations.history?channel=${channelId}&oldest=${oldest}&limit=500`, {
    headers: { 'Authorization': `Bearer ${botToken}` },
  });
  const histJson = await histRes.json().catch(() => ({}));
  if (!histJson.ok) throw new Error('Slack history failed: ' + (histJson.error || 'unknown') + ' — ensure bot has channels:history scope and is in the channel.');

  const messages = (histJson.messages || [])
    .filter(m => m.type === 'message' && !m.subtype && m.text)
    .map(m => ({ user: m.user || m.username || '', ts: m.ts, text: m.text }));

  if (!messages.length) {
    return { candidates: { tasks: [], decisions: [], risks: [] }, messageCount: 0, elapsedMs: Date.now() - t0 };
  }

  // Map user IDs to names (best-effort; cache could be added later)
  const userIds = [...new Set(messages.map(m => m.user).filter(u => u && /^U/.test(u)))];
  const userNames = {};
  for (const uid of userIds.slice(0, 50)) {
    try {
      const uRes = await fetch(`https://slack.com/api/users.info?user=${uid}`, { headers: { 'Authorization': `Bearer ${botToken}` } });
      const uJson = await uRes.json().catch(() => ({}));
      if (uJson.ok && uJson.user) userNames[uid] = uJson.user.real_name || uJson.user.name || uid;
    } catch (e) { /* skip */ }
  }
  const labeled = messages.map(m => ({
    speaker: userNames[m.user] || m.user || 'unknown',
    when: new Date(parseFloat(m.ts) * 1000).toISOString().split('T')[0],
    text: m.text,
  }));

  // Snapshot for grounding — Claude only proposes items NOT already in the project
  const snapshot = buildProjectSnapshot(project);
  const existingTasks = (snapshot.tasks || []).map(t => t.task).filter(Boolean);
  const existingDecisions = (snapshot.decisions || []).map(d => d.decision).filter(Boolean);

  const system = `You scan Slack conversations and propose project-tracker additions. Today: ${new Date().toISOString().split('T')[0]}.

For each candidate item, classify it as one of: TASK (work to do, with an owner if mentioned), DECISION (a concrete choice/commitment), or RISK (something that could go wrong / blocker / concern).

CRITICAL RULES:
- Do NOT propose items that duplicate things ALREADY in the project (lists below).
- Do NOT propose chit-chat, questions, or generic discussion as tasks.
- Owner names: map to a Slack speaker name when possible.
- Dates: extract from natural language ("by Friday", "next week" → calendar date).
- If nothing concrete is found, return empty arrays.

Respond with ONLY valid JSON in this exact shape:
{
  "tasks": [ { "text": "<concrete task>", "owner": "<name or team>", "dueDate": "YYYY-MM-DD or empty", "context": "<one-line context>" } ],
  "decisions": [ { "text": "<decision>", "decidedBy": "<name>", "date": "YYYY-MM-DD", "context": "<one-line>" } ],
  "risks": [ { "text": "<risk>", "owner": "<name>", "impact": "High|Med|Low", "mitigation": "<plan or empty>" } ]
}`;

  const user = `PROJECT: ${project.name || 'Unnamed'}

ALREADY IN PROJECT — DO NOT PROPOSE DUPLICATES:
Tasks: ${JSON.stringify(existingTasks.slice(0, 50))}
Decisions: ${JSON.stringify(existingDecisions.slice(0, 30))}

SLACK MESSAGES (last ${lookbackDays} days, oldest first):
${JSON.stringify(labeled.slice(0, 200), null, 2)}

Now propose new candidate tasks, decisions, and risks. Return JSON only.`;

  const raw = await callClaude(system, user, 2500);
  const parsed = parseJsonSafe(raw, { tasks: [], decisions: [], risks: [] });
  return {
    candidates: {
      tasks: Array.isArray(parsed.tasks) ? parsed.tasks : [],
      decisions: Array.isArray(parsed.decisions) ? parsed.decisions : [],
      risks: Array.isArray(parsed.risks) ? parsed.risks : [],
    },
    messageCount: messages.length,
    channelId,
    elapsedMs: Date.now() - t0,
  };
}

// ════════════════════════════════════════════════════════════════
// PM AI SCAN — health score + duplicate detection + gap analysis
// ════════════════════════════════════════════════════════════════
async function pmAiScan({ project, settings, lastScanAt }) {
  const t0 = Date.now();
  const snapshot = buildProjectSnapshot(project);

  // ── Bonus pass: if a Slack channel is configured, also scan it since the last
  //    health check for candidate tasks/risks the user may have missed.
  let slackCandidates = null;
  const slackChannel = (settings && settings.slackChannel) || (project.aiSettings && project.aiSettings.slackChannel) || '';
  if (slackChannel && process.env.SLACK_BOT_TOKEN) {
    try {
      // Default to 7 days back if no prior scan timestamp
      const sinceMs = lastScanAt ? new Date(lastScanAt).getTime() : (Date.now() - 7 * 86400000);
      const daysBack = Math.max(1, Math.min(30, Math.ceil((Date.now() - sinceMs) / 86400000)));
      const slackRes = await readSlackChannel({ channel: slackChannel, days: daysBack, project });
      slackCandidates = {
        channel: slackChannel.replace(/^#/, ''),
        daysScanned: daysBack,
        messageCount: slackRes.messageCount || 0,
        tasks: slackRes.candidates?.tasks || [],
        decisions: slackRes.candidates?.decisions || [],
        risks: slackRes.candidates?.risks || [],
      };
    } catch (e) {
      // Don't fail the whole scan if Slack errors
      slackCandidates = { error: e.message };
    }
  }

  // Fast deterministic pass: find candidate duplicate pairs via token-overlap similarity
  const tasks = (project.tasks || []).filter(t => t.task && t.task.trim());
  const tokenize = s => String(s || '').toLowerCase().replace(/[^a-z0-9 ]/g, ' ').split(/\s+/).filter(w => w.length > 3);
  const candidatePairs = [];
  for (let i = 0; i < tasks.length; i++) {
    const tiTokens = new Set(tokenize(tasks[i].task));
    if (tiTokens.size < 2) continue;
    for (let j = i + 1; j < tasks.length; j++) {
      const tjTokens = new Set(tokenize(tasks[j].task));
      if (tjTokens.size < 2) continue;
      let overlap = 0;
      tiTokens.forEach(w => { if (tjTokens.has(w)) overlap++; });
      const sim = overlap / Math.min(tiTokens.size, tjTokens.size);
      if (sim >= 0.6) candidatePairs.push({ a: tasks[i].id, b: tasks[j].id, sim });
    }
  }

  // Deterministic gap signals
  const today = new Date(); today.setHours(0,0,0,0);
  const todayMs = today.getTime();
  const tasksNoOwner   = tasks.filter(t => !t.owner || !t.owner.trim());
  const tasksNoDueDate = tasks.filter(t => !t.dueDate && t.status !== 'Done');
  const pastDue        = tasks.filter(t => t.dueDate && t.status !== 'Done' && new Date(t.dueDate+'T00:00:00').getTime() < todayMs);
  const blocked        = tasks.filter(t => t.status === 'Blocked');
  const openCritAIs    = (project.actionItems || []).filter(a => a.status !== 'Done' && a.priority === 'Critical');
  const openHighRisks  = (project.risks || []).filter(r => (r.status === 'Open' || !r.status) && (r.impact || '').toLowerCase() === 'high');

  const system = `You are an experienced PM auditing a project tracker. Analyze the project state and return ONLY valid JSON in this exact shape:
{
  "healthScore": 0-100 (overall project health),
  "healthSummary": "one-sentence explanation",
  "duplicates": [
    { "ids": [<task id>, <task id>], "reason": "<why these look like duplicates>", "recommendation": "<merge|keep both|clarify>" }
  ],
  "gaps": [
    { "type": "<category>", "message": "<specific issue>", "count": <optional number> }
  ],
  "suggestions": ["<actionable next step>", ...]
}

Rules for healthScore:
- 80-100 (Healthy): on track, few past-due/blockers, owners assigned, risks managed.
- 60-79 (Watch): some past-due or blockers, gaps in ownership/dates, manageable risks.
- 0-59 (At risk): many past-due, multiple blockers/critical AIs, missing owners, high-impact risks unmitigated.

Rules for duplicates:
- Consider the candidate pairs the user-side similarity check identified, AND any others you spot.
- Only include true duplicates (same work, redundant) — NOT related tasks that should both exist.
- Recommendation: "merge" (keep one), "keep both" (false positive — explain), "clarify" (rename to distinguish).

Rules for gaps: surface ONLY real issues — be specific and quantify ("8 tasks have no owner", "3 tasks past due >7 days", etc.).

Suggestions: 3-6 prioritized, actionable next-step bullets the PM can take TODAY.`;

  const user = `PROJECT SNAPSHOT:
\`\`\`json
${JSON.stringify(snapshot, null, 2)}
\`\`\`

CANDIDATE DUPLICATE PAIRS (token-overlap >= 60%):
${JSON.stringify(candidatePairs, null, 2)}

DETERMINISTIC SIGNALS:
- Tasks without owner: ${tasksNoOwner.length}
- Tasks without due date (not Done): ${tasksNoDueDate.length}
- Past-due tasks: ${pastDue.length}
- Blocked tasks: ${blocked.length}
- Critical open action items: ${openCritAIs.length}
- High-impact open risks: ${openHighRisks.length}

Now run the PM audit. Return ONLY the JSON object — no prose.`;

  const raw = await callClaude(system, user, 2500);
  const parsed = parseJsonSafe(raw, {
    healthScore: 60, healthSummary: 'Unable to parse AI response.',
    duplicates: [], gaps: [], suggestions: [],
  });

  return {
    healthScore: Math.max(0, Math.min(100, parseInt(parsed.healthScore, 10) || 60)),
    healthSummary: parsed.healthSummary || '',
    duplicates: Array.isArray(parsed.duplicates) ? parsed.duplicates : [],
    gaps: Array.isArray(parsed.gaps) ? parsed.gaps : [],
    suggestions: Array.isArray(parsed.suggestions) ? parsed.suggestions : [],
    candidatePairs,
    slackCandidates,
    elapsedMs: Date.now() - t0,
  };
}

// ════════════════════════════════════════════════════════════════
// PROJECT UPDATE — Claude writes a structured status update
// (gated behind a fresh pm-ai-scan; uses its findings as context)
// ════════════════════════════════════════════════════════════════
async function projectUpdate({ audience, healthScan, lastScanFindings, project }) {
  const t0 = Date.now();
  const snapshot = buildProjectSnapshot(project);

  const today = new Date(); today.setHours(0,0,0,0);
  const todayMs = today.getTime();
  const weekAgoMs = todayMs - 7 * 86400000;
  const weekAheadMs = todayMs + 7 * 86400000;

  const completed = (project.tasks || [])
    .filter(t => t.status === 'Done' && t.completedAt && new Date(t.completedAt).getTime() >= weekAgoMs)
    .map(t => ({ task: t.task, owner: t.owner, completedAt: t.completedAt }));
  const upcoming  = (project.tasks || [])
    .filter(t => t.dueDate && t.status !== 'Done' && new Date(t.dueDate+'T00:00:00').getTime() <= weekAheadMs && new Date(t.dueDate+'T00:00:00').getTime() >= todayMs)
    .map(t => ({ task: t.task, owner: t.owner, dueDate: t.dueDate, status: t.status }));
  const pastDue   = (project.tasks || [])
    .filter(t => t.dueDate && t.status !== 'Done' && new Date(t.dueDate+'T00:00:00').getTime() < todayMs)
    .map(t => ({ task: t.task, owner: t.owner, dueDate: t.dueDate, status: t.status }));
  const blockers  = (project.tasks || []).filter(t => t.status === 'Blocked' || t.flaggedRisk).map(t => ({ task: t.task, owner: t.owner, notes: t.notes }));
  const openRisks = (project.risks || []).filter(r => r.status === 'Open' || r.status === 'Monitoring' || !r.status)
    .map(r => ({ risk: r.risk, owner: r.owner, impact: r.impact, mitigation: r.mitigation }));

  const audienceTone = audience === 'leadership'
    ? 'Executive — bottom-line first, business impact framing, max 250 words. No operational detail.'
    : audience === 'stakeholders'
    ? 'External stakeholders — high-level, no internal jargon, no team names unless necessary. Polite, confident.'
    : 'Team — operational detail, name owners by team, include specifics. Crisp, action-oriented.';

  const system = `You are a Principal Strategy Manager writing a project status update. Today: ${new Date().toISOString().split('T')[0]}.

Audience: ${audienceTone}

Output a clean plain-text update with these sections, in order:
1) HEADER — Project name, current health score (from healthScan), one-line health summary.
2) ✅ COMPLETED LAST WEEK — bullets from the completed list. If empty say "Nothing completed in the window."
3) 📅 PIPELINE (NEXT 7 DAYS) — bullets from the upcoming list. Include owner and due date.
4) ❗ PAST DUE — bullets from the pastDue list (if any).
5) ⚠️ RISKS &  BLOCKERS — bullets from openRisks + blockers, prioritized by impact.
6) 💡 AI HEALTH CHECK INSIGHTS — 2-3 of the most important gaps/suggestions from the lastScanFindings.
7) CALL TO ACTION — 1-2 lines on what the recipient should do next.

Use plain text. Bullets are "•". No markdown headers (use ALL CAPS for section labels). Keep it scannable.`;

  const user = `PROJECT: ${project.name || 'Unnamed'}
LEAD: ${project.lead || 'TBD'}
TARGET LAUNCH: ${project.dueDate || 'TBD'}

LATEST HEALTH SCAN:
${JSON.stringify(healthScan || {}, null, 2)}

HEALTH SCAN FINDINGS (use to spot gaps/suggestions to call out):
${JSON.stringify(lastScanFindings || {}, null, 2)}

COMPLETED LAST 7 DAYS:
${JSON.stringify(completed, null, 2)}

UPCOMING NEXT 7 DAYS:
${JSON.stringify(upcoming, null, 2)}

PAST DUE:
${JSON.stringify(pastDue, null, 2)}

OPEN RISKS:
${JSON.stringify(openRisks, null, 2)}

BLOCKERS:
${JSON.stringify(blockers, null, 2)}

PROJECT SNAPSHOT (for context only — don't quote verbatim):
${JSON.stringify(snapshot, null, 2)}

Now write the update.`;

  const update = await callClaude(system, user, 2000);
  return { update, elapsedMs: Date.now() - t0 };
}

// ════════════════════════════════════════════════════════════════
// SEND LEAD REMINDER — Friday "don't forget to run health check + send update"
// ════════════════════════════════════════════════════════════════
async function sendLeadReminder({ project, settings, fbProjectId }) {
  const t0 = Date.now();
  const leadEmail = (project.leadEmail || '').trim();
  const leadName = (project.lead || '').trim() || 'Project Lead';
  if (!leadEmail) throw new Error('Project Lead Email not set on Project Setup tab.');

  const projectUrl = (fbProjectId)
    ? `https://moxie-ops-project-plans.vercel.app/#${fbProjectId}`
    : '';
  const lastScan = project.pmAiLastScan;
  const lastWhen = lastScan && lastScan.at
    ? new Date(lastScan.at).toLocaleDateString('en-US', { weekday:'short', month:'short', day:'numeric' })
    : 'never';
  const daysAgo = lastScan && lastScan.at
    ? Math.round((Date.now() - new Date(lastScan.at).getTime()) / 86400000)
    : null;
  const staleText = daysAgo === null ? 'No health check has been run yet.' : (daysAgo > 7 ? `It's been ${daysAgo} days since the last health check.` : `Last health check was ${lastWhen}.`);

  const subject = `Friday reminder — Run Health Check + Send Update · ${project.name || 'Project'}`;
  const html = `<div style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;max-width:560px;margin:0 auto;color:#1F2937;">
    <div style="background:#1F1A47;color:#fff;padding:18px 22px;border-radius:8px 8px 0 0;">
      <div style="font-size:12px;text-transform:uppercase;letter-spacing:1px;opacity:.8;">📅 Friday Reminder</div>
      <div style="font-size:20px;font-weight:800;margin-top:4px;">${escapeHtmlServer(project.name || 'Project')}</div>
    </div>
    <div style="background:#fff;border:1px solid #E2E8F0;border-top:none;padding:18px 22px;border-radius:0 0 8px 8px;">
      <p style="font-size:14px;margin:0 0 12px 0;">Hi ${escapeHtmlServer(leadName.split(' ')[0])} —</p>
      <p style="font-size:13px;line-height:1.6;margin:0 0 14px 0;">Quick heads-up to close out the week strong:</p>
      <ol style="font-size:13px;line-height:1.7;margin:0 0 16px 18px;padding:0;">
        <li><strong>🩺 Run a Project Health Check</strong> — surfaces duplicates, gaps, and gives an AI-graded health score. ${escapeHtmlServer(staleText)}</li>
        <li><strong>📣 Send the Project Update</strong> — dashboard snapshot to your stakeholders + Slack channel.</li>
      </ol>
      ${projectUrl ? `<div style="text-align:center;margin:16px 0;">
        <a href="${projectUrl}" style="display:inline-block;background:#6B4EFF;color:#fff;text-decoration:none;font-weight:700;padding:10px 22px;border-radius:6px;font-size:13px;">→ Open Project Tracker</a>
      </div>` : ''}
      <p style="font-size:12px;color:#64748B;margin:12px 0 0 0;line-height:1.5;">It takes ~3 minutes: Run Health Check → make any quick fixes → click "Send Dashboard Snapshot to Email" + "Send to Slack Channel". Done.</p>
    </div>
  </div>`;

  const result = { emailSent: false, slackSent: false, errors: [] };

  // 1) Email the lead
  try {
    await sendEmail({
      to: [leadEmail],
      subject,
      html,
      fromName: settings?.fromName || 'Moxie Project Tracker',
      projectName: project.name || '',
    });
    result.emailSent = true;
  } catch (e) {
    result.errors.push('Email: ' + e.message);
  }

  // 2) Slack DM the lead — look up by email
  const botToken = process.env.SLACK_BOT_TOKEN;
  if (botToken) {
    try {
      const lookRes = await fetch(`https://slack.com/api/users.lookupByEmail?email=${encodeURIComponent(leadEmail)}`, {
        headers: { 'Authorization': `Bearer ${botToken}` },
      });
      const lookJson = await lookRes.json().catch(() => ({}));
      if (lookJson.ok && lookJson.user && lookJson.user.id) {
        const dmText = `:wave: *Friday reminder for ${project.name || 'your project'}*\n\n` +
          `Two quick things to close the week:\n` +
          `1. :stethoscope: Run a Project Health Check (${staleText.toLowerCase()})\n` +
          `2. :loudspeaker: Send the project update to your team\n\n` +
          (projectUrl ? `<${projectUrl}|→ Open Project Tracker>` : '');
        const dmRes = await fetch('https://slack.com/api/chat.postMessage', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json; charset=utf-8',
            'Authorization': `Bearer ${botToken}`,
          },
          body: JSON.stringify({ channel: lookJson.user.id, text: dmText, mrkdwn: true }),
        });
        const dmJson = await dmRes.json().catch(() => ({}));
        if (dmJson.ok) result.slackSent = true;
        else result.errors.push('Slack DM: ' + (dmJson.error || 'unknown'));
      } else {
        result.errors.push('Slack DM: lead email not found in workspace (' + (lookJson.error || 'no user') + ')');
      }
    } catch (e) {
      result.errors.push('Slack DM: ' + e.message);
    }
  }

  result.elapsedMs = Date.now() - t0;
  return result;
}

function escapeHtmlServer(s) {
  return String(s || '').replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
}

// ════════════════════════════════════════════════════════════════
// ANALYZE FLOW STEP — risks / gaps / suggested next step
// ════════════════════════════════════════════════════════════════
async function analyzeFlowStep({ stepText, stepType, priorSteps, project }) {
  const t0 = Date.now();
  if (!stepText || !stepText.trim()) throw new Error('Step description required.');

  const system = `You are a process designer helping a project manager build a step-by-step process flow chart for a MedSpa SaaS company called Moxie.

For the given step, surface 3-5 specific risks, gaps, or failure modes that could occur. Be concrete — name the EXACT failure (e.g., "PSM forgets to set feature flag" not "configuration error").

Also propose what the most likely NEXT step in the process would be, given the prior steps.

Respond with ONLY valid JSON in this shape:
{
  "risks": ["<concrete risk 1>", "<concrete risk 2>", ...],
  "suggestedNext": "<one-line suggestion for the next step, OR empty if you can't tell>"
}`;

  const user = `PROJECT: ${project.name || 'Unnamed'}
${project.description ? 'DESCRIPTION: ' + project.description : ''}

PRIOR STEPS IN THE FLOW:
${(priorSteps || []).map((s, i) => `${i+1}. ${s.text} [${s.type || 'manual'}]`).join('\n') || '(none yet — this is the first step)'}

CURRENT STEP TO ANALYZE:
"${stepText}" [${stepType || 'manual'}]

Return the JSON now.`;

  const raw = await callClaude(system, user, 800);
  const parsed = parseJsonSafe(raw, { risks: [], suggestedNext: '' });
  return {
    risks: Array.isArray(parsed.risks) ? parsed.risks : [],
    suggestedNext: parsed.suggestedNext || '',
    elapsedMs: Date.now() - t0,
  };
}

// Build a compact, Claude-friendly snapshot of the project (drops verbose fields, trims notes).
function buildProjectSnapshot(project) {
  const today = new Date(); today.setHours(0,0,0,0);
  const todayMs = today.getTime();
  const trim = (s, n) => { s = String(s || ''); return s.length > n ? s.slice(0, n) + '…' : s; };

  const stakeholders = Object.entries(project.stakeholders || {})
    .filter(([_, s]) => s && s.included)
    .map(([key, s]) => ({ team: key, name: s.contactName || '', email: s.contactEmail || '' }));

  const milestones = (project.milestones || []).map(m => ({
    id: m.id, name: m.name, date: m.targetDate || m.date || '', status: m.status || ''
  }));

  const tasks = (project.tasks || []).map(t => {
    const isPastDue = t.dueDate && t.status !== 'Done' && new Date(t.dueDate+'T00:00:00').getTime() < todayMs;
    return {
      task: trim(t.task, 200),
      owner: t.owner || '',
      status: t.status || '',
      dueDate: t.dueDate || '',
      pastDue: isPastDue,
      flagged: !!t.flaggedRisk,
      milestoneId: t.milestoneId || null,
      notes: trim(t.notes, 150),
    };
  });

  const actionItems = (project.actionItems || []).map(a => ({
    item: trim(a.item, 200),
    owner: a.owner || '',
    status: a.status || '',
    priority: a.priority || '',
    dueDate: a.dueDate || '',
    milestoneId: a.milestoneId || null,
  }));

  const risks = (project.risks || []).map(r => ({
    risk: trim(r.risk, 200),
    owner: r.owner || '',
    impact: r.impact || '',
    likelihood: r.likelihood || '',
    status: r.status || '',
    mitigation: trim(r.mitigation, 150),
  }));

  const decisions = (project.decisions || []).slice(-20).map(d => ({
    decision: trim(d.decision, 200),
    date: d.date || '',
    decidedBy: d.decidedBy || '',
  }));

  const launchChecklist = (project.launchChecklist || []).map(c => ({
    name: c.name, required: !!c.required, checked: !!c.checked, na: !!c.na,
  }));

  const recentMeetings = (project.meetings || []).slice(-5).map(m => ({
    title: m.title || '', date: m.date || '', notes: trim(m.notes, 200),
  }));

  return {
    name: project.name || '',
    lead: project.lead || '',
    dueDate: project.dueDate || '',
    description: trim(project.description, 300),
    ragStatus: project.ragStatus || {},
    stakeholders,
    milestones,
    tasks,
    actionItems,
    risks,
    decisions,
    launchChecklist,
    recentMeetings,
  };
}

async function pinProjectToSlack({ project, settings, fbProjectId }) {
  const botToken = process.env.SLACK_BOT_TOKEN;
  if (!botToken) throw new Error('Slack bot token not configured.');
  const channel = (settings && settings.slackChannel) || process.env.SLACK_CHANNEL;
  if (!channel) throw new Error('No Slack channel set. Open AI Assistant → Settings → fill in Slack Channel.');

  const projectUrl = `https://moxie-ops-project-plans.vercel.app/#${fbProjectId || ''}`;
  const dueDate = project.dueDate
    ? new Date(project.dueDate + 'T00:00:00').toLocaleDateString('en-US', { month:'short', day:'numeric', year:'numeric' })
    : 'TBD';
  const rag = (project.ragStatus && project.ragStatus.status) || 'Green';
  const ragEmoji = rag === 'Red' ? ':red_circle:' : rag === 'Amber' ? ':large_yellow_circle:' : ':large_green_circle:';

  // Count POCs by team
  const stakeholders = project.stakeholders || {};
  const pocs = Object.entries(stakeholders)
    .filter(([_, s]) => s && s.included && s.contactName)
    .map(([key, s]) => `${key}: ${s.contactName}`)
    .slice(0, 8);

  const text = `📌 *${project.name}* — Project Plan
${ragEmoji} *Health:* ${rag} ${project.ragStatus?.reason ? `· _${project.ragStatus.reason}_` : ''}
:dart: *Lead:* ${project.lead || 'TBD'}
:calendar: *Target Launch:* ${dueDate}
:link: *Live Project Plan:* <${projectUrl}|Open in tracker>
${project.description ? '\n_' + project.description.slice(0, 200) + '_' : ''}
${pocs.length ? '\n*Team POCs:*\n• ' + pocs.join('\n• ') : ''}

_Live updates auto-sync. Open the link anytime to see current state._`;

  // 1) Post the message (with auto-join recovery on not_in_channel)
  const attemptPost = async () => {
    const res = await fetch('https://slack.com/api/chat.postMessage', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json; charset=utf-8',
        'Authorization': `Bearer ${botToken}`,
      },
      body: JSON.stringify({ channel, text, mrkdwn: true }),
    });
    const json = await res.json().catch(() => ({}));
    return { ok: res.ok && json.ok, json };
  };

  let pr = await attemptPost();
  if (!pr.ok && pr.json.error === 'not_in_channel') {
    const joinResult = await ensureBotInChannel(botToken, channel);
    if (joinResult.joined) {
      pr = await attemptPost();
    } else {
      const hint = joinResult.reason === 'private_channel'
        ? ' — bot cannot auto-join private channels. Run /invite @<bot-name> in the channel.'
        : joinResult.reason === 'missing_scope'
        ? ' — to auto-join, add channels:join scope to your bot and reinstall. Or run /invite @<bot-name>.'
        : ' — invite the bot first: /invite @<bot-name>';
      throw new Error('Slack post failed: not_in_channel' + hint);
    }
  }
  if (!pr.ok) {
    throw new Error('Slack post failed: ' + (pr.json.error || 'unknown'));
  }
  const postJson = pr.json;
  const ts = postJson.ts;          // timestamp of the message
  const channelId = postJson.channel; // resolved channel ID

  const result = { posted: true, channel, ts, projectUrl, pinned: false, topicSet: false, warnings: [] };

  // 2) Try to pin the message (needs pins:write scope)
  try {
    const pinRes = await fetch('https://slack.com/api/pins.add', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json; charset=utf-8',
        'Authorization': `Bearer ${botToken}`,
      },
      body: JSON.stringify({ channel: channelId, timestamp: ts }),
    });
    const pinJson = await pinRes.json().catch(() => ({}));
    if (pinJson.ok) {
      result.pinned = true;
    } else if (pinJson.error === 'missing_scope') {
      result.warnings.push('Auto-pin requires the `pins:write` scope. Pin the message manually in Slack (Slack menu → Pin to channel) or add the scope and reinstall.');
    } else if (pinJson.error === 'already_pinned') {
      result.pinned = true;
    } else {
      result.warnings.push('Pin failed: ' + (pinJson.error || 'unknown'));
    }
  } catch (e) {
    result.warnings.push('Pin attempt error: ' + e.message);
  }

  // 3) Try to set the channel topic to include the project URL (optional)
  try {
    const topic = `📌 ${project.name} · ${projectUrl}`;
    const topicRes = await fetch('https://slack.com/api/conversations.setTopic', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json; charset=utf-8',
        'Authorization': `Bearer ${botToken}`,
      },
      body: JSON.stringify({ channel: channelId, topic }),
    });
    const topicJson = await topicRes.json().catch(() => ({}));
    if (topicJson.ok) {
      result.topicSet = true;
    } else if (topicJson.error === 'missing_scope') {
      // silently skip — topic is nice-to-have
    } else if (topicJson.error !== 'no_permission' && topicJson.error !== 'not_in_channel') {
      result.warnings.push('Channel topic update skipped: ' + (topicJson.error || 'unknown'));
    }
  } catch (e) {
    // silently ignore
  }

  // 4) Add channel bookmarks (links at top of channel) — needs bookmarks:write scope
  // Adds: 1) Project Plan URL, 2) all links from project.links, dedupe by URL.
  // Get existing bookmarks first to avoid duplicates.
  result.bookmarksAdded = 0;
  result.bookmarksSkipped = 0;
  try {
    // List existing channel bookmarks
    const listRes = await fetch(`https://slack.com/api/bookmarks.list?channel_id=${encodeURIComponent(channelId)}`, {
      method: 'GET',
      headers: { 'Authorization': `Bearer ${botToken}` },
    });
    const listJson = await listRes.json().catch(() => ({}));
    if (!listJson.ok && listJson.error === 'missing_scope') {
      result.warnings.push('Channel bookmarks need `bookmarks:write` scope. Add it to the bot and reinstall.');
    } else {
      const existing = new Set((listJson.bookmarks || []).map(b => (b.link || '').toLowerCase()));

      // Build the bookmark list: project plan first, then project.links
      const wanted = [];
      wanted.push({
        title: `📋 ${project.name} Project Plan`,
        link: projectUrl,
        emoji: ':pushpin:',
      });
      (project.links || []).forEach(l => {
        if (!l || !l.url) return;
        wanted.push({
          title: (l.title || l.url).slice(0, 80),
          link: l.url,
          emoji: l.category ? ':link:' : ':link:',
        });
      });

      // Add each bookmark, skipping ones already present
      for (const b of wanted) {
        if (existing.has((b.link || '').toLowerCase())) {
          result.bookmarksSkipped++;
          continue;
        }
        try {
          const addRes = await fetch('https://slack.com/api/bookmarks.add', {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json; charset=utf-8',
              'Authorization': `Bearer ${botToken}`,
            },
            body: JSON.stringify({
              channel_id: channelId,
              title: b.title,
              type: 'link',
              link: b.link,
              emoji: b.emoji,
            }),
          });
          const addJson = await addRes.json().catch(() => ({}));
          if (addJson.ok) result.bookmarksAdded++;
          else if (addJson.error === 'missing_scope') {
            result.warnings.push('Bookmark add failed — `bookmarks:write` scope missing.');
            break;
          } else {
            result.warnings.push(`Bookmark "${b.title}" failed: ${addJson.error || 'unknown'}`);
          }
        } catch (e) {
          result.warnings.push(`Bookmark "${b.title}" error: ${e.message}`);
        }
      }
    }
  } catch (e) {
    result.warnings.push('Bookmarks step error: ' + e.message);
  }

  return result;
}

// ════════════════════════════════════════════════════════════════
// CREATE SLACK CHANNEL + INVITE POCs (called from New Project wizard)
// ════════════════════════════════════════════════════════════════
// Required Slack bot scopes:
//   channels:manage (public) OR groups:write (private)  ← for conversations.create
//   users:read, users:read.email                        ← for users.lookupByEmail
//   channels:write.invites OR groups:write              ← for conversations.invite
//   bookmarks:write                                     ← optional, for project link bookmark
async function createSlackChannel({ channelName, isPrivate, pocEmails, projectName, projectUrl, project, settings }) {
  const botToken = process.env.SLACK_BOT_TOKEN;
  if (!botToken) throw new Error('Slack bot token not configured.');
  if (!channelName) throw new Error('channelName is required.');

  // Slack channel names: lowercase, no spaces, only a-z 0-9 - _ ; ≤ 80 chars
  const safeName = String(channelName).toLowerCase().replace(/[^a-z0-9_-]/g, '-').replace(/-+/g, '-').slice(0, 80);

  // 1) Create the channel
  const createRes = await fetch('https://slack.com/api/conversations.create', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json; charset=utf-8',
      'Authorization': `Bearer ${botToken}`,
    },
    body: JSON.stringify({ name: safeName, is_private: !!isPrivate }),
  });
  const createJson = await createRes.json().catch(() => ({}));

  let channelId;
  let channelNameFinal = safeName;
  if (createJson.ok) {
    channelId = createJson.channel.id;
    channelNameFinal = createJson.channel.name;
  } else if (createJson.error === 'name_taken') {
    // Channel already exists — try to look up its ID so we can still invite
    const listRes = await fetch('https://slack.com/api/conversations.list?types=public_channel,private_channel&limit=1000', {
      headers: { 'Authorization': `Bearer ${botToken}` },
    });
    const listJson = await listRes.json().catch(() => ({}));
    const found = (listJson.channels || []).find(c => c.name === safeName);
    if (!found) {
      throw new Error(`Channel #${safeName} exists but bot cannot see it. Invite the bot to that channel first, or pick a different name.`);
    }
    channelId = found.id;
    channelNameFinal = found.name;
  } else if (createJson.error === 'missing_scope') {
    throw new Error('Slack channel create failed — bot needs the `channels:manage` scope (or `groups:write` for private channels). Add it in the Slack app settings and reinstall.');
  } else {
    throw new Error('Slack channel create failed: ' + (createJson.error || 'unknown'));
  }

  // 2) Set the channel topic / purpose to make the project obvious
  try {
    if (projectUrl) {
      await fetch('https://slack.com/api/conversations.setTopic', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json; charset=utf-8',
          'Authorization': `Bearer ${botToken}`,
        },
        body: JSON.stringify({ channel: channelId, topic: `📌 ${projectName || 'Project'} · ${projectUrl}` }),
      });
    }
    if (projectName) {
      await fetch('https://slack.com/api/conversations.setPurpose', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json; charset=utf-8',
          'Authorization': `Bearer ${botToken}`,
        },
        body: JSON.stringify({ channel: channelId, purpose: `${projectName} project plan, status updates, and discussion.` }),
      });
    }
  } catch (e) {
    // non-fatal
  }

  // 3) Resolve POC emails → Slack user IDs, then invite in a single batch
  const emails = (Array.isArray(pocEmails) ? pocEmails : [])
    .map(e => String(e || '').trim())
    .filter(Boolean);

  const invitedUserIds = [];
  const failedInvites = [];
  for (const email of emails) {
    try {
      const lookRes = await fetch(`https://slack.com/api/users.lookupByEmail?email=${encodeURIComponent(email)}`, {
        headers: { 'Authorization': `Bearer ${botToken}` },
      });
      const lookJson = await lookRes.json().catch(() => ({}));
      if (lookJson.ok && lookJson.user && lookJson.user.id) {
        invitedUserIds.push(lookJson.user.id);
      } else {
        failedInvites.push({ email, reason: lookJson.error || 'not_found' });
      }
    } catch (e) {
      failedInvites.push({ email, reason: e.message });
    }
  }

  let invitedCount = 0;
  if (invitedUserIds.length) {
    // Slack's invite endpoint accepts a comma-separated list of user IDs (up to 1000)
    const inviteRes = await fetch('https://slack.com/api/conversations.invite', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json; charset=utf-8',
        'Authorization': `Bearer ${botToken}`,
      },
      body: JSON.stringify({ channel: channelId, users: invitedUserIds.join(',') }),
    });
    const inviteJson = await inviteRes.json().catch(() => ({}));
    if (inviteJson.ok) {
      invitedCount = invitedUserIds.length;
    } else if (inviteJson.error === 'already_in_channel') {
      invitedCount = invitedUserIds.length; // count as success
    } else if (inviteJson.errors && Array.isArray(inviteJson.errors)) {
      // Some succeeded, some didn't — count the ones not in errors as invited
      invitedCount = invitedUserIds.length - inviteJson.errors.length;
      inviteJson.errors.forEach(e => failedInvites.push({ userId: e.user, reason: e.error }));
    } else {
      failedInvites.push({ batch: invitedUserIds.length, reason: inviteJson.error || 'unknown' });
    }
  }

  // 4) Add the project plan as a channel bookmark (best-effort)
  if (projectUrl) {
    try {
      await fetch('https://slack.com/api/bookmarks.add', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json; charset=utf-8',
          'Authorization': `Bearer ${botToken}`,
        },
        body: JSON.stringify({
          channel_id: channelId,
          title: `📋 ${projectName || 'Project'} Plan`,
          type: 'link',
          link: projectUrl,
          emoji: ':pushpin:',
        }),
      });
    } catch (e) {
      // ignore
    }
  }

  // 5) Post a kickoff message
  try {
    const kickoffText = `:tada: *${projectName || 'New Project'}* channel is live!\n\n` +
      (projectUrl ? `📋 *Project plan:* <${projectUrl}|Open tracker>\n` : '') +
      `👥 *POCs invited:* ${invitedCount}\n` +
      `\nUse this channel for project updates, decisions, and async discussion.`;
    await fetch('https://slack.com/api/chat.postMessage', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json; charset=utf-8',
        'Authorization': `Bearer ${botToken}`,
      },
      body: JSON.stringify({ channel: channelId, text: kickoffText, mrkdwn: true }),
    });
  } catch (e) {
    // ignore
  }

  // Construct the deep link to the channel (best-guess slack:// or web URL)
  const channelUrl = `https://slack.com/app_redirect?channel=${channelId}`;

  return {
    ok: true,
    channelId,
    channelName: channelNameFinal,
    channelUrl,
    invitedCount,
    failedInvites,
    isPrivate: !!isPrivate,
  };
}

// ════════════════════════════════════════════════════════════════
// HELPERS — Slack, Email, Drive, PPTX
// ════════════════════════════════════════════════════════════════
// Try to have the bot auto-join a public channel (needs channels:join scope).
// Returns true if joined / already in, false otherwise (e.g., private channel).
async function ensureBotInChannel(botToken, channel) {
  // First find the channel ID — conversations.join needs the ID not the name
  // For simplicity, we try the channel name directly; Slack often accepts both
  try {
    const res = await fetch('https://slack.com/api/conversations.join', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json; charset=utf-8',
        'Authorization': `Bearer ${botToken}`,
      },
      body: JSON.stringify({ channel: channel.replace(/^#/, '') }),
    });
    const json = await res.json().catch(() => ({}));
    if (json.ok) return { joined: true };
    if (json.error === 'already_in_channel') return { joined: true };
    if (json.error === 'method_not_supported_for_channel_type') {
      return { joined: false, reason: 'private_channel' };
    }
    if (json.error === 'missing_scope') {
      return { joined: false, reason: 'missing_scope' };
    }
    if (json.error === 'channel_not_found') {
      // Try resolving via conversations.list to get the ID
      const listRes = await fetch('https://slack.com/api/conversations.list?types=public_channel,private_channel&limit=1000', {
        method: 'GET',
        headers: { 'Authorization': `Bearer ${botToken}` },
      });
      const listJson = await listRes.json().catch(() => ({}));
      if (listJson.ok && Array.isArray(listJson.channels)) {
        const target = listJson.channels.find(c => c.name === channel.replace(/^#/, ''));
        if (target) {
          if (target.is_member) return { joined: true };
          const retryRes = await fetch('https://slack.com/api/conversations.join', {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json; charset=utf-8',
              'Authorization': `Bearer ${botToken}`,
            },
            body: JSON.stringify({ channel: target.id }),
          });
          const retryJson = await retryRes.json().catch(() => ({}));
          if (retryJson.ok) return { joined: true };
          if (retryJson.error === 'method_not_supported_for_channel_type') return { joined: false, reason: 'private_channel' };
          return { joined: false, reason: retryJson.error || 'unknown' };
        }
      }
      return { joined: false, reason: 'channel_not_found' };
    }
    return { joined: false, reason: json.error || 'unknown' };
  } catch (e) {
    return { joined: false, reason: e.message };
  }
}

async function postToSlack({ text, settings }) {
  // Prefer Bot Token (chat.postMessage API) if available — works for any channel the bot is in.
  // Fall back to Webhook URL — channel-locked but no install required.
  const botToken = process.env.SLACK_BOT_TOKEN;
  const webhook  = process.env.SLACK_WEBHOOK_URL;

  if (botToken) {
    const channel = (settings && settings.slackChannel) || process.env.SLACK_CHANNEL;
    if (!channel) throw new Error('SLACK_BOT_TOKEN set but no channel — set SLACK_CHANNEL env var or fill Slack Channel in Settings.');

    const attemptPost = async () => {
      const res = await fetch('https://slack.com/api/chat.postMessage', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json; charset=utf-8',
          'Authorization': `Bearer ${botToken}`,
        },
        body: JSON.stringify({ channel, text, mrkdwn: true }),
      });
      const json = await res.json().catch(() => ({}));
      return { ok: res.ok && json.ok, json, status: res.status, statusText: res.statusText };
    };

    let r = await attemptPost();
    if (!r.ok && r.json.error === 'not_in_channel') {
      // Auto-recovery: try to join the channel, then retry
      const joinResult = await ensureBotInChannel(botToken, channel);
      if (joinResult.joined) {
        r = await attemptPost();
      } else {
        const hint = joinResult.reason === 'private_channel'
          ? ` — bot can't auto-join private channels. Invite manually: /invite @<bot> in ${channel}`
          : joinResult.reason === 'missing_scope'
          ? ` — to enable auto-join, add channels:join scope to your Slack bot and reinstall, OR invite manually: /invite @<bot> in ${channel}`
          : ` — invite manually: /invite @<bot> in ${channel}`;
        throw new Error('Slack chat.postMessage failed: not_in_channel' + hint);
      }
    }
    if (!r.ok) {
      throw new Error('Slack chat.postMessage failed: ' + (r.json.error || r.statusText));
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

  // Auto-prefix subject with [Project Name] unless it already starts with [ProjectName]
  if (projectName) {
    const prefix = `[${projectName}]`;
    if (!subject || !subject.startsWith(prefix)) {
      subject = `${prefix} ${subject || ''}`.trim();
    }
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
  // Never cache API responses — guarantees fresh data on every click
  res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate');
  res.setHeader('Pragma', 'no-cache');
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
      case 'schedule-meeting':
        result = await scheduleMeetingAction({ ...payload, project, settings });
        break;
      case 'list-calendar-events':
        result = await listCalendarEvents({ ...payload, project });
        break;
      case 'pin-project-to-slack':
        result = await pinProjectToSlack({ ...payload, project, settings, fbProjectId });
        break;
      case 'create-slack-channel':
        result = await createSlackChannel({ ...payload, project, settings });
        break;
      case 'ask-dashboard':
        result = await askDashboard({ ...payload, project });
        break;
      case 'suggest-rag':
        result = await suggestRag({ ...payload, project });
        break;
      case 'meeting-prep-brief':
        result = await meetingPrepBrief({ ...payload, project });
        break;
      case 'suggest-task-teams':
        result = await suggestTaskTeams({ ...payload, project });
        break;
      case 'draft-sop':
        result = await draftSop({ ...payload, project });
        break;
      case 'draft-standup':
        result = await draftStandup({ ...payload, project });
        break;
      case 'extract-decisions':
        result = await extractDecisions({ ...payload, project });
        break;
      case 'onboarding-brief':
        result = await onboardingBrief({ ...payload, project });
        break;
      case 'draft-email-reply':
        result = await draftEmailReply({ ...payload, project });
        break;
      case 'build-proposal':
        result = await buildProposal({ ...payload, project });
        break;
      case 'read-slack-channel':
        result = await readSlackChannel({ ...payload, project });
        break;
      case 'pm-ai-scan':
        result = await pmAiScan({ ...payload, project });
        break;
      case 'project-update':
        result = await projectUpdate({ ...payload, project });
        break;
      case 'send-lead-reminder':
        result = await sendLeadReminder({ ...payload, project, settings, fbProjectId });
        break;
      case 'analyze-flow-step':
        result = await analyzeFlowStep({ ...payload, project });
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
