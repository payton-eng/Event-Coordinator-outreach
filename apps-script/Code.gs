/**
 * Magix Production – DFW event planner outreach
 *
 * Lives inside the "Outreach – DFW Event Planners" Google Sheet
 * (Extensions → Apps Script). Sends a short email sequence from your Gmail,
 * follows up in the same thread, and stops as soon as someone replies.
 *
 * Needs the Gmail advanced service: in the Apps Script editor click
 * Services (+) → Gmail API → Add.
 *
 * Everything you'd want to change (wording, limits, signature) lives in the
 * Settings and Templates tabs of the sheet, not in this code.
 */

const TAB_LEADS = 'Leads';
const TAB_IMPORT = 'Apollo Import';
const TAB_SETTINGS = 'Settings';
const TAB_TEMPLATES = 'Templates';

const LEAD_HEADERS = [
  'First Name', 'Last Name', 'Email', 'Title', 'Company', 'City', 'Website',
  'LinkedIn', 'Status', 'Step', 'Last Sent', 'Notes', 'Thread ID', 'Campaign',
];

const TEMPLATE_HEADERS = ['Step', 'Wait Days', 'Subject', 'Body', 'Campaign'];

// Leads and template rows with a blank Campaign belong to this one.
const DEFAULT_CAMPAIGN = 'Planners';

// Apollo's CSV column names (first match wins) for each Leads column.
const APOLLO_COLUMNS = {
  'First Name': ['First Name'],
  'Last Name': ['Last Name'],
  'Email': ['Email', 'Work Email', 'Email Address'],
  'Title': ['Title', 'Job Title'],
  'Company': ['Company Name for Emails', 'Company', 'Company Name'],
  'City': ['City'],
  'Website': ['Website', 'Company Website'],
  'LinkedIn': ['Person Linkedin Url', 'LinkedIn Url', 'Linkedin Url', 'LinkedIn'],
};

const DEFAULT_SETTINGS = [
  ['Setting', 'Value', 'What it does'],
  ['From Name', 'Payton | Magix Production', 'The name planners see in their inbox.'],
  ['Signature',
    'Payton\nMagix Production | Event AV & Production\nDallas–Fort Worth · magixproduction.com\n[YOUR PHONE]\n[YOUR MAILING ADDRESS]',
    'Added to the bottom of every email. Replace the [BRACKETS] first — nothing sends until you do. Ctrl+Enter makes a new line inside a cell.'],
  ['Opt-out Line',
    'Not the right person, or rather not hear from me? Just reply "remove" and I won\'t email again.',
    'Required by law (CAN-SPAM). Keep some version of it.'],
  ['Paused', 'NO', 'Set to YES to stop all sending right away.'],
  ['Daily Limit Week 1', 10, 'Max emails per day during the first week of sending.'],
  ['Daily Limit Week 2', 20, 'Max emails per day during the second week.'],
  ['Daily Limit After', 25, 'Max emails per day after that. Keep this low to protect the Gmail account.'],
  ['Send Start Hour', 9, 'Earliest hour to send (24-hour clock: 9 = 9am).'],
  ['Send End Hour', 16, 'No emails at or after this hour (16 = 4pm).'],
  ['Max Per Run', 4, 'Emails per hourly run, so sending is spread across the day.'],
  ['Weekdays Only', 'YES', 'YES = never send on Saturday or Sunday.'],
  ['Only Verified Emails', 'YES', 'When importing from Apollo, skip emails Apollo has not verified.'],
  ['Reply Alert Email', '', 'Optional. When someone replies, also send a heads-up to this address.'],
];

const DEFAULT_TEMPLATES = [
  TEMPLATE_HEADERS,
  [1, 0, 'AV for {{Company|your}} events',
    'Hi {{First Name}},\n\n' +
    'I\'m Payton, owner of Magix Production here in DFW. Before starting my company, I worked hotel AV with Encore at the Gaylord Texan, the Omni, the Adolphus and the Crescent Court.\n\n' +
    'Now my team handles sound, lighting, staging, LED walls and projection, video production and livestreaming for corporate events, galas and conferences. Same big-hotel standards, without the big-hotel pricing.\n\n' +
    'Do you have anything coming up in the next few months where a second AV quote would be useful?',
    'Planners'],
  [2, 3, '',
    'Hi {{First Name}}, quick follow-up.\n\n' +
    'Having worked on the in-house side, I know many venues allow outside AV, and it\'s often less expensive for the same (or better) production. Happy to put together a side-by-side quote for your next event so you can compare. No obligation.\n\n' +
    'For reference, before starting Magix I was the AV Coordinator for the Texas Rangers, overseeing AV for everything on the field and around the stadium grounds, from Jonas Brothers and P!nk concerts to corporate events in the conference rooms. Our biggest show since starting Magix was the Badass Business Summit.\n\n' +
    'Worth a look?',
    'Planners'],
  [3, 6, '',
    'Hi {{First Name}}, I\'ll stop filling your inbox after this one.\n\n' +
    'If you ever need an AV partner in DFW, or a backup when your usual vendor is booked, just reply here and I\'ll get you a detailed quote within 3 days.',
    'Planners'],
];

// Added to an existing Templates tab by "Set up tabs" if it has no Associations rows yet.
const ASSOCIATION_TEMPLATES = [
  [1, 0, 'AV for {{Company|your}} conference and gala',
    'Hi {{First Name}},\n\n' +
    'I\'m Payton, owner of Magix Production here in DFW. I was the AV Coordinator for the Texas Rangers, worked hotel AV at the Gaylord Texan, the Omni and the Adolphus, and have produced meetings for the National Speakers Association\'s North Texas chapter.\n\n' +
    'We handle the full show for annual conferences and awards galas: sound, lighting, staging, LED walls, video, and livestreaming so members who can\'t make it in person can still watch live.\n\n' +
    'Is the AV for {{Company|your}} next conference or gala already locked in, or would a second quote be useful?',
    'Associations'],
  [2, 3, '',
    'Hi {{First Name}}, quick follow-up.\n\n' +
    'For associations, livestreaming is often the easiest win: members who can\'t travel still see the keynotes and awards, sponsors get extra visibility, and the recording becomes content you can share all year. We run it alongside the in-room production with one crew, so you\'re not coordinating two vendors.\n\n' +
    'For reference, our biggest show to date was the Badass Business Summit. Happy to put together a quote for your next conference or gala. No obligation.\n\n' +
    'Worth a look?',
    'Associations'],
  [3, 6, '',
    'Hi {{First Name}}, I\'ll stop filling your inbox after this one.\n\n' +
    'If you\'re starting to plan next year\'s conference or awards gala, just reply here and I\'ll get you a detailed quote within 3 days.',
    'Associations'],
];

const UNSUBSCRIBE_RE = /\b(remove|unsubscribe|stop emailing|not interested|take me off|do not contact|don't contact)\b/i;
const BOUNCE_FROM_RE = /mailer-daemon|postmaster/i;

/* ---------------------------------------------------------------- menu --- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Outreach')
    .addItem('1. Set up tabs', 'setupTabs')
    .addItem('2. Import leads from Apollo tab', 'importFromApollo')
    .addItem('3. Send myself a test email', 'sendTestEmail')
    .addSeparator()
    .addItem('Start automatic sending', 'startAutomation')
    .addItem('Stop automatic sending', 'stopAutomation')
    .addItem('Run once now', 'runNow')
    .addToUi();
}

function setupTabs() {
  const ss = SpreadsheetApp.getActive();

  let leads = ss.getSheetByName(TAB_LEADS);
  if (!leads) leads = ss.insertSheet(TAB_LEADS);
  const lastCol = Math.max(leads.getLastColumn(), 1);
  const existing = leads.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const missing = LEAD_HEADERS.filter(h => existing.indexOf(h) === -1);
  if (missing.length) {
    const start = existing.filter(String).length + 1;
    leads.getRange(1, start, 1, missing.length).setValues([missing]);
  }
  leads.setFrozenRows(1);

  if (!ss.getSheetByName(TAB_SETTINGS)) {
    const s = ss.insertSheet(TAB_SETTINGS);
    s.getRange(1, 1, DEFAULT_SETTINGS.length, 3).setValues(DEFAULT_SETTINGS);
    s.getRange('A1:C1').setFontWeight('bold');
    s.setColumnWidth(1, 170).setColumnWidth(2, 380).setColumnWidth(3, 420);
    s.getRange(1, 2, DEFAULT_SETTINGS.length, 2).setWrap(true);
    s.setFrozenRows(1);
  }

  let t = ss.getSheetByName(TAB_TEMPLATES);
  if (!t) {
    t = ss.insertSheet(TAB_TEMPLATES);
    t.getRange(1, 1, DEFAULT_TEMPLATES.length, TEMPLATE_HEADERS.length).setValues(DEFAULT_TEMPLATES);
    t.getRange(1, 1, 1, TEMPLATE_HEADERS.length).setFontWeight('bold');
    t.setColumnWidth(3, 220).setColumnWidth(4, 600);
    t.setFrozenRows(1);
  }
  // Older Templates tabs were made before campaigns existed.
  const tHeaders = t.getRange(1, 1, 1, Math.max(t.getLastColumn(), 1)).getValues()[0].map(h => String(h).trim());
  if (tHeaders.indexOf('Campaign') === -1) {
    t.getRange(1, tHeaders.filter(String).length + 1).setValue('Campaign').setFontWeight('bold');
  }
  let addedAssociations = false;
  if (!getTemplates_()[campaignKey_('Associations')]) {
    const col = templateColumns_(t);
    const rows = ASSOCIATION_TEMPLATES.map(r => {
      const out = new Array(t.getLastColumn()).fill('');
      TEMPLATE_HEADERS.forEach((h, i) => { out[col[h]] = r[i]; });
      return out;
    });
    t.getRange(t.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    addedAssociations = true;
  }
  t.getDataRange().setWrap(true).setVerticalAlignment('top');

  SpreadsheetApp.getUi().alert(
    'Done! You have Leads, Settings and Templates tabs.' +
    (addedAssociations ? '\n\nAdded the 3 Associations emails to the bottom of the Templates tab.' : '') +
    '\n\nIf you haven\'t yet: open the Settings tab and replace [YOUR PHONE] and [YOUR MAILING ADDRESS] in the Signature.');
}

/* -------------------------------------------------------------- import --- */

function importFromApollo() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(TAB_IMPORT);
  if (!src) {
    ui.alert('Could not find a tab named "' + TAB_IMPORT + '". Import your Apollo CSV into a tab with that exact name.');
    return;
  }
  const settings = getSettings_();
  const onlyVerified = isYes_(settings['Only Verified Emails']);
  const values = src.getDataRange().getValues();
  if (values.length < 2) {
    ui.alert('The "' + TAB_IMPORT + '" tab is empty.');
    return;
  }

  const srcHeaders = values[0].map(h => String(h).trim().toLowerCase());
  const srcIndex = {};
  Object.keys(APOLLO_COLUMNS).forEach(target => {
    const name = APOLLO_COLUMNS[target].find(n => srcHeaders.indexOf(n.toLowerCase()) !== -1);
    srcIndex[target] = name ? srcHeaders.indexOf(name.toLowerCase()) : -1;
  });
  const statusIdx = srcHeaders.indexOf('email status');
  if (srcIndex['Email'] === -1 || srcIndex['First Name'] === -1) {
    ui.alert('The Apollo tab needs "First Name" and "Email" columns. Did the CSV import correctly?');
    return;
  }

  const byCampaign = getTemplates_();
  const campaigns = Object.keys(byCampaign).map(k => byCampaign[k][0].campaign);
  const answer = ui.prompt('Which campaign are these leads for?',
    'Type one of: ' + campaigns.join(', ') + '\n(Leave blank for ' + DEFAULT_CAMPAIGN + '.)',
    ui.ButtonSet.OK_CANCEL);
  if (answer.getSelectedButton() !== ui.Button.OK) return;
  const typed = answer.getResponseText().trim() || DEFAULT_CAMPAIGN;
  const campaign = campaigns.find(c => campaignKey_(c) === campaignKey_(typed));
  if (!campaign) {
    ui.alert('There are no emails for a campaign called "' + typed + '" in the Templates tab. ' +
      'Nothing was imported. Try again with one of: ' + campaigns.join(', '));
    return;
  }

  const leads = loadLeads_();
  const seen = new Set(leads.rows.map(r => String(r[leads.col['Email']]).trim().toLowerCase()).filter(String));
  const toAdd = [];
  let skippedNoEmail = 0, skippedUnverified = 0, skippedDuplicate = 0;

  values.slice(1).forEach(row => {
    const email = String(row[srcIndex['Email']]).trim();
    const first = String(row[srcIndex['First Name']]).trim();
    if (!email || email.indexOf('@') === -1 || !first) { skippedNoEmail++; return; }
    if (onlyVerified && statusIdx !== -1) {
      const status = String(row[statusIdx]).toLowerCase();
      if (status.indexOf('verified') === -1 || status.indexOf('unverified') !== -1) { skippedUnverified++; return; }
    }
    if (seen.has(email.toLowerCase())) { skippedDuplicate++; return; }
    seen.add(email.toLowerCase());

    const out = leads.headers.map(() => '');
    Object.keys(srcIndex).forEach(target => {
      if (srcIndex[target] !== -1) out[leads.col[target]] = String(row[srcIndex[target]]).trim();
    });
    out[leads.col['Status']] = 'new';
    out[leads.col['Campaign']] = campaign;
    toAdd.push(out);
  });

  if (toAdd.length) {
    leads.sheet.getRange(leads.sheet.getLastRow() + 1, 1, toAdd.length, leads.headers.length).setValues(toAdd);
  }
  ui.alert(
    'Added ' + toAdd.length + ' new lead(s) to the Leads tab for the ' + campaign + ' campaign.\n\n' +
    'Skipped:\n' +
    '• ' + skippedDuplicate + ' already in Leads\n' +
    '• ' + skippedUnverified + ' with an unverified email\n' +
    '• ' + skippedNoEmail + ' missing a name or email');
}

/* ---------------------------------------------------------------- test --- */

function sendTestEmail() {
  const ui = SpreadsheetApp.getUi();
  const ctx = buildContext_();
  const problems = validate_(ctx);
  if (problems.length) { ui.alert('Fix these first:\n\n• ' + problems.join('\n• ')); return; }

  let count = 0;
  Object.keys(ctx.templates).forEach(key => {
    const steps = ctx.templates[key];
    const campaign = steps[0].campaign;
    // Preview with a real lead from this campaign when there is one.
    const firstLead = ctx.leads.rows.find(r => r[ctx.leads.col['Email']] && campaignKey_(r[ctx.leads.col['Campaign']]) === key);
    const lead = firstLead ? leadObject_(ctx.leads, firstLead)
      : { 'First Name': 'Jamie', 'Company': 'Sample Organization' };
    steps.forEach((t, i) => {
      const subject = '[TEST ' + campaign + ' ' + (i + 1) + '/' + steps.length + '] ' +
        (fill_(t.subject, lead) || 'Re: ' + fill_(steps[0].subject, lead));
      sendRaw_({ to: ctx.me, subject: subject, body: buildBody_(ctx, t, lead), fromName: ctx.settings['From Name'], me: ctx.me });
      count++;
    });
  });
  ui.alert('Sent ' + count + ' test email(s) to ' + ctx.me + ', one set per campaign. Check your inbox!');
}

/* ---------------------------------------------------------- automation --- */

function startAutomation() {
  const ui = SpreadsheetApp.getUi();
  const problems = validate_(buildContext_());
  if (problems.length) { ui.alert('Fix these first:\n\n• ' + problems.join('\n• ')); return; }
  removeTriggers_();
  ScriptApp.newTrigger('hourlyRun').timeBased().everyHours(1).create();
  ui.alert('Automatic sending is ON.\n\nEvery hour during your sending window it checks for replies, ' +
    'then sends a few emails. To pause, use Outreach → Stop automatic sending, or set Paused to YES in Settings.');
}

function stopAutomation() {
  removeTriggers_();
  SpreadsheetApp.getUi().alert('Automatic sending is OFF.');
}

function runNow() {
  const summary = run_(true);
  SpreadsheetApp.getUi().alert(summary);
}

/** Called by the hourly trigger. */
function hourlyRun() {
  Logger.log(run_(false));
}

function run_(ignoreWindow) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) return 'Another run is already in progress. Try again in a minute.';
  try {
    const ctx = buildContext_();
    const problems = validate_(ctx);
    if (problems.length) return 'Nothing sent. Fix these first:\n\n• ' + problems.join('\n• ');
    const replies = checkReplies_(ctx);
    const sent = sendDue_(ctx, ignoreWindow);
    return 'Replies/bounces found: ' + replies + '\n' + sent;
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------------------------------------- replies --- */

function checkReplies_(ctx) {
  const leads = ctx.leads;
  const me = ctx.me.toLowerCase();
  const today = dayNumber_(new Date(), ctx.tz);
  let found = 0;

  leads.rows.forEach((row, i) => {
    const status = String(row[leads.col['Status']]).toLowerCase();
    const threadId = String(row[leads.col['Thread ID']]).trim();
    if (!threadId || (status !== 'sent' && status !== 'done')) return;
    const lastSent = row[leads.col['Last Sent']];
    if (lastSent instanceof Date && today - dayNumber_(lastSent, ctx.tz) > 30) return;

    let thread;
    try {
      thread = Gmail.Users.Threads.get('me', threadId, { format: 'metadata', metadataHeaders: ['From'] });
    } catch (e) {
      return; // thread deleted by hand; leave the row alone
    }
    for (const msg of thread.messages || []) {
      const from = header_(msg, 'From');
      if (from.toLowerCase().indexOf(me) !== -1) continue;
      if (BOUNCE_FROM_RE.test(from)) {
        setCell_(leads, i, 'Status', 'bounced');
      } else {
        const newStatus = UNSUBSCRIBE_RE.test(msg.snippet || '') ? 'unsubscribed' : 'replied';
        setCell_(leads, i, 'Status', newStatus);
        alertReply_(ctx, leadObject_(leads, row), newStatus, msg.snippet || '');
      }
      found++;
      break;
    }
  });

  // Bounces that Gmail delivers outside the original thread.
  const byEmail = {};
  leads.rows.forEach((row, i) => {
    const status = String(row[leads.col['Status']]).toLowerCase();
    if (status === 'sent' || status === 'done') byEmail[String(row[leads.col['Email']]).trim().toLowerCase()] = i;
  });
  GmailApp.search('from:(mailer-daemon OR postmaster) newer_than:3d', 0, 50).forEach(thread => {
    thread.getMessages().forEach(m => {
      const emails = m.getPlainBody().toLowerCase().match(/[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}/g) || [];
      emails.forEach(e => {
        if (byEmail[e] !== undefined) {
          setCell_(leads, byEmail[e], 'Status', 'bounced');
          delete byEmail[e];
          found++;
        }
      });
    });
  });

  return found;
}

function alertReply_(ctx, lead, status, snippet) {
  const to = String(ctx.settings['Reply Alert Email'] || '').trim();
  if (!to) return;
  const who = (lead['First Name'] + ' ' + lead['Last Name']).trim() + (lead['Company'] ? ' (' + lead['Company'] + ')' : '');
  GmailApp.sendEmail(to,
    (status === 'replied' ? 'New reply: ' : 'Unsubscribe: ') + who,
    who + ' replied to your outreach:\n\n"' + snippet + '"\n\nOpen ' + ctx.me + ' to answer.');
}

/* ------------------------------------------------------------- sending --- */

function sendDue_(ctx, ignoreWindow) {
  const s = ctx.settings;
  if (isYes_(s['Paused'])) return 'Paused (Settings → Paused is YES). Nothing sent.';

  const now = new Date();
  if (!ignoreWindow) {
    const dow = Number(Utilities.formatDate(now, ctx.tz, 'u')); // 1 = Monday … 7 = Sunday
    const hour = Number(Utilities.formatDate(now, ctx.tz, 'H'));
    if (isYes_(s['Weekdays Only']) && dow >= 6) return 'Weekend. Nothing sent.';
    if (hour < Number(s['Send Start Hour']) || hour >= Number(s['Send End Hour'])) return 'Outside sending hours. Nothing sent.';
  }

  const props = PropertiesService.getScriptProperties();
  const todayKey = 'sent:' + Utilities.formatDate(now, ctx.tz, 'yyyy-MM-dd');
  const sentToday = Number(props.getProperty(todayKey) || 0);
  const limit = dailyLimit_(ctx, props, now);
  const budget = Math.min(limit - sentToday, Number(s['Max Per Run']) || 4);
  if (budget <= 0) return 'Daily limit reached (' + sentToday + '/' + limit + '). Nothing more sent today.';

  const queue = buildQueue_(ctx, now);
  const batch = queue.slice(0, budget);
  let sent = 0;
  const errors = [];

  batch.forEach((item, n) => {
    try {
      sendStep_(ctx, item.index, item.step);
      sent++;
      props.setProperty(todayKey, String(sentToday + sent));
      if (!props.getProperty('firstSendDay')) props.setProperty('firstSendDay', String(dayNumber_(now, ctx.tz)));
    } catch (e) {
      errors.push(leadObject_(ctx.leads, ctx.leads.rows[item.index])['Email'] + ': ' + e.message);
      setCell_(ctx.leads, item.index, 'Notes', 'Send error: ' + e.message);
    }
    // Small random pause so emails don't go out in a burst.
    if (n < batch.length - 1) Utilities.sleep(20000 + Math.floor(Math.random() * 25000));
  });

  return 'Sent ' + sent + ' email(s). Today: ' + (sentToday + sent) + '/' + limit + '. ' +
    'Still waiting in line: ' + (queue.length - sent) + '.' +
    (errors.length ? '\n\nErrors:\n• ' + errors.join('\n• ') : '');
}

/** Follow-ups that are due come first, then brand-new leads. */
function buildQueue_(ctx, now) {
  const leads = ctx.leads;
  const today = dayNumber_(now, ctx.tz);
  const followUps = [];
  const fresh = [];

  leads.rows.forEach((row, i) => {
    const status = String(row[leads.col['Status']]).trim().toLowerCase();
    const email = String(row[leads.col['Email']]).trim();
    if (!email || !String(row[leads.col['First Name']]).trim()) return;
    if (status === '' || status === 'new') {
      if (templatesFor_(ctx, row).length) fresh.push({ index: i, step: 1 });
    } else if (status === 'sent') {
      const step = Number(row[leads.col['Step']]) || 0;
      const next = templatesFor_(ctx, row)[step]; // index `step` holds step+1
      const lastSent = row[leads.col['Last Sent']];
      if (!next || !(lastSent instanceof Date) || !row[leads.col['Thread ID']]) return;
      if (today - dayNumber_(lastSent, ctx.tz) >= next.wait) followUps.push({ index: i, step: step + 1 });
    }
  });
  return followUps.concat(fresh);
}

function sendStep_(ctx, index, stepNumber) {
  const leads = ctx.leads;
  const row = leads.rows[index];
  const lead = leadObject_(leads, row);
  const steps = templatesFor_(ctx, row);
  const template = steps[stepNumber - 1];
  const body = buildBody_(ctx, template, lead);
  const msg = { to: lead['Email'], body: body, fromName: ctx.settings['From Name'], me: ctx.me };

  if (stepNumber === 1) {
    msg.subject = fill_(template.subject, lead);
  } else {
    // Reply in the same thread so it shows up under the first email.
    const threadId = String(row[leads.col['Thread ID']]);
    const thread = Gmail.Users.Threads.get('me', threadId, { format: 'metadata', metadataHeaders: ['Message-ID', 'Subject'] });
    const first = thread.messages[0];
    const originalSubject = header_(first, 'Subject');
    msg.subject = template.subject ? fill_(template.subject, lead)
      : (/^re:/i.test(originalSubject) ? originalSubject : 'Re: ' + originalSubject);
    msg.threadId = threadId;
    msg.inReplyTo = header_(first, 'Message-ID');
  }

  const result = sendRaw_(msg);
  const isLast = stepNumber >= steps.length;
  setCell_(leads, index, 'Thread ID', result.threadId);
  setCell_(leads, index, 'Step', stepNumber);
  setCell_(leads, index, 'Last Sent', new Date());
  setCell_(leads, index, 'Status', isLast ? 'done' : 'sent');
}

function sendRaw_(o) {
  const headers = [
    'From: ' + encodeName_(o.fromName) + ' <' + o.me + '>',
    'To: ' + o.to,
    'Subject: ' + encodeHeader_(o.subject),
    'MIME-Version: 1.0',
    'Content-Type: text/plain; charset=UTF-8',
    'Content-Transfer-Encoding: base64',
  ];
  if (o.inReplyTo) headers.push('In-Reply-To: ' + o.inReplyTo, 'References: ' + o.inReplyTo);
  const body = Utilities.base64Encode(o.body, Utilities.Charset.UTF_8).replace(/(.{76})/g, '$1\r\n');
  const resource = { raw: Utilities.base64EncodeWebSafe(headers.join('\r\n') + '\r\n\r\n' + body) };
  if (o.threadId) resource.threadId = o.threadId;
  return Gmail.Users.Messages.send(resource, 'me');
}

/* ------------------------------------------------------------- helpers --- */

function buildContext_() {
  const ss = SpreadsheetApp.getActive();
  return {
    me: Session.getEffectiveUser().getEmail(),
    tz: ss.getSpreadsheetTimeZone(),
    settings: getSettings_(),
    templates: getTemplates_(),
    leads: loadLeads_(),
  };
}

function validate_(ctx) {
  const problems = [];
  if (typeof Gmail === 'undefined') {
    problems.push('Turn on the Gmail service: in the Apps Script editor click Services (+) → Gmail API → Add.');
  }
  const sig = String(ctx.settings['Signature'] || '');
  if (/\[[A-Z ]+\]/.test(sig)) problems.push('Replace the [BRACKETS] in Settings → Signature with your real info.');
  const keys = Object.keys(ctx.templates);
  if (!keys.length) problems.push('The Templates tab has no emails in it.');
  keys.forEach(key => {
    const first = ctx.templates[key][0];
    if (!first.subject) problems.push('Email step 1 for ' + first.campaign + ' in the Templates tab needs a Subject.');
  });
  return problems;
}

function getSettings_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(TAB_SETTINGS);
  if (!sheet) throw new Error('Missing Settings tab. Run Outreach → 1. Set up tabs.');
  const out = {};
  sheet.getDataRange().getValues().slice(1).forEach(r => { if (r[0]) out[String(r[0]).trim()] = r[1]; });
  return out;
}

/** Returns { campaignKey: [step 1, step 2, …] } from the Templates tab. */
function getTemplates_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(TAB_TEMPLATES);
  if (!sheet) throw new Error('Missing Templates tab. Run Outreach → 1. Set up tabs.');
  const col = templateColumns_(sheet);
  const out = {};
  sheet.getDataRange().getValues().slice(1).forEach(r => {
    if (!String(r[col['Body']]).trim()) return;
    const campaign = (col['Campaign'] !== undefined && String(r[col['Campaign']]).trim()) || DEFAULT_CAMPAIGN;
    const key = campaignKey_(campaign);
    (out[key] = out[key] || []).push({
      campaign: campaign,
      step: Number(r[col['Step']]),
      wait: Number(r[col['Wait Days']]) || 0,
      subject: String(r[col['Subject']]).trim(),
      body: String(r[col['Body']]),
    });
  });
  Object.keys(out).forEach(k => out[k].sort((a, b) => a.step - b.step));
  return out;
}

/** Column positions in the Templates tab, by header name. */
function templateColumns_(sheet) {
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0].map(h => String(h).trim());
  const col = {};
  TEMPLATE_HEADERS.forEach((h, i) => {
    const at = headers.indexOf(h);
    if (at !== -1) col[h] = at;
    else if (h !== 'Campaign') col[h] = i; // tabs from before headers were checked
  });
  return col;
}

function templatesFor_(ctx, row) {
  return ctx.templates[campaignKey_(row[ctx.leads.col['Campaign']])] || [];
}

function campaignKey_(name) {
  return (String(name || '').trim() || DEFAULT_CAMPAIGN).toLowerCase();
}

function loadLeads_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(TAB_LEADS);
  if (!sheet) throw new Error('Missing Leads tab. Run Outreach → 1. Set up tabs.');
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const col = {};
  headers.forEach((h, i) => { if (h && col[h] === undefined) col[h] = i; });
  LEAD_HEADERS.forEach(h => {
    if (col[h] === undefined) throw new Error('Leads tab is missing the "' + h + '" column. Run Outreach → 1. Set up tabs.');
  });
  return { sheet: sheet, headers: headers, col: col, rows: values.slice(1) };
}

function setCell_(leads, index, header, value) {
  leads.sheet.getRange(index + 2, leads.col[header] + 1).setValue(value);
  leads.rows[index][leads.col[header]] = value;
}

function leadObject_(leads, row) {
  const out = {};
  leads.headers.forEach((h, i) => { if (h) out[h] = String(row[i] === undefined ? '' : row[i]).trim(); });
  return out;
}

function buildBody_(ctx, template, lead) {
  const parts = [fill_(template.body, lead).trim(), String(ctx.settings['Signature'] || '').trim()];
  const optOut = String(ctx.settings['Opt-out Line'] || '').trim();
  if (optOut) parts.push(optOut);
  return parts.filter(String).join('\n\n');
}

/** Replaces {{Column Name}} or {{Column Name|fallback}} with the lead's value. */
function fill_(text, lead) {
  return String(text || '').replace(/\{\{\s*([^}|]+?)\s*(?:\|([^}]*))?\}\}/g, (m, key, fallback) => {
    const value = lead[key] ? String(lead[key]).trim() : '';
    return value || (fallback || '').trim();
  });
}

function dailyLimit_(ctx, props, now) {
  const s = ctx.settings;
  const first = props.getProperty('firstSendDay');
  const days = first ? dayNumber_(now, ctx.tz) - Number(first) : 0;
  if (days < 7) return Number(s['Daily Limit Week 1']) || 10;
  if (days < 14) return Number(s['Daily Limit Week 2']) || 20;
  return Number(s['Daily Limit After']) || 25;
}

/** Whole days since 1970 in the sheet's time zone, for "how many days ago" math. */
function dayNumber_(date, tz) {
  const [y, m, d] = Utilities.formatDate(date, tz, 'yyyy-MM-dd').split('-').map(Number);
  return Math.round(Date.UTC(y, m - 1, d) / 86400000);
}

function header_(msg, name) {
  const h = ((msg.payload && msg.payload.headers) || []).find(x => x.name.toLowerCase() === name.toLowerCase());
  return h ? h.value : '';
}

function encodeHeader_(s) {
  return /^[\x20-\x7e]*$/.test(s) ? s : '=?UTF-8?B?' + Utilities.base64Encode(s, Utilities.Charset.UTF_8) + '?=';
}

function encodeName_(s) {
  return /^[\x20-\x7e]*$/.test(s) ? '"' + String(s).replace(/"/g, '') + '"' : encodeHeader_(s);
}

function isYes_(v) {
  return v === true || /^(yes|y|true)$/i.test(String(v).trim());
}

function removeTriggers_() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'hourlyRun')
    .forEach(t => ScriptApp.deleteTrigger(t));
}
