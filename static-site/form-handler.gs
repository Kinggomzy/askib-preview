/**
 * askib.in form handler: a free Google Apps Script web app.
 *
 * Receives the enquiry forms from askib.in, emails each one to the office and
 * saves it as a row in the Google Sheet this script belongs to. Runs on the
 * company's own Google account; nothing to pay and no outside service.
 * Setup steps: static-site/README.md.
 */

// ── Settings ─────────────────────────────────────────────────────────────────
const OFFICE_EMAIL = 'info@askib.in' // where enquiries are sent
const SHEET_NAME = 'Enquiries' // tab that collects them (made automatically)
const MIN_FILL_MS = 3000 // faster than this is a bot
const MAX_PER_EMAIL_PER_HOUR = 5 // stops one sender flooding the inbox
const MAX_FIELDS = 120
const MAX_VALUE_LENGTH = 5000

function doPost(e) {
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}')

    // Bots: the hidden field is filled in, or the form was sent impossibly fast.
    // Answer "ok" so they do not learn anything, but send and save nothing.
    if (data.botcheck || !(Number(data.fillMs) >= MIN_FILL_MS)) return reply({ ok: true })

    const email = clean(data.replyTo, 200)
    const name = clean(data.name, 200)
    const form = clean(data.form, 100) || 'Website enquiry'
    const reference = clean(data.reference, 40)
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || !name) return reply({ ok: false, error: 'missing name or email' })
    const fields = (Array.isArray(data.fields) ? data.fields : [])
      .slice(0, MAX_FIELDS)
      .map((f) => [clean(f && f[0], 120), clean(f && f[1], MAX_VALUE_LENGTH)])
      .filter((f) => f[0] && f[1])

    // One sender cannot send more than a handful an hour.
    const cache = CacheService.getScriptCache()
    const key = 'n:' + email.toLowerCase()
    const count = Number(cache.get(key) || 0)
    if (count >= MAX_PER_EMAIL_PER_HOUR) return reply({ ok: false, error: 'too many enquiries, please try again later' })
    cache.put(key, String(count + 1), 3600)

    save(form, reference, email, fields, clean(data.page, 500))
    if (MailApp.getRemainingDailyQuota() > 0) {
      MailApp.sendEmail({
        to: OFFICE_EMAIL,
        replyTo: email,
        name: 'askib.in website',
        subject: clean(data.subject, 200) || form + ' · ' + reference,
        body: fields.map((f) => f[0] + ': ' + f[1]).join('\n') + '\n\nReference: ' + reference + '\nPage: ' + clean(data.page, 500),
        htmlBody: html(form, reference, fields, clean(data.page, 500)),
      })
    }
    return reply({ ok: true, reference: reference })
  } catch (err) {
    console.error(err)
    return reply({ ok: false, error: 'server error' })
  }
}

/** Opening the web app address in a browser shows this, to check it is live. */
function doGet() {
  return ContentService.createTextOutput('askib.in form handler is running.')
}

function save(form, reference, email, fields, page) {
  const book = SpreadsheetApp.getActiveSpreadsheet()
  let sheet = book.getSheetByName(SHEET_NAME)
  if (!sheet) {
    sheet = book.insertSheet(SHEET_NAME)
    sheet.appendRow(['Received', 'Reference', 'Form', 'Name', 'Email', 'Phone', 'Company', 'City', 'Details', 'Page'])
    sheet.setFrozenRows(1)
    sheet.getRange('1:1').setFontWeight('bold')
  }
  const get = (label) => (fields.find((f) => f[0] === label) || [, ''])[1]
  const shown = ['Your name', 'Email', 'Phone', 'Company', 'City']
  const details = fields
    .filter((f) => shown.indexOf(f[0]) < 0)
    .map((f) => f[0] + ': ' + f[1])
    .join('\n')
  // A leading = + - @ would make the sheet run the text as a formula.
  const safe = (v) => (/^[=+\-@]/.test(v) ? "'" + v : v)
  sheet.appendRow([new Date(), reference, form, get('Your name'), email, get('Phone'), get('Company'), get('City'), details, page].map((v) => (typeof v === 'string' ? safe(v) : v)))
}

function html(form, reference, fields, page) {
  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
  const rows = fields
    .map(
      (f) =>
        '<tr><td style="padding:6px 12px 6px 0;color:#555;vertical-align:top;white-space:nowrap">' +
        esc(f[0]) +
        '</td><td style="padding:6px 0;white-space:pre-wrap">' +
        esc(f[1]) +
        '</td></tr>',
    )
    .join('')
  return (
    '<div style="font-family:Arial,sans-serif;font-size:14px;color:#111">' +
    '<h2 style="margin:0 0 4px">' + esc(form) + '</h2>' +
    '<p style="margin:0 0 16px;color:#555">Reference ' + esc(reference) + '. Reply to this email to answer the customer.</p>' +
    '<table style="border-collapse:collapse">' + rows + '</table>' +
    '<p style="margin-top:16px;color:#888;font-size:12px">Sent from ' + esc(page) + '</p></div>'
  )
}

function clean(v, max) {
  return String(v == null ? '' : v).trim().slice(0, max)
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON)
}
