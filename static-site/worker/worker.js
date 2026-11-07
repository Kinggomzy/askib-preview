/**
 * askib.in form handler: a Cloudflare Worker (free plan) that uses the
 * company's own Microsoft 365.
 *
 * For each enquiry from the website it:
 *   1. emails it to info@askib.in, sent from info@askib.in, with the customer as reply-to
 *   2. adds a row to the "Enquiries" sheet of the company's Excel file on OneDrive/SharePoint
 *
 * It talks to Microsoft 365 through Microsoft Graph with an app registration
 * (client id + secret). Setup: static-site/worker/README.md.
 *
 * Settings (Cloudflare → Worker → Settings → Variables):
 *   TENANT_ID, CLIENT_ID            from the app registration
 *   CLIENT_SECRET                   from the app registration (add as a Secret)
 *   MAILBOX                         info@askib.in
 *   EXCEL_SHARE_URL                 the Excel file's share link
 *   ALLOWED_ORIGINS                 https://askib.in,https://www.askib.in
 *   WORKBOOK_NAME (optional)        file name to give the workbook, default "ASKIB Website Enquiries.xlsx"
 *
 * On its first enquiry it renames the workbook and adds an "Enquiries" sheet
 * (a formatted table) and a "How to use" sheet. Existing sheets are left alone.
 */

const SHEET = 'Enquiries'
const TABLE = 'EnquiriesTable'
const GUIDE = 'How to use'
const HEADERS = ['Received', 'Reference', 'Form', 'Status', 'Name', 'Company', 'Email', 'Phone', 'City', 'Details', 'Page', 'Notes']
/** Column widths in points, in the order of HEADERS. */
const COLUMN_WIDTHS = [110, 110, 120, 70, 130, 150, 180, 110, 90, 380, 180, 220]
const GUIDE_LINES = [
  'ASKIB website enquiries',
  '',
  'Every form sent from askib.in (contact, quotes, service requests, consultations) is added to the Enquiries sheet as a new row, and emailed to info@askib.in.',
  'New rows start with Status "New". Change it as you follow up, for example: Contacted, Quoted, Won, Lost, Spam.',
  'Use the Notes column for anything the team should know. Filter or sort with the arrows in the header row.',
  'Do not rename the Enquiries sheet or its table (EnquiriesTable), and do not delete the header row: the website writes to them.',
  'The Reference matches the one the customer saw on screen and the one in the email subject.',
]
const MIN_FILL_MS = 3000
const MAX_PER_EMAIL_PER_HOUR = 5
const MAX_FIELDS = 120
const MAX_VALUE = 5000
const MAX_BODY = 200_000

// Kept between requests while the worker stays warm.
let token = null // { value, expires }
let workbook = null // promise of the workbook's Graph path
const recent = new Map() // email → [timestamps]

export default {
  async fetch(request, env) {
    const origin = request.headers.get('Origin') || ''
    const allowed = (env.ALLOWED_ORIGINS || 'https://askib.in').split(',').map((s) => s.trim())
    const cors = {
      'Access-Control-Allow-Origin': allowed.includes(origin) ? origin : allowed[0],
      'Access-Control-Allow-Methods': 'POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
      Vary: 'Origin',
    }
    const json = (obj, status = 200) => new Response(JSON.stringify(obj), { status, headers: { ...cors, 'Content-Type': 'application/json' } })

    if (request.method === 'OPTIONS') return new Response(null, { status: 204, headers: cors })
    if (request.method === 'GET') return new Response('askib.in form handler is running.', { headers: cors })
    if (request.method !== 'POST') return json({ ok: false, error: 'method not allowed' }, 405)
    if (!allowed.includes(origin)) return json({ ok: false, error: 'forbidden' }, 403)

    let data
    try {
      const text = await request.text()
      if (text.length > MAX_BODY) return json({ ok: false, error: 'too large' }, 413)
      data = JSON.parse(text)
    } catch {
      return json({ ok: false, error: 'bad request' }, 400)
    }

    // Bots: hidden field filled, or sent impossibly fast. Say ok, do nothing.
    if (data.botcheck || !(Number(data.fillMs) >= MIN_FILL_MS)) return json({ ok: true })

    const email = clean(data.replyTo, 200)
    const name = clean(data.name, 200)
    const form = clean(data.form, 100) || 'Website enquiry'
    const reference = clean(data.reference, 40)
    const page = clean(data.page, 500)
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || !name) return json({ ok: false, error: 'missing name or email' }, 400)
    const fields = (Array.isArray(data.fields) ? data.fields : [])
      .slice(0, MAX_FIELDS)
      .map((f) => [clean(f && f[0], 120), clean(f && f[1], MAX_VALUE)])
      .filter((f) => f[0] && f[1])

    const now = Date.now()
    const key = email.toLowerCase()
    const times = (recent.get(key) || []).filter((t) => now - t < 3600_000)
    if (times.length >= MAX_PER_EMAIL_PER_HOUR) return json({ ok: false, error: 'too many enquiries, please try again later' }, 429)
    times.push(now)
    recent.set(key, times)
    if (recent.size > 5000) recent.clear()

    const subject = clean(data.subject, 200) || `${form} · ${reference}`
    const [mail, row] = await Promise.allSettled([
      sendMail(env, { subject, email, name, form, reference, fields, page }),
      addRow(env, { form, reference, email, fields, page }),
    ])
    if (mail.status === 'rejected') console.error('mail failed:', mail.reason?.message)
    if (row.status === 'rejected') console.error('excel failed:', row.reason?.message)
    // Either one reaching the office is enough; only fail if both did.
    if (mail.status === 'rejected' && row.status === 'rejected') return json({ ok: false, error: 'could not deliver' }, 502)
    return json({ ok: true, reference })
  },
}

// ── Microsoft Graph ───────────────────────────────────────────────────────────
async function graph(env, path, init = {}) {
  const res = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    ...init,
    headers: { Authorization: `Bearer ${await accessToken(env)}`, 'Content-Type': 'application/json', ...(init.headers || {}) },
  })
  if (!res.ok) throw new Error(`${init.method || 'GET'} ${path.split('?')[0]} → ${res.status} ${(await res.text()).slice(0, 300)}`)
  return res.status === 204 || res.status === 202 ? null : res.json()
}

let pendingToken = null
function accessToken(env) {
  if (token && token.expires > Date.now() + 60_000) return token.value
  // Mail and Excel ask at the same moment; share one request.
  pendingToken ??= fetchToken(env).finally(() => (pendingToken = null))
  return pendingToken
}

async function fetchToken(env) {
  const res = await fetch(`https://login.microsoftonline.com/${env.TENANT_ID}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: env.CLIENT_ID,
      client_secret: env.CLIENT_SECRET,
      scope: 'https://graph.microsoft.com/.default',
      grant_type: 'client_credentials',
    }),
  })
  const body = await res.json()
  if (!res.ok) throw new Error(`token → ${res.status} ${body.error_description || body.error}`)
  token = { value: body.access_token, expires: Date.now() + body.expires_in * 1000 }
  return token.value
}

async function sendMail(env, { subject, email, name, form, reference, fields, page }) {
  const mailbox = env.MAILBOX || 'info@askib.in'
  await graph(env, `/users/${encodeURIComponent(mailbox)}/sendMail`, {
    method: 'POST',
    body: JSON.stringify({
      message: {
        subject,
        body: { contentType: 'HTML', content: mailHtml(form, reference, fields, page) },
        toRecipients: [{ emailAddress: { address: mailbox } }],
        replyTo: [{ emailAddress: { address: email, name } }],
      },
      saveToSentItems: false,
    }),
  })
}

/** Graph paths of the workbook and its file, found from the share link once. */
function workbookPath(env) {
  // One lookup (and first-run setup) at a time, shared by concurrent enquiries.
  workbook ??= findWorkbook(env).catch((err) => {
    workbook = null
    throw err
  })
  return workbook
}

async function findWorkbook(env) {
  const id = 'u!' + btoa(env.EXCEL_SHARE_URL).replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-')
  const item = await graph(env, `/shares/${id}/driveItem?$select=id,name,parentReference`)
  const file = `/drives/${item.parentReference.driveId}/items/${item.id}`
  await setUpWorkbook(env, file, item.name)
  return `${file}/workbook`
}

const ws = (name) => `worksheets/${encodeURIComponent(name)}`
const col = (i) => String.fromCharCode(65 + i) // 0 → A

/**
 * First run only: renames the file, adds the Enquiries sheet with a formatted
 * table, and a "How to use" sheet. Never deletes or changes existing sheets.
 */
async function setUpWorkbook(env, file, currentName) {
  const wb = `${file}/workbook`
  const tables = await graph(env, `${wb}/tables?$select=name`)
  if (tables.value.some((t) => t.name === TABLE)) return

  const wanted = env.WORKBOOK_NAME || 'ASKIB Website Enquiries.xlsx'
  if (currentName !== wanted) await graph(env, file, { method: 'PATCH', body: JSON.stringify({ name: wanted }) })

  const sheets = (await graph(env, `${wb}/worksheets?$select=name`)).value.map((s) => s.name)
  if (!sheets.includes(SHEET)) await graph(env, `${wb}/worksheets/add`, { method: 'POST', body: JSON.stringify({ name: SHEET }) })
  const last = col(HEADERS.length - 1)
  await graph(env, `${wb}/${ws(SHEET)}/range(address='A1:${last}1')`, { method: 'PATCH', body: JSON.stringify({ values: [HEADERS] }) })
  const table = await graph(env, `${wb}/${ws(SHEET)}/tables/add`, {
    method: 'POST',
    body: JSON.stringify({ address: `${SHEET}!A1:${last}1`, hasHeaders: true }),
  })
  await graph(env, `${wb}/tables/${table.id}`, { method: 'PATCH', body: JSON.stringify({ name: TABLE, style: 'TableStyleMedium2' }) })
  await graph(env, `${wb}/${ws(SHEET)}/freezePanes/freezeRows`, { method: 'POST', body: JSON.stringify({ count: 1 }) })
  for (const [i, width] of COLUMN_WIDTHS.entries()) {
    const c = col(i)
    await graph(env, `${wb}/${ws(SHEET)}/range(address='${c}:${c}')/format`, {
      method: 'PATCH',
      body: JSON.stringify({ columnWidth: width, wrapText: HEADERS[i] === 'Details' || HEADERS[i] === 'Notes', verticalAlignment: 'Top' }),
    })
  }

  if (!sheets.includes(GUIDE)) {
    await graph(env, `${wb}/worksheets/add`, { method: 'POST', body: JSON.stringify({ name: GUIDE }) })
    await graph(env, `${wb}/${ws(GUIDE)}/range(address='A1:A${GUIDE_LINES.length}')`, {
      method: 'PATCH',
      body: JSON.stringify({ values: GUIDE_LINES.map((l) => [l]) }),
    })
    await graph(env, `${wb}/${ws(GUIDE)}/range(address='A1')/format/font`, { method: 'PATCH', body: JSON.stringify({ bold: true, size: 14 }) })
    await graph(env, `${wb}/${ws(GUIDE)}/range(address='A:A')/format`, { method: 'PATCH', body: JSON.stringify({ columnWidth: 600, wrapText: true }) })
  }
  await graph(env, `${wb}/${ws(SHEET)}`, { method: 'PATCH', body: JSON.stringify({ position: 0 }) })
}

async function addRow(env, { form, reference, email, fields, page }) {
  const wb = await workbookPath(env)
  const get = (label) => (fields.find((f) => f[0] === label) || [])[1] || ''
  const shown = ['Your name', 'Email', 'Phone', 'Company', 'City']
  const details = fields
    .filter((f) => !shown.includes(f[0]))
    .map((f) => `${f[0]}: ${f[1]}`)
    .join('\n')
  const received = new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata', dateStyle: 'medium', timeStyle: 'short' })
  const byHeader = {
    Received: received,
    Reference: reference,
    Form: form,
    Status: 'New',
    Name: get('Your name'),
    Company: get('Company'),
    Email: email,
    Phone: get('Phone'),
    City: get('City'),
    Details: details,
    Page: page,
    Notes: '',
  }
  // A leading = + - @ would make Excel run the text as a formula.
  const safe = (v) => (/^[=+\-@]/.test(v) ? `'${v}` : v)
  const values = HEADERS.map((h) => safe(byHeader[h]))
  await graph(env, `${wb}/tables/${TABLE}/rows/add`, { method: 'POST', body: JSON.stringify({ values: [values] }) })
}

function mailHtml(form, reference, fields, page) {
  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
  const rows = fields
    .map(
      (f) =>
        `<tr><td style="padding:6px 12px 6px 0;color:#555;vertical-align:top;white-space:nowrap">${esc(f[0])}</td>` +
        `<td style="padding:6px 0;white-space:pre-wrap">${esc(f[1])}</td></tr>`,
    )
    .join('')
  return (
    `<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#111">` +
    `<h2 style="margin:0 0 4px">${esc(form)}</h2>` +
    `<p style="margin:0 0 16px;color:#555">Reference ${esc(reference)}. Reply to this email to answer the customer.</p>` +
    `<table style="border-collapse:collapse">${rows}</table>` +
    `<p style="margin-top:16px;color:#888;font-size:12px">Sent from ${esc(page)}</p></div>`
  )
}

function clean(v, max) {
  return String(v == null ? '' : v).trim().slice(0, max)
}
