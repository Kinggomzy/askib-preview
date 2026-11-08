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
/** Testimonials carry one photo, already shrunk in the browser (about 100–300 KB). */
const MAX_TESTIMONIAL_BODY = 1_500_000
const MAX_PHOTO_BYTES = 1_000_000

// Kept between requests while the worker stays warm.
let token = null // { value, expires }
let workbook = null // promise of { wb, drive, parent }: the workbook's Graph path and its folder
let testimonialSheet = null // promise: the Testimonials sheet exists
const recent = new Map() // email → [timestamps]

export default {
  async fetch(request, env) {
    const origin = request.headers.get('Origin') || ''
    const allowed = (env.ALLOWED_ORIGINS || 'https://askib.in').split(',').map((s) => s.trim())
    const cors = {
      'Access-Control-Allow-Origin': allowed.includes(origin) ? origin : allowed[0],
      'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type',
      Vary: 'Origin',
    }
    const json = (obj, status = 200) => new Response(JSON.stringify(obj), { status, headers: { ...cors, 'Content-Type': 'application/json' } })

    if (request.method === 'OPTIONS') return new Response(null, { status: 204, headers: cors })
    if (request.method === 'GET') {
      const { pathname } = new URL(request.url)
      if (pathname === '/testimonials') return approvedTestimonials(request, env)
      const photo = pathname.match(/^\/testimonials\/photo\/([A-Za-z0-9-]{4,40})\.jpg$/)
      if (photo) return testimonialPhoto(request, env, photo[1])
      return new Response('askib.in form handler is running.', { headers: cors })
    }
    if (request.method !== 'POST') return json({ ok: false, error: 'method not allowed' }, 405)
    if (!allowed.includes(origin)) return json({ ok: false, error: 'forbidden' }, 403)

    let data
    try {
      const text = await request.text()
      if (text.length > MAX_TESTIMONIAL_BODY) return json({ ok: false, error: 'too large' }, 413)
      data = JSON.parse(text)
      if (data.kind !== 'testimonial' && text.length > MAX_BODY) return json({ ok: false, error: 'too large' }, 413)
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

    if (data.kind === 'testimonial') {
      const t = data.testimonial || {}
      const entry = {
        reference,
        name,
        email,
        company: clean(t.company, 200),
        role: clean(t.role, 120),
        city: clean(t.city, 120),
        phone: clean(t.phone, 40),
        service: clean(t.service, 120),
        rating: Math.min(5, Math.max(0, Math.round(Number(t.rating) || 0))),
        text: clean(t.text, 2000),
        page,
      }
      if (!entry.company || entry.text.length < 10) return json({ ok: false, error: 'missing company or testimonial' }, 400)
      let photo = null
      if (data.photo) {
        photo = jpegBytes(data.photo)
        if (!photo) return json({ ok: false, error: 'the photo must be a JPEG under 1 MB' }, 400)
      }
      const [saved, mail] = await Promise.allSettled([
        saveTestimonial(env, entry, photo),
        sendTestimonialMail(env, entry, photo, data.photo),
      ])
      if (saved.status === 'rejected') console.error('testimonial excel failed:', saved.reason?.message)
      if (mail.status === 'rejected') console.error('testimonial mail failed:', mail.reason?.message)
      if (saved.status === 'rejected' && mail.status === 'rejected') return json({ ok: false, error: 'could not deliver' }, 502)
      return json({ ok: true, reference })
    }

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
  return { wb: `${file}/workbook`, drive: item.parentReference.driveId, parent: item.parentReference.id }
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
  const { wb } = await workbookPath(env)
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

const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

function mailHtml(form, reference, fields, page) {
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

// ── Testimonials ──────────────────────────────────────────────────────────────
// Clients send testimonials from askib.in/testimonial. Each one is added to the
// "Testimonials" sheet with Status "Pending", its photo is saved in the
// "Testimonial photos" folder next to the workbook, and info@askib.in gets an
// email with the photo attached. Nothing is shown on the website until someone
// changes Status to "Approved" in the sheet.
const T_SHEET = 'Testimonials'
const T_TABLE = 'TestimonialsTable'
const T_HEADERS = ['Received', 'Reference', 'Status', 'Name', 'Company', 'Role', 'City', 'Rating', 'Service', 'Testimonial', 'Photo', 'Show on homepage', 'Order', 'Email', 'Phone', 'Notes']
const T_WIDTHS = [110, 110, 80, 130, 160, 120, 90, 55, 110, 380, 160, 90, 50, 180, 110, 200]
const PHOTO_FOLDER = 'Testimonial photos'
const APPROVED = /^(approved|published|yes|live)$/i
/** How long the website's copy of the approved list and photos is kept before asking Excel again. */
const LIST_CACHE_SECONDS = 300
const PHOTO_CACHE_SECONDS = 86400

function jpegBytes(dataUrl) {
  const m = String(dataUrl).match(/^data:image\/jpeg;base64,([A-Za-z0-9+/=]+)$/)
  if (!m) return null
  let bin
  try {
    bin = atob(m[1])
  } catch {
    return null
  }
  if (bin.length > MAX_PHOTO_BYTES || bin.charCodeAt(0) !== 0xff || bin.charCodeAt(1) !== 0xd8) return null
  const bytes = new Uint8Array(bin.length)
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i)
  return bytes
}

function ensureTestimonialSheet(env, wb) {
  testimonialSheet ??= setUpTestimonialSheet(env, wb).catch((err) => {
    testimonialSheet = null
    throw err
  })
  return testimonialSheet
}

async function setUpTestimonialSheet(env, wb) {
  const tables = await graph(env, `${wb}/tables?$select=name`)
  if (tables.value.some((t) => t.name === T_TABLE)) return
  const sheets = (await graph(env, `${wb}/worksheets?$select=name`)).value.map((s) => s.name)
  if (!sheets.includes(T_SHEET)) await graph(env, `${wb}/worksheets/add`, { method: 'POST', body: JSON.stringify({ name: T_SHEET }) })
  const last = col(T_HEADERS.length - 1)
  await graph(env, `${wb}/${ws(T_SHEET)}/range(address='A1:${last}1')`, { method: 'PATCH', body: JSON.stringify({ values: [T_HEADERS] }) })
  const table = await graph(env, `${wb}/${ws(T_SHEET)}/tables/add`, {
    method: 'POST',
    body: JSON.stringify({ address: `${T_SHEET}!A1:${last}1`, hasHeaders: true }),
  })
  await graph(env, `${wb}/tables/${table.id}`, { method: 'PATCH', body: JSON.stringify({ name: T_TABLE, style: 'TableStyleMedium4' }) })
  await graph(env, `${wb}/${ws(T_SHEET)}/freezePanes/freezeRows`, { method: 'POST', body: JSON.stringify({ count: 1 }) })
  for (const [i, width] of T_WIDTHS.entries()) {
    const c = col(i)
    await graph(env, `${wb}/${ws(T_SHEET)}/range(address='${c}:${c}')/format`, {
      method: 'PATCH',
      body: JSON.stringify({ columnWidth: width, wrapText: ['Testimonial', 'Notes'].includes(T_HEADERS[i]), verticalAlignment: 'Top' }),
    })
  }
  // A drop-down on Status so approving is one click.
  try {
    await graph(env, `${wb}/${ws(T_SHEET)}/range(address='C2:C2000')/dataValidation`, {
      method: 'PATCH',
      body: JSON.stringify({ rule: { list: { inCellDropDown: true, source: 'Pending,Approved,Rejected' } } }),
    })
  } catch (err) {
    console.error('status drop-down not added:', err.message)
  }
  if (sheets.includes(GUIDE)) {
    const used = await graph(env, `${wb}/${ws(GUIDE)}/usedRange?$select=rowCount`)
    const start = (used?.rowCount || GUIDE_LINES.length) + 2
    const lines = [
      'Testimonials',
      'Clients send testimonials from https://askib.in/testimonial. Each one arrives on the Testimonials sheet with Status "Pending" and is emailed to info@askib.in with the photo attached.',
      'To show one on the website, change its Status to "Approved". It appears within about 5 minutes. Change it back to "Pending" or "Rejected" to take it down.',
      'Type "Yes" in "Show on homepage" to feature it on the homepage too. Use "Order" (1, 2, 3…) to choose which come first.',
      'You can edit the wording (for example to fix a typo) before approving. Photos are in the "Testimonial photos" folder next to this file, named by Reference. Replace a photo by uploading a new file with the same name.',
      'To add one yourself: add a row, give it a unique Reference (e.g. TST-MANUAL-01), fill Name, Company and Testimonial, and either upload TST-MANUAL-01.jpg to the photos folder or paste a public image link in Photo.',
    ]
    await graph(env, `${wb}/${ws(GUIDE)}/range(address='A${start}:A${start + lines.length - 1}')`, {
      method: 'PATCH',
      body: JSON.stringify({ values: lines.map((l) => [l]) }),
    })
    await graph(env, `${wb}/${ws(GUIDE)}/range(address='A${start}')/format/font`, { method: 'PATCH', body: JSON.stringify({ bold: true, size: 13 }) })
  }
}

async function saveTestimonial(env, t, photo) {
  const info = await workbookPath(env)
  await ensureTestimonialSheet(env, info.wb)
  let photoLink = ''
  if (photo) {
    const item = await graph(env, `/drives/${info.drive}/items/${info.parent}:/${encodeURIComponent(PHOTO_FOLDER)}/${t.reference}.jpg:/content`, {
      method: 'PUT',
      headers: { 'Content-Type': 'image/jpeg' },
      body: photo,
    })
    photoLink = item?.webUrl || `${PHOTO_FOLDER}/${t.reference}.jpg`
  }
  const received = new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata', dateStyle: 'medium', timeStyle: 'short' })
  const byHeader = {
    Received: received,
    Reference: t.reference,
    Status: 'Pending',
    Name: t.name,
    Company: t.company,
    Role: t.role,
    City: t.city,
    Rating: t.rating || '',
    Service: t.service,
    Testimonial: t.text,
    Photo: photoLink,
    'Show on homepage': '',
    Order: '',
    Email: t.email,
    Phone: t.phone,
    Notes: '',
  }
  const safe = (v) => (typeof v === 'string' && /^[=+\-@]/.test(v) ? `'${v}` : v)
  await graph(env, `${info.wb}/tables/${T_TABLE}/rows/add`, {
    method: 'POST',
    body: JSON.stringify({ values: [T_HEADERS.map((h) => safe(byHeader[h]))] }),
  })
}

async function sendTestimonialMail(env, t, photo, dataUrl) {
  const mailbox = env.MAILBOX || 'info@askib.in'
  const stars = t.rating ? '★'.repeat(t.rating) + '☆'.repeat(5 - t.rating) : 'not given'
  const rows = [
    ['Name', t.name],
    ['Company', t.company],
    ['Role', t.role],
    ['City', t.city],
    ['Email', t.email],
    ['Phone', t.phone],
    ['Service used', t.service],
    ['Rating', stars],
    ['Photo', photo ? 'attached' : 'none'],
  ].filter((r) => r[1])
  const html =
    `<div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;color:#111">` +
    `<h2 style="margin:0 0 4px">New testimonial, waiting for approval</h2>` +
    `<p style="margin:0 0 16px;color:#555">Reference ${esc(t.reference)}. It is <b>not</b> on the website yet. To publish it, open the ` +
    `workbook, go to the Testimonials sheet and change its Status to <b>Approved</b>.</p>` +
    `<blockquote style="margin:0 0 16px;padding:12px 16px;border-left:4px solid #c9a227;background:#f7f5ee;white-space:pre-wrap">${esc(t.text)}</blockquote>` +
    `<table style="border-collapse:collapse">${rows.map((r) => `<tr><td style="padding:4px 12px 4px 0;color:#555">${esc(r[0])}</td><td>${esc(r[1])}</td></tr>`).join('')}</table></div>`
  await graph(env, `/users/${encodeURIComponent(mailbox)}/sendMail`, {
    method: 'POST',
    body: JSON.stringify({
      message: {
        subject: `Testimonial from ${t.name} (${t.company}) · ${t.reference}`,
        body: { contentType: 'HTML', content: html },
        toRecipients: [{ emailAddress: { address: mailbox } }],
        replyTo: [{ emailAddress: { address: t.email, name: t.name } }],
        attachments: photo
          ? [
              {
                '@odata.type': '#microsoft.graph.fileAttachment',
                name: `${t.reference}.jpg`,
                contentType: 'image/jpeg',
                contentBytes: dataUrl.split(',')[1],
              },
            ]
          : [],
      },
      saveToSentItems: false,
    }),
  })
}

/** Approved rows, newest first after any Order numbers. Never includes email or phone. */
async function readApproved(env) {
  const { wb } = await workbookPath(env)
  await ensureTestimonialSheet(env, wb)
  const [head, body] = await Promise.all([
    graph(env, `${wb}/tables/${T_TABLE}/headerRowRange?$select=values`),
    graph(env, `${wb}/tables/${T_TABLE}/dataBodyRange?$select=values`),
  ])
  const cols = head.values[0].map((x) => String(x).trim())
  const at = (row, name) => {
    const i = cols.indexOf(name)
    return i < 0 ? '' : String(row[i] ?? '').trim()
  }
  const list = []
  for (const [n, row] of (body?.values || []).entries()) {
    if (!APPROVED.test(at(row, 'Status'))) continue
    const reference = at(row, 'Reference').replace(/[^A-Za-z0-9-]/g, '')
    const text = at(row, 'Testimonial')
    const name = at(row, 'Name')
    if (!name || !text) continue
    const photoCell = at(row, 'Photo')
    let photo = ''
    if (/^https:\/\//i.test(photoCell) && !/sharepoint\.com|onedrive\.live\.com|1drv\.ms/i.test(photoCell)) photo = photoCell
    else if (reference && photoCell) photo = `/testimonials/photo/${reference}.jpg`
    const order = Number(at(row, 'Order'))
    list.push({
      id: reference || `row-${n + 1}`,
      name,
      company: at(row, 'Company'),
      role: at(row, 'Role'),
      city: at(row, 'City'),
      service: at(row, 'Service'),
      rating: Math.min(5, Math.max(0, Number(at(row, 'Rating')) || 0)),
      text,
      photo,
      featured: /^(yes|y|true|1)$/i.test(at(row, 'Show on homepage')),
      order: Number.isFinite(order) && order > 0 ? order : null,
      n,
    })
  }
  list.sort((a, b) => (a.order ?? 1e9) - (b.order ?? 1e9) || b.n - a.n)
  return list.map(({ n, ...x }) => x)
}

async function approvedTestimonials(request, env) {
  const cache = caches.default
  const key = new Request(new URL('/testimonials', request.url).toString())
  const hit = await cache.match(key)
  if (hit) return withOpenCors(hit)
  let list = []
  try {
    list = await readApproved(env)
  } catch (err) {
    console.error('testimonials read failed:', err.message)
    return withOpenCors(new Response(JSON.stringify({ ok: false, items: [] }), { status: 503, headers: { 'Content-Type': 'application/json' } }))
  }
  const res = new Response(JSON.stringify({ ok: true, items: list }), {
    headers: { 'Content-Type': 'application/json', 'Cache-Control': `public, max-age=${LIST_CACHE_SECONDS}` },
  })
  await cache.put(key, res.clone())
  return withOpenCors(res)
}

async function testimonialPhoto(request, env, reference) {
  const cache = caches.default
  const key = new Request(new URL(`/testimonials/photo/${reference}.jpg`, request.url).toString())
  const hit = await cache.match(key)
  if (hit) return withOpenCors(hit)
  // Only photos of approved testimonials are served.
  const list = await readApproved(env).catch(() => [])
  if (!list.some((t) => t.id === reference && t.photo.startsWith('/testimonials/photo/'))) return withOpenCors(new Response('not found', { status: 404 }))
  const { drive, parent } = await workbookPath(env)
  const meta = await graph(env, `/drives/${drive}/items/${parent}:/${encodeURIComponent(PHOTO_FOLDER)}/${reference}.jpg?$select=id,@microsoft.graph.downloadUrl`).catch(() => null)
  const url = meta?.['@microsoft.graph.downloadUrl']
  if (!url) return withOpenCors(new Response('not found', { status: 404 }))
  const file = await fetch(url)
  if (!file.ok) return withOpenCors(new Response('not found', { status: 404 }))
  const res = new Response(file.body, {
    headers: { 'Content-Type': 'image/jpeg', 'Cache-Control': `public, max-age=${PHOTO_CACHE_SECONDS}` },
  })
  await cache.put(key, res.clone())
  return withOpenCors(res)
}

function withOpenCors(res) {
  const out = new Response(res.body, res)
  out.headers.set('Access-Control-Allow-Origin', '*')
  return out
}

function clean(v, max) {
  return String(v == null ? '' : v).trim().slice(0, max)
}
