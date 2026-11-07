/*
 * Working forms and search for the static copy of askib.in (GitHub Pages).
 *
 * The static copy has no server, so the React build shows a "preview copy"
 * notice where each form would be. This script swaps that notice for a plain
 * form, and adds search over a prebuilt index.
 *
 * Forms are sent to our own form handler (static-site/worker/worker.js, a
 * free Cloudflare Worker), which emails the enquiry from info@askib.in through
 * Microsoft 365 and adds it to the enquiries Excel workbook on OneDrive.
 * Paste the worker's address below; see static-site/worker/README.md.
 */
;(() => {
  'use strict'

  // ── Settings ───────────────────────────────────────────────────────────────
  const FORM_ENDPOINT = '' // ← paste the form handler's address here (https://….workers.dev)
  const OFFICE_EMAIL = 'info@askib.in'
  const OFFICE_PHONE = '+91 91637 90000'
  const BASE = '' // the site is served from the domain root
  /** Submissions sent faster than this after the page opened are treated as bots. */
  const MIN_FILL_MS = 3000

  // ── Form definitions (mirrors apps/web/src/enquiries/formTypes.ts) ─────────
  const contactFields = [
    { name: 'name', label: 'Your name', type: 'text', required: true, half: true, auto: 'name' },
    { name: 'company', label: 'Company', type: 'text', half: true, auto: 'organization' },
    { name: 'email', label: 'Email', type: 'email', required: true, half: true, auto: 'email' },
    { name: 'phone', label: 'Phone', type: 'tel', required: true, half: true, auto: 'tel' },
    { name: 'gstin', label: 'GSTIN (optional)', type: 'text', half: true },
    { name: 'city', label: 'City', type: 'text', half: true, auto: 'address-level2' },
  ]

  const opt = (pairs) => pairs.map(([value, label]) => ({ value, label }))

  const forms = {
    general: {
      title: 'General enquiry',
      skipContact: ['gstin'],
      header: [],
      messageLabel: 'Message',
      messageRequired: true,
    },
    calibration: {
      title: 'Calibration quote',
      itemLabel: 'Instrument',
      header: [
        { name: 'requiredBy', label: 'Required by', type: 'date', half: true, future: true },
        {
          name: 'atSite',
          label: 'Where should calibration happen?',
          type: 'select',
          half: true,
          options: opt([
            ['At your laboratory', 'At your laboratory'],
            ['At our site', 'At our site'],
            ['Not sure', 'Not sure'],
          ]),
        },
        { name: 'deliveryLocation', label: 'City / site location', type: 'text' },
      ],
      item: [
        { name: 'description', label: 'Instrument', type: 'text', required: true, placeholder: 'e.g. Pressure gauge' },
        { name: 'make', label: 'Make', type: 'text', half: true },
        { name: 'model', label: 'Model', type: 'text', half: true },
        { name: 'range', label: 'Range', type: 'text', half: true, placeholder: 'e.g. 0 to 40 bar' },
        { name: 'accuracy', label: 'Accuracy / least count', type: 'text', half: true },
        { name: 'serialNumber', label: 'Serial / ID no.', type: 'text', half: true },
        { name: 'quantity', label: 'Quantity', type: 'number', required: true, half: true },
      ],
    },
    testing: {
      title: 'Testing quote',
      itemLabel: 'Sample',
      header: [
        { name: 'projectName', label: 'Project name', type: 'text', half: true },
        { name: 'requiredBy', label: 'Report needed by', type: 'date', half: true, future: true },
        { name: 'deliveryLocation', label: 'Site / city', type: 'text' },
      ],
      item: [
        { name: 'description', label: 'Material', type: 'text', required: true, placeholder: 'e.g. Concrete cubes' },
        { name: 'testRequired', label: 'Tests required', type: 'text', placeholder: 'e.g. Compressive strength' },
        { name: 'standard', label: 'Standard', type: 'text', half: true, placeholder: 'e.g. IS 516' },
        { name: 'quantity', label: 'Number of samples', type: 'number', required: true, half: true },
      ],
    },
    product: {
      title: 'Product quote',
      itemLabel: 'Product',
      header: [
        { name: 'deliveryLocation', label: 'Delivery city / site', type: 'text' },
        { name: 'installationNeeded', label: 'Installation needed', type: 'checkbox', half: true },
        { name: 'calibrationNeeded', label: 'Calibration certificate needed', type: 'checkbox', half: true },
      ],
      item: [
        { name: 'description', label: 'Product', type: 'text', required: true },
        { name: 'specifications', label: 'Specifications / model', type: 'text' },
        { name: 'quantity', label: 'Quantity', type: 'number', required: true, half: true },
      ],
    },
    repair: {
      title: 'Repair request',
      itemLabel: 'Instrument',
      header: [
        {
          name: 'urgency',
          label: 'Urgency',
          type: 'select',
          half: true,
          options: opt([
            ['Normal', 'Normal'],
            ['Urgent', 'Urgent'],
            ['Breakdown (critical)', 'Breakdown (critical)'],
          ]),
        },
        { name: 'deliveryLocation', label: 'City / site', type: 'text', half: true },
      ],
      item: [
        { name: 'description', label: 'Instrument', type: 'text', required: true },
        {
          name: 'serviceType',
          label: 'Service needed',
          type: 'select',
          half: true,
          options: opt([
            ['Repair', 'Repair'],
            ['Annual maintenance contract (AMC)', 'Annual maintenance contract (AMC)'],
            ['Preventive maintenance visit', 'Preventive maintenance visit'],
          ]),
        },
        { name: 'quantity', label: 'Quantity', type: 'number', required: true, half: true },
        { name: 'make', label: 'Make', type: 'text', half: true },
        { name: 'model', label: 'Model', type: 'text', half: true },
        { name: 'serialNumber', label: 'Serial no.', type: 'text', half: true },
        { name: 'fault', label: 'What is the problem?', type: 'textarea', placeholder: 'Describe the fault. Needed for repairs.' },
      ],
    },
    consultation: {
      title: 'Consultation request',
      submitLabel: 'Request the appointment',
      sectionLabel: 'Your appointment',
      header: [
        {
          name: 'topic',
          label: 'What is it about?',
          type: 'select',
          required: true,
          half: true,
          options: opt(
            ['Testing', 'Calibration', 'Repair and maintenance', 'Products and equipment', 'Engineering consultancy', 'Training', 'Something else'].map((x) => [x, x]),
          ),
        },
        {
          name: 'meetingMode',
          label: 'How should we meet?',
          type: 'select',
          required: true,
          half: true,
          options: opt(['At our office', 'At your site', 'Video call', 'Phone call'].map((x) => [x, x])),
        },
        { name: 'preferredDate', label: 'Preferred date', type: 'date', required: true, future: true, half: true },
        {
          name: 'preferredSlot',
          label: 'Preferred time',
          type: 'select',
          half: true,
          options: opt(['Morning (10 am to 1 pm)', 'Afternoon (1 pm to 4 pm)', 'Late afternoon (4 pm to 6 pm)'].map((x) => [x, x])),
        },
        { name: 'siteAddress', label: 'Site address (for site visits)', type: 'text' },
      ],
    },
  }
  // Request a service asks the same questions as a repair request.
  forms.service = { ...forms.repair, title: 'Service request' }

  /** Which form each page shows. */
  const FORM_FOR_PATH = {
    '/contact': 'general',
    '/book-a-consultation': 'consultation',
    '/request-a-service': 'service',
    '/request-a-quote/general': 'general',
    '/request-a-quote/calibration': 'calibration',
    '/request-a-quote/testing': 'testing',
    '/request-a-quote/product': 'product',
    '/request-a-quote/repair': 'repair',
  }

  // ── Helpers ────────────────────────────────────────────────────────────────
  const OWN = 'data-askib-static'
  const NOTICE_TEXT = 'Forms and search are switched off'

  const h = (tag, attrs = {}, ...children) => {
    const el = document.createElement(tag)
    for (const [k, v] of Object.entries(attrs)) {
      if (v === undefined || v === null || v === false) continue
      if (k === 'class') el.className = v
      else if (k === 'text') el.textContent = v
      else if (k.startsWith('on')) el.addEventListener(k.slice(2), v)
      else el.setAttribute(k, v === true ? '' : v)
    }
    for (const c of children.flat()) if (c != null) el.append(c)
    return el
  }

  const pagePath = () => {
    let p = location.pathname.slice(BASE.length).replace(/\.html$/, '').replace(/\/index$/, '').replace(/\/+$/, '')
    return p || '/'
  }

  const today = () => {
    const d = new Date()
    d.setMinutes(d.getMinutes() - d.getTimezoneOffset())
    return d.toISOString().slice(0, 10)
  }

  const reference = () => {
    const d = new Date()
    const ymd = `${String(d.getFullYear()).slice(2)}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`
    const rnd = Math.random().toString(36).slice(2, 6).toUpperCase()
    return `WEB-${ymd}-${rnd}`
  }

  let fieldSeq = 0

  function fieldEl(def, name, prefill) {
    const id = `askib-f${++fieldSeq}`
    const wrap = h('div', { class: `askib-field${def.half ? ' askib-half' : ''}${def.type === 'checkbox' ? ' askib-check' : ''}` })
    if (def.type === 'checkbox') {
      wrap.append(h('label', { for: id }, h('input', { type: 'checkbox', id, name, value: 'Yes' }), h('span', { text: def.label })))
      return wrap
    }
    const label = h('label', { for: id, text: def.label + (def.required ? ' *' : '') })
    let input
    const common = { id, name, required: def.required, placeholder: def.placeholder, autocomplete: def.auto, 'aria-describedby': `${id}-err` }
    if (def.type === 'select') {
      input = h('select', common, h('option', { value: '', text: 'Select…' }), def.options.map((o) => h('option', { value: o.value, text: o.label })))
    } else if (def.type === 'textarea') {
      input = h('textarea', { ...common, rows: 4 })
    } else {
      input = h('input', {
        ...common,
        type: def.type,
        min: def.type === 'number' ? 1 : def.future ? today() : undefined,
        inputmode: def.type === 'number' ? 'numeric' : undefined,
      })
    }
    if (prefill) input.value = prefill
    wrap.append(label, input, h('p', { class: 'askib-err', id: `${id}-err`, 'aria-live': 'polite' }))
    return wrap
  }

  // ── Forms ──────────────────────────────────────────────────────────────────
  async function productName(id) {
    if (!id) return ''
    try {
      const r = await fetch(`${BASE}/static-site/products.json`)
      const map = await r.json()
      return map[id] || ''
    } catch {
      return ''
    }
  }

  function buildForm(key) {
    const def = forms[key]
    const startedAt = Date.now()
    const form = h('form', { class: 'askib-form', novalidate: true, [OWN]: 'form' })

    const section = (title, ...kids) => h('fieldset', { class: 'askib-section' }, h('legend', { text: title }), h('div', { class: 'askib-grid' }, ...kids))

    const contacts = contactFields.filter((f) => !(def.skipContact || []).includes(f.name))
    form.append(section('Your details', contacts.map((f) => fieldEl(f, f.name))))

    if (def.header.length) {
      form.append(section(def.sectionLabel || 'Requirements', def.header.map((f) => fieldEl(f, f.name))))
    }

    let rows
    if (def.item) {
      rows = h('div', { class: 'askib-rows' })
      const addRow = (prefill = {}) => {
        const n = rows.children.length + 1
        const row = h(
          'fieldset',
          { class: 'askib-row' },
          h('legend', { text: `${def.itemLabel} ${n}` }),
          h('div', { class: 'askib-grid' }, def.item.map((f) => fieldEl(f, `${def.itemLabel} ${n}: ${f.label}`, prefill[f.name]))),
        )
        if (n > 1) {
          row.append(
            h('button', {
              type: 'button',
              class: 'askib-link',
              text: 'Remove',
              onclick: () => {
                row.remove()
                renumber()
              },
            }),
          )
        }
        rows.append(row)
      }
      const renumber = () =>
        [...rows.children].forEach((row, i) => {
          row.querySelector('legend').textContent = `${def.itemLabel} ${i + 1}`
          row.querySelectorAll('[name]').forEach((el) => {
            el.name = el.name.replace(/^.+? \d+: /, `${def.itemLabel} ${i + 1}: `)
          })
        })
      addRow()
      if (key === 'product') {
        const id = new URLSearchParams(location.search).get('product')
        productName(id).then((name) => {
          const first = rows.querySelector('input')
          if (name && first && !first.value) first.value = name
        })
      }
      form.append(
        h(
          'div',
          { class: 'askib-section' },
          h('p', { class: 'askib-legend', text: `${def.itemLabel}s` }),
          rows,
          h('button', { type: 'button', class: 'askib-secondary', text: `+ Add another ${def.itemLabel.toLowerCase()}`, onclick: () => rows.children.length < 50 && addRow() }),
        ),
      )
    }

    form.append(
      h(
        'div',
        { class: 'askib-section' },
        fieldEl({ name: 'message', label: def.item ? 'Anything else we should know' : def.messageLabel || 'Message', type: 'textarea', required: def.messageRequired }, 'message'),
        def.item || key === 'consultation'
          ? h('p', { class: 'askib-help' }, 'Drawings, photos or a list of items? Send the enquiry, then reply to our email with the files attached, or email them to ', h('a', { href: `mailto:${OFFICE_EMAIL}`, text: OFFICE_EMAIL }), '.')
          : null,
      ),
    )

    // Honeypot: hidden from people, ticked by bots. The script drops these too.
    form.append(
      h('div', { class: 'askib-hp', 'aria-hidden': 'true' }, h('label', {}, 'Leave this empty ', h('input', { type: 'checkbox', name: 'botcheck', tabindex: '-1', autocomplete: 'off' }))),
    )

    const consentId = `askib-f${++fieldSeq}`
    form.append(
      h(
        'div',
        { class: 'askib-field askib-check' },
        h(
          'label',
          { for: consentId },
          h('input', { type: 'checkbox', id: consentId, name: 'consent', value: 'Yes', required: true, 'aria-describedby': `${consentId}-err` }),
          h('span', {}, 'I agree to be contacted about this enquiry. See the ', h('a', { href: `${BASE}/privacy-policy`, text: 'privacy policy' }), '.'),
        ),
        h('p', { class: 'askib-err', id: `${consentId}-err`, 'aria-live': 'polite' }),
      ),
    )

    const status = h('div', { class: 'askib-status', role: 'status', 'aria-live': 'polite' })
    const button = h('button', { type: 'submit', class: 'askib-primary', text: def.submitLabel || 'Send enquiry' })
    form.append(status, button)

    form.addEventListener('submit', async (e) => {
      e.preventDefault()
      status.className = 'askib-status'
      status.textContent = ''
      if (!validate(form)) {
        status.className = 'askib-status askib-bad'
        status.textContent = 'Please check the highlighted fields.'
        form.querySelector('[aria-invalid="true"]')?.focus()
        return
      }
      if (!FORM_ENDPOINT) {
        status.className = 'askib-status askib-bad'
        status.replaceChildren('This form is not connected yet. Please email ', h('a', { href: `mailto:${OFFICE_EMAIL}`, text: OFFICE_EMAIL }), ` or call ${OFFICE_PHONE}.`)
        return
      }
      const data = new FormData(form)
      if (data.get('botcheck') || Date.now() - startedAt < MIN_FILL_MS) {
        // Looks automated: pretend it worked and send nothing.
        done(form, reference())
        return
      }
      const ref = reference()
      const fields = []
      const payload = {
        form: def.title,
        reference: ref,
        subject: `${def.title} from ${data.get('name')}${data.get('company') ? ` (${data.get('company')})` : ''} · ${ref}`,
        replyTo: data.get('email'),
        name: data.get('name'),
        page: location.href,
        fillMs: Date.now() - startedAt,
        botcheck: '',
        fields,
      }
      for (const [k, v] of data) {
        if (k === 'botcheck' || k === 'consent') continue
        const value = String(v).trim()
        if (!value) continue
        const label = k.includes(':') ? k : [...contactFields, ...def.header, { name: 'message', label: 'Message' }].find((f) => f.name === k)?.label || k
        fields.push([label, value])
      }
      button.disabled = true
      button.textContent = 'Sending…'
      try {
        // text/plain keeps this a "simple" request, so the browser needs no preflight.
        const res = await fetch(FORM_ENDPOINT, {
          method: 'POST',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: JSON.stringify(payload),
        })
        const json = await res.json().catch(() => ({}))
        if (!res.ok || json.ok !== true) throw new Error(json.error || `HTTP ${res.status}`)
        done(form, ref)
      } catch (err) {
        console.error('[askib forms]', err)
        status.className = 'askib-status askib-bad'
        status.replaceChildren('Sorry, your enquiry could not be sent. Please try again, or email ', h('a', { href: `mailto:${OFFICE_EMAIL}`, text: OFFICE_EMAIL }), ` or call ${OFFICE_PHONE}.`)
        button.disabled = false
        button.textContent = def.submitLabel || 'Send enquiry'
      }
    })
    return form
  }

  function validate(form) {
    let ok = true
    for (const el of form.querySelectorAll('input, select, textarea')) {
      if (el.name === 'botcheck') continue
      const err = el.getAttribute('aria-describedby') && document.getElementById(el.getAttribute('aria-describedby'))
      let msg = ''
      const v = el.type === 'checkbox' ? el.checked : el.value.trim()
      const label = el.closest('.askib-field')?.querySelector('label')?.textContent.replace(/\s*\*$/, '').toLowerCase() || 'this field'
      if (el.required && !v) msg = el.name === 'consent' ? 'Please agree so we can reply' : el.tagName === 'SELECT' ? `Choose ${label}` : `Enter ${label}`
      else if (v && el.type === 'email' && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v)) msg = 'Enter a valid email address'
      else if (v && el.type === 'tel' && (v.replace(/\D/g, '').length < 8 || /[^\d\s()+.-]/.test(v))) msg = 'Enter a valid phone number'
      else if (v && el.type === 'number' && !(Number(v) >= 1)) msg = `Check ${label}`
      else if (v && el.type === 'date' && el.min && v < el.min) msg = 'Choose today or a later date'
      if (err) err.textContent = msg
      if (msg) {
        el.setAttribute('aria-invalid', 'true')
        ok = false
      } else el.removeAttribute('aria-invalid')
    }
    return ok
  }

  function done(form, ref) {
    const box = h(
      'div',
      { class: 'askib-done', role: 'status', tabindex: '-1', [OWN]: 'form' },
      h('h3', { text: 'Thank you. We have your enquiry.' }),
      h('p', {}, 'Your reference is ', h('strong', { text: ref }), '. Please quote it if you contact us about this request.'),
      h('p', { text: 'Our team will reply by phone or email, usually within one working day.' }),
      h('p', {}, h('a', { href: `${BASE}/`, text: 'Back to the homepage' })),
    )
    box._anchor = form._anchor
    form.replaceWith(box)
    box.focus()
  }

  // ── Search ─────────────────────────────────────────────────────────────────
  let index
  const loadIndex = () =>
    (index ??= fetch(`${BASE}/static-site/search-index.json`)
      .then((r) => r.json())
      .catch(() => []))

  const norm = (s) => s.toLowerCase().normalize('NFKD').replace(/[̀-ͯ]/g, '')

  async function search(q, limit = 30) {
    const words = norm(q).split(/[^a-z0-9.]+/).filter((w) => w.length > 1)
    if (!words.length) return []
    const pages = await loadIndex()
    const out = []
    for (const p of pages) {
      const t = norm(p.t)
      const d = norm(p.d || '')
      const k = norm(p.k || '')
      let score = 0
      let all = true
      for (const w of words) {
        const s = (t.includes(w) ? 6 : 0) + (k.includes(w) ? 3 : 0) + (d.includes(w) ? 2 : 0)
        if (!s) all = false
        score += s
      }
      if (all && score) out.push({ p, score: score - p.u.split('/').length * 0.1 })
    }
    return out.sort((a, b) => b.score - a.score).slice(0, limit).map((x) => x.p)
  }

  const resultLink = (p, cls) => h('a', { href: BASE + p.u, class: cls }, h('span', { class: 'askib-r-title', text: p.t }), p.d ? h('span', { class: 'askib-r-sum', text: p.d }) : null)

  /** Fills the header's quick-results list (React renders it empty on the static copy). */
  function hookHeaderSearch() {
    const lists = document.querySelectorAll('header [role="listbox"], [role="dialog"] [role="listbox"]')
    for (const list of lists) {
      const input = list.closest('[role="dialog"]')?.querySelector('input') || list.parentElement?.parentElement?.querySelector('input')
      if (!input) continue
      let own = list.querySelector(`[${OWN}="quick"]`)
      if (!own) {
        own = h('div', { [OWN]: 'quick' })
        list.prepend(own)
      }
      const q = input.value.trim()
      if (own.dataset.q === q) continue
      own.dataset.q = q
      search(q, 8).then((res) => {
        if (own.dataset.q !== q) return
        own.replaceChildren(...res.map((p) => resultLink(p, 'askib-quick')))
        // React's own "no matches" line is the first plain child; show it only when we found nothing too.
        const none = [...list.children].find((c) => c !== own && c.tagName !== 'A')
        if (none) none.style.display = res.length ? 'none' : ''
      })
    }
  }

  async function renderSearchPage() {
    if (pagePath() !== '/search') return
    const main = document.querySelector('main')
    if (!main || main.querySelector(`[${OWN}="search"]`)) return
    if (!hydrated(main)) return later()
    for (const c of main.children) c.style.display = 'none'
    const q = new URLSearchParams(location.search).get('q') || ''
    document.title = q ? `Search: ${q} | ASKIB` : 'Search | ASKIB'
    const input = h('input', { type: 'search', name: 'q', value: q, placeholder: 'Search services, products, standards…', 'aria-label': 'Search the site', autofocus: true })
    const list = h('div', { class: 'askib-results' })
    const page = h(
      'section',
      { class: 'askib-search-page', [OWN]: 'search' },
      h('h1', { text: 'Search' }),
      h('form', { role: 'search', action: `${BASE}/search`, method: 'get', class: 'askib-search-form' }, input, h('button', { type: 'submit', class: 'askib-primary', text: 'Search' })),
      list,
    )
    main.append(page)
    if (!q) return
    const res = await search(q)
    list.replaceChildren(
      h('p', { class: 'askib-help', text: res.length ? `${res.length} result${res.length === 1 ? '' : 's'} for “${q}”` : `Nothing matched “${q}”. Try fewer or different words.` }),
      ...res.map((p) => resultLink(p, 'askib-result')),
    )
  }

  // ── Wiring ─────────────────────────────────────────────────────────────────
  /** True once React has hydrated this element; changing the page before that breaks hydration. */
  const hydrated = (el) => Object.keys(el).some((k) => k.startsWith('__reactFiber'))
  let retry
  const later = () => {
    clearTimeout(retry)
    retry = setTimeout(schedule, 150)
  }

  /** Hides each "preview copy" notice and puts the page's form after it. */
  function mountForms() {
    for (const msg of document.querySelectorAll('[role="alert"]')) {
      if (!msg.textContent.includes(NOTICE_TEXT) || msg.dataset.askibDone) continue
      if (!hydrated(msg)) {
        later()
        continue
      }
      const key = FORM_FOR_PATH[pagePath()]
      msg.dataset.askibDone = '1'
      if (!key) {
        // A page whose form cannot work without the server (e.g. order tracking).
        const text = [...msg.querySelectorAll('div')].pop()
        if (text) text.textContent = `This is not available online yet. Please email ${OFFICE_EMAIL} or call ${OFFICE_PHONE}.`
        continue
      }
      msg.style.display = 'none'
      const form = buildForm(key)
      form._anchor = msg
      msg.after(form)
    }
  }

  /** Removes our nodes once the React content they belong to has gone (client-side page change). */
  function cleanUp() {
    for (const el of document.querySelectorAll(`[${OWN}="form"]`)) {
      if (el._anchor && !el._anchor.isConnected) el.remove()
    }
    if (pagePath() !== '/search') {
      for (const el of document.querySelectorAll(`[${OWN}="search"]`)) {
        for (const c of el.parentElement.children) c.style.display = ''
        el.remove()
      }
    }
  }

  let queued = false
  const tick = () => {
    queued = false
    cleanUp()
    mountForms()
    renderSearchPage()
    hookHeaderSearch()
  }
  const schedule = () => {
    if (!queued) {
      queued = true
      requestAnimationFrame(tick)
    }
  }

  const start = () => {
    // Give React a moment to hydrate before touching the page.
    setTimeout(() => {
      tick()
      new MutationObserver(schedule).observe(document.body, { childList: true, subtree: true })
      document.addEventListener('input', schedule, true)
    }, 0)
  }
  if (document.readyState === 'complete') start()
  else window.addEventListener('load', start)
})()
