/*
 * Search, filters and sorting for the catalogue pages of the static askib.in:
 *
 *   /spare-parts  spare parts from static-site/spare-parts.csv (edit it in Excel)
 *   /products     products, from catalogue-products.json (made by build.mjs)
 *   /services     services, from catalogue-services.json (made by build.mjs)
 *
 * Everything runs in the browser. Filters are kept in the address, so a
 * filtered list can be bookmarked or shared. Loaded by site.js on those pages.
 */
;(() => {
  'use strict'

  const S = window.askibStatic || {}
  const BASE = S.BASE || ''
  const OFFICE_EMAIL = S.OFFICE_EMAIL || 'info@askib.in'
  const PAGE_SIZE = 24
  const OWN = 'data-askib-static'
  const YEAR_NOW = new Date().getFullYear()

  // ── Small helpers ──────────────────────────────────────────────────────────
  const h = (tag, attrs = {}, ...children) => {
    const el = document.createElement(tag)
    for (const [k, v] of Object.entries(attrs)) {
      if (v === undefined || v === null || v === false) continue
      if (k === 'class') el.className = v
      else if (k === 'text') el.textContent = v
      else if (k === 'html') el.innerHTML = v
      else if (k.startsWith('on')) el.addEventListener(k.slice(2), v)
      else el.setAttribute(k, v === true ? '' : v)
    }
    for (const c of children.flat(Infinity)) if (c != null && c !== false) el.append(c)
    return el
  }
  const pagePath = () => (location.pathname.slice(BASE.length).replace(/\.html$/, '').replace(/\/index$/, '').replace(/\/+$/, '') || '/')
  const hydrated = (el) => Object.keys(el).some((k) => k.startsWith('__reactFiber'))
  const norm = (s) => String(s ?? '').toLowerCase().normalize('NFKD').replace(/[̀-ͯ]/g, '')
  /** Part numbers compare without spaces, dots, dashes or slashes: "CTM-20/01" = "ctm 2001". */
  const code = (s) => norm(s).replace(/[^a-z0-9]/g, '')
  const inr = (n) => '₹' + Number(n).toLocaleString('en-IN', { maximumFractionDigits: 0 })
  const plural = (n, one, many) => `${n.toLocaleString('en-IN')} ${n === 1 ? one : many}`
  const uniq = (xs) => [...new Set(xs.filter((x) => x !== '' && x != null))]
  const ICON = {
    filter: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M3 5h18l-7 8.5V20l-4-2v-4.5z" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"/></svg>',
    search: '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="11" cy="11" r="6.5" fill="none" stroke="currentColor" stroke-width="1.8"/><path d="m16 16 4.5 4.5" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/></svg>',
    close: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M6 6l12 12M18 6 6 18" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/></svg>',
    grid: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 4h7v7H4zM13 4h7v7h-7zM4 13h7v7H4zM13 13h7v7h-7z" fill="none" stroke="currentColor" stroke-width="1.6"/></svg>',
    list: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 6h16M4 12h16M4 18h16" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/></svg>',
    part: '<svg viewBox="0 0 48 48" aria-hidden="true"><path d="M24 5 40.5 14.5v19L24 43 7.5 33.5v-19z" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linejoin="round"/><circle cx="24" cy="24" r="7.5" fill="none" stroke="currentColor" stroke-width="2.4"/></svg>',
  }
  /** Like el.replaceChildren, but takes arrays and skips empty values. */
  const put = (el, ...kids) => el.replaceChildren(...kids.flat(Infinity).filter((k) => k != null && k !== false && k !== ''))
  const icon = (name, cls = 'ak-ico') => h('span', { class: cls, html: ICON[name] })

  // ── CSV (as saved by Excel: commas, quotes, line breaks inside quotes) ──────
  function parseCsv(text) {
    const rows = []
    let row = []
    let cell = ''
    let quoted = false
    text = text.replace(/^﻿/, '')
    for (let i = 0; i < text.length; i++) {
      const c = text[i]
      if (quoted) {
        if (c === '"' && text[i + 1] === '"') {
          cell += '"'
          i++
        } else if (c === '"') quoted = false
        else cell += c
      } else if (c === '"') quoted = true
      else if (c === ',') {
        row.push(cell)
        cell = ''
      } else if (c === '\n' || c === '\r') {
        if (c === '\r' && text[i + 1] === '\n') i++
        row.push(cell)
        rows.push(row)
        row = []
        cell = ''
      } else cell += c
    }
    if (cell || row.length) {
      row.push(cell)
      rows.push(row)
    }
    const head = (rows.shift() || []).map((x) => x.trim())
    return rows
      .filter((r) => r.some((x) => x.trim()))
      .map((r) => Object.fromEntries(head.map((k, i) => [k, (r[i] ?? '').trim()])))
  }

  const list = (s) => uniq(String(s || '').split(/\s*;\s*/).map((x) => x.trim()))
  const num = (s) => {
    const n = Number(String(s ?? '').replace(/[₹,\s]/g, ''))
    return s === '' || s == null || !Number.isFinite(n) ? null : n
  }
  function years(s) {
    const m = String(s || '').match(/(\d{4})\s*(?:[-–to]+\s*(\d{4}|now|present|onwards|on)?)?/i)
    if (!m) return [null, null]
    const from = Number(m[1])
    const to = m[2] ? (/\d/.test(m[2]) ? Number(m[2]) : YEAR_NOW) : /[-–]|onwards/.test(s) ? YEAR_NOW : from
    return [from, to]
  }

  const months = (n) => (n % 12 === 0 && n >= 12 ? `${n / 12} year${n > 12 ? 's' : ''}` : `${n} month${n === 1 ? '' : 's'}`)
  const toMonths = (s) => parseFloat(s) * (/year/.test(s) ? 12 : 1)
  const UNIVERSAL = /^(universal|all|any|all makes)$/i
  const AVAIL_ORDER = ['In stock', 'Low stock', 'Made to order', 'On order', 'Enquire', 'Out of stock']
  function availability(s, qty) {
    const t = norm(s)
    if (/out/.test(t)) return 'Out of stock'
    if (/low/.test(t)) return 'Low stock'
    if (/made|build/.test(t)) return 'Made to order'
    if (/order|import|lead/.test(t)) return 'On order'
    if (/in stock|stock|ready|available/.test(t)) return qty === 0 ? 'Out of stock' : 'In stock'
    if (qty != null) return qty > 0 ? 'In stock' : 'Out of stock'
    return 'Enquire'
  }
  const CONDITION_ORDER = ['New', 'Refurbished', 'Used', 'Reconditioned']
  const condition = (s) => {
    const t = norm(s)
    if (!t) return ''
    if (/refurb/.test(t)) return 'Refurbished'
    if (/recon/.test(t)) return 'Reconditioned'
    if (/used|second/.test(t)) return 'Used'
    if (/new/.test(t)) return 'New'
    return s.trim()
  }

  // ── Spare parts: one row of spare-parts.csv → one item ──────────────────────
  function partFromRow(r, i) {
    const g = (k) => r[k] ?? ''
    const makes = list(g('Fits make'))
    const models = list(g('Fits model'))
    const yrs = list(g('Fits years'))
    const n = Math.max(makes.length, models.length, yrs.length)
    const compat = []
    for (let k = 0; k < n; k++) {
      const [from, to] = years(yrs[k] ?? yrs[yrs.length - 1] ?? '')
      compat.push({ make: makes[k] ?? makes[makes.length - 1] ?? '', model: models[k] ?? '', from, to })
    }
    const price = num(g('Price (INR)'))
    const mrp = num(g('List price (INR)'))
    const qty = num(g('Stock qty'))
    const avail = availability(g('Availability'), qty)
    const specs = list(g('Specifications')).map((x) => {
      const at = x.indexOf(':')
      return at > 0 ? [x.slice(0, at).trim(), x.slice(at + 1).trim()] : ['', x]
    })
    const partNo = g('Part number')
    const oem = list(g('OEM / alternate numbers'))
    const made = num(g('Year of manufacture'))
    const added = Date.parse(g('Date added')) || 0
    const item = {
      id: g('ID') || code(partNo) || String(i + 1),
      title: g('Part name') || partNo || 'Spare part',
      summary: g('Description'),
      eyebrow: [g('Brand'), g('Category')].filter(Boolean).join(' · '),
      image: g('Image URL') ? { url: g('Image URL'), alt: g('Part name') } : null,
      price,
      mrp: mrp && price && mrp > price ? mrp : null,
      partNo,
      oem,
      compat: compat.filter((c) => c.make || c.model),
      qty,
      lead: num(g('Lead time (days)')),
      made,
      date: added,
      pop: num(g('Popularity')) || 0,
      featured: /^(y|yes|true|1)$/i.test(g('Featured')),
      order: i,
      facets: {
        category: list(g('Category')),
        subcategory: list(g('Sub-category')),
        equipment: list(g('Equipment type')),
        brand: list(g('Brand')),
        condition: uniq([condition(g('Condition'))]),
        genuine: list(g('Genuine or compatible')),
        availability: [avail],
        location: list(g('Location')),
        origin: list(g('Country of origin')),
        warranty: num(g('Warranty (months)')) ? [months(num(g('Warranty (months)')))] : [],
        make: uniq(compat.map((c) => c.make)),
        model: uniq(compat.map((c) => c.model)),
        shipping: list(g('Delivery')),
      },
      nums: { made, warranty: num(g('Warranty (months)')), moq: num(g('Minimum order')), weight: num(g('Weight (kg)')) },
      details: [
        ['Part number', partNo],
        ['OEM / alternate numbers', oem.join(', ')],
        ['Category', [g('Category'), g('Sub-category')].filter(Boolean).join(' › ')],
        ['For equipment', g('Equipment type')],
        ['Brand', g('Brand')],
        ['Condition', condition(g('Condition'))],
        ['Genuine or compatible', g('Genuine or compatible')],
        ['Year of manufacture', made ? String(made) : ''],
        ['Material', g('Material')],
        ['Dimensions', g('Dimensions')],
        ['Weight', num(g('Weight (kg)')) ? `${num(g('Weight (kg)'))} kg` : ''],
        ['Availability', avail + (qty ? ` (${qty} available)` : '') + (num(g('Lead time (days)')) && avail !== 'In stock' ? ` · usually ${num(g('Lead time (days)'))} days` : '')],
        ['Ships from', g('Location')],
        ['Delivery', g('Delivery')],
        ['Country of origin', g('Country of origin')],
        ['Warranty', num(g('Warranty (months)')) ? months(num(g('Warranty (months)'))) : ''],
        ['Minimum order', num(g('Minimum order')) > 1 ? String(num(g('Minimum order'))) : ''],
        ...specs.map(([k, v]) => [k || 'Specification', v]),
      ].filter((d) => d[1]),
      keywords: g('Keywords'),
    }
    item.codes = uniq([partNo, ...oem].map(code))
    item.hay = norm([item.title, item.summary, item.eyebrow, partNo, oem.join(' '), g('Sub-category'), g('Equipment type'), g('Keywords'), g('Material'), compat.map((c) => `${c.make} ${c.model}`).join(' '), specs.flat().join(' ')].join(' '))
    item.cta = { href: `${BASE}/request-a-quote/product?item=${encodeURIComponent(`${item.title}${partNo ? ` (part no. ${partNo})` : ''}`)}`, label: 'Request a quote' }
    return item
  }

  // ── Products and services: the lists build.mjs saved from the export ──────
  function fromExport(data) {
    return data.items.map((x, i) => {
      const money = (t) => num(String(t || '').match(/[\d,]+(?:\.\d+)?/)?.[0])
      const labels = {}
      for (const g of data.filters) {
        labels[g.id] = (x.facets[g.id] || []).map((v) => g.options.find((o) => o.value === v)?.label ?? v)
      }
      const category = labels.range?.[0] || labels.kind?.[0] || labels.area?.[0] || ''
      return {
        id: String(x.id),
        title: x.title,
        href: x.href ? BASE + x.href : null,
        summary: x.summary,
        eyebrow: x.eyebrow || category,
        image: x.image,
        price: typeof x.price === 'number' ? x.price : null,
        priceText: x.priceText,
        mrp: money(x.was),
        off: x.off,
        partNo: x.model || '',
        oem: [],
        compat: [],
        date: Number(x.id) || 0,
        pop: x.featured ? 1 : 0,
        featured: Boolean(x.featured),
        order: i,
        facets: labels,
        nums: {},
        details: [],
        codes: uniq([code(x.model)]),
        hay: norm([x.title, x.summary, x.eyebrow, x.search, Object.values(labels).flat().join(' ')].join(' ')),
        cta: data.kind === 'products' ? { href: `${BASE}/request-a-quote/product?product=${x.id}`, label: 'Request a quote' } : { href: `${BASE}/request-a-service`, label: 'Request this service' },
      }
    })
  }

  // ── Page set-ups ───────────────────────────────────────────────────────────
  const SPARE_FACETS = [
    { id: 'category', label: 'Category', order: 'alpha' },
    { id: 'subcategory', label: 'Sub-category', order: 'alpha', dependent: true },
    { id: 'equipment', label: 'For equipment', order: 'alpha' },
    { id: 'brand', label: 'Brand', order: 'alpha' },
    { id: 'condition', label: 'Condition', order: CONDITION_ORDER },
    { id: 'genuine', label: 'Genuine or compatible', order: 'alpha' },
    { id: 'availability', label: 'Availability', order: AVAIL_ORDER },
    { id: 'made', label: 'Year of manufacture', type: 'range', unit: 'year' },
    { id: 'warranty', label: 'Warranty', order: (a, b) => toMonths(a) - toMonths(b) },
    { id: 'location', label: 'Ships from', order: 'alpha' },
    { id: 'origin', label: 'Country of origin', order: 'alpha' },
    { id: 'shipping', label: 'Delivery', order: 'alpha' },
  ]
  const PAGES = {
    '/spare-parts': {
      kind: 'parts',
      noun: ['spare part', 'spare parts'],
      title: 'Spare parts',
      eyebrow: 'Spares & replacement parts',
      intro: 'Find parts for testing machines, measuring instruments and lab equipment. Search by part number or OEM number, or pick your machine’s make, model and year.',
      placeholder: 'Part name, part number or OEM number…',
      finder: true,
      facets: SPARE_FACETS,
      price: true,
      sorts: ['relevance', 'featured', 'price-asc', 'price-desc', 'newest', 'popular', 'stock', 'name-asc', 'name-desc', 'year-desc'],
    },
    '/products': {
      kind: 'products',
      source: 'catalogue-products.json',
      noun: ['product', 'products'],
      placeholder: 'Search products, models or standards…',
      facets: [
        { id: 'range', label: 'Category', order: 'alpha' },
        { id: 'brand', label: 'Brand', order: 'alpha' },
        { id: 'availability', label: 'Availability' },
        { id: 'industry', label: 'Industry', order: 'alpha' },
        { id: 'standard', label: 'Standard', order: 'alpha' },
        { id: 'tag', label: 'Topic', order: 'alpha' },
      ],
      price: true,
      sorts: ['relevance', 'featured', 'price-asc', 'price-desc', 'name-asc', 'name-desc', 'newest'],
    },
    '/services': {
      kind: 'services',
      source: 'catalogue-services.json',
      noun: ['service', 'services'],
      placeholder: 'Search tests, calibration, standards…',
      facets: [
        { id: 'kind', label: 'Service type', order: 'alpha' },
        { id: 'area', label: 'Service area', order: 'alpha' },
        { id: 'where', label: 'Where' },
        { id: 'industry', label: 'Industry', order: 'alpha' },
        { id: 'standard', label: 'Standard', order: 'alpha' },
        { id: 'tag', label: 'Topic', order: 'alpha' },
      ],
      price: true,
      sorts: ['relevance', 'featured', 'price-asc', 'price-desc', 'name-asc', 'name-desc', 'newest'],
    },
  }
  const SORTS = {
    relevance: 'Best match',
    featured: 'Featured',
    'price-asc': 'Price: low to high',
    'price-desc': 'Price: high to low',
    newest: 'Newest first',
    popular: 'Most popular',
    stock: 'In stock first',
    'name-asc': 'Name: A to Z',
    'name-desc': 'Name: Z to A',
    'year-desc': 'Year of manufacture: newest',
  }
  const PRICE_PRESETS = [
    [0, 1000, 'Under ₹1,000'],
    [1000, 5000, '₹1,000 – ₹5,000'],
    [5000, 25000, '₹5,000 – ₹25,000'],
    [25000, 100000, '₹25,000 – ₹1 lakh'],
    [100000, null, 'Over ₹1 lakh'],
  ]

  // ── Search scoring ─────────────────────────────────────────────────────────
  /** 0 when the item does not match every word; higher is a better match. */
  function score(item, words, raw) {
    if (!words.length) return 1
    let total = 0
    const title = norm(item.title)
    const whole = code(raw)
    if (whole.length >= 3 && item.codes.some((c) => c === whole)) total += 100
    for (const w of words) {
      const cw = code(w)
      let s = 0
      if (cw.length >= 2 && item.codes.some((c) => c === cw)) s += 60
      else if (cw.length >= 3 && item.codes.some((c) => c.startsWith(cw))) s += 30
      else if (cw.length >= 3 && item.codes.some((c) => c.includes(cw))) s += 15
      if (new RegExp(`(^|[^a-z0-9])${w.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}`).test(title)) s += 12
      else if (title.includes(w)) s += 8
      if (!s && item.hay.includes(w)) s += 3
      if (!s) {
        // Plural or singular ("gauges" finds "gauge").
        const stem = w.replace(/(es|s)$/, '')
        if (stem.length > 2 && item.hay.includes(stem)) s += 2
      }
      if (!s) return 0
      total += s
    }
    return total
  }

  // ── The browser ────────────────────────────────────────────────────────────
  function mountBrowser(host, cfg, items, sample) {
    // Facets that actually split the list, with their options.
    const valueSets = {}
    for (const it of items) for (const [k, vs] of Object.entries(it.facets)) for (const v of vs || []) (valueSets[k] ??= new Set()).add(v)
    const facets = []
    if (cfg.finder) {
      facets.push({ id: 'make', label: 'Machine make', order: 'alpha', fit: true }, { id: 'model', label: 'Machine model', order: 'alpha', fit: true, dependent: true }, { id: 'year', label: 'Machine year', type: 'range', unit: 'year', fit: true })
    }
    for (const f of cfg.facets) facets.push(f)
    const yearBounds = (() => {
      const ys = items.flatMap((i) => i.compat.flatMap((c) => [c.from, c.to])).filter(Boolean)
      return ys.length ? [Math.min(...ys), Math.max(...ys)] : null
    })()
    const madeBounds = (() => {
      const ys = items.map((i) => i.nums.made).filter(Boolean)
      return ys.length ? [Math.min(...ys), Math.max(...ys)] : null
    })()
    const usable = facets.filter((f) => {
      if (f.id === 'year') return Boolean(yearBounds)
      if (f.id === 'made') return Boolean(madeBounds)
      // A list filter is only offered when it splits the list (machine fit: whenever there is data).
      return (valueSets[f.id]?.size || 0) > (f.fit ? 0 : 1)
    })
    const listFacets = usable.filter((f) => f.type !== 'range')
    const prices = items.map((i) => i.price).filter((p) => typeof p === 'number')
    const hasPrice = cfg.price && prices.length > 0
    const sorts = cfg.sorts.filter((s) => (s.startsWith('price') ? hasPrice : s === 'year-desc' ? Boolean(madeBounds) : s === 'popular' ? items.some((i) => i.pop) : s === 'stock' ? valueSets.availability?.size > 1 : true))

    // State, read from and written to the address.
    const state = { q: '', sort: '', view: 'grid', sel: {}, price: [null, null], year: [null, null], made: [null, null], limit: PAGE_SIZE, part: '' }
    const rangeOf = (s) => {
      const m = String(s || '').match(/^(\d*)-(\d*)$/)
      return m ? [m[1] ? Number(m[1]) : null, m[2] ? Number(m[2]) : null] : [null, null]
    }
    function readUrl() {
      const p = new URLSearchParams(location.search)
      state.q = p.get('q') || ''
      state.sort = sorts.includes(p.get('sort')) ? p.get('sort') : ''
      state.view = p.get('view') === 'list' ? 'list' : 'grid'
      state.sel = {}
      for (const f of listFacets) {
        const v = p.get(f.id)
        if (v) state.sel[f.id] = v.split('|').filter((x) => valueSets[f.id]?.has(x))
      }
      state.price = rangeOf(p.get('price'))
      state.year = rangeOf(p.get('year'))
      state.made = rangeOf(p.get('made'))
      state.part = p.get('part') || ''
    }
    function writeUrl(push) {
      const p = new URLSearchParams()
      if (sample) p.set('demo', '1')
      if (state.q) p.set('q', state.q)
      for (const f of listFacets) if (state.sel[f.id]?.length) p.set(f.id, state.sel[f.id].join('|'))
      for (const k of ['price', 'year', 'made']) if (state[k][0] != null || state[k][1] != null) p.set(k, `${state[k][0] ?? ''}-${state[k][1] ?? ''}`)
      if (state.sort) p.set('sort', state.sort)
      if (state.view === 'list') p.set('view', 'list')
      if (state.part) p.set('part', state.part)
      const qs = p.toString()
      const url = location.pathname + (qs ? `?${qs.replace(/%7C/g, '|').replace(/%2C/g, ',')}` : '') + location.hash
      if (url !== location.pathname + location.search + location.hash) history[push ? 'pushState' : 'replaceState'](history.state, '', url)
    }

    const words = () => norm(state.q).split(/[^a-z0-9.]+/).filter((w) => w.length > 0)
    const inRange = (v, [lo, hi]) => v != null && (lo == null || v >= lo) && (hi == null || v <= hi)
    const overlaps = (c, [lo, hi]) => c.from == null || ((hi == null || c.from <= hi) && (lo == null || (c.to ?? c.from) >= lo))

    /** Does the item pass every filter except `skip`? */
    function passes(it, skip) {
      if (hasPrice && skip !== 'price' && (state.price[0] != null || state.price[1] != null) && !inRange(it.price, state.price)) return false
      if (skip !== 'made' && (state.made[0] != null || state.made[1] != null) && !inRange(it.nums.made, state.made)) return false
      for (const f of listFacets) {
        if (f.fit || f.id === skip) continue
        const pick = state.sel[f.id]
        if (pick?.length && !pick.some((v) => it.facets[f.id]?.includes(v))) return false
      }
      // Machine fit: one compatibility line must match make, model and year together.
      const mk = skip === 'make' ? [] : state.sel.make || []
      const md = skip === 'model' ? [] : state.sel.model || []
      const yr = skip === 'year' ? [null, null] : state.year
      const yearOn = yr[0] != null || yr[1] != null
      if (mk.length || md.length || yearOn) {
        // A line with no model fits every model of its make.
        if (!it.compat.some((c) => (!mk.length || mk.includes(c.make)) && (!md.length || !c.model || md.includes(c.model)) && (!yearOn || overlaps(c, yr)))) return false
      }
      return true
    }

    let scores = new Map()
    function results() {
      const ws = words()
      scores = new Map()
      const out = []
      for (const it of items) {
        const s = score(it, ws, state.q)
        if (!s) continue
        scores.set(it, s)
        if (passes(it)) out.push(it)
      }
      const sort = state.sort || (ws.length ? 'relevance' : 'featured')
      const byName = (a, b) => a.title.localeCompare(b.title, 'en', { numeric: true, sensitivity: 'base' })
      const nullsLast = (f, dir) => (a, b) => {
        const x = f(a)
        const y = f(b)
        if (x == null && y == null) return byName(a, b)
        if (x == null) return 1
        if (y == null) return -1
        return dir * (x - y) || byName(a, b)
      }
      const stockRank = (i) => {
        const k = AVAIL_ORDER.indexOf(i.facets.availability?.[0])
        return k < 0 ? 9 : k
      }
      const cmp = {
        relevance: (a, b) => scores.get(b) - scores.get(a) || Number(b.featured) - Number(a.featured) || a.order - b.order,
        featured: (a, b) => Number(b.featured) - Number(a.featured) || a.order - b.order,
        'price-asc': nullsLast((i) => i.price, 1),
        'price-desc': nullsLast((i) => i.price, -1),
        newest: (a, b) => b.date - a.date || a.order - b.order,
        popular: (a, b) => b.pop - a.pop || a.order - b.order,
        stock: (a, b) => stockRank(a) - stockRank(b) || a.order - b.order,
        'name-asc': byName,
        'name-desc': (a, b) => byName(b, a),
        'year-desc': nullsLast((i) => i.nums.made, -1),
      }[sort]
      return out.sort(cmp)
    }
    const count = (fid, v) => {
      let n = 0
      for (const it of scores.keys()) if (it.facets[fid]?.includes(v) && passes(it, fid)) n++
      return n
    }

    // ── Layout ──
    const resultsLive = h('p', { class: 'ak-count', 'aria-live': 'polite' })
    const search = h('input', { type: 'search', class: 'ak-search-input', placeholder: cfg.placeholder, 'aria-label': `Search ${cfg.noun[1]}`, autocomplete: 'off', list: `${cfg.kind}-suggest`, enterkeyhint: 'search' })
    const suggest = h('datalist', { id: `${cfg.kind}-suggest` }, uniq(items.flatMap((i) => [i.title, i.partNo, ...i.oem])).slice(0, 2000).map((v) => h('option', { value: v })))
    const sortSel = h('select', { class: 'ak-select', 'aria-label': 'Sort by' }, h('option', { value: '', text: 'Sort: Recommended' }), sorts.map((s) => h('option', { value: s, text: `Sort: ${SORTS[s]}` })))
    const filtersBtn = h('button', { type: 'button', class: 'ak-btn ak-filters-btn', 'aria-haspopup': 'dialog', 'aria-expanded': 'false' }, icon('filter'), h('span', { text: 'Filters' }), h('span', { class: 'ak-badge', hidden: true }))
    const viewGrid = h('button', { type: 'button', class: 'ak-icon-btn', 'aria-label': 'Show as grid', title: 'Grid' }, icon('grid'))
    const viewList = h('button', { type: 'button', class: 'ak-icon-btn', 'aria-label': 'Show as list', title: 'List' }, icon('list'))
    const chips = h('div', { class: 'ak-chips' })
    const grid = h('ul', { class: 'ak-results' })
    const more = h('div', { class: 'ak-more' })
    const empty = h('div', { class: 'ak-empty', hidden: true })
    const sideBody = h('div', { class: 'ak-side-body' })
    const showBtn = h('button', { type: 'button', class: 'ak-primary ak-show' })
    const closeBtn = h('button', { type: 'button', class: 'ak-icon-btn ak-close', 'aria-label': 'Close filters' }, icon('close'))
    const clearSide = h('button', { type: 'button', class: 'ak-link', text: 'Clear all' })
    const side = h(
      'aside',
      { class: 'ak-side', 'aria-label': 'Filters', id: `${cfg.kind}-filters`, tabindex: '-1' },
      h('div', { class: 'ak-side-head' }, h('h2', { text: 'Filters' }), clearSide, closeBtn),
      sideBody,
      h('div', { class: 'ak-side-foot' }, showBtn),
    )
    const scrim = h('div', { class: 'ak-scrim', hidden: true })
    filtersBtn.setAttribute('aria-controls', side.id)

    // Finder: "Select your machine" (spare parts only).
    let finder = null
    const finderSel = {}
    if (cfg.finder && usable.some((f) => f.fit)) {
      const mk = h('select', { class: 'ak-select', 'aria-label': 'Machine make' })
      const md = h('select', { class: 'ak-select', 'aria-label': 'Machine model' })
      const yr = h('select', { class: 'ak-select', 'aria-label': 'Year of the machine' })
      Object.assign(finderSel, { mk, md, yr })
      const go = h('button', { type: 'submit', class: 'ak-primary', text: 'Find parts' })
      const reset = h('button', { type: 'button', class: 'ak-link', text: 'Reset' })
      finder = h(
        'form',
        { class: 'ak-finder', role: 'search', 'aria-label': 'Find parts for your machine' },
        h('p', { class: 'ak-finder-title' }, h('strong', { text: 'Find parts for your machine' }), h('span', { text: 'Choose the make, then the model and year.' })),
        h('div', { class: 'ak-finder-row' }, mk, md, yr, go, reset),
      )
      mk.addEventListener('change', () => fillFinder(mk.value))
      finder.addEventListener('submit', (e) => {
        e.preventDefault()
        state.sel.make = mk.value ? [mk.value] : []
        state.sel.model = md.value ? [md.value] : []
        state.year = yr.value ? [Number(yr.value), Number(yr.value)] : [null, null]
        update(true)
        resultsLive.scrollIntoView({ behavior: 'smooth', block: 'start' })
      })
      reset.addEventListener('click', () => {
        state.sel.make = []
        state.sel.model = []
        state.year = [null, null]
        update(true)
      })
    }
    function fillFinder(make) {
      if (!finder) return
      const { mk, md, yr } = finderSel
      const makes = uniq(items.flatMap((i) => i.compat.map((c) => c.make))).filter((m) => !UNIVERSAL.test(m)).sort((a, b) => a.localeCompare(b))
      put(mk, h('option', { value: '', text: 'Make' }), makes.map((m) => h('option', { value: m, text: m })))
      mk.value = make || ''
      const rows = items.flatMap((i) => i.compat).filter((c) => !make || c.make === make)
      const models = uniq(rows.map((c) => c.model)).sort((a, b) => a.localeCompare(b, 'en', { numeric: true }))
      put(md, h('option', { value: '', text: make ? 'Any model' : 'Model' }), models.map((m) => h('option', { value: m, text: m })))
      md.disabled = !make
      const ys = rows.flatMap((c) => (c.from ? [c.from, c.to ?? c.from] : []))
      const lo = ys.length ? Math.min(...ys) : yearBounds?.[0]
      const hi = ys.length ? Math.max(...ys) : yearBounds?.[1]
      const opts = []
      if (lo) for (let y = hi; y >= lo; y--) opts.push(h('option', { value: String(y), text: String(y) }))
      put(yr, h('option', { value: '', text: 'Any year' }), opts)
    }

    const root = h(
      'div',
      { class: 'ak-cat', [OWN]: 'catalogue' },
      sample ? h('p', { class: 'ak-sample', role: 'note' }, h('strong', { text: 'Sample data. ' }), 'These parts are made up to show how search and filters work. Customers do not see this list; it appears only with ?demo=1 in the address.') : null,
      finder,
      h('div', { class: 'ak-bar' }, h('div', { class: 'ak-search' }, icon('search'), search, suggest), sortSel, h('div', { class: 'ak-views', role: 'group', 'aria-label': 'Layout' }, viewGrid, viewList), filtersBtn),
      h('div', { class: 'ak-body' }, side, h('div', { class: 'ak-main' }, h('div', { class: 'ak-summary' }, resultsLive, chips), grid, more, empty)),
      scrim,
    )
    put(host, root)

    // ── Rendering ──
    function facetOptions(f) {
      const vs = [...(valueSets[f.id] || [])]
      if (Array.isArray(f.order)) return vs.sort((a, b) => (f.order.indexOf(a) + 1 || 99) - (f.order.indexOf(b) + 1 || 99) || a.localeCompare(b))
      if (typeof f.order === 'function') return vs.sort(f.order)
      if (f.order === 'numeric') return vs.sort((a, b) => parseFloat(a) - parseFloat(b))
      if (f.order === 'alpha') return vs.sort((a, b) => a.localeCompare(b, 'en', { numeric: true, sensitivity: 'base' }))
      return vs
    }
    const open = {} // facet id → expanded "show all"
    const finds = {} // facet id → text typed in its own little search
    const collapsed = {}

    function rangeGroup(f, key, bounds) {
      const [lo, hi] = state[key]
      const fieldset = h('fieldset', { class: 'ak-group' })
      const legend = h('legend', { class: 'ak-group-title', text: f.label })
      if (f.unit === 'year') {
        const mkSel = (val, label) => {
          const s = h('select', { class: 'ak-select ak-select-sm', 'aria-label': `${f.label}: ${label}` }, h('option', { value: '', text: label }))
          for (let y = bounds[1]; y >= bounds[0]; y--) s.append(h('option', { value: String(y), text: String(y) }))
          s.value = val == null ? '' : String(val)
          return s
        }
        const a = mkSel(lo, 'From')
        const b = mkSel(hi, 'To')
        const set = () => {
          let x = a.value ? Number(a.value) : null
          let y = b.value ? Number(b.value) : null
          if (x != null && y != null && x > y) [x, y] = [y, x]
          state[key] = [x, y]
          update(true)
        }
        a.addEventListener('change', set)
        b.addEventListener('change', set)
        fieldset.append(legend, h('div', { class: 'ak-range' }, a, h('span', { 'aria-hidden': 'true', text: '–' }), b))
        if (key === 'year') fieldset.append(h('p', { class: 'ak-hint', text: 'Parts that fit any machine made in these years.' }))
      }
      return fieldset
    }

    function priceGroup() {
      const fieldset = h('fieldset', { class: 'ak-group' }, h('legend', { class: 'ak-group-title', text: 'Price (₹)' }))
      const mk = (val, label) => h('input', { type: 'number', min: '0', inputmode: 'numeric', class: 'ak-input', placeholder: label, 'aria-label': `Price ${label.toLowerCase()}`, value: val ?? '' })
      const a = mk(state.price[0], 'Min')
      const b = mk(state.price[1], 'Max')
      let timer
      const set = () => {
        clearTimeout(timer)
        timer = setTimeout(() => {
          const x = a.value === '' ? null : Math.max(0, Number(a.value))
          const y = b.value === '' ? null : Math.max(0, Number(b.value))
          state.price = x != null && y != null && x > y ? [y, x] : [x, y]
          update(true, true)
        }, 450)
      }
      a.addEventListener('input', set)
      b.addEventListener('input', set)
      const presets = h(
        'div',
        { class: 'ak-presets' },
        PRICE_PRESETS.filter(([lo]) => prices.some((p) => p >= lo)).map(([lo, hi, label]) => {
          const on = state.price[0] === lo && state.price[1] === hi
          return h('button', {
            type: 'button',
            class: `ak-pill${on ? ' is-on' : ''}`,
            'aria-pressed': String(on),
            text: label,
            onclick: () => {
              state.price = on ? [null, null] : [lo, hi]
              update(true)
            },
          })
        }),
      )
      const priced = [...scores.keys()].filter((i) => typeof i.price === 'number' && passes(i, 'price')).length
      fieldset.append(h('div', { class: 'ak-range' }, a, h('span', { 'aria-hidden': 'true', text: '–' }), b), presets, h('p', { class: 'ak-hint', text: `Items priced on enquiry are hidden while a price is set. ${priced} with a price.` }))
      return fieldset
    }

    function listGroup(f) {
      const all = facetOptions(f).map((v) => ({ v, n: count(f.id, v), on: (state.sel[f.id] || []).includes(v) }))
      let opts = f.dependent ? all.filter((o) => o.n > 0 || o.on) : all
      if (!opts.length) return null
      const fieldset = h('fieldset', { class: 'ak-group' })
      const title = h('button', { type: 'button', class: 'ak-group-title ak-toggle', 'aria-expanded': String(!collapsed[f.id]) }, h('span', { text: f.label }), (state.sel[f.id] || []).length ? h('span', { class: 'ak-badge', text: String(state.sel[f.id].length) }) : null)
      title.addEventListener('click', () => {
        collapsed[f.id] = !collapsed[f.id]
        render()
      })
      fieldset.append(h('legend', { class: 'ak-sr', text: f.label }), title)
      if (collapsed[f.id]) return fieldset
      if (opts.length > 8) {
        const find = h('input', { type: 'search', class: 'ak-input ak-input-sm', placeholder: `Find a ${f.label.toLowerCase()}`, 'aria-label': `Find in ${f.label}`, value: finds[f.id] || '' })
        find.addEventListener('input', () => {
          finds[f.id] = find.value
          render()
          const again = sideBody.querySelector(`[aria-label="Find in ${f.label}"]`)
          again?.focus()
          again?.setSelectionRange(find.value.length, find.value.length)
        })
        fieldset.append(find)
        if (finds[f.id]) opts = opts.filter((o) => norm(o.v).includes(norm(finds[f.id])))
      }
      const shown = open[f.id] || finds[f.id] ? opts : opts.slice(0, 6)
      const ul = h('ul', { class: 'ak-options' })
      for (const o of shown) {
        const id = `ak-${cfg.kind}-${f.id}-${code(o.v) || 'x'}-${Math.random().toString(36).slice(2, 6)}`
        const box = h('input', { type: 'checkbox', id, checked: o.on })
        box.addEventListener('change', () => {
          const cur = new Set(state.sel[f.id] || [])
          box.checked ? cur.add(o.v) : cur.delete(o.v)
          state.sel[f.id] = [...cur]
          if (f.id === 'make') state.sel.model = (state.sel.model || []).filter((m) => items.some((i) => i.compat.some((c) => c.model === m && (!cur.size || cur.has(c.make)))))
          update(true)
        })
        ul.append(h('li', { class: o.n === 0 && !o.on ? 'is-zero' : '' }, h('label', { for: id }, box, h('span', { class: 'ak-opt', text: o.v }), h('span', { class: 'ak-n', text: String(o.n) }))))
      }
      fieldset.append(ul)
      if (opts.length > 6 && !finds[f.id]) {
        fieldset.append(
          h('button', {
            type: 'button',
            class: 'ak-link',
            text: open[f.id] ? 'Show fewer' : `Show all ${opts.length}`,
            onclick: () => {
              open[f.id] = !open[f.id]
              render()
            },
          }),
        )
      }
      return fieldset
    }

    function chipList() {
      const out = []
      const add = (label, clear) => out.push({ label, clear })
      if (state.q) add(`“${state.q}”`, () => (state.q = ''))
      if (state.sel.make?.length || state.sel.model?.length || state.year[0] != null || state.year[1] != null) {
        for (const v of state.sel.make || []) add(`Make: ${v}`, () => (state.sel.make = state.sel.make.filter((x) => x !== v)))
        for (const v of state.sel.model || []) add(`Model: ${v}`, () => (state.sel.model = state.sel.model.filter((x) => x !== v)))
        if (state.year[0] != null || state.year[1] != null) add(`Machine year: ${yearText(state.year)}`, () => (state.year = [null, null]))
      }
      for (const f of listFacets) {
        if (f.fit) continue
        for (const v of state.sel[f.id] || []) add(`${f.label}: ${v}`, () => (state.sel[f.id] = state.sel[f.id].filter((x) => x !== v)))
      }
      if (state.made[0] != null || state.made[1] != null) add(`Made: ${yearText(state.made)}`, () => (state.made = [null, null]))
      if (state.price[0] != null || state.price[1] != null) add(`Price: ${state.price[0] != null ? inr(state.price[0]) : '₹0'} – ${state.price[1] != null ? inr(state.price[1]) : 'any'}`, () => (state.price = [null, null]))
      return out
    }
    const yearText = ([a, b]) => (a != null && b != null ? (a === b ? String(a) : `${a}–${b}`) : a != null ? `${a} or later` : `${b} or earlier`)

    function priceBlock(it) {
      if (typeof it.price === 'number') {
        const off = it.mrp && it.mrp > it.price ? Math.round((1 - it.price / it.mrp) * 100) : 0
        return h('p', { class: 'ak-price' }, h('strong', { text: inr(it.price) }), off ? h('s', { text: inr(it.mrp), 'aria-label': `List price ${inr(it.mrp)}` }) : null, off ? h('span', { class: 'ak-off', text: `${off}% off` }) : null)
      }
      return h('p', { class: 'ak-price ak-price-ask', text: it.priceText && !/[\d]/.test(it.priceText) ? it.priceText : 'Price on request' })
    }
    const AVAIL_TONE = { 'In stock': 'good', 'Low stock': 'warn', 'Made to order': 'info', 'On order': 'info', Enquire: 'muted', 'Out of stock': 'bad' }

    function fitsLine(it, max = 2) {
      if (!it.compat.length) return null
      const mk = state.sel.make || []
      const md = state.sel.model || []
      const sorted = [...it.compat].sort((a, b) => Number(mk.includes(b.make) || md.includes(b.model)) - Number(mk.includes(a.make) || md.includes(a.model)))
      const text = (c) => [c.make, c.model].filter(Boolean).join(' ') + (c.from ? ` (${c.from === c.to ? c.from : `${c.from}–${c.to === YEAR_NOW ? 'now' : c.to}`})` : '')
      const shown = sorted.slice(0, max).map(text).join(', ')
      return h('p', { class: 'ak-fits' }, h('span', { class: 'ak-k', text: 'Fits ' }), shown, it.compat.length > max ? ` +${it.compat.length - max} more` : '')
    }

    function card(it) {
      const badges = []
      const avail = it.facets.availability?.[0]
      if (cfg.kind === 'parts') {
        for (const c of it.facets.condition || []) badges.push(h('span', { class: `ak-tag ${c === 'New' ? 'tone-good' : 'tone-info'}`, text: c }))
        for (const g of it.facets.genuine || []) badges.push(h('span', { class: 'ak-tag', text: g }))
      }
      if (avail && avail !== 'Enquire') badges.push(h('span', { class: `ak-tag tone-${AVAIL_TONE[avail] || 'muted'}`, text: avail + (avail === 'Low stock' && it.qty ? ` · ${it.qty} left` : '') }))
      const openDetail = cfg.kind === 'parts'
      const titleLink = openDetail
        ? h('a', { href: `?part=${encodeURIComponent(it.id)}`, text: it.title, onclick: (e) => { e.preventDefault(); showPart(it) } })
        : h('a', { href: it.href || '#', text: it.title })
      const img = it.image?.url
        ? h('img', { src: it.image.url.startsWith('/') ? BASE + it.image.url : it.image.url, alt: it.image.alt || '', loading: 'lazy', decoding: 'async' })
        : icon('part', 'ak-ph')
      const media = h('div', { class: `ak-media${it.image?.url ? '' : ' is-empty'}` }, img)
      if (openDetail) {
        media.addEventListener('click', () => showPart(it))
        media.style.cursor = 'pointer'
      }
      return h(
        'li',
        { class: 'ak-card' },
        media,
        h(
          'div',
          { class: 'ak-card-body' },
          it.eyebrow ? h('p', { class: 'ak-eyebrow', text: it.eyebrow }) : null,
          h('h3', { class: 'ak-title' }, titleLink),
          it.partNo ? h('p', { class: 'ak-code' }, h('span', { class: 'ak-k', text: cfg.kind === 'parts' ? 'Part no. ' : 'Model ' }), it.partNo, it.oem.length ? h('span', { class: 'ak-k', text: ` · OEM ${it.oem.slice(0, 2).join(', ')}${it.oem.length > 2 ? '…' : ''}` }) : null) : null,
          fitsLine(it),
          it.summary && (state.view === 'list' || cfg.kind !== 'parts') ? h('p', { class: 'ak-sum', text: it.summary }) : null,
          badges.length ? h('p', { class: 'ak-tags' }, badges) : null,
          h('div', { class: 'ak-card-foot' }, priceBlock(it), h('a', { class: 'ak-secondary', href: it.cta.href, text: it.cta.label })),
        ),
      )
    }

    // Part details in a dialog, linkable with ?part=…
    const dialog = h('dialog', { class: 'ak-dialog', 'aria-labelledby': `${cfg.kind}-dlg-title` })
    root.append(dialog)
    dialog.addEventListener('close', () => {
      if (state.part) {
        state.part = ''
        writeUrl(false)
      }
    })
    dialog.addEventListener('click', (e) => e.target === dialog && dialog.close())
    function showPart(it) {
      state.part = it.id
      writeUrl(true)
      const fits = it.compat.length
        ? h(
            'table',
            { class: 'ak-fit-table' },
            h('caption', { text: 'Fits these machines' }),
            h('thead', {}, h('tr', {}, h('th', { text: 'Make' }), h('th', { text: 'Model' }), h('th', { text: 'Years' }))),
            h('tbody', {}, it.compat.map((c) => h('tr', {}, h('td', { text: c.make || '–' }), h('td', { text: c.model || 'All models' }), h('td', { text: c.from ? (c.from === c.to ? String(c.from) : `${c.from}–${c.to === YEAR_NOW ? 'now' : c.to}`) : 'All years' })))),
          )
        : null
      put(dialog, 
        h(
          'div',
          { class: 'ak-dlg' },
          h('button', { type: 'button', class: 'ak-icon-btn ak-dlg-close', 'aria-label': 'Close', onclick: () => dialog.close() }, icon('close')),
          h('div', { class: 'ak-dlg-media' }, it.image?.url ? h('img', { src: it.image.url.startsWith('/') ? BASE + it.image.url : it.image.url, alt: it.image.alt || '' }) : icon('part', 'ak-ph')),
          h(
            'div',
            { class: 'ak-dlg-body' },
            it.eyebrow ? h('p', { class: 'ak-eyebrow', text: it.eyebrow }) : null,
            h('h2', { id: `${cfg.kind}-dlg-title`, text: it.title }),
            it.summary ? h('p', { text: it.summary }) : null,
            priceBlock(it),
            h('dl', { class: 'ak-specs' }, it.details.map(([k, v]) => [h('dt', { text: k }), h('dd', { text: v })])),
            fits,
            h('div', { class: 'ak-dlg-actions' }, h('a', { class: 'ak-primary', href: it.cta.href, text: it.cta.label }), h('button', { type: 'button', class: 'ak-secondary', text: 'Copy link', onclick: (e) => copyLink(e.currentTarget) })),
          ),
        ),
      )
      if (!dialog.open) dialog.showModal()
    }
    async function copyLink(btn) {
      try {
        await navigator.clipboard.writeText(location.href)
        btn.textContent = 'Link copied'
      } catch {
        btn.textContent = 'Copy the address bar'
      }
    }

    // Mobile drawer
    const narrow = matchMedia('(max-width: 63.99rem)')
    // On phones the filters open as a drawer. It is moved to <body> while open, because
    // the page's own animated wrappers would otherwise trap a fixed panel inside them.
    const portal = h('div', { class: 'ak-cat ak-portal' })
    const home = h('span', { hidden: true })
    function openDrawer() {
      side.before(home)
      portal.append(scrim, side)
      document.body.append(portal)
      side.classList.add('is-open')
      scrim.hidden = false
      document.documentElement.classList.add('ak-lock')
      filtersBtn.setAttribute('aria-expanded', 'true')
      side.setAttribute('role', 'dialog')
      side.setAttribute('aria-modal', 'true')
      closeBtn.focus()
    }
    function closeDrawer() {
      if (!side.classList.contains('is-open')) return
      side.classList.remove('is-open')
      scrim.hidden = true
      if (home.isConnected) {
        home.replaceWith(side)
        root.append(scrim)
      }
      portal.remove()
      document.documentElement.classList.remove('ak-lock')
      filtersBtn.setAttribute('aria-expanded', 'false')
      side.removeAttribute('role')
      side.removeAttribute('aria-modal')
      filtersBtn.focus()
    }
    filtersBtn.addEventListener('click', openDrawer)
    closeBtn.addEventListener('click', closeDrawer)
    showBtn.addEventListener('click', closeDrawer)
    scrim.addEventListener('click', closeDrawer)
    side.addEventListener('keydown', (e) => e.key === 'Escape' && closeDrawer())
    narrow.addEventListener?.('change', () => !narrow.matches && closeDrawer())

    // Controls
    let typing
    search.addEventListener('input', () => {
      clearTimeout(typing)
      typing = setTimeout(() => {
        state.q = search.value.trim()
        update(false, true)
      }, 180)
    })
    search.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault()
        clearTimeout(typing)
        state.q = search.value.trim()
        update(true)
      }
    })
    sortSel.addEventListener('change', () => {
      state.sort = sortSel.value
      update(true)
    })
    viewGrid.addEventListener('click', () => {
      state.view = 'grid'
      update(false)
    })
    viewList.addEventListener('click', () => {
      state.view = 'list'
      update(false)
    })
    const clearAll = () => {
      state.sel = {}
      state.price = [null, null]
      state.year = [null, null]
      state.made = [null, null]
      state.q = ''
      update(true)
    }
    clearSide.addEventListener('click', clearAll)
    window.addEventListener('popstate', () => {
      readUrl()
      render()
      syncPart()
    })

    function update(push, keepFocus) {
      state.limit = PAGE_SIZE
      writeUrl(push)
      render(keepFocus)
    }

    function render() {
      const shown = results()
      if (document.activeElement !== search) search.value = state.q
      sortSel.value = state.sort
      fillFinder(state.sel.make?.[0] || '')
      if (finder) {
        finderSel.md.value = state.sel.model?.length === 1 ? state.sel.model[0] : ''
        finderSel.yr.value = state.year[0] != null && state.year[0] === state.year[1] ? String(state.year[0]) : ''
      }
      viewGrid.setAttribute('aria-pressed', String(state.view === 'grid'))
      viewList.setAttribute('aria-pressed', String(state.view === 'list'))
      grid.className = `ak-results is-${state.view}`

      // Sidebar: machine fit first, then price, then the rest.
      const fitEls = []
      const otherEls = []
      for (const f of usable) {
        const el = f.id === 'year' ? rangeGroup(f, 'year', yearBounds) : f.id === 'made' ? rangeGroup(f, 'made', madeBounds) : listGroup(f)
        if (el) (f.fit ? fitEls : otherEls).push(el)
      }
      if (fitEls.length) fitEls.unshift(h('p', { class: 'ak-side-sub', text: 'Your machine' }))
      put(sideBody, ...fitEls, ...(hasPrice ? [priceGroup()] : []), ...otherEls)

      // Chips and counts
      const cs = chipList()
      const active = cs.length - (state.q ? 1 : 0)
      const badge = filtersBtn.querySelector('.ak-badge')
      badge.hidden = !active
      badge.textContent = String(active)
      clearSide.hidden = !cs.length
      resultsLive.textContent = shown.length === items.length && !cs.length ? `${plural(items.length, cfg.noun[0], cfg.noun[1])}` : `${plural(shown.length, cfg.noun[0], cfg.noun[1])} found`
      showBtn.textContent = shown.length ? `Show ${plural(shown.length, 'result', 'results')}` : 'No results: change filters'
      put(chips, 
        ...cs.map((c) =>
          h('button', {
            type: 'button',
            class: 'ak-chip',
            'aria-label': `Remove ${c.label}`,
            onclick: () => {
              c.clear()
              update(true)
            },
          }, h('span', { text: c.label }), icon('close')),
        ),
        cs.length > 1 ? h('button', { type: 'button', class: 'ak-link', text: 'Clear all', onclick: clearAll }) : null,
      )

      // Results
      put(grid, ...shown.slice(0, state.limit).map(card))
      put(more, 
        shown.length > state.limit
          ? h('button', {
              type: 'button',
              class: 'ak-secondary',
              text: `Show more (${shown.length - state.limit} left)`,
              onclick: () => {
                state.limit += PAGE_SIZE
                render()
              },
            })
          : shown.length > PAGE_SIZE
            ? h('p', { class: 'ak-hint', text: `Showing all ${shown.length}.` })
            : null,
      )
      empty.hidden = shown.length > 0
      if (!shown.length) {
        const last = cs[cs.length - 1]
        put(empty, 
          h('h3', { text: items.length ? `No ${cfg.noun[1]} match` : `Our ${cfg.noun[1]} list is being prepared` }),
          h('p', { text: items.length ? 'Try fewer words, a part number without spaces, or remove a filter.' : 'Tell us what you need and we will find it for you.' }),
          h(
            'div',
            { class: 'ak-empty-actions' },
            last ? h('button', { type: 'button', class: 'ak-secondary', text: `Remove “${last.label}”`, onclick: () => { last.clear(); update(true) } }) : null,
            cs.length > 1 ? h('button', { type: 'button', class: 'ak-secondary', text: 'Clear all filters', onclick: clearAll }) : null,
            h('a', {
              class: 'ak-primary',
              href: `${BASE}/request-a-quote/product${state.q ? `?item=${encodeURIComponent(state.q)}` : ''}`,
              text: cfg.kind === 'parts' ? 'Ask us to source this part' : 'Ask us about it',
            }),
          ),
          h('p', { class: 'ak-hint' }, 'Or email the part number and a photo to ', h('a', { href: `mailto:${OFFICE_EMAIL}?subject=Spare%20part%20enquiry`, text: OFFICE_EMAIL }), '.'),
        )
      }
    }
    function syncPart() {
      const it = state.part && items.find((i) => i.id === state.part)
      if (it) showPart(it)
      else if (dialog.open) dialog.close()
    }

    readUrl()
    render()
    syncPart()
  }

  // ── Mounting ───────────────────────────────────────────────────────────────
  async function loadItems(cfg, demo) {
    if (cfg.kind === 'parts') {
      const file = demo ? 'spare-parts.sample.csv' : 'spare-parts.csv'
      const res = await fetch(`${BASE}/static-site/${file}?v=${Math.floor(Date.now() / 300000)}`)
      if (!res.ok) return []
      return parseCsv(await res.text()).map(partFromRow)
    }
    const res = await fetch(`${BASE}/static-site/${cfg.source}`)
    if (!res.ok) return null
    return fromExport(await res.json())
  }

  let mounted = ''
  let busy = false
  async function tryMount() {
    const path = pagePath()
    const cfg = PAGES[path]
    if (!cfg) {
      mounted = ''
      return
    }
    if (busy || (mounted === path && document.querySelector(`[${OWN}="catalogue"]`))) return
    const main = document.querySelector('main')
    if (!main || !hydrated(main)) return
    let host
    if (cfg.kind === 'parts') {
      host = h('section', { class: 'ak-page' })
    } else {
      // Replace the built-in list (search box + sidebar) with this one.
      const box = main.querySelector('[role="search"]')
      if (!box) return
      let grid = box.parentElement
      if (grid?.parentElement?.querySelector(':scope > aside')) grid = grid.parentElement
      if (!grid) return
      host = h('div', { class: 'ak-host' })
      grid.style.display = 'none'
      grid.after(host)
      host._hidden = grid
      addSparePartsLink(main)
    }
    busy = true
    const demo = cfg.kind === 'parts' && new URLSearchParams(location.search).get('demo') === '1'
    let items
    try {
      items = await loadItems(cfg, demo)
    } catch (err) {
      console.error('[askib catalogue]', err)
      items = null
    }
    busy = false
    if (items == null) {
      // Could not load: put the built-in list back.
      if (host._hidden) host._hidden.style.display = ''
      host.remove()
      return
    }
    if (cfg.kind === 'parts') {
      for (const c of main.children) c.style.display = 'none'
      document.title = 'Spare parts | ASKIB'
      host.append(
        h('header', { class: 'ak-head' }, h('p', { class: 'ak-eyebrow', text: cfg.eyebrow }), h('h1', { text: cfg.title }), h('p', { class: 'ak-lead', text: cfg.intro })),
      )
      const body = h('div', {})
      host.append(body)
      main.append(host)
      mountBrowser(body, cfg, items, demo)
    } else {
      mountBrowser(host, cfg, items, false)
    }
    host.setAttribute(OWN, 'catalogue')
    mounted = path
  }

  function addSparePartsLink(main) {
    if (pagePath() !== '/products' || main.querySelector('[data-ak-spares]')) return
    const pill = [...main.querySelectorAll('ul a[href*="/products/"]')][0]?.closest('ul')
    if (!pill) return
    const sample = pill.querySelector('li')
    const li = sample ? sample.cloneNode(true) : h('li')
    const a = li.querySelector('a') || li.appendChild(h('a'))
    a.href = `${BASE}/spare-parts`
    a.textContent = 'Spare parts'
    li.dataset.akSpares = '1'
    pill.append(li)
  }

  function cleanUp() {
    if (PAGES[pagePath()]) return
    for (const el of document.querySelectorAll(`[${OWN}="catalogue"]`)) {
      if (el._hidden) el._hidden.style.display = ''
      if (el.classList.contains('ak-page')) for (const c of el.parentElement.children) c.style.display = ''
      el.remove()
    }
    for (const el of document.querySelectorAll('.ak-portal')) el.remove()
    document.documentElement.classList.remove('ak-lock')
  }

  let queued = false
  const tick = () => {
    queued = false
    cleanUp()
    tryMount()
  }
  const schedule = () => {
    if (!queued) {
      queued = true
      requestAnimationFrame(tick)
    }
  }
  const link = document.createElement('link')
  link.rel = 'stylesheet'
  link.href = `${BASE}/static-site/catalogue.css?v=1`
  document.head.append(link)
  tick()
  new MutationObserver(schedule).observe(document.body, { childList: true, subtree: true })
  setTimeout(schedule, 300)
})()
