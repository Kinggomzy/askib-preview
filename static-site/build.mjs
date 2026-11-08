// Turns on forms and search in the static copy of askib.in.
//
//   node static-site/build.mjs [site folder]     (default: the folder above this one)
//
// Safe to run again after every new export. It:
//   1. adds static-site/site.css and site.js to every page,
//   2. writes static-site/search-index.json (title, description and headings of each page),
//   3. writes static-site/products.json (product id → name, to prefill "Request a quote" links),
//   4. adds a /search page, built from the not-found page, that site.js fills with results,
//   5. writes catalogue-products.json and catalogue-services.json (the lists catalogue.js
//      searches, filters and sorts), read from the exported /products and /services pages,
//   6. adds the /spare-parts page (filled from spare-parts.csv by catalogue.js) and the
//      /testimonial page (the link to send clients; site.js shows the form), and puts the
//      spare parts in the search index.
// No dependencies: Node 18+.
import fs from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

const here = path.dirname(fileURLToPath(import.meta.url))
const root = path.resolve(process.argv[2] ?? path.join(here, '..'))
const outDir = path.join(root, 'static-site')
/** Set when the site is served from a sub-folder, e.g. /askib-platform. */
const base = (process.env.STATIC_BASE_PATH ?? '').replace(/\/$/, '')
const MARK = '<!-- askib-static-site -->'
const TAGS = `${MARK}<link rel="stylesheet" href="${base}/static-site/site.css"><script src="${base}/static-site/site.js" defer></script>`

/** Pages left out of search: they have nothing to find or need a server. */
const NOT_SEARCHED = [/^\/404$/, /^\/search$/, /^\/testimonial$/, /^\/spare-parts$/, /^\/cart$/, /^\/checkout$/, /^\/track-order$/, /^\/image-credits$/]

function* htmlFiles(dir) {
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    if (e.name.startsWith('.') || e.name === '_next' || e.name === 'static-site' || (dir === root && e.name === 'api')) continue
    const p = path.join(dir, e.name)
    if (e.isDirectory()) yield* htmlFiles(p)
    else if (e.name.endsWith('.html')) yield p
  }
}

const decode = (s) =>
  s
    .replace(/<[^>]+>/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#x27;|&#39;/g, "'")
    .replace(/&nbsp;/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()

const urlOf = (file) => {
  const rel = '/' + path.relative(root, file).split(path.sep).join('/')
  return rel.replace(/(\/index)?\.html$/, '') || '/' // without the base path
}

// 1. Inject the tags.
let injected = 0
const pages = new Map()
for (const file of htmlFiles(root)) {
  let html = fs.readFileSync(file, 'utf8')
  if (!html.includes(MARK)) {
    html = html.replace('</head>', `${TAGS}</head>`)
    fs.writeFileSync(file, html)
    injected++
  }
  const url = urlOf(file)
  // x.html and x/index.html are the same page; keep one.
  if (!pages.has(url)) pages.set(url, html)
}

// 2. Search index.
const index = []
const products = {}
for (const [u, html] of pages) {
  const main = html.slice(html.indexOf('<main'), html.indexOf('</main>') + 7)
  for (const m of main.matchAll(/href="[^"]*\/request-a-quote\/product\?product=(\d+)"/g)) {
    const h1 = main.match(/<h1[^>]*>([\s\S]*?)<\/h1>/)
    if (h1 && !products[m[1]] && u.startsWith('/products/')) products[m[1]] = decode(h1[1])
  }
  if (NOT_SEARCHED.some((re) => re.test(u))) continue
  const title = decode(html.match(/<title>([\s\S]*?)<\/title>/)?.[1] ?? '')
    .replace(/\s*\|\s*ASKIB\s*$/i, '')
    .trim()
  if (!title || /not found/i.test(title)) continue
  const desc = decode(html.match(/<meta name="description" content="([^"]*)"/)?.[1] ?? '')
  const heads = [...main.matchAll(/<h[1-3][^>]*>([\s\S]*?)<\/h[1-3]>/g)].map((m) => decode(m[1])).filter(Boolean)
  index.push({ u, t: title, d: desc, k: [...new Set(heads)].join(' · ').slice(0, 400) })
}
// Spare parts: the page itself and every part, so the site search finds part numbers.
const csvPath = path.join(outDir, 'spare-parts.csv')
const parts = fs.existsSync(csvPath) ? parseCsv(fs.readFileSync(csvPath, 'utf8')) : []
index.push({
  u: '/spare-parts',
  t: 'Spare parts',
  d: 'Spare and replacement parts for testing machines, measuring instruments and lab equipment. Search by part number, make, model and year.',
  k: 'spares · replacement parts · part number · OEM',
})
for (const [i, r] of parts.entries()) {
  const id = r['ID'] || String(i + 1)
  const no = r['Part number']
  index.push({
    u: `/spare-parts?part=${encodeURIComponent(id)}`,
    t: `${r['Part name'] || no}${no ? ` (${no})` : ''}`,
    d: r['Description'] || '',
    k: [no, r['OEM / alternate numbers'], r['Category'], r['Brand'], r['Fits make'], r['Fits model'], r['Keywords']].filter(Boolean).join(' · ').slice(0, 400),
  })
}
index.sort((a, b) => a.u.localeCompare(b.u))
fs.mkdirSync(outDir, { recursive: true })
fs.writeFileSync(path.join(outDir, 'search-index.json'), JSON.stringify(index))
fs.writeFileSync(path.join(outDir, 'products.json'), JSON.stringify(products))

// site.js needs the base path too.
const js = path.join(outDir, 'site.js')
fs.writeFileSync(js, fs.readFileSync(js, 'utf8').replace(/const BASE = '[^']*'/, `const BASE = '${base}'`))

// 3. The /search page.
const notFound = fs.readFileSync(path.join(root, '404.html'), 'utf8')
fs.writeFileSync(path.join(root, 'search.html'), notFound)
fs.mkdirSync(path.join(root, 'search'), { recursive: true })
fs.writeFileSync(path.join(root, 'search', 'index.html'), notFound)

// 4. Catalogue lists for catalogue.js, from the React props saved in each listing page.
function browseProps(html) {
  const flight = [...html.matchAll(/self\.__next_f\.push\(\[1,("(?:[^"\\]|\\.)*")\]\)/g)].map((m) => JSON.parse(m[1])).join('')
  const line = flight.split('\n').find((l) => l.includes('"items":[') && l.includes('"filters":['))
  if (!line) return null
  const tree = JSON.parse(line.slice(line.indexOf(':') + 1))
  const find = (n) => {
    if (!n || typeof n !== 'object') return null
    if (Array.isArray(n.items) && Array.isArray(n.filters)) return n
    for (const v of Array.isArray(n) ? n : Object.values(n)) {
      const r = find(v)
      if (r) return r
    }
    return null
  }
  return find(tree)
}
const undef = (v) => (v === '$undefined' ? undefined : v)
let catalogues = 0
for (const [kind, url] of [['products', '/products'], ['services', '/services']]) {
  const html = pages.get(url)
  const props = html && browseProps(html)
  if (!props) {
    console.warn(`No list found on ${url}; catalogue-${kind}.json not written.`)
    continue
  }
  const items = props.items.map((x) => ({
    id: x.id,
    title: x.title,
    href: undef(x.href),
    summary: undef(x.summary),
    eyebrow: undef(x.eyebrow),
    image: x.image && undef(x.image) ? { url: x.image.url, alt: undef(x.image.alt) || '' } : null,
    price: typeof x.sort?.price === 'number' ? x.sort.price : null,
    priceText: undef(x.price?.text),
    was: undef(x.price?.was),
    off: undef(x.price?.off),
    search: undef(x.search),
    facets: x.facets || {},
  }))
  const filters = props.filters.map((g) => ({ id: g.id, label: g.label, options: g.options }))
  fs.writeFileSync(path.join(outDir, `catalogue-${kind}.json`), JSON.stringify({ kind, filters, items }))
  catalogues++
}

// 5. The /spare-parts and /testimonial pages, from the not-found page.
const shell = (title, description, robots) =>
  notFound
    .replace(/<title>[\s\S]*?<\/title>/, `<title>${title} | ASKIB</title>`)
    .replace(/<meta name="description" content="[^"]*"\s*\/?>/, `<meta name="description" content="${description}"/>`)
    .replace(/<meta name="robots" content="[^"]*"\s*\/?>/g, robots ? `<meta name="robots" content="${robots}"/>` : '')
for (const [slug, html] of [
  ['spare-parts', shell('Spare parts', 'Spare and replacement parts for testing machines, measuring instruments and lab equipment. Search by part number, make, model and year.', '')],
  ['testimonial', shell('Share your experience', 'Tell us how ASKIB Engineers helped you.', 'noindex')],
]) {
  fs.writeFileSync(path.join(root, `${slug}.html`), html)
  fs.mkdirSync(path.join(root, slug), { recursive: true })
  fs.writeFileSync(path.join(root, slug, 'index.html'), html)
}

console.log(`Tags added to ${injected} pages. Search index: ${index.length} pages. Products: ${Object.keys(products).length}. Catalogue lists: ${catalogues}. Spare parts: ${parts.length}.`)

/** Reads a CSV saved by Excel (commas, quotes, line breaks inside quotes). */
function parseCsv(text) {
  const rows = []
  let row = []
  let cell = ''
  let quoted = false
  text = text.replace(/^\uFEFF/, '')
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
  return rows.filter((r) => r.some((x) => x.trim())).map((r) => Object.fromEntries(head.map((k, i) => [k, (r[i] ?? '').trim()])))
}
