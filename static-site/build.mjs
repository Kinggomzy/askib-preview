// Turns on forms and search in the static copy of askib.in.
//
//   node static-site/build.mjs [site folder]     (default: the folder above this one)
//
// Safe to run again after every new export. It:
//   1. adds static-site/site.css and site.js to every page,
//   2. writes static-site/search-index.json (title, description and headings of each page),
//   3. writes static-site/products.json (product id → name, to prefill "Request a quote" links),
//   4. adds a /search page, built from the not-found page, that site.js fills with results.
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
const NOT_SEARCHED = [/^\/404$/, /^\/search$/, /^\/cart$/, /^\/checkout$/, /^\/track-order$/, /^\/image-credits$/]

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

console.log(`Tags added to ${injected} pages. Search index: ${index.length} pages. Products: ${Object.keys(products).length}.`)
