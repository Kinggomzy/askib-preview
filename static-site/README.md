# Forms and search on the static site

askib.in is a static copy of the website, so it has no server to receive forms
or run searches. The files here add both back:

- `site.js` swaps each "Forms are switched off" notice for a working form and
  sends it to our own form handler in `worker/` (a free Cloudflare Worker). The
  handler emails the enquiry from info@askib.in through Microsoft 365 and adds a
  row to the "ASKIB Website Enquiries" Excel workbook. Spam is filtered with a
  hidden honeypot field, a minimum fill time and a limit of 5 enquiries an hour
  per email address.
- `site.css` styles the forms and search results with the site's own colours.
- `search-index.json` and `products.json` are generated; do not edit them.
- `build.mjs` adds the two files to every page, regenerates the index and adds
  the `/search` page. Run it after every new export: `node static-site/build.mjs` (scripts/static-preview.mjs does this for you).

## Connecting the forms

See [worker/README.md](worker/README.md). Until the handler's address is pasted
into `FORM_ENDPOINT` in `site.js`, the forms show the office email and phone instead.

## Testimonials from clients

Send clients the link **https://askib.in/testimonial**. You can fill in their
name and company for them, so the page greets them by name:

    https://askib.in/testimonial?name=Priya%20Sen&company=Eastern%20Concrete%20Labs

They give their name, company, role (optional), city, email (never shown),
a star rating, what we helped with, the testimonial and a photo, and must tick
the consent box. The photo is shrunk in their browser (longest side 800 px,
JPEG, under ~700 KB) before it is sent.

The form handler (`worker/`) then:

- adds a row to the **Testimonials** sheet of the enquiries workbook with
  Status **Pending**,
- saves the photo as `Testimonial photos/<Reference>.jpg` in the same OneDrive
  folder as the workbook,
- emails info@askib.in with the testimonial and the photo attached.

**Nothing appears on the website until you approve it.** To approve, change
Status to **Approved** in the sheet (it has a drop-down). Within about five
minutes it shows on /testimonials, with the photo beside it. Type **Yes** in
"Show on homepage" to feature it on the homepage too, and use **Order** (1, 2,
3…) to put the best first. Change Status back to Pending or Rejected to take it
down. You may correct typos in the wording before approving.

To add one yourself (e.g. from a letter or WhatsApp message), add a row with a
unique Reference such as `TST-MANUAL-01`, fill Name, Company and Testimonial,
set Status to Approved, and either upload `TST-MANUAL-01.jpg` to the photos
folder (type anything in Photo) or paste a public image link in Photo.

Every testimonial on the site, old or new, shows a round photo beside the name,
or the person's initials when there is no photo.

## Spare parts

`/spare-parts` lists the parts in **`spare-parts.csv`**. Open it in Excel, add
one row per part, save as CSV, and upload it to this folder on GitHub (Add file
→ Upload files). The page reads it directly, so no rebuild is needed; run
`node static-site/build.mjs` afterwards only to add the new parts to the site search.

Columns (all optional except Part name; leave a cell empty if unknown):

| Column | Example | Used for |
|---|---|---|
| ID | SP-001 | the part's link (`/spare-parts?part=SP-001`); keep it unique and unchanged |
| Part name | Hydraulic pump assembly | title, search |
| Part number | HP-2000-01 | search (spaces, dashes and dots are ignored) |
| OEM / alternate numbers | OEM-77821; HX2000P | search; separate several with `;` |
| Category / Sub-category | Hydraulics / Pumps | filters |
| Equipment type | Compression testing machine | "For equipment" filter |
| Brand | Hydrotest | filter (maker of the part) |
| Condition | New, Used, Refurbished | filter |
| Genuine or compatible | Genuine OEM, Compatible | filter |
| Fits make / Fits model / Fits years | Hydrotest / CTM-2000; CTM-3000 / 2012-2020; 2016- | "Find parts for your machine", make, model and year filters (see below) |
| Year of manufacture | 2019 | year range filter (useful for used parts and machines) |
| Price (INR) / List price (INR) | 48500 / 54000 | price filter and sort; list price shows a strike-through and % off. Empty = "Price on request" |
| Availability / Stock qty / Lead time (days) | In stock / 3 / 14 | filter and "In stock first" sort. Words understood: In stock, Low stock, Made to order, On order, Out of stock |
| Location | Kolkata | "Ships from" filter |
| Delivery | Pan-India courier | filter |
| Country of origin | India | filter |
| Warranty (months) | 12 | filter |
| Minimum order, Material, Dimensions, Weight (kg) | | details |
| Specifications | Flow: 2.2 L/min; Max pressure: 700 bar | details and search (`label: value`, separated by `;`) |
| Image URL | /api/media/file/pump.jpg or https://… | picture |
| Description, Keywords | | details and search (keywords are hidden, e.g. "gasket oring") |
| Date added | 2026-11-08 | "Newest first" sort |
| Popularity | 120 | "Most popular" sort (e.g. units sold or enquiries) |
| Featured | Yes | shown first by default |

**Machines a part fits:** put several in the same order across the three Fits
columns, separated by `;`. A single make or year applies to every model, a year
range can be open-ended (`2016-` = 2016 to now), and an empty model means
"all models of that make".

To see the page with made-up example parts (customers never see these), open
`/spare-parts?demo=1`; the examples are in `spare-parts.sample.csv`.

## Search, filters and sorting on /products, /services and /spare-parts

`catalogue.js` (loaded by site.js on those pages) replaces the built-in list with
one that has keyword and part-number search with suggestions, filters with
result counts, active-filter chips with Clear all, price presets and ranges,
year ranges, sorting (best match, featured, price, newest, most popular, in
stock first, name, year), grid or list view, and a filter drawer on phones.
Every choice is kept in the address, so a filtered list can be bookmarked or
sent to a customer, e.g. `/spare-parts?make=Hydrotest&model=CTM-2000&year=2018-2018`.
The product and service lists come from `catalogue-products.json` and
`catalogue-services.json`, which build.mjs writes from the exported pages.
