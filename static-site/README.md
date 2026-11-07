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
  the `/search` page. Run it after every new export: `node static-site/build.mjs`.

## Connecting the forms

See [worker/README.md](worker/README.md). Until the handler's address is pasted
into `FORM_ENDPOINT` in `site.js`, the forms show the office email and phone instead.
