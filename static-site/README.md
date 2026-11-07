# Forms and search on the static site

askib.in is a static copy of the website, so it has no server to receive forms
or run searches. The files here add both back:

- `site.js` swaps each "Forms are switched off" notice for a working form. Forms
  are emailed through [Web3Forms](https://web3forms.com) (free plan, 250 a month).
  Spam is filtered with a hidden honeypot field and a minimum fill time.
- `site.css` styles the forms and search results with the site's own colours.
- `search-index.json` and `products.json` are generated; do not edit them.
- `build.mjs` adds the two files to every page, regenerates the index and adds
  the `/search` page. Run it after every new export: `node static-site/build.mjs`.

## Connecting the forms

1. Go to https://web3forms.com, enter `info@askib.in` and press "Create Access Key".
2. Copy the key from the email that arrives.
3. Paste it into `WEB3FORMS_ACCESS_KEY` near the top of `site.js` and push.

Until a key is pasted, the forms show the office email and phone number instead of sending.
