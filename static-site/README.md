# Forms and search on the static site

askib.in is a static copy of the website, so it has no server to receive forms
or run searches. The files here add both back:

- `site.js` swaps each "Forms are switched off" notice for a working form and
  sends it to our own Google Apps Script (`form-handler.gs`). The script emails
  the enquiry to info@askib.in and saves it as a row in a Google Sheet. It is
  free (a Google account can send 100 emails a day, Workspace 1,500).
  Spam is filtered with a hidden honeypot field, a minimum fill time and a limit
  of 5 enquiries an hour per email address.
- `site.css` styles the forms and search results with the site's own colours.
- `search-index.json` and `products.json` are generated; do not edit them.
- `build.mjs` adds the two files to every page, regenerates the index and adds
  the `/search` page. Run it after every new export: `node static-site/build.mjs`.

## Connecting the forms (about 5 minutes, once)

1. Signed in to the Google account that should own the enquiries (ideally the
   one for info@askib.in), open https://sheets.new to make a new Google Sheet.
   Name it "askib.in enquiries".
2. In the sheet: **Extensions → Apps Script**. Delete what is in the editor,
   paste in the whole of `form-handler.gs`, and press the save icon.
3. **Deploy → New deployment**. Click the gear next to "Select type" and choose
   **Web app**. Set "Execute as" to **Me** and "Who has access" to **Anyone**.
   Press **Deploy**, then **Authorize access** and allow it (Google warns that
   the app is unverified because you wrote it: Advanced → Go to … → Allow).
4. Copy the **Web app URL** (it ends in `/exec`) and paste it into
   `FORM_ENDPOINT` near the top of `site.js`, then push.

Opening that URL in a browser should say "askib.in form handler is running."
Until it is pasted, the forms show the office email and phone number instead.

If you later edit the script, use **Deploy → Manage deployments → Edit →
New version** so the same URL keeps working.
