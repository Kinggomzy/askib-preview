# askib.in form handler (Microsoft 365)

`worker.js` receives the website's enquiry forms and, for each one:

- emails it to **info@askib.in**, sent from info@askib.in, with the customer as
  reply-to, so pressing Reply in Outlook answers the customer;
- adds a row to the **Enquiries** sheet of the company's Excel workbook on OneDrive.

On its first enquiry it sets the workbook up: renames it to
**ASKIB Website Enquiries.xlsx**, adds an **Enquiries** sheet (a formatted table:
Received, Reference, Form, Status, Name, Company, Email, Phone, City, Details,
Page, Notes, with the header row frozen) and a **How to use** sheet. Existing
sheets are never changed or deleted.

Everything used is free: Cloudflare Workers free plan (100,000 requests a day)
and the Microsoft 365 you already have.

## Setup (about 15 minutes, once)

### 1. Let the handler use Microsoft 365 (Microsoft Entra)

Sign in as a Microsoft 365 admin.

1. Open https://entra.microsoft.com → **Applications → App registrations → New registration**.
   Name: `askib.in website forms`. Supported account types: **this organizational directory only**.
   Leave Redirect URI empty. Press **Register**.
2. On the app's **Overview**, copy **Application (client) ID** and **Directory (tenant) ID**.
3. **API permissions → Add a permission → Microsoft Graph → Application permissions**.
   Tick **Mail.Send** and **Files.ReadWrite.All**, press **Add permissions**, then
   **Grant admin consent for Askib** and confirm.
4. **Certificates & secrets → New client secret**. Description `website`, expiry
   24 months. Copy the **Value** straight away (it is shown only once).
   Put a reminder in your calendar to make a new one before it expires.

### 2. Run the handler (Cloudflare, free)

1. Sign up at https://dash.cloudflare.com/sign-up (free plan).
2. **Workers & Pages → Create → Create Worker**. Name it `askib-forms`, press **Deploy**.
3. **Edit code**, delete what is there, paste the whole of `worker.js`, press **Deploy**.
4. **Settings → Variables and Secrets → Add**, one by one:

   | Name | Type | Value |
   |---|---|---|
   | `TENANT_ID` | Text | Directory (tenant) ID from step 1.2 |
   | `CLIENT_ID` | Text | Application (client) ID from step 1.2 |
   | `CLIENT_SECRET` | **Secret** | the secret Value from step 1.4 |
   | `MAILBOX` | Text | `info@askib.in` |
   | `EXCEL_SHARE_URL` | Text | the workbook's share link (Share → Copy link in Excel) |
   | `ALLOWED_ORIGINS` | Text | `https://askib.in,https://www.askib.in` |

   Press **Deploy** after adding them.
5. Copy the worker's address from its overview, e.g. `https://askib-forms.<name>.workers.dev`.
   Opening it in a browser should say "askib.in form handler is running."

### 3. Connect the website

Paste the worker's address into `FORM_ENDPOINT` near the top of
`static-site/site.js` and push. Send one test enquiry from askib.in/contact:
it should arrive in the info@askib.in inbox and the workbook gets its new sheets.

## If something goes wrong

Cloudflare → `askib-forms` → **Logs** (Live) shows each request and any
Microsoft error. Common ones:

- `token → 401`: wrong client ID, tenant ID or secret, or the secret expired.
- `sendMail → 403` / `shares → 403`: admin consent was not granted in step 1.3.
- `shares → 404`: the share link is wrong; copy it again from Excel.

If only one of email or Excel fails, the visitor still sees "Thank you" and the
other one still reaches you.

## Tightening access (optional)

Mail.Send as an application permission can send from any mailbox in the
organisation. To limit the app to info@askib.in only, an Exchange admin can add
an application access policy:
https://learn.microsoft.com/graph/auth-limit-mailbox-access

## Testimonials

The same handler takes testimonials from askib.in/testimonial (see
`static-site/README.md` for how to approve them). On the first one it adds a
**Testimonials** sheet (Status drop-down: Pending, Approved, Rejected) and adds
a few lines about it to the **How to use** sheet. Photos go into a
**Testimonial photos** folder next to the workbook. No new permissions are
needed: it uses the same Mail.Send and Files.ReadWrite.All.

It also answers two public addresses the website reads:

- `GET /testimonials`: approved testimonials only (name, company, role, city,
  rating, text, photo link; never email or phone), cached for 5 minutes.
- `GET /testimonials/photo/<Reference>.jpg`: the photo of an approved
  testimonial, cached for a day.

After updating `worker.js` in Cloudflare (Edit code → paste → Deploy), open
`https://<your worker>/testimonials`; it should show `{"ok":true,"items":[]}`.
