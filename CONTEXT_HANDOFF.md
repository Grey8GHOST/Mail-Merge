# Mail Merge Add-in — AI Session Handover
**Last updated:** 2026-08-18  
**Project folder (local):** `/Users/Leighton/ASD/`  
**GitHub Pages:** `https://leighton-grey.github.io/mail-merge-addin/`  
**Developer:** Leighton — leighton@prometheanit.com / Promethean IT  
**M365 Tenant:** Sunshine Sachs (ssmandl.com)

---

## CRITICAL RULES FOR NEXT SESSION

1. **All deliverables MUST be saved to `/Users/Leighton/ASD/` automatically** — do not wait to be asked. Every file created (docx, zip, commented code) goes to ASD in the same step as SendUserFile.
2. **The project folder is `/Users/Leighton/ASD/`** — not `/Users/Leighton/mail-merge-addin/`. That old path is no longer used.
3. **LICENSE_ENFORCEMENT = false** — never change this without explicit instruction from Leighton. The add-in must continue working via Admin Centre and EAC.

---

## 1. Project Overview

A Microsoft Outlook Mail Merge add-in (Office.js) that lets users send personalised bulk emails directly from Outlook via Microsoft Graph API. All email processing (template merging, sending) is 100% client-side. No recipient data or email content ever touches a Promethean IT server.

### Architecture layers
- **Add-in (client-side):** taskpane.html / taskpane.js / taskpane.css — the full UI and merge engine
- **Auth dialog:** auth-dialog.html — MSAL OAuth flow in a separate Office dialog window
- **Licensing layer (server-side, Azure Functions):** handles AppSource subscriptions only. Email data never touches it.

---

## 2. Entra ID / Azure App Registration

| Field | Value |
|-------|-------|
| Client ID | `d06ae3cf-a7da-4264-b20e-ab8d70c06977` |
| Tenant ID | `04ca7ab1-691d-45a2-928f-574b3e3300a0` |
| Authority | `https://login.microsoftonline.com/04ca7ab1-691d-45a2-928f-574b3e3300a0` |
| Redirect URI (SPA) | `https://leighton-grey.github.io/mail-merge-addin/taskpane.html` |
| Redirect URI (SPA) | `https://leighton-grey.github.io/mail-merge-addin/auth-dialog.html` |
| Scopes (Delegated) | `Mail.Send`, `User.Read`, `Contacts.Read`, `People.Read.All` |

All permissions are **Delegated** (not Application). Admin consent granted for all four scopes.

---

## 3. File Inventory — `/Users/Leighton/ASD/`

### Core add-in files
| File | Description |
|------|-------------|
| `taskpane.html` | Main add-in UI — tabs: Recipients / Compose / Options |
| `taskpane.css` | All styles, PRMT brand colours, v2.6.0 |
| `taskpane.js` | Full merge engine v2.5.0 — CSV, Graph API, MSAL, scheduling, rate limiting, per-recipient attachments, license check |
| `auth-dialog.html` | MSAL OAuth dialog — dark glass card, pulse rings, progress bar, 6s handleRedirectPromise timeout fix |
| `privacy.html` | Live privacy policy page (referenced by manifest.json) |
| `terms.html` | Live terms of use page (referenced by manifest.json) |
| `user-guide.html` | Self-contained HTML user guide with simulated app screenshots |

### Manifests
| File | Description |
|------|-------------|
| `manifest.json` | Teams JSON v1.17 — primary manifest for Admin Centre deployment |
| `manifest.xml` | XML manifest — for Exchange/EAC deployment |
| `manifest-personal.xml` | XML manifest — personal/sideloaded variant |

All manifests point to `https://leighton-grey.github.io/mail-merge-addin/` for all URLs.

### Licensing project
| File | Description |
|------|-------------|
| `mail-merge-licensing.zip` | Complete Azure Functions project — fully commented source code |

### Documentation
| File | Description |
|------|-------------|
| `MailMerge-LicensingGuide.docx` | Architecture + Azure setup + AppSource submission + step-by-step Azure SWA and AWS S3+CloudFront deployment guides + AppSource readiness checklist |
| `MailMerge-Learnings.docx` | Senior-to-junior code walkthrough, component breakdown, 3 presentation highlights, Q&A prep |
| `MailMerge-CodeExplained.docx` | Every licensing file rendered with colour-coded comment/code blocks. **Partially complete** — covers licensing project only. Core add-in files (taskpane.js, taskpane.html, taskpane.css) NOT yet included. See Pending Tasks. |
| `Mail_Merge_User_Guide.docx` | Word doc user guide (older, from an earlier session) |
| `README.md` | Basic project readme |
| `CONTEXT_HANDOFF.md` | This file |

### Deployment zips
| File | Description |
|------|-------------|
| `mail-merge-addin.zip` | manifest.json v2.5.0 + all icon PNGs — upload to M365 Admin Centre |
| `mail-merge-dev.zip` | Dev/test variant |

### Assets
```
assets/
  icon-color.png    192×192 — purple gradient rounded square, white envelope + merge lines
  icon-outline.png  32×32 — white-on-transparent outline
  icon-16.png       16×16
  icon-32.png       32×32
  icon-80.png       80×80
```

---

## 4. taskpane.js — Key Constants (top of file)

```javascript
// ── License enforcement (currently OFF) ────────────────────────────────────
const LICENSE_ENFORCEMENT      = false;   // ← flip to true to gate the product
const LICENSE_API_URL          = "https://YOUR_FUNCTION_APP.azurewebsites.net/api/license";
const LICENSE_CACHE_KEY        = "mailmerge_license_token";
const LICENSE_CACHE_EXPIRY_KEY = "mailmerge_license_expiry";
```

**LICENSE_ENFORCEMENT = false** means:
- `checkLicense()` runs silently on every startup and logs the result
- No features are blocked, no overlay is shown
- Users see zero change from current behaviour
- Admin Centre and EAC deployments are completely unaffected

To enable licensing: set `LICENSE_ENFORCEMENT = true`, replace `LICENSE_API_URL` with the deployed Azure Functions URL, push to GitHub.

### Cache-busting versions in taskpane.html
```
taskpane.css?v=2.6.0
taskpane.js?v=2.4.1
auth-dialog.html?v=5
```
Bump the `v=` suffix whenever making changes so Outlook reloads the file.

### Manifest version
Current: `2.5.0` in manifest.json. M365 Admin Centre rejects manifest uploads if version is unchanged. Always increment before re-zipping.

---

## 5. Licensing Layer — Azure Functions Project

### What it is
A Node.js Azure Functions v4 project (`mail-merge-licensing/`) that handles AppSource subscription licensing. Lives inside `mail-merge-licensing.zip` in ASD.

### Source files inside the zip
```
mail-merge-licensing/
  package.json                      Project config (no node-fetch — uses native Node 18 fetch)
  host.json                         Azure Functions runtime config (routePrefix: "api")
  local.settings.json.example       Environment variables template
  src/
    shared/
      jwtHelper.js                  Issues/verifies 24-hour license JWT tokens
      tableStorage.js               ALL Azure Table Storage operations (subscriptions, users)
      fulfillmentClient.js          Microsoft SaaS Fulfillment API v2 client
    functions/
      licenseCheck.js               GET /api/license — bouncer, called by add-in on startup
      activate.js                   GET /api/activate — AppSource landing page after purchase
      saasWebhook.js                POST /api/webhook — all Microsoft subscription lifecycle events
```

### Environment variables required (Azure Function App settings)
```
AZURE_STORAGE_CONNECTION_STRING    Connection string from Storage Account → Access keys
MARKETPLACE_CLIENT_ID              Entra app client ID for Fulfillment API OAuth
MARKETPLACE_CLIENT_SECRET          Entra app client secret
MARKETPLACE_TENANT_ID              Your Azure tenant ID
LICENSE_JWT_SECRET                 Long random string — signs the 24h license tokens
ACTIVATION_SUCCESS_URL             Optional redirect after activation (can be blank)
ALLOWED_ORIGINS                    Comma-separated allowed origins for CORS
                                   e.g. https://leighton-grey.github.io,https://grey8ghost.github.io
```

### Bugs fixed in this session
1. **Critical:** `getSubscription("unknown", subscriptionId)` in saasWebhook.js — webhook events for Unsubscribed/Suspended/Reinstated/ChangePlan/ChangeQuantity used "unknown" as the tenantId PartitionKey, which never matched any real record. Cancelled subscriptions stayed "Active" forever. **Fixed:** added `getSubscriptionById(subscriptionId)` which does a cross-partition scan by RowKey alone.
2. **node-fetch@3 incompatible** — ESM-only, can't be `require()`'d in CommonJS. **Fixed:** removed it; code already used native Node 18 `fetch`.
3. **Unused import** — `TableServiceClient` imported but never used in tableStorage.js. **Fixed:** removed.
4. **OData filter injection** — raw string interpolation in filter queries. **Fixed:** switched to `odata` template tag from @azure/data-tables SDK throughout.
5. **No TableClient caching** — new client created on every storage call. **Fixed:** added `const _clients = new Map()` module-level cache.
6. **Dead code** — `idToken` variable declared and assigned null twice, never used in checkLicense(). **Fixed:** removed.
7. **Missing webhook validation** — anyone could POST fake events. **Fixed:** added `validateWebhookAuth()` checking JWT issuer. TODO: upgrade to full JWKS signature verification before AppSource production go-live.

### How the license check works
1. Add-in calls `GET /api/license?userId=...&tenantId=...&email=...` on startup
2. licenseCheck.js looks up tenant in Azure Table Storage
3. If Active subscription found: returns `{ licensed: true, plan: "pro", token: "eyJ...", expiresAt: "..." }`
4. Add-in caches the JWT in localStorage for 24 hours
5. On next startup within 24h: uses cached token without calling the API
6. If `LICENSE_ENFORCEMENT = true` and not licensed: `showLicenseGate()` shows non-dismissible overlay

---

## 6. Unique Product Differentiators

These features are confirmed unique — no competitor has all of them:

1. **Inline CSV editing** — recipients edited live in a table inside the taskpane. ALL competitors (SecureMailMerge, MailMerge365, EmailMerge365, Mail Merge Toolkit) require editing in an external spreadsheet.
2. **Type-emails mode** — recipients entered as plain email addresses, no spreadsheet at all
3. **Import from Outlook To field** — addresses already in compose window pulled into merge list
4. **Per-recipient attachments** — CSV `attachment` column matched to individual files. No competitor documents this feature.
5. **Advanced merge tokens** — pipe filters, conditionals `{{if:col=val:yes:no}}`, group merge `{{merge_table}}`, `{{greeting_line}}` auto-generation
6. **Full data privacy** — zero server-side processing of email or recipient data. Everything through Microsoft Graph on the user's own account.

### Competitor pricing context
| Competitor | Price |
|-----------|-------|
| EmailMerge365 | ~$100/user/year |
| MailMerge365 | ~$180/user/year |
| SecureMailMerge | ~$180/user/year |
| SwiftMerge | ~$180/user/year |

**Recommended pricing:** Free tier (200 emails/day) → Pro $49/user/year → Enterprise ~$35/user/year at 100+ users.

---

## 7. Hosting Options Summary

The add-in's static files can be hosted anywhere serving HTTPS. The Azure Functions licensing layer stays on Azure regardless.

| Option | Cost | SLA | Notes |
|--------|------|-----|-------|
| GitHub Pages | Free | None | Current hosting. Not recommended for commercial product. |
| Azure Static Web Apps | Free tier | 99.95% | **Recommended.** GitHub auto-deploy, custom domain, Microsoft ecosystem. |
| AWS S3 + CloudFront | ~$1-5/mo | 99.99% | Production-grade. Good if already in AWS. |
| Vercel / Netlify | Free tier | 99.99% | Simple GitHub integration. Fully compatible. |

Full step-by-step instructions for both Azure SWA and AWS S3+CloudFront are in `MailMerge-LicensingGuide.docx` sections 10 and 11.

### What changes when switching hosting
- All manifest URLs (3 files) — find/replace old domain with new domain
- `ALLOWED_ORIGINS` in Azure Function App environment variables — add new domain
- Entra App Registration — add new domain as redirect URI
- `LICENSE_API_URL` in taskpane.js — stays pointing to Azure Functions (does NOT move to AWS)

---

## 8. AppSource Submission — Readiness Status

### Code: COMPLETE (with one TODO)
All Azure Functions code is production-ready. LICENSE_ENFORCEMENT = false keeps everything inert.

### Outstanding before AppSource go-live
| Item | Status |
|------|--------|
| `privacy.html` | ✅ EXISTS in ASD — live at GitHub Pages URL |
| `terms.html` | ✅ EXISTS in ASD — live at GitHub Pages URL |
| AppSource icons (icon-color.png, icon-outline.png) | ✅ Exist in assets/ |
| Webhook JWKS full signature verification | ⚠️ TODO — currently checks issuer only. Upgrade before production. |
| `LICENSE_API_URL` — replace placeholder | ⚠️ TODO — set to actual Azure Functions URL when deploying |
| `LICENSE_ENFORCEMENT = true` | ⚠️ Only when SaaS offer is live and tested |
| `ALLOWED_ORIGINS` set to production domain | ⚠️ TODO when hosting migrated |
| Partner Center SaaS offer | ⚠️ NOT STARTED — needs Partner Center account |
| AppSource add-in listing | ⚠️ NOT STARTED |

### AppSource Partner Center setup (when ready)
- Landing page URL: `https://YOUR_FUNCTION_APP.azurewebsites.net/api/activate`
- Webhook URL: `https://YOUR_FUNCTION_APP.azurewebsites.net/api/webhook`
- Entra app Client ID: `d06ae3cf-a7da-4264-b20e-ab8d70c06977`
- Microsoft takes ~3% of revenue

---

## 9. Product Roadmap (Missing Features for Enterprise)

| Priority | Feature | Why | Approach |
|----------|---------|-----|----------|
| 1 | Open/click tracking | Every competitor has it. #1 question from sales/marketing teams. | Azure Function or Cloudflare Worker receives tracking pixel pings |
| 2 | Org-wide opt-out sync | Current suppression list is per-user localStorage. GDPR requires shared suppression. | SharePoint list via Graph API |
| 3 | Shared templates | Templates in per-user localStorage. Enterprise teams need shared library. | SharePoint list via Graph API |
| 4 | Bounce/NDR handling | Post-send NDRs not processed. | Monitor inbox for NDRs via Graph, auto-update suppression list |
| 5 | Admin dashboard | Aggregate usage logs for IT admins. | SharePoint/Azure log writes, read-only dashboard |

---

## 10. Build Scripts (in cloud workspace `/home/claude/`)

These JS scripts build the Word documents. Run with `node scriptname.js`.

| Script | Output |
|--------|--------|
| `build-licensing-doc.js` | `MailMerge-LicensingGuide.docx` — sections 1-12 including step-by-step deployment guides |
| `build-code-explained.js` | `MailMerge-CodeExplained.docx` — licensing project files only (taskpane files NOT yet added) |

**Note:** These scripts exist in the Claude cloud workspace, not in ASD. They are lost when the session ends. The built `.docx` outputs ARE in ASD.

---

## 11. Pending / Incomplete Tasks

### High priority — interrupted this session
1. **taskpane.js commenting** — the agent was interrupted mid-run (credits ran low). `taskpane-commented.html` and `taskpane-commented.css` ARE complete at `/home/claude/` in the cloud workspace. `taskpane-commented.js` was NOT completed. On resume: spawn an agent to read `/mnt/user-data/uploads/mail-merge-addin/taskpane.js` (5765 lines), add line-by-line comments, write to `/home/claude/taskpane-commented.js`.

2. **Expand MailMerge-CodeExplained.docx** — currently only covers the 9 licensing project files. Needs to be expanded to include taskpane.html, taskpane.css, and taskpane.js (with comments). Update `build-code-explained.js` to add these three files to the FILES manifest, rebuild, deliver via SendUserFile, commit to ASD.

3. **Commit commented source files to ASD** — once taskpane-commented.js is done, all three commented files (taskpane-commented.html, taskpane-commented.css, taskpane-commented.js) should be sent to user and committed to ASD.

### Medium priority
4. **Azure Functions deployment** — Leighton has not yet deployed the Azure Functions project. Steps are in MailMerge-LicensingGuide.docx section 5.
5. **GitHub push** — changes to taskpane.js (license check code added), manifests may not all be pushed to GitHub yet. Leighton needs to git push.
6. **M365 Admin Centre zip upload** — confirm the v2.5.0 manifest zip has been uploaded.

---

## 12. Auth Dialog Technical Notes

The `auth-dialog.html` is a separate browser context from the taskpane — no shared cookies or storage.

- Uses `sessionStorage` (not `localStorage`) for MSAL cache
- `handleRedirectPromise()` can hang in dialog context — **fixed with 6-second `Promise.race()` timeout**
- Token passed back to taskpane via `Office.context.ui.messageParent(JSON.stringify({ type: "token", token }))`
- Taskpane listens with `dialog.addEventHandler(Office.EventType.DialogMessageReceived, handler)`

Auth flow:
1. MSAL loads from jsDelivr CDN (alcdn.msauth.net as fallback)
2. `Office.onReady` fires (5s fallback)
3. `msalInstance.initialize()`
4. `handleRedirectPromise()` with 6s timeout
5. If null → try `acquireTokenSilent` for cached accounts
6. If that fails → `acquireTokenRedirect` (navigates to Microsoft login)
7. On return → `handleRedirectPromise()` gets token → `messageParent` sends it to taskpane

---

## 13. CSS Architecture

All styles in `taskpane.css` using CSS custom properties:
- App chrome (header, tabs, focus rings): `--accent: #534AB7` (purple)
- Action buttons: `--prmt-orange: #F58220` (orange)
- Preview button: `--prmt-black: #1C1C1E` (dark charcoal)
- Button system: `.btn-primary` (orange), `.btn-preview` (charcoal→orange hover), `.btn-secondary` (white→orange hover), `.btn-danger` (red)

---

## 14. What a New Session Should Do First

1. Ask Leighton what he wants to work on
2. If continuing the commenting/CodeExplained work:
   - Check if `/home/claude/taskpane-commented.js` exists (it won't — cloud workspace resets between sessions)
   - Spawn an agent to re-run the taskpane.js commenting task (read from `/mnt/user-data/uploads/mail-merge-addin/taskpane.js`)
   - Then update `build-code-explained.js` to include taskpane.html, taskpane.css, taskpane.js
   - Rebuild MailMerge-CodeExplained.docx and commit to ASD
3. If working on new features: check the roadmap in section 9
4. If deploying: follow the go-live checklist in MailMerge-LicensingGuide.docx section 12

---

*Promethean IT · Mail Merge Add-in · Internal handover document*
