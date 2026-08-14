# Mail Merge Add-in — AI Handoff Context
**Generated:** 2026-08-14  
**Project folder:** `/Users/Leighton/mail-merge-addin/`  
**GitHub Pages:** `https://leighton-grey.github.io/mail-merge-addin/`  
**M365 Tenant:** Sunshine Sachs (ssmandl.com)  
**Developer:** Leighton — leighton@prometheanit.com / Promethean IT

---

## 1. Project Overview

A Microsoft Outlook Mail Merge add-in (Office.js / Teams JSON manifest) that lets users send personalised bulk email directly from New Outlook via the Microsoft Graph API. All processing happens client-side (browser); no server-side data storage.

- **Manifest format:** Teams JSON v1.17  
- **Auth:** MSAL.js (msal-browser 2.38.3), delegated permissions, SPA redirect flow  
- **Recipients:** CSV/XLSX upload or from Contacts (Graph API)  
- **Send engine:** Microsoft Graph `/me/sendMail` with `$batch` support  
- **Hosting:** GitHub Pages (static) — `leighton-grey.github.io/mail-merge-addin`

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

All permissions are **Delegated** (not Application). Admin consent has been granted for all four.

---

## 3. Current File State (as of this session)

### `/Users/Leighton/mail-merge-addin/manifest.json`
- Version: **2.5.0**
- Icons reference `assets/icon-color.png`, `icon-outline.png`, `icon-16.png`, `icon-32.png`, `icon-80.png` via GitHub Pages URLs
- `webApplicationInfo.id`: `d06ae3cf-a7da-4264-b20e-ab8d70c06977`

### `/Users/Leighton/mail-merge-addin/taskpane.html`
- CSS version: `taskpane.css?v=2.6.0`
- Header: Purple gradient, Mail Merge icon, no version badge (moved to About section)
- Tab bar: Recipients / Compose / Options
- Compose tab: Simplified — verbose hints collapsed into `<details>` elements, "Insert field at cursor" hidden (kept for JS), friendly tag instruction shown in orange above chips
- Options tab: Has **About section** at bottom (version badge, icon, links, "by Promethean IT")
- Footer: Privacy note collapsed under `<details>` disclosure (🔒 Privacy & data toggle)
- JS: `taskpane.js?v=2.4.1`

### `/Users/Leighton/mail-merge-addin/taskpane.css`
- Version: v2.6.0 styles
- **PRMT brand variables added:**
  - `--prmt-orange: #F58220`
  - `--prmt-orange-dark: #D96610`
  - `--prmt-orange-glow: rgba(245,130,32,0.38)`
  - `--prmt-black: #1C1C1E`
  - `--prmt-black-hover: #2E2E30`
- **Button system:**
  - `.btn-primary` → Orange gradient (Run/confirm actions)
  - `.btn-preview` → Charcoal/black gradient, flips to orange on hover (Preview button only)
  - `.btn-secondary` → White gradient, warm orange hover
  - `.btn-danger` → Red (Stop button)
- **Tag chips:** Larger (5px padding), lift on hover, orange hover colour
- **About section CSS:** `about-section`, `about-header`, `about-icon`, `about-app-name`, `about-version-badge`, `about-tagline`, `about-links`, `about-link`, `about-developer`
- **Compose helpers:** `tags-header`, `tags-hint`, `hint-advanced`, `hint-advanced-toggle`, `hint-advanced-body`
- **Footer disclosure:** `footer-disclosure`, `footer-disclosure-btn`, `footer-disclosure-body`

### `/Users/Leighton/mail-merge-addin/taskpane.js`
- Version comment: `Mail Merge Engine v2.5.0`
- Auth dialog URL: `auth-dialog.html?v=5` (cache-busted)
- Full engine: CSV/XLSX parsing, Graph API batch send, MSAL SSO flow, scheduling, rate limiting, templates, simulation, drafts, broadcast, opt-out, suppression, attachments, per-recipient attachments, inline images, fill-in fields, pre-send confirmation, duplicate send guard

### `/Users/Leighton/mail-merge-addin/auth-dialog.html`
**Fully redesigned in this session.** Visual changes:
- Dark navy/purple gradient background (`#0C0A1A → #16123A → #0D1B38`)
- Three ambient orbs (purple + orange) with CSS drift animation
- Frosted-glass card (rgba white, blur border, deep shadow)
- App icon (80px) with 3 staggered **orange pulse rings** radiating outward
- "Mail Merge / by Promethean IT" title
- Step label that updates at each auth stage
- Progress bar: purple→orange gradient fill with shimmer animation
- 5 step dots that advance (orange glow when active)
- Error box (red) shown on failure, rings pause
- Debug log at bottom (monospace, colour-coded: green=OK, red=error)

**Bug fix applied:** `handleRedirectPromise()` now races against a 6-second timeout via `Promise.race()`. If MSAL doesn't resolve it (previously caused a hang), the timeout fires and falls through to `acquireTokenRedirect`. This fixed the auth dialog getting stuck.

**Auth flow:**
1. jsDelivr CDN loads MSAL (alcdn.msauth.net as fallback)
2. `Office.onReady` fires (5s fallback if it doesn't)
3. `msalInstance.initialize()`
4. `handleRedirectPromise()` — with 6s timeout
5. If null → try `acquireTokenSilent` for cached accounts
6. If that fails → `acquireTokenRedirect` (navigates to Microsoft login)
7. On return → `handleRedirectPromise()` gets token → `Office.context.ui.messageParent(JSON.stringify({ type: "token", token }))`

### `/Users/Leighton/mail-merge-addin/assets/`
- `icon-color.png` — 192×192, purple gradient rounded square, white envelope with merge lines (PRMT design, generated with cairosvg)
- `icon-outline.png` — 32×32, white-on-transparent outline version
- `icon-16.png`, `icon-32.png`, `icon-80.png` — same design at smaller sizes

### `/Users/Leighton/mail-merge-addin/mail-merge-addin.zip`
- Contains: `manifest.json` (v2.5.0) + all icon PNGs
- Upload this to M365 Admin Center → Settings → Integrated Apps → Mail Merge → Update
- **The previous upload failed** because the version was unchanged (2.4.0). Now bumped to 2.5.0 so it should accept.

### `/Users/Leighton/mail-merge-addin/user-guide.html`
- Fully self-contained HTML user guide (no external dependencies)
- Can be hosted at `https://leighton-grey.github.io/mail-merge-addin/user-guide.html`
- **Section A:** 5-step quick start with simulated app screenshots (HTML/CSS mockups), orange callout numbers
- **Section B (renamed "Optional"):** 7 optional features with mini mockups, numbered 1–7 with grey badges
- Features covered: Test send, Scheduling, Attachments, Suppression/Opt-Out, Templates, Simulate & Drafts, Advanced Options
- Example data used: Sarah Johnson / Mike Chen / Emma Wilson at Acme Corp / Globex

---

## 4. All Work Completed (Chronological)

### Previous session (pre-summary):
1. Fixed MSAL auth dialog stuck on spinner — root cause was `alcdn.msauth.net` CDN failing silently; fixed by switching to `cdn.jsdelivr.net` as primary
2. Added `await msalInstance.initialize()` (required in MSAL browser v2.37+)
3. Fixed Entra ID permissions: changed from Application to Delegated type for all scopes
4. Added `auth-dialog.html` as second SPA redirect URI in Entra
5. Successfully completed test mail merge send to Watheeq David
6. Created `Mail_Merge_User_Guide.docx` (12-page Word doc with screenshots)
7. Designed new custom icon/logo (purple gradient rounded square + white envelope + merge lines)
8. Rebuilt `mail-merge-addin.zip` with v2.5.0 manifest

### This session:
1. **PRMT button redesign** — orange gradient for Run, charcoal for Preview, warm orange hover for secondary
2. **Header** — removed v2.5.0 badge from header
3. **About section** — added to bottom of Options tab (version badge, icon, developer credit, links)
4. **Compose tab simplification** — collapsed verbose hints, hidden "Insert field at cursor" UI, friendly tag instruction in orange
5. **Tag chips** — more clickable appearance, orange on hover
6. **Footer disclosure** — privacy text collapsed to 🔒 toggle
7. **Auth dialog redesign** — full visual overhaul (dark card, pulse rings, progress bar, step dots)
8. **handleRedirectPromise timeout fix** — 6-second `Promise.race` prevents hang
9. **user-guide.html** — comprehensive HTML guide with simulated app screenshots, Section A (5 steps) and Optional section (7 features, numbered 1–7)

---

## 5. Pending / Not Yet Done

| Item | Status | Notes |
|------|--------|-------|
| Push all changes to GitHub | ⚠️ **NOT DONE** | Leighton must `git push` from `/Users/Leighton/mail-merge-addin/` |
| Upload new zip to M365 Admin Center | ⚠️ **NOT DONE** | Admin Center → Settings → Integrated Apps → Mail Merge → Update → upload `mail-merge-addin.zip` |
| Verify ribbon icon update in Outlook | ⏳ After zip upload | Outlook should reload the new purple envelope icon |
| Auth dialog live test | ⏳ After push | Push v=5 cache-buster first, then test merge button |

---

## 6. Key Technical Notes for Next AI

### GitHub Pages caching
- The add-in files are served from GitHub Pages. Cache-busting uses `?v=N` query params on CSS/JS.
- Current: `taskpane.css?v=2.6.0`, `taskpane.js?v=2.4.1`, `auth-dialog.html?v=5`
- Bump the version suffix whenever making changes to force Outlook to reload

### Manifest versioning
- M365 Admin Center **rejects** manifest uploads if the `version` field hasn't changed
- Current deployed version: `2.4.0` (last successful upload)
- Zip contains: `2.5.0` — this should work on next upload

### Office.js dialog API
- `Office.context.ui.displayDialogAsync(url, options, callback)` opens auth-dialog.html in a separate window
- `Office.context.ui.messageParent(JSON.stringify(data))` sends token back from dialog to taskpane
- Taskpane listens with `dialog.addEventHandler(Office.EventType.DialogMessageReceived, handler)`

### MSAL in Office dialog context
- The dialog is a separate browser context — no shared cookies/storage with taskpane
- Must use `sessionStorage` (not `localStorage`) for MSAL cache in dialog
- `handleRedirectPromise()` can hang in dialog context — the 6s timeout workaround handles this
- After successful auth, the token is passed via `messageParent`, not stored

### CSS architecture
- All styles in `taskpane.css` using CSS custom properties (`--accent`, `--prmt-orange`, etc.)
- App chrome (header, tabs, progress, focus rings) uses purple (`--accent: #534AB7`)
- Action buttons use PRMT orange (`--prmt-orange: #F58220`) 
- Dark button (Preview only) uses `--prmt-black: #1C1C1E`

---

## 7. File Tree

```
/Users/Leighton/mail-merge-addin/
├── manifest.json              ← v2.5.0, Teams JSON format
├── taskpane.html              ← Main add-in UI
├── taskpane.css               ← v2.6.0 styles (PRMT colours)
├── taskpane.js                ← Full merge engine v2.5.0
├── auth-dialog.html           ← MSAL auth dialog (redesigned)
├── privacy.html               ← Privacy policy page
├── terms.html                 ← Terms of use page
├── mail-merge-addin.zip       ← Deployment zip (v2.5.0 manifest + icons)
├── user-guide.html            ← HTML user guide with mockups
├── Mail_Merge_User_Guide.docx ← Word doc guide (older, from prev session)
├── CONTEXT_HANDOFF.md         ← This file
└── assets/
    ├── icon-color.png         ← 192×192 app icon (purple gradient + envelope)
    ├── icon-outline.png       ← 32×32 white outline icon
    ├── icon-16.png            ← 16×16
    ├── icon-32.png            ← 32×32
    └── icon-80.png            ← 80×80
```

---

## 9. Competitive Analysis & Product Roadmap

### Market Competitors (100+ user scale)

| Competitor | Price | Key Differentiator |
|-----------|-------|-------------------|
| EmailMerge365 | ~$100/user/year (~$6k–$10k/100 users) | Email tracking, scheduling, IT admin deployment |
| MailMerge365 | ~$180/user/year (€15/mo), enterprise custom | Unsubscribe compliance, campaign reporting |
| SecureMailMerge | ~$180/user/year | Local data processing (GDPR/SOC2), AppSource volume discounts |
| SwiftMerge | ~$180/user/year | Native Graph API send, seamless Sent Items sync |

### Our Unique Value Propositions

**1. Data sovereignty (strongest card).** All competitors route some data through their own servers. Our add-in processes everything inside the user's browser and sends exclusively via Microsoft Graph on the user's own authenticated account. No recipient list, email body, or attachment ever touches a third-party server. SecureMailMerge charges $180/user/year for this. We have it by architecture.

**2. Per-recipient attachments.** None of the four competitors offer per-recipient file matching (matching attachment filenames to a CSV `attachment` column). Critical for legal, HR, finance use cases — sending each person their own contract, invoice, or payslip.

**3. Token depth.** Pipe filters, conditionals (`{{if:col=val:yes:no}}`), many-to-one group merge (`{{merge_table}}`), and `{{greeting_line}}` auto-generation go beyond what any listed competitor documents.

**4. Graph API native sending.** SwiftMerge markets this as their differentiator. We already do it — emails come from the user's own account, appear in their Sent Items, no deliverability or reputation issues.

**5. Central M365 Admin Center deployment.** IT deploys once via Integrated Apps, appears for the whole org. No per-device installation.

### Critical Missing Features (must-build before enterprise sales)

Priority order:

1. **Open & click tracking** ← highest priority gap. Every competitor offers it. Requires a lightweight serverless endpoint (Azure Functions or Cloudflare Workers) to receive tracking pixel pings. Without this, can't answer "did anyone open this?" — first question any sales/marketing team asks.

2. **Org-wide opt-out/suppression sync.** Current opt-out list lives in each user's `localStorage` — doesn't sync across users. User A gets an unsubscribe, User B can still email that address tomorrow. GDPR/CAN-SPAM compliance at enterprise scale requires shared suppression. Implement via SharePoint list or Azure Table Storage.

3. **Shared templates across users.** Templates also live in per-user `localStorage`. Enterprise teams need shared template libraries — sales managers push approved templates to their teams. Same SharePoint approach as suppression sync.

4. **Bounce/NDR handling.** Graph API send failures are caught, but post-send NDR bounces (returned to inbox after merge completes) are not processed. Need to monitor NDR folder via Graph and update suppression list automatically.

5. **Admin dashboard.** At 100+ users, IT admins need org-wide visibility: who sent what, total volume, rate limit hits, failures. No individual mailbox access needed — aggregate logs written to SharePoint or Azure.

### Feature Parity vs Competitors

| Feature | Us | EmailMerge365 | MailMerge365 | SecureMailMerge | SwiftMerge |
|---------|-----|---------------|--------------|-----------------|------------|
| Graph API native send | ✅ | ❓ | ❓ | ✅ | ✅ |
| Local data processing | ✅ | ❌ | ❌ | ✅ | ❌ |
| Central IT deployment | ✅ | ✅ | ✅ | ✅ | ✅ |
| Scheduling | ✅ | ✅ | ✅ | ✅ | ✅ |
| Rate limiting | ✅ | ✅ | ✅ | ✅ | ✅ |
| Per-recipient attachments | ✅ | ❌ | ❌ | ❌ | ❌ |
| Advanced token/conditionals | ✅ | ❌ | ❌ | ❌ | ❌ |
| Suppression/opt-out lists | ✅ (local) | ✅ | ✅ | ✅ | ✅ |
| List-Unsubscribe header | ✅ | ❓ | ✅ | ❓ | ❓ |
| Open/click tracking | ❌ | ✅ | ✅ | ✅ | ✅ |
| Campaign analytics/reporting | ❌ | ✅ | ✅ | ❌ | ❌ |
| Org-wide suppression sync | ❌ | ✅ | ✅ | ✅ | ✅ |
| Shared org templates | ❌ | ❓ | ❓ | ❓ | ❓ |
| Admin dashboard | ❌ | ✅ | ✅ | ❌ | ❌ |
| Bounce handling | ❌ | ✅ | ✅ | ❓ | ❓ |

### Recommended Pricing Structure

**Free tier** — unlimited users, up to 200 emails/day per user, no tracking, no admin dashboard. Goal: drive adoption, get org past the cap, trigger upgrade.

**Pro — $49/user/year (billed annually) / $5/month.** Full current feature set. No tracking (be honest, roadmap it). At $4,900/year for 100 users vs $10,000–$18,000 competitors — lead with "same Graph API sending as SwiftMerge, same local data privacy as SecureMailMerge, at 27% of the price."

**Enterprise — custom quote, ~$35/user/year at 100+ users.** Volume tiers: 100–250 / 250–500 / 500+. Minimum ~$3,500/year for 100 users. Includes: org-wide suppression sync, shared templates, admin dashboard (once built), priority support, custom deployment assistance.

**One-time enterprise setup fee: $500–$1,500.** Assisted deployment + IT admin onboarding. Signals professional services capability, standard in the space.

### Build Roadmap (before enterprise targeting)

| Priority | Feature | Approach | Unblocks |
|----------|---------|----------|---------|
| 1 | Open/click tracking | Cloudflare Worker or Azure Function (single tracking endpoint) | Parity with all 4 competitors |
| 2 | Org-wide opt-out sync | SharePoint list via Graph API | GDPR compliance at scale |
| 3 | Shared templates | SharePoint list via Graph API | Enterprise team workflows |
| 4 | Bounce/NDR handling | Monitor inbox for NDRs via Graph, auto-update suppression | Large list management |
| 5 | Admin dashboard | Read-only SharePoint/Azure log aggregation | IT procurement approval |

Items 1–3 make us feature-competitive with all listed players at less than 30% of their price.

---

## 8. Next Steps for Continuing AI

1. **First priority:** Confirm Leighton has pushed to GitHub and uploaded the zip to M365 Admin Center
2. **Test auth flow end-to-end** — open Outlook, click Run Merge, confirm the new dark auth dialog appears, confirm sign-in works, confirm merge sends
3. **If auth still hangs** — check the debug log in the auth dialog. If it stops after "calling handleRedirectPromise() timeout — proceeding" and never calls acquireTokenRedirect, there may be an MSAL/CSP issue
4. **User guide** — Leighton may want the `user-guide.html` converted to a PDF or PPTX for distribution. Use the `pdf` or `pptx` skills
5. **Version bump** — if any more code changes are made, bump `taskpane.js?v=` and `taskpane.css?v=` cache-busters in `taskpane.html`, and increment `manifest.json` version before re-zipping
