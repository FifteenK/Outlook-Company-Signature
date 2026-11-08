# Outlook Company Signature

An event-based Office.js add-in that automatically adds a remotely-managed company sign-off block whenever a user opens a new compose window in Outlook. No task pane, no user interaction — it just works silently in the background.

The block is the same for everyone: company logo, name, address, phone, website, social links, an optional campaign banner, and a legal disclaimer. Each person types their own name above it as they write. That gives a small business one consistent sign-off it controls centrally — swap in a seasonal sale banner, update the address or change the disclaimer, and every email from the team picks it up on the next compose.

If you're the only person sending email, this achieves nothing over pasting a template into your own Outlook signature settings. It earns its keep once there's a team to keep consistent.

Signature content is served from Cloudflare Pages, meaning it can be updated without touching the add-in or re-sideloading the manifest.

-----

## How It Works

1. User opens a new compose window in Outlook
1. Outlook fires the `OnNewMessageCompose` event
1. `commands.js` fetches `signature.html` from Cloudflare Pages
1. The signature is injected via `setSignatureAsync()` (falls back to `body.setAsync()` on older builds)
1. No UI, no button, no task pane

-----

## Repo Structure

```
Add-in (what you deploy)
├── manifest.xml       # Add-in manifest — sideloaded into Outlook
├── commands.html      # Hidden runtime page that loads Office.js and commands.js
├── commands.js        # Event handler that fetches and inserts the sign-off
├── signature.html     # Company sign-off block — edit this to change it for everyone
├── assets/            # Add-in icons, logo and campaign banner
└── _headers           # Cloudflare Pages CORS header config

Everything else (not needed to deploy)
├── docs/              # Project website, published with GitHub Pages
├── skills/            # Notes for extending the POC (per-user signatures, templating, caching)
└── readme.md
```

Only the top group is needed to run the add-in. `docs/` is the project's own website, and you can delete it from your fork.

-----

## Requirements

This add-in relies on Outlook's [event-based activation](https://learn.microsoft.com/office/dev/add-ins/develop/event-based-activation) feature (`OnNewMessageCompose`, Mailbox requirement set 1.10), not on Outlook for Mac specifically. Per Microsoft's [supported-clients table](https://learn.microsoft.com/office/dev/add-ins/develop/event-based-activation#supported-events), that means:

- **Windows** — both the new Outlook for Windows and classic Outlook for Windows support it (classic requires Windows 10 version 1903 / Build 18362 or later, or Windows Server 2019 version 1903+)
- **Mac** — only the **new Outlook for Mac**. Legacy/classic Outlook for Mac does not support event-based activation at all, so the add-in will not fire on it
- **Outlook on the web** also supports it
- A Microsoft 365 account signed into Outlook (event-based activation isn't available on POP/IMAP accounts)
- A [Cloudflare Pages](https://pages.cloudflare.com) account (free tier)
- A GitHub account

**Known gap on classic Outlook for Windows**: `commands.js` currently uses `async`/`await`. Classic Outlook for Windows builds older than Version 2403 (Build 17425.20000) run event handlers in a JS runtime limited to ES2016, which doesn't support `async`/`await` — the handler will time out and the signature won't be injected. Newer classic builds (2403+), the new Outlook for Windows, and Outlook on the web are unaffected. If you need to support older classic Windows builds, rewrite the handler using `.then()`/callback chains instead of `async`/`await` (a version of this exists in this repo's history at commit `a5aaf97`, later reverted).

-----

## Deploy

### 1. Fork or clone this repo

```bash
git clone https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git
cd YOUR_REPO_NAME
```

### 2. Customise the signature

Edit `signature.html` with your company's details. It's a company sign-off block, not a personal signature — nobody's name goes in it:

- **Logo** — replace `assets/logo.png` (shown at 160 × 40, export at 320 × 80 for sharp rendering)
- **Company name, address, phone and website**
- **Social links** — LinkedIn, Instagram and Facebook are included as text links; add or remove as needed
- **Campaign banner** (optional) — replace `assets/banner.png` (shown at 560 × 100, export at 1120 × 200) and its link, or delete that table row
- **Disclaimer** — your registered company details and confidentiality wording

All image and link URLs must be absolute HTTPS URLs.

### 3. Deploy to Cloudflare Pages

1. Log in to [Cloudflare Dashboard](https://dash.cloudflare.com) → **Workers & Pages** → **Create**
1. Connect your GitHub account and select this repository
1. Set build configuration:
- **Build command**: *(leave empty)*
- **Build output directory**: `/`
1. Click **Save and Deploy**

Your files will be live at `https://YOUR_PROJECT_NAME.pages.dev`.

The `_headers` file automatically sets permissive CORS headers on `commands.js` and `signature.html` — no extra configuration needed.

### 4. Update the manifest

Open `manifest.xml` and replace every instance of `YOUR_DOMAIN` with your Cloudflare Pages URL:

```xml
<bt:Url id="Commands.Url" DefaultValue="https://YOUR_PROJECT_NAME.pages.dev/commands.js"/>
```

Also replace the `<Id>` field with a fresh GUID. On macOS:

```bash
uuidgen
```

Update `commands.js` too:

```js
const SIGNATURE_URL = "https://YOUR_PROJECT_NAME.pages.dev/signature.html";
```

Push the changes — Cloudflare Pages will redeploy automatically.

### 5. Sideload the add-in

1. Open Outlook (new Outlook for Mac, or Outlook on Windows/web — see [Requirements](#requirements))
1. **Insert → Add-ins → My Add-ins → ··· → Add a Custom Add-in → Add from File…**
1. Select your local `manifest.xml`
1. Open a new compose window — the signature will appear automatically

The menu path above is the same across new Outlook for Mac, Outlook on Windows, and Outlook on the web. On legacy/classic Outlook for Mac, event-based activation isn't supported, so this add-in won't fire there regardless of sideloading.

-----

## Updating the Signature

Edit `signature.html` (or replace `assets/banner.png` for a new campaign) and push to `main`. Cloudflare Pages redeploys in seconds and everyone's next email uses the new block. No manifest changes, no re-sideloading required.

-----

## Troubleshooting

**Signature doesn’t appear after sideloading**

- Open `commands.js` and `signature.html` URLs directly in a browser to confirm they’re live and returning the expected content
- Check that all `YOUR_DOMAIN` placeholders in `manifest.xml` have been replaced
- Confirm the `<Id>` in `manifest.xml` is a valid GUID

**Add-in loads but signature is blank**

- Open browser DevTools on the Cloudflare Pages URL and check for CORS errors
- Verify the `_headers` file is present in the repo root and has deployed correctly

**`setSignatureAsync` has no effect**

- Outlook respects the user’s native signature settings — if a default signature is already configured in Outlook preferences, the API will not overwrite it. Remove the native signature under **Outlook → Preferences → Signatures** (Mac) or **File → Options → Mail → Signatures** (Windows) and try again.

**Add-in installs but signature never appears on Outlook for Windows**

- If this is classic (not new) Outlook for Windows, check **Help → About Outlook** for the build number. Builds older than Version 2403 (Build 17425.20000) can't run the `async`/`await` code in `commands.js` and will silently time out — see [Requirements](#requirements)

**Manifest rejected silently by Outlook**

- Validate the XML is well-formed (no unclosed tags, no stray placeholder text)
- Ensure the Cloudflare Pages URLs are HTTPS

-----

## Architecture Notes

This is a proof of concept. The sideloaded manifest approach works for development and internal testing without requiring Microsoft 365 admin access or AppSource publishing. For a production rollout, the manifest would be deployed via the Microsoft 365 admin centre, making the add-in available to all users in an organisation automatically.

-----

## About

This is the DIY end of the spectrum: one company sign-off for the whole team, no moving parts, fork it and it's yours. If you're weighing that against a commercial signature management platform instead — per-user signatures, an admin dashboard, centralised deployment — [SigHQ](https://sighq.app) is a good place to see what that space looks like.

-----

## License

MIT
