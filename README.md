# Voice Edge Auth Sync — single source, dual target (Firefox MV2 + Chrome MV3)

One source tree builds both browsers. `background.js` is platform-neutral
(runtime feature detection); only the manifest and a couple of tiny helpers
differ, and those differences are resolved at build time by `build.mjs`.

## Layout

```
src/                      shared source
  background.js           platform-neutral (see "How neutrality works")
  browser-polyfill.js     no-op on Firefox; chrome->browser shim on Chromium
  content-doubao.js       Chromium-only page hook (Firefox uses filterResponseData)
  content-deepseek.js     shared
  content-m365.js         shared
  content-sharepoint.js   shared
  popup.{html,js,css}     shared
  icons/                  shared
manifest.base.json        shared manifest fields (name/version/description/icons)
manifest.firefox.json     MV2 overrides
manifest.chrome.json      MV3 overrides
build.mjs                 merges base+override, copies only referenced files
dist/<target>/            build output
```

## Build

```bash
node build.mjs              # both targets -> dist/firefox, dist/chrome
node build.mjs firefox      # one target
node build.mjs chrome --zip # build + zip -> dist/voice-edge-auth-sync-chrome.zip
node build.mjs all --zip    # both + zip each
```

npm equivalents: `npm run build`, `build:firefox`, `build:chrome`,
`package`, `package:firefox`, `package:chrome`, `package:source`.

`build.mjs` copies **only the files a target's manifest references**, so the
Firefox build never ships `content-doubao.js` and each output is minimal.

## Release (automated)

Cutting a release is a single command:

```bash
npm run release --tag=vX.Y.Z    # must match manifest.base.json "version"
```

That tags the commit and pushes the tag, which triggers
`.github/workflows/build.yml`. The workflow then, in order:

1. **Creates a draft GitHub Release** for the tag (idempotent).
2. **Verifies** the tag matches the `version` in `manifest.base.json`
   (AMO refuses to sign the same version twice, so a mismatch fails fast).
3. **Builds both targets** — `dist/firefox/` (MV2) and the zipped
   `dist/voice-edge-auth-sync-chrome.zip` (MV3).
4. **Signs Firefox** via AMO self-hosted signing (`web-ext`, `channel:
unlisted`) — `bin/sign.mjs` signs the built `dist/firefox` and downloads
   the `.xpi` to `dist/signed/`.
5. **Publishes both assets** to the release, named after the version:
   - `voice-edge-auth-sync-<version>-firefox.xpi` — **signed**, installable
     directly in Firefox (self-hosted, not listed on AMO).
   - `voice-edge-auth-sync-<version>-chrome.zip` — **unsigned** (Chrome needs
     no signature; load unpacked, or drag the zip into `chrome://extensions`).
6. **Un-drafts the release** once every asset uploaded successfully.

### Required repository secrets (Firefox signing only)

Create an AMO API credential at
<https://addons.mozilla.org/developers/addon/api/key/> and add two repo
secrets. Chrome needs none.

| Secret           | Value                        |
| ---------------- | ---------------------------- |
| `AMO_API_KEY`    | AMO API **key** (JWT issuer) |
| `AMO_API_SECRET` | AMO API **secret**           |

`bin/sign.mjs` is resilient: transient AMO/network errors are retried with
backoff, and if a version was already uploaded (e.g. a prior run died waiting
for review) it **recovers the signed file from the AMO API** instead of
re-signing — because AMO permanently consumes a version number once uploaded.

To sign locally: `WEB_EXT_API_KEY=… WEB_EXT_API_SECRET=… npm run sign`
(builds `dist/firefox` first, then writes the `.xpi` to `dist/signed/`).

## How neutrality works (no `#ifdef`, just runtime detection)

| Concern         | Firefox (MV2)                                     | Chrome (MV3)                                     | Neutral mechanism in `background.js`                                                                                             |
| --------------- | ------------------------------------------------- | ------------------------------------------------ | -------------------------------------------------------------------------------------------------------------------------------- |
| API namespace   | native `browser.*`                                | `chrome.*` (+shim)                               | `browser-polyfill.js`: no-op on FF, aliases `chrome`+bridges `onMessage` on Chrome                                               |
| Toolbar action  | `browser_action`                                  | `action`                                         | `const actionApi = browser.action \|\| browser.browserAction;`                                                                   |
| ttwid fallback  | `tabs.executeScript({code})`                      | `scripting.executeScript({func})`                | `if (browser.scripting?.executeScript) … else …`                                                                                 |
| Doubao capture  | `webRequest.filterResponseData` (response stream) | page-world fetch/XHR hook → `DOUBAO_CAPTURE` msg | `if (browser.webRequest.filterResponseData) …register stream… else …register message…` — both call the same `ingestDoubaoBody()` |
| Background type | persistent page (`scripts[]`)                     | service worker                                   | manifest-only; `importScripts()` is `try/caught` so it's inert on FF                                                             |

The DeepSeek header sniff (`onBeforeSendHeaders`) and the M365 token sniff
(`onBeforeRequest` + `requestBody`) are observer-only — no `blocking` — and work
unchanged on both.

## Load unpacked (no signing needed for local use)

- **Chrome**: `chrome://extensions` → Developer mode → Load unpacked → `dist/chrome`.
- **Firefox**: `about:debugging` → This Firefox → Load Temporary Add-on → pick
  `dist/firefox/manifest.json` (or use `web-ext run` / `web-ext build`).

## Fixed extension ID (Chrome)

Native messaging pins the extension by ID (`allowed_origins`), so the Chrome ID
**must not drift** — otherwise the native host rejects the connection and auth
sync silently dies. `manifest.chrome.json` therefore embeds a `"key"` field
(an RSA **public** key, safe to commit — the private key is never needed for
`load unpacked` and is not used here). It pins the ID to:

```
lbhjpomjdikihplfkgdmdkgjdjnklmak
```

This holds regardless of load path, machine, or repacking. Firefox needs no
equivalent: its ID is already fixed by `browser_specific_settings.gecko.id`.

To rotate the ID (rarely needed), regenerate the key and recompute the ID:

```bash
# new public key for the "key" field in manifest.chrome.json
openssl genrsa 2048 | openssl rsa -pubout -outform DER 2>/dev/null | base64 -w0

# the resulting fixed ID (feed the SAME DER public key in)
openssl genrsa 2048 | openssl rsa -pubout -outform DER 2>/dev/null \
  | python3 -c "import sys,hashlib;h=hashlib.sha256(sys.stdin.buffer.read()).hexdigest()[:32];print(''.join(chr(97+int(c,16)) for c in h))"
```

## Native messaging host (differs per browser!)

Same host binary, but the allow-list key differs:

- **Firefox** uses `allowed_extensions: ["voice-edge-auth-sync@local"]`.
- **Chrome** uses
  `allowed_origins: ["chrome-extension://lbhjpomjdikihplfkgdmdkgjdjnklmak/"]`
  (the fixed ID above — verify it in `chrome://extensions` after loading).

Host manifest install paths:

| OS      | Firefox                                                       | Chrome                                                              |
| ------- | ------------------------------------------------------------- | ------------------------------------------------------------------- |
| macOS   | `~/Library/Application Support/Mozilla/NativeMessagingHosts/` | `~/Library/Application Support/Google/Chrome/NativeMessagingHosts/` |
| Linux   | `~/.mozilla/native-messaging-hosts/`                          | `~/.config/google-chrome/NativeMessagingHosts/`                     |
| Windows | `HKCU\Software\Mozilla\NativeMessagingHosts\<name>`           | `HKCU\Software\Google\Chrome\NativeMessagingHosts\<name>`           |

The M365 WebSocket bridge (`ws://127.0.0.1:5003/ws`) is identical on both.

## Chrome service-worker notes

The background is a service worker on Chrome. It stays alive in practice via two
long-lived connections (native-messaging port + M365 WebSocket) that reset the
idle timer; M365 auth/settings persist in `storage.local` and rehydrate on wake.
Transient per-flow state (Doubao/DeepSeek/Qwen in-progress) is short-lived and
re-derived on the next request — same behaviour as Firefox.

## Caveat: page-CSP inline injection

`content-doubao.js` (like `content-m365.js`) injects its page-world hook via an
inline `<script>`. If Doubao ships a strict CSP blocking inline scripts, capture
degrades silently and you can fall back to a manual popup sync; switch the hook
to a `web_accessible_resources` file loaded via `src=` if that ever happens.
