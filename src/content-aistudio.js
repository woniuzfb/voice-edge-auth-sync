/*
 * Voice Edge · AI Studio Playground (MakerSuiteService) page-world bridge.
 *
 * Architecture:
 *   bridge (:5003/ws) <--> background.js <--> content-aistudio.js (isolated)
 *     <--> PAGE_HOOK (main world on https://aistudio.google.com/*)
 *
 * Responsibilities in PAGE_HOOK:
 *   1. Hook XMLHttpRequest to passively capture MakerSuite dynamic headers:
 *        - X-AIStudio-G1-Tier
 *        - X-AIStudio-Visit-Id
 *        - X-Goog-Ext-519733851-bin
 *        - X-Goog-Api-Key
 *        - X-Goog-AuthUser
 *   2. Hook window.default_MakerSuite's high-level snapshot wrapper (whose source
 *      contains `.snapshot({`, `content`, and `yield`) to capture `args[0]`
 *      (the official WAA/BotGuard service instance) into `hookState.service`.
 *   3. Perform a zero-quota silent bootstrap ("Reply with OK" + Run click +
 *      XHR.send interception before network transmission + Stop click + textarea
 *      clear) whenever a protected call arrives and `hookState.service` has not
 *      yet been exposed by user interaction.
 *   4. Serialize WAA proof generation: sync `pagePrompt` into `ms-prompt-box textarea`
 *      using the native HTMLTextAreaElement.prototype.value setter + InputEvent/Event,
 *      await `hookState.snapshot(digest)`, and clear `textarea` in `finally`.
 *   5. Compute fresh 3-part SAPISIDHASH Authorization header on every request,
 *      stream XHR chunks (`AIS_HEAD`, `AIS_CHUNK`, `AIS_END`, `AIS_ERROR`) without
 *      splitting UTF-16 surrogate pairs, and fire the 2.2s companion `CountTokens`.
 */
(function () {
  "use strict";

  const AIS_HOOK_VERSION = "aistudio-hook-v3";

  function PAGE_HOOK() {
    if (
      window.__veAisHook &&
      window.__veAisHook.ready &&
      window.__veAisHook.version === "aistudio-hook-v3"
    ) {
      return;
    }
    const prevHook = window.__veAisHook || null;

    const MAKERSUITE_HOST = "alkalimakersuite-pa.clients6.google.com";
    const APPLETS_HOST = "alkalimakersuiteapplets.pa.clients6.google.com";
    const MAKERSUITE_BASE =
      "https://alkalimakersuite-pa.clients6.google.com/$rpc/google.internal.alkali.applications.makersuite.v1.MakerSuiteService";
    const GENERATE_CONTENT_SUFFIX = "/GenerateContent";
    const COUNT_TOKENS_URL = MAKERSUITE_BASE + "/CountTokens";
    const GENERATE_ACCESS_TOKEN_URL = MAKERSUITE_BASE + "/GenerateAccessToken";
    const UPLOAD_INIT_URL = "https://aistudio.google.com/_/upload";
    const TIERED_RPC_SUFFIXES = [
      "/CodeAssistantOffline",
      "/GenerateContent",
      "/CountTokens",
    ];
    const DRIVE_RPC_SUFFIXES = [
      "/GenerateAccessToken",
      "/ListPrompts",
      "/CreatePrompt",
      "/UpdatePrompt",
      "/GetAppFolder",
      "/ResolveDriveResource",
    ];
    const ORIGIN = "https://aistudio.google.com";

    const hookState = {
      ready: true,
      version: "aistudio-hook-v3",
      snapshotKey: (prevHook && prevHook.snapshotKey) || "",
      service: (prevHook && prevHook.service) || null,
      snapshot: (prevHook && prevHook.snapshot) || null,
      headers: Object.assign(
        {
          "x-aistudio-g1-tier": "",
          "x-aistudio-visit-id": "",
          "x-goog-ext-519733851-bin": "",
          "x-goog-api-key": "",
          "x-goog-authuser": "0",
        },
        (prevHook && prevHook.headers) || {},
      ),
      cookies: Object.assign(
        {
          SAPISID: "",
          "__Secure-1PAPISID": "",
          "__Secure-3PAPISID": "",
        },
        (prevHook && prevHook.cookies) || {},
      ),
      drive: Object.assign(
        {
          status: "unknown",
          checkedAt: 0,
          error: "",
        },
        (prevHook && prevHook.drive) || {},
      ),
      driveProbeInFlight: null,
      activeRequests: (prevHook && prevHook.activeRequests) || new Map(),
      bootstrapIntercept: false,
      bootstrapResolve: null,
      bootstrapPromise: null,
      proofQueue: (prevHook && prevHook.proofQueue) || Promise.resolve(),
    };
    window.__veAisHook = hookState;

    function post(payload) {
      try {
        window.postMessage(
          Object.assign({ __veAis: true, dir: "fromPage" }, payload),
          "*",
        );
      } catch (_) {}
    }

    function detectTimezone() {
      try {
        return Intl.DateTimeFormat().resolvedOptions().timeZone || "UTC";
      } catch (_) {
        return "UTC";
      }
    }

    function detectAuthUserFromUrl() {
      const m = /^\/u\/(\d+)\//.exec(location.pathname || "");
      if (m && m[1]) return m[1];
      try {
        const params = new URLSearchParams(location.search || "");
        const au = params.get("authuser");
        if (au !== null && /^\d+$/.test(au.trim())) {
          return au.trim();
        }
      } catch (_) {}
      return "";
    }

    function detectApiKeyFromDom() {
      try {
        const html = document.documentElement
          ? document.documentElement.innerHTML
          : "";
        const m = /"WIu0Nc"\s*:\s*"([^"]+)"/.exec(html);
        if (m && m[1]) return m[1];
      } catch (_) {}
      return "";
    }

    function generateVisitId() {
      try {
        if (typeof crypto !== "undefined" && crypto.randomUUID) {
          return "v1_" + btoa(crypto.randomUUID());
        }
      } catch (_) {}
      return "v1_" + btoa(String(Date.now()) + "-" + Math.random());
    }

    function readCookieMap() {
      const out = Object.assign({}, hookState.cookies);
      try {
        const raw = String(document.cookie || "");
        for (const part of raw.split(";")) {
          const idx = part.indexOf("=");
          if (idx <= 0) continue;
          const k = part.slice(0, idx).trim();
          const v = part.slice(idx + 1).trim();
          if (k && v) out[k] = v;
        }
      } catch (_) {}
      return out;
    }

    function hasAuthCookies() {
      const c = readCookieMap();
      return Boolean(
        c.SAPISID || c["__Secure-1PAPISID"] || c["__Secure-3PAPISID"],
      );
    }

    function shortCookieFingerprint(raw) {
      const s = String(raw || "").trim();
      if (!s) return "";
      let h = 0x811c9dc5;
      for (let i = 0; i < s.length; i++) {
        h ^= s.charCodeAt(i);
        h = Math.imul(h, 0x01000193) >>> 0;
      }
      return h.toString(16).padStart(8, "0");
    }

    function isDriveRpcUrl(url) {
      const u = String(url || "");
      return DRIVE_RPC_SUFFIXES.some((suffix) => u.endsWith(suffix));
    }

    function recordDriveResult(statusCode, responseText, silent) {
      const status = Number(statusCode || 0);
      const bodyStr = String(responseText || "");
      let nextStatus = "";
      let nextErr = "";
      if (status === 200) {
        nextStatus = "authorized";
      } else if (
        status === 401 ||
        status === 403 ||
        bodyStr.includes("unauthorized_client")
      ) {
        nextStatus = "unauthorized";
        nextErr = bodyStr.includes("unauthorized_client")
          ? "OAuth error: unauthorized_client"
          : `HTTP ${status}`;
      }
      if (!nextStatus) return;
      const changed =
        hookState.drive.status !== nextStatus ||
        hookState.drive.error !== nextErr;
      hookState.drive = {
        status: nextStatus,
        checkedAt: Date.now(),
        error: nextErr,
      };
      if (changed && !silent) {
        emitReady({ event: "drive-auth-updated" });
      }
    }

    function emitReady(extra) {
      const u = detectAuthUserFromUrl();
      if (u) {
        hookState.headers["x-goog-authuser"] = u;
      }
      if (!hookState.headers["x-goog-api-key"]) {
        const domKey = detectApiKeyFromDom();
        if (domKey) hookState.headers["x-goog-api-key"] = domKey;
      }
      if (!hookState.headers["x-aistudio-visit-id"]) {
        hookState.headers["x-aistudio-visit-id"] = generateVisitId();
      }
      const c = readCookieMap();
      const sapisidHash = shortCookieFingerprint(
        c.SAPISID || c["__Secure-3PAPISID"] || c["__Secure-1PAPISID"] || "",
      );
      post(
        Object.assign(
          {
            type: "AIS_READY",
            hookVersion: hookState.version,
            tabUrl: location.href,
            timezone: detectTimezone(),
            service: hookState.service ? "captured" : "pending",
            snapshotKey: hookState.snapshotKey || "",
            auth: {
              ok:
                location.hostname === "aistudio.google.com" && hasAuthCookies(),
              authUser: u || hookState.headers["x-goog-authuser"] || "0",
              tier: hookState.headers["x-aistudio-g1-tier"] || "",
              hasApiKey: Boolean(hookState.headers["x-goog-api-key"]),
              hasExtBin: Boolean(hookState.headers["x-goog-ext-519733851-bin"]),
              sapisidHash,
            },
            drive: Object.assign({}, hookState.drive),
          },
          extra || {},
        ),
      );
    }

    // ---- 1. Hook window.default_MakerSuite for WAA service -------------------
    function tryInstallSnapshotHook() {
      if (hookState.snapshot) return true;
      const makerSuite = window.default_MakerSuite;
      if (!makerSuite || typeof makerSuite !== "object") return false;
      for (const key of Object.keys(makerSuite)) {
        let original;
        try {
          const candidate = makerSuite[key];
          if (typeof candidate !== "function") continue;
          original = candidate.__veAisOriginal || candidate;
          const source = Function.prototype.toString.call(original);
          if (
            !source.includes(".snapshot({") ||
            !source.includes("content") ||
            !source.includes("yield")
          ) {
            continue;
          }
        } catch (_) {
          continue;
        }
        hookState.snapshotKey = key;
        hookState.snapshot = function (digest) {
          return original.call(makerSuite, hookState.service, digest);
        };
        const wrapper = function (...args) {
          const service = args[0];
          if (
            service &&
            (typeof service === "object" || typeof service === "function")
          ) {
            hookState.service = service;
            try {
              makerSuite[key] = original;
            } catch (_) {}
            emitReady({ event: "service-captured" });
          }
          return original.apply(this, args);
        };
        wrapper.__veAisOriginal = original;
        makerSuite[key] = wrapper;
        return true;
      }
      return false;
    }

    tryInstallSnapshotHook();
    const snapshotPollTimer = setInterval(() => {
      if (tryInstallSnapshotHook()) {
        clearInterval(snapshotPollTimer);
      }
    }, 100);
    setTimeout(() => clearInterval(snapshotPollTimer), 60000);

    // ---- 2. Hook XMLHttpRequest to capture dynamic headers & bootstrap -------
    try {
      const XHR = window.XMLHttpRequest;
      if (XHR && XHR.prototype && !XHR.prototype.__veAisHooked) {
        const origOpen = XHR.prototype.open;
        const origSetHeader = XHR.prototype.setRequestHeader;
        const origSend = XHR.prototype.send;

        XHR.prototype.open = function (method, url) {
          try {
            this.__veAisMethod = String(method || "GET").toUpperCase();
            this.__veAisUrl = String(url || "");
            this.__veAisReqHeaders = {};
          } catch (_) {}
          return origOpen.apply(this, arguments);
        };

        XHR.prototype.setRequestHeader = function (name, value) {
          try {
            if (!this.__veAisSynthetic && name && value != null) {
              const lower = String(name).toLowerCase().trim();
              const val = String(value).trim();
              if (this.__veAisReqHeaders) {
                this.__veAisReqHeaders[lower] = val;
              }
              if (
                this.__veAisUrl &&
                (this.__veAisUrl.includes(MAKERSUITE_HOST) ||
                  this.__veAisUrl.includes(APPLETS_HOST)) &&
                lower in hookState.headers &&
                val
              ) {
                hookState.headers[lower] = val;
              }
            }
          } catch (_) {}
          return origSetHeader.apply(this, arguments);
        };

        XHR.prototype.send = function (body) {
          try {
            const url = String(this.__veAisUrl || "");
            if (
              !this.__veAisSynthetic &&
              (url.includes(MAKERSUITE_HOST) || url.includes(APPLETS_HOST))
            ) {
              const hdrs = this.__veAisReqHeaders || {};
              let changed = false;
              for (const k of Object.keys(hookState.headers)) {
                if (hdrs[k] && hookState.headers[k] !== hdrs[k]) {
                  hookState.headers[k] = hdrs[k];
                  changed = true;
                }
              }
              if (url.endsWith("/GetAiStudioBenefitTier")) {
                this.addEventListener("load", () => {
                  try {
                    const parsed = JSON.parse(this.responseText || "[]");
                    const code = Array.isArray(parsed) ? parsed[0] : null;
                    if (code === 1)
                      hookState.headers["x-aistudio-g1-tier"] = "TIER1";
                    else if (code === 2)
                      hookState.headers["x-aistudio-g1-tier"] = "TIER2";
                    else if (code === 3)
                      hookState.headers["x-aistudio-g1-tier"] = "TIER0";
                  } catch (_) {}
                });
              }
              if (isDriveRpcUrl(url)) {
                this.addEventListener("load", () => {
                  try {
                    recordDriveResult(this.status, this.responseText);
                  } catch (_) {}
                });
              }
              const bodyText = typeof body === "string" ? body : "";
              if (
                hookState.bootstrapIntercept &&
                url.endsWith(GENERATE_CONTENT_SUFFIX)
              ) {
                hookState.bootstrapIntercept = false;
                const resolver = hookState.bootstrapResolve;
                hookState.bootstrapResolve = null;
                if (typeof resolver === "function") {
                  resolver({
                    url,
                    headers: Object.assign({}, hdrs),
                    body: bodyText,
                  });
                }
                // Abort bootstrap GenerateContent locally without hitting network
                setTimeout(() => {
                  try {
                    if (typeof this.onerror === "function") {
                      this.onerror(new ProgressEvent("error"));
                    }
                    this.dispatchEvent(new ProgressEvent("error"));
                  } catch (_) {}
                }, 0);
                return;
              }
              if (
                (hookState.bootstrapIntercept ||
                  hookState.bootstrapPromise ||
                  bodyText.includes('"Reply with OK"')) &&
                (url.endsWith("/CountTokens") ||
                  url.endsWith("/GenerateTitle") ||
                  url.endsWith("/CreatePrompt") ||
                  url.endsWith("/UpdatePrompt"))
              ) {
                // Suppress any side-effect RPCs triggered by the silent bootstrap prompt
                setTimeout(() => {
                  try {
                    if (typeof this.onerror === "function") {
                      this.onerror(new ProgressEvent("error"));
                    }
                    this.dispatchEvent(new ProgressEvent("error"));
                  } catch (_) {}
                }, 0);
                return;
              }
              if (changed) {
                emitReady({ event: "headers-updated" });
              }
            }
          } catch (_) {}
          return origSend.apply(this, arguments);
        };

        XHR.prototype.__veAisHooked = true;
      }
    } catch (_) {}

    // ---- 3. DOM Helpers, Silent Bootstrap & Proof Generation -----------------
    function visible(element) {
      if (!element) return false;
      if (typeof element.checkVisibility === "function") {
        return element.checkVisibility({ visibilityProperty: true });
      }
      return Boolean(
        element.offsetWidth ||
        element.offsetHeight ||
        element.getClientRects().length,
      );
    }

    function uniqueVisible(selector, label) {
      const items = Array.from(document.querySelectorAll(selector)).filter(
        visible,
      );
      if (items.length > 1) {
        throw new Error((label || selector) + " matched multiple elements");
      }
      return items[0] || null;
    }

    function buttonEnabled(button) {
      return Boolean(
        button &&
        !button.matches(":disabled") &&
        !button.closest('[aria-disabled="true"]'),
      );
    }

    function dismissKnownOverlays() {
      const selectors = [
        'ms-g1-welcome-dialog button[aria-label="Close dialog"]',
        'button[aria-label="Close guided tour"]',
        "#glue-cookie-notification-bar-1 .glue-cookie-notification-bar__reject",
      ];
      let clicked = 0;
      for (const selector of selectors) {
        try {
          const btn = uniqueVisible(selector, selector);
          if (btn && buttonEnabled(btn)) {
            btn.click();
            clicked += 1;
          }
        } catch (_) {}
      }
      return clicked;
    }

    function setTextareaPrompt(textarea, text) {
      const desc = Object.getOwnPropertyDescriptor(
        HTMLTextAreaElement.prototype,
        "value",
      );
      const setValue = desc && desc.set;
      if (!setValue) {
        textarea.value = text;
      } else {
        setValue.call(textarea, text);
      }
      textarea.dispatchEvent(
        new InputEvent("input", {
          bubbles: true,
          inputType: text ? "insertText" : "deleteContent",
          data: text || null,
        }),
      );
      if (text) {
        textarea.dispatchEvent(new Event("change", { bubbles: true }));
      }
      return textarea.value === text;
    }

    function sleep(ms) {
      return new Promise((resolve) => setTimeout(resolve, ms));
    }

    async function waitForCondition(fn, timeoutMs, label) {
      const deadline = Date.now() + timeoutMs;
      while (Date.now() < deadline) {
        try {
          const res = fn();
          if (res) return res;
        } catch (_) {}
        await sleep(100);
      }
      throw new Error("Timed out waiting for: " + label);
    }

    function navigateToPlaygroundSpa() {
      hookState.service = null;
      hookState.snapshot = null;
      tryInstallSnapshotHook();
      try {
        const navLink =
          document.querySelector('a[href*="/prompts/new_chat"]') ||
          Array.from(document.querySelectorAll("a")).find((a) =>
            /^\s*Playground\s*$/i.test(String(a.textContent || "")),
          );
        if (navLink && typeof navLink.click === "function") {
          navLink.click();
          return true;
        }
      } catch (_) {}
      try {
        const u =
          detectAuthUserFromUrl() ||
          hookState.headers["x-goog-authuser"] ||
          "0";
        const targetPath = `/u/${u}/prompts/new_chat`;
        if (location.pathname !== targetPath) {
          history.pushState({}, "", targetPath);
          window.dispatchEvent(new PopStateEvent("popstate"));
          return true;
        }
      } catch (_) {}
      return false;
    }

    async function ensurePromptTextarea(timeoutMs) {
      const existing = uniqueVisible(
        "ms-prompt-box textarea",
        "prompt textarea",
      );
      if (existing) return existing;
      let navigated = false;
      if (!location.pathname.includes("/prompts/")) {
        navigated = navigateToPlaygroundSpa();
      }
      const startTs = Date.now();
      const deadline = startTs + (timeoutMs || 10000);
      while (Date.now() < deadline) {
        try {
          dismissKnownOverlays();
          const area = uniqueVisible(
            "ms-prompt-box textarea",
            "prompt textarea",
          );
          if (area) return area;
          if (!navigated && Date.now() - startTs >= 1200) {
            navigated = navigateToPlaygroundSpa();
          }
        } catch (_) {}
        await sleep(100);
      }
      throw new Error("Timed out waiting for: ms-prompt-box textarea");
    }

    async function ensureBootstrapped() {
      if (location.hostname === "accounts.google.com") {
        const err = new Error("Google authentication required on AI Studio");
        err.code = "auth_required";
        err.httpStatus = 401;
        throw err;
      }
      if (location.pathname.includes("/docs/available-regions")) {
        const err = new Error("AI Studio region not supported");
        err.code = "region_not_supported";
        err.httpStatus = 403;
        throw err;
      }
      if (
        hookState.service &&
        hookState.snapshot &&
        hookState.headers["x-goog-api-key"]
      ) {
        return;
      }
      if (hookState.bootstrapPromise) {
        return hookState.bootstrapPromise;
      }
      hookState.bootstrapPromise = (async () => {
        try {
          await ensurePromptTextarea(25000);
          await waitForCondition(
            () =>
              performance
                .getEntriesByType("resource")
                .some((e) =>
                  String(e.name || "").endsWith("/GetLoggingContext"),
                ) || Boolean(hookState.headers["x-goog-ext-519733851-bin"]),
            15000,
            "GetLoggingContext",
          ).catch(() => {});
          dismissKnownOverlays();
          await waitForCondition(
            () => tryInstallSnapshotHook(),
            15000,
            "window.default_MakerSuite snapshot hook",
          );
          if (hookState.service && hookState.headers["x-goog-api-key"]) {
            return;
          }

          const textarea = await ensurePromptTextarea(8000);
          if (!textarea) {
            throw new Error("AI Studio prompt textarea not found");
          }
          const interceptPromise = new Promise((resolve, reject) => {
            hookState.bootstrapIntercept = true;
            const t = setTimeout(() => {
              hookState.bootstrapIntercept = false;
              hookState.bootstrapResolve = null;
              reject(
                new Error("Bootstrap GenerateContent intercept timed out"),
              );
            }, 15000);
            hookState.bootstrapResolve = (data) => {
              clearTimeout(t);
              resolve(data);
            };
          });

          setTextareaPrompt(textarea, "Reply with OK");
          const runBtn = await waitForCondition(
            () => {
              const b = uniqueVisible("ms-run-button button", "Run button");
              return b && buttonEnabled(b) ? b : null;
            },
            10000,
            "ms-run-button enabled",
          );
          runBtn.click();

          await interceptPromise;

          // Yield once for the synthetic XHR error event to settle, then click Stop if still spinning
          await sleep(20);
          const stopBtn = uniqueVisible("ms-run-button button", "Stop button");
          if (
            stopBtn &&
            stopBtn.querySelector(".spin") &&
            buttonEnabled(stopBtn)
          ) {
            stopBtn.click();
          }
          const cleanArea = uniqueVisible(
            "ms-prompt-box textarea",
            "prompt textarea",
          );
          if (cleanArea) {
            setTextareaPrompt(cleanArea, "");
          }
          if (!hookState.service) {
            throw new Error("WAA service was not captured during bootstrap");
          }
          emitReady({ event: "bootstrap-complete" });
        } finally {
          hookState.bootstrapIntercept = false;
          hookState.bootstrapResolve = null;
          hookState.bootstrapPromise = null;
        }
      })();
      return hookState.bootstrapPromise;
    }

    function normalizeNewlines(value) {
      return String(value || "")
        .replace(/\r\n/g, "\n")
        .replace(/\r/g, "\n");
    }

    function generateProof(rawPagePrompt, digest) {
      const run = async () => {
        await ensureBootstrapped();
        const prompt = normalizeNewlines(rawPagePrompt);
        const deadline = Date.now() + 8000;
        while (true) {
          let textarea = uniqueVisible(
            "ms-prompt-box textarea",
            "prompt textarea",
          );
          if (!textarea) {
            textarea = await ensurePromptTextarea(
              Math.max(1500, deadline - Date.now()),
            ).catch(() => null);
          }
          if (!textarea) {
            throw new Error("AI Studio prompt textarea not visible");
          }
          const synced = setTextareaPrompt(textarea, prompt);
          let proof = null;
          try {
            if (synced) {
              proof = await hookState.snapshot(digest);
            }
          } finally {
            try {
              setTextareaPrompt(textarea, "");
            } catch (_) {}
          }
          if (typeof proof === "string" && proof.length > 0) {
            if (!proof.startsWith("!")) {
              throw new Error("Invalid WAA proof prefix");
            }
            return proof;
          }
          if (Date.now() >= deadline) {
            throw new Error("Failed to synchronize prompt state for WAA proof");
          }
          await sleep(100);
        }
      };
      const next = hookState.proofQueue.then(run, run);
      hookState.proofQueue = next.catch(() => {});
      return next;
    }

    // ---- 4. 3-Part SAPISIDHASH Authorization & Request Headers ---------------
    async function sha1Hex(input) {
      const bytes = new TextEncoder().encode(input);
      const digest = await crypto.subtle.digest("SHA-1", bytes);
      const arr = new Uint8Array(digest);
      let hex = "";
      for (let i = 0; i < arr.length; i++) {
        hex += arr[i].toString(16).padStart(2, "0");
      }
      return hex;
    }

    async function buildAuthorizationHeader() {
      const cookies = readCookieMap();
      const timestamp = Math.floor(Date.now() / 1000);
      const defs = [
        ["SAPISID", "SAPISIDHASH"],
        ["__Secure-1PAPISID", "SAPISID1PHASH"],
        ["__Secure-3PAPISID", "SAPISID3PHASH"],
      ];
      const parts = [];
      for (const [cookieName, scheme] of defs) {
        const value = String(cookies[cookieName] || "").trim();
        if (!value) continue;
        const hash = await sha1Hex(`${timestamp} ${value} ${ORIGIN}`);
        parts.push(`${scheme} ${timestamp}_${hash}`);
      }
      if (parts.length === 0) {
        const err = new Error(
          "Missing Google SAPISID cookies on aistudio.google.com",
        );
        err.code = "auth_required";
        err.httpStatus = 401;
        throw err;
      }
      return parts.join(" ");
    }

    function isTieredRpcUrl(url) {
      const u = String(url || "");
      return TIERED_RPC_SUFFIXES.some((suffix) => u.endsWith(suffix));
    }

    async function buildOrderedRequestHeaders(
      omitExtBin,
      targetUrl,
      includeTierOverride,
    ) {
      const authHeader = await buildAuthorizationHeader();
      const u = detectAuthUserFromUrl();
      if (u) {
        hookState.headers["x-goog-authuser"] = u;
      }
      if (!hookState.headers["x-goog-api-key"]) {
        const domKey = detectApiKeyFromDom();
        if (domKey) hookState.headers["x-goog-api-key"] = domKey;
      }
      if (!hookState.headers["x-aistudio-visit-id"]) {
        hookState.headers["x-aistudio-visit-id"] = generateVisitId();
      }
      const sendTier =
        typeof includeTierOverride === "boolean"
          ? includeTierOverride
          : isTieredRpcUrl(targetUrl);
      const pairs = [
        [
          "X-AIStudio-G1-Tier",
          sendTier ? hookState.headers["x-aistudio-g1-tier"] : "",
        ],
        ["X-AIStudio-Visit-Id", hookState.headers["x-aistudio-visit-id"]],
        [
          "X-Goog-Ext-519733851-bin",
          omitExtBin ? "" : hookState.headers["x-goog-ext-519733851-bin"],
        ],
        ["X-Goog-Api-Key", hookState.headers["x-goog-api-key"]],
        ["X-Goog-AuthUser", hookState.headers["x-goog-authuser"] || "0"],
        ["Authorization", authHeader],
        ["Content-Type", "application/json+protobuf"],
        ["X-User-Agent", "grpc-web-javascript/0.1"],
      ];
      return pairs.filter(([, val]) => Boolean(String(val || "").trim()));
    }

    // ---- 5. Streaming, Unary XHR & /_/upload Execution -----------------------
    function base64ToUint8Array(b64) {
      const clean = String(b64 || "").replace(/\s+/g, "");
      const bin = atob(clean);
      const out = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) {
        out[i] = bin.charCodeAt(i);
      }
      return out;
    }

    async function decodeBase64ToUint8Array(b64) {
      const clean = String(b64 || "").replace(/\s+/g, "");
      if (!clean) return new Uint8Array(0);
      try {
        const res = await fetch(
          "data:application/octet-stream;base64," + clean,
        );
        if (res.ok) {
          return new Uint8Array(await res.arrayBuffer());
        }
      } catch (_) {}
      return base64ToUint8Array(clean);
    }

    function uint8ArrayToBase64(bytes) {
      const chunkSize = 0x8000;
      const chunks = [];
      for (let i = 0; i < bytes.length; i += chunkSize) {
        chunks.push(
          String.fromCharCode.apply(null, bytes.subarray(i, i + chunkSize)),
        );
      }
      return btoa(chunks.join(""));
    }

    async function sha256Base64(bytes) {
      const digest = await crypto.subtle.digest("SHA-256", bytes);
      return uint8ArrayToBase64(new Uint8Array(digest));
    }

    async function executeAisUpload(msg) {
      const id = String(msg.id || "");
      if (!id) return;
      try {
        if (location.hostname !== "aistudio.google.com" || !hasAuthCookies()) {
          const err = new Error("Google authentication required on AI Studio");
          err.code = "auth_required";
          err.httpStatus = 401;
          throw err;
        }
        const bytes = await decodeBase64ToUint8Array(
          msg.dataBase64 || msg.data || "",
        );
        const mimeType = String(
          msg.mimeType || "application/octet-stream",
        ).trim();
        const fileName = String(msg.fileName || "attachment").trim();
        const hashHeader = "sha256=" + (await sha256Base64(bytes));
        const authUser =
          detectAuthUserFromUrl() ||
          hookState.headers["x-goog-authuser"] ||
          "0";

        const startHeaders = {
          "X-Goog-Upload-Header-Content-Length": String(bytes.byteLength),
          "X-Goog-Upload-Header-Content-Type": mimeType,
          "X-Goog-Hash": hashHeader,
          "X-Goog-Upload-Protocol": "resumable",
          "X-Goog-Upload-Command": "start",
        };
        if (Number(authUser) > 0) {
          startHeaders["X-Goog-AuthUser"] = String(authUser);
        }

        const startRes = await fetch(UPLOAD_INIT_URL, {
          method: "POST",
          credentials: "include",
          headers: startHeaders,
        });
        if (!startRes.ok) {
          const errText = await startRes.text().catch(() => "");
          const err = new Error(
            `/_/upload start failed (HTTP ${startRes.status}): ${errText.slice(0, 200)}`,
          );
          err.code = "upload_start_failed";
          err.httpStatus = startRes.status || 502;
          throw err;
        }
        const uploadUrl =
          startRes.headers.get("x-goog-upload-url") ||
          startRes.headers.get("X-Goog-Upload-URL") ||
          "";
        if (!uploadUrl) {
          const err = new Error(
            "/_/upload start response missing X-Goog-Upload-URL",
          );
          err.code = "upload_missing_url";
          err.httpStatus = 502;
          throw err;
        }

        const finalizeHeaders = {
          "Content-Type": mimeType,
          "X-Goog-Hash": hashHeader,
          "X-Goog-Upload-Command": "upload, finalize",
          "X-Goog-Upload-Offset": "0",
        };
        if (Number(authUser) > 0) {
          finalizeHeaders["X-Goog-AuthUser"] = String(authUser);
        }
        const finalizeRes = await fetch(uploadUrl, {
          method: "POST",
          credentials: "include",
          headers: finalizeHeaders,
          body: bytes,
        });
        const rawText = (await finalizeRes.text().catch(() => "")).trim();
        if (!finalizeRes.ok || !rawText) {
          const err = new Error(
            `/_/upload finalize failed (HTTP ${finalizeRes.status}): ${rawText.slice(0, 200)}`,
          );
          err.code = "upload_finalize_failed";
          err.httpStatus = finalizeRes.status || 502;
          throw err;
        }
        post({
          type: "AIS_HEAD",
          id,
          status: 200,
          headers: { "content-type": "text/plain; charset=utf-8" },
        });
        post({
          type: "AIS_CHUNK",
          id,
          seq: 0,
          data: rawText,
        });
        post({
          type: "AIS_END",
          id,
          seq: 1,
          status: 200,
          uploadPath: rawText,
          fileName,
          mimeType,
          size: bytes.byteLength,
        });
      } catch (err) {
        post({
          type: "AIS_ERROR",
          id,
          code: String((err && err.code) || "upload_failed"),
          message: String((err && err.message) || err),
          httpStatus: Number((err && err.httpStatus) || 502),
        });
      }
    }

    async function executeAisRequest(msg) {
      const id = String(msg.id || "");
      if (!id) return;
      try {
        let bodyStr = String(msg.body || "");
        const proofField = Number(msg.proofField || 0);
        if (proofField > 0 && msg.digest) {
          const proof = await generateProof(
            msg.pagePrompt || "",
            String(msg.digest),
          );
          const parsedBody = JSON.parse(bodyStr);
          if (!Array.isArray(parsedBody)) {
            throw new Error("Protected request body must be a JSON array");
          }
          while (parsedBody.length < proofField) {
            parsedBody.push(null);
          }
          parsedBody[proofField - 1] = proof;
          bodyStr = JSON.stringify(parsedBody);
        } else if (msg.requireBootstrap) {
          await ensureBootstrapped();
        }

        const headerPairs = await buildOrderedRequestHeaders(
          Boolean(msg.omitExtBin),
          String(msg.url || ""),
          typeof msg.includeTierHeader === "boolean"
            ? msg.includeTierHeader
            : undefined,
        );
        const xhr = new XMLHttpRequest();
        xhr.__veAisSynthetic = true;
        hookState.activeRequests.set(id, xhr);

        let seq = 0;
        let sentOffset = 0;
        let headSent = false;

        const sendHeadIfNeeded = () => {
          if (headSent || xhr.readyState < 2) return;
          headSent = true;
          const rawLines = String(xhr.getAllResponseHeaders() || "")
            .split(/\r\n/)
            .filter(Boolean);
          const respHeaders = {};
          for (const line of rawLines) {
            const idx = line.indexOf(":");
            if (idx > 0) {
              respHeaders[line.slice(0, idx).trim().toLowerCase()] = line
                .slice(idx + 1)
                .trim();
            }
          }
          post({
            type: "AIS_HEAD",
            id,
            status: xhr.status || 200,
            headers: respHeaders,
          });
        };

        const flushChunks = (isFinal) => {
          sendHeadIfNeeded();
          const text = String(xhr.responseText || "");
          while (sentOffset < text.length) {
            let end = Math.min(text.length, sentOffset + 65536);
            if (!isFinal && end < text.length) {
              const code = text.charCodeAt(end - 1);
              if (code >= 0xd800 && code <= 0xdbff) {
                end -= 1;
              }
            }
            if (end <= sentOffset) break;
            const chunk = text.slice(sentOffset, end);
            sentOffset = end;
            post({
              type: "AIS_CHUNK",
              id,
              seq: seq++,
              data: chunk,
            });
          }
        };

        xhr.open(String(msg.method || "POST"), String(msg.url || ""));
        xhr.withCredentials = true;
        for (const [hName, hVal] of headerPairs) {
          xhr.setRequestHeader(hName, hVal);
        }

        xhr.onreadystatechange = () => {
          try {
            if (xhr.readyState >= 2) sendHeadIfNeeded();
            if (xhr.readyState === 3) flushChunks(false);
          } catch (_) {}
        };
        xhr.onprogress = () => {
          try {
            flushChunks(false);
          } catch (_) {}
        };
        xhr.onload = () => {
          hookState.activeRequests.delete(id);
          try {
            if (isDriveRpcUrl(msg.url)) {
              recordDriveResult(xhr.status, xhr.responseText);
            }
            flushChunks(true);
            post({
              type: "AIS_END",
              id,
              seq,
              status: xhr.status || 200,
            });
          } catch (err) {
            post({
              type: "AIS_ERROR",
              id,
              code: "stream_finalize_error",
              message: String((err && err.message) || err),
            });
          }
        };
        xhr.onerror = () => {
          hookState.activeRequests.delete(id);
          post({
            type: "AIS_ERROR",
            id,
            code: "network_error",
            message: "NetworkError when attempting to fetch MakerSuite RPC",
            httpStatus: xhr.status || 502,
          });
        };
        xhr.onabort = () => {
          hookState.activeRequests.delete(id);
          post({
            type: "AIS_ERROR",
            id,
            code: "aborted",
            message: "AbortError: The operation was aborted",
            httpStatus: 499,
          });
        };

        xhr.send(bodyStr);

        if (
          msg.companionCountTokens &&
          hookState.drive.status !== "authorized"
        ) {
          probeDriveAuth(false).catch(() => {});
        }

        // 2200ms companion CountTokens after GenerateContent
        if (msg.companionCountTokens && msg.pagePrompt) {
          const companionPrompt = normalizeNewlines(msg.pagePrompt);
          const companionModel = String(
            msg.model || "models/gemini-4-argon",
          ).startsWith("models/")
            ? String(msg.model || "models/gemini-4-argon")
            : "models/" + String(msg.model || "gemini-4-argon");
          const companionBody = JSON.stringify([
            companionModel,
            [[[[null, companionPrompt]], "user"]],
          ]);
          setTimeout(() => {
            try {
              const companion = new XMLHttpRequest();
              companion.__veAisSynthetic = true;
              companion.open("POST", COUNT_TOKENS_URL);
              companion.withCredentials = true;
              for (const [hName, hVal] of headerPairs) {
                companion.setRequestHeader(hName, hVal);
              }
              companion.send(companionBody);
            } catch (_) {}
          }, 2200);
        }
      } catch (err) {
        hookState.activeRequests.delete(id);
        post({
          type: "AIS_ERROR",
          id,
          code: String((err && err.code) || "page_hook_error"),
          message: String((err && err.message) || err),
          httpStatus: Number((err && err.httpStatus) || 500),
        });
      }
    }

    async function probeDriveAuth(silent) {
      if (hookState.driveProbeInFlight) {
        return hookState.driveProbeInFlight;
      }
      if (location.hostname !== "aistudio.google.com" || !hasAuthCookies()) {
        return hookState.drive;
      }
      if (!hookState.headers["x-goog-api-key"]) {
        const domKey = detectApiKeyFromDom();
        if (domKey) hookState.headers["x-goog-api-key"] = domKey;
      }
      if (!hookState.headers["x-goog-api-key"]) {
        return hookState.drive;
      }
      hookState.driveProbeInFlight = (async () => {
        try {
          const headerPairs = await buildOrderedRequestHeaders(true);
          await new Promise((resolve) => {
            const xhr = new XMLHttpRequest();
            xhr.__veAisSynthetic = true;
            xhr.open("POST", GENERATE_ACCESS_TOKEN_URL);
            xhr.withCredentials = true;
            for (const [hName, hVal] of headerPairs) {
              xhr.setRequestHeader(hName, hVal);
            }
            xhr.onload = () => {
              try {
                recordDriveResult(
                  xhr.status,
                  xhr.responseText,
                  Boolean(silent),
                );
              } catch (_) {}
              resolve();
            };
            xhr.onerror = () => resolve();
            xhr.onabort = () => resolve();
            xhr.send(JSON.stringify(["users/me"]));
          });
        } catch (_) {
        } finally {
          hookState.driveProbeInFlight = null;
        }
        return hookState.drive;
      })();
      return hookState.driveProbeInFlight;
    }

    window.addEventListener("message", (ev) => {
      const d = ev.data;
      if (!d || d.__veAis !== true || d.dir !== "toPage") return;
      if (d.cookies && typeof d.cookies === "object") {
        Object.assign(hookState.cookies, d.cookies);
      }
      if (d.type === "AIS_PING") {
        if (hookState.drive.status === "unknown" && hasAuthCookies()) {
          probeDriveAuth(true)
            .catch(() => {})
            .finally(() => {
              emitReady({ event: "ping" });
            });
        } else {
          emitReady({ event: "ping" });
        }
        return;
      }
      if (d.type === "AIS_PROBE") {
        (async () => {
          if (d.warmup) {
            await ensureBootstrapped();
          }
          await probeDriveAuth(true);
          emitReady({ id: d.id || "", event: "probe-ok" });
        })().catch((err) => {
          emitReady({
            id: d.id || "",
            event: "probe-error",
            probeError: String((err && err.message) || err),
          });
        });
        return;
      }
      if (d.type === "AIS_WARMUP") {
        ensureBootstrapped()
          .then(() => probeDriveAuth(true).catch(() => {}))
          .then(() => {
            emitReady({ id: d.id || "", event: "warmup-ok" });
          })
          .catch((err) => {
            post({
              type: "AIS_ERROR",
              id: d.id || "",
              code: String((err && err.code) || "warmup_failed"),
              message: String((err && err.message) || err),
              httpStatus: Number((err && err.httpStatus) || 500),
            });
          });
        return;
      }
      if (d.type === "AIS_UPLOAD") {
        executeAisUpload(d);
        return;
      }
      if (d.type === "AIS_STREAM" || d.type === "AIS_CALL") {
        executeAisRequest(d);
        return;
      }
      if (d.type === "AIS_ABORT") {
        const id = String(d.id || "");
        const xhr = hookState.activeRequests.get(id);
        if (xhr) {
          hookState.activeRequests.delete(id);
          try {
            xhr.abort();
          } catch (_) {}
        }
      }
    });

    // Canonicalize bare /prompts/... paths without /u/<N>/ to /u/0/prompts/...
    // so Google multi-account routing, GetLoggingContext, WAA proof, and X-Goog-AuthUser
    // stay strictly locked to slot 0 instead of falling into an unlinked default session.
    try {
      if (
        location.hostname === "aistudio.google.com" &&
        /^\/prompts(\/|$)/.test(location.pathname || "") &&
        !detectAuthUserFromUrl()
      ) {
        const targetUrl =
          location.origin +
          "/u/0" +
          location.pathname +
          (location.search || "") +
          (location.hash || "");
        location.replace(targetUrl);
        return;
      }
    } catch (_) {}

    try {
      if (document.documentElement) {
        document.documentElement.setAttribute("data-ve-ais-hook", "ready");
      }
    } catch (_) {}
    post({ type: "AIS_HOOK_BOOTED", hookVersion: hookState.version });
    if (
      !prevHook &&
      hookState.drive.status === "unknown" &&
      hasAuthCookies() &&
      Boolean(
        detectAuthUserFromUrl() ||
        hookState.headers["x-goog-ext-519733851-bin"],
      )
    ) {
      probeDriveAuth(true)
        .catch(() => {})
        .finally(() => {
          emitReady({ event: "init" });
        });
    } else {
      emitReady({ event: "init" });
    }
  }

  // ---- Isolated-world Content Script -----------------------------------------
  let pageHookBooted = false;
  const pendingToPage = [];

  function flushPendingToPage() {
    while (pendingToPage.length > 0) {
      const item = pendingToPage.shift();
      window.postMessage(
        Object.assign({ __veAis: true, dir: "toPage" }, item),
        "*",
      );
    }
  }

  function markPageHookBooted() {
    pageHookBooted = true;
    try {
      if (document.documentElement) {
        document.documentElement.removeAttribute("data-ve-ais-hook");
      }
    } catch (_) {}
    flushPendingToPage();
  }

  function findCspNonce() {
    try {
      const el = document.querySelector("script[nonce]");
      if (!el) return "";
      return String(el.nonce || el.getAttribute("nonce") || "").trim();
    } catch (_) {
      return "";
    }
  }

  function injectPageHook(nonce) {
    if (pageHookBooted) return true;
    const code = "(" + PAGE_HOOK.toString() + ")();";

    // Strategy 1: Firefox MV2 wrappedJSObject.eval (runs directly in main world,
    // immune to CSP nonce and TrustedTypes restrictions).
    try {
      if (
        typeof window.wrappedJSObject === "object" &&
        window.wrappedJSObject &&
        typeof window.wrappedJSObject.eval === "function"
      ) {
        window.wrappedJSObject.eval(code);
        if (
          document.documentElement &&
          document.documentElement.getAttribute("data-ve-ais-hook") === "ready"
        ) {
          markPageHookBooted();
          return true;
        }
      }
    } catch (_) {}

    // Strategy 2: <script nonce="..."> with TrustedTypes policy support.
    try {
      const resolvedNonce = nonce || findCspNonce();
      const s = document.createElement("script");
      if (resolvedNonce) {
        s.nonce = resolvedNonce;
        s.setAttribute("nonce", resolvedNonce);
      }
      let scriptPayload = code;
      if (
        window.trustedTypes &&
        typeof window.trustedTypes.createPolicy === "function"
      ) {
        try {
          if (!window.__veAisTrustedPolicy) {
            window.__veAisTrustedPolicy = window.trustedTypes.createPolicy(
              "ve-aistudio-hook",
              { createScript: (input) => input },
            );
          }
          scriptPayload = window.__veAisTrustedPolicy.createScript(code);
        } catch (_) {}
      }
      s.textContent = scriptPayload;
      (document.head || document.documentElement).appendChild(s);
      s.remove();
      if (
        document.documentElement &&
        document.documentElement.getAttribute("data-ve-ais-hook") === "ready"
      ) {
        markPageHookBooted();
      }
      return pageHookBooted;
    } catch (_) {
      return false;
    }
  }

  window.addEventListener("message", (ev) => {
    const d = ev.data;
    if (!d || d.__veAis !== true || d.dir !== "fromPage") return;
    if (d.type === "AIS_HOOK_BOOTED") {
      markPageHookBooted();
      return;
    }
    try {
      browser.runtime.sendMessage(d).catch(() => {});
    } catch (_) {}
  });

  browser.runtime.onMessage.addListener((msg) => {
    if (!msg || !msg.__veAisToPage) return undefined;
    if (!pageHookBooted) {
      injectPageHook(findCspNonce());
    }
    const payload = msg.payload || {};
    if (pageHookBooted) {
      window.postMessage(
        Object.assign({ __veAis: true, dir: "toPage" }, payload),
        "*",
      );
    } else {
      pendingToPage.push(payload);
    }
    return Promise.resolve({
      ok: true,
      booted: pageHookBooted,
      tabUrl: location.href,
    });
  });

  if (!injectPageHook("")) {
    try {
      browser.runtime
        .sendMessage({
          __veAis: true,
          type: "AIS_NEED_MAIN_WORLD",
          code: "(" + PAGE_HOOK.toString() + ")();",
        })
        .catch(() => {});
    } catch (_) {}
  }

  if (!pageHookBooted && typeof MutationObserver === "function") {
    const obs = new MutationObserver(() => {
      if (pageHookBooted) {
        obs.disconnect();
        return;
      }
      const nonce = findCspNonce();
      if (nonce) {
        obs.disconnect();
        injectPageHook(nonce);
      }
    });
    if (document.documentElement) {
      obs.observe(document.documentElement, {
        childList: true,
        subtree: true,
      });
      setTimeout(() => obs.disconnect(), 15000);
    }
  }
})();
