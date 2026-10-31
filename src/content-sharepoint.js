(() => {
  "use strict";

  function inferSiteUrlFromLocation() {
    try {
      const u = new URL(location.href);
      if (u.protocol !== "https:" || !u.hostname.endsWith(".sharepoint.com"))
        return "";
      const m = u.pathname.match(/^(\/sites\/[^/]+)/i);
      return m ? u.origin + m[1] : "";
    } catch (_) {
      return "";
    }
  }

  function cleanSiteUrl(value) {
    const raw = String(value || "").trim() || inferSiteUrlFromLocation();
    const url = new URL(raw);
    if (url.protocol !== "https:" || !url.hostname.endsWith(".sharepoint.com"))
      throw new Error("SharePoint site URL is invalid");
    const m = url.pathname.match(/^(\/sites\/[^/]+)(?:\/|$)/i);
    if (!m)
      throw new Error(
        "SharePoint site URL must be https://tenant.sharepoint.com/sites/site",
      );
    return url.origin + m[1].replace(/\/$/, "");
  }

  function normalizeFolder(value) {
    let folder = String(value || "")
      .trim()
      .replace(/^\/+|\/+$/g, "");
    // Accept both the documented library-relative form (General) and older
    // extension builds that stored Shared Documents/General.
    folder = folder.replace(/^Shared Documents(?:\/|$)/i, "");
    const parts = folder.split("/").filter(Boolean);
    if (parts.some((part) => part === "." || part === ".."))
      throw new Error("SharePoint upload folder is invalid");
    return parts.join("/");
  }

  function odataString(value) {
    return String(value).replace(/'/g, "''");
  }

  async function readJson(response) {
    const text = await response.text();
    let data = {};
    try {
      data = text ? JSON.parse(text) : {};
    } catch (_) {
      data = { raw: text };
    }
    if (!response.ok) {
      const detail =
        data?.error?.message?.value ||
        data?.error?.message ||
        data.raw ||
        response.statusText ||
        `HTTP ${response.status}`;
      throw new Error(
        `SharePoint ${response.status}: ${String(detail).replace(/\s+/g, " ").slice(0, 500)}`,
      );
    }
    return data;
  }

  async function spFetch(url, options = {}) {
    const response = await fetch(url, {
      credentials: "include",
      cache: "no-store",
      ...options,
      headers: {
        Accept: "application/json;odata=nometadata",
        ...(options.headers || {}),
      },
    });
    return readJson(response);
  }

  // SharePoint Online returns transient 503 (and occasionally 429/500/502/504)
  // under concurrent multi-file load — an HTML throttling page, not a durable
  // failure. Retry those, and bare network faults, with exponential backoff +
  // jitter, honoring Retry-After when present. Every retried request in the
  // upload path is idempotent (contextinfo POST; addUsingPath uses
  // overwrite=true; the read-backs are GETs), so replay is safe. 4xx other
  // than the transient ones below (auth, not-found, bad request) is NOT retried
  // — it will not fix itself and must surface immediately.
  //   408 Request Timeout — server-side timeout, safe to replay an idempotent op.
  //   423 Locked          — SharePoint co-authoring / background job holds a
  //                         short-lived lock on the item; clears on its own.
  const RETRYABLE_STATUS = new Set([408, 423, 429, 500, 502, 503, 504]);
  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
  function backoffDelay(base, attempt) {
    return base * 2 ** attempt + Math.floor(Math.random() * 300);
  }
  // Upper bound on how long we will honor a server Retry-After. SharePoint
  // throttling frequently returns Retry-After of 60–300s; sleeping the full
  // value unbounded (a) can exceed the background completion budget
  // (sharePointUploadTimeoutMs, now capped at 600s) so the waiter times out
  // while a request is still sleeping, and (b) stalls the whole batch on one
  // hot request. Cap the honored delay and let normal exponential backoff take
  // over past the cap; the request still retries, just sooner.
  const RETRY_AFTER_MAX_MS = 60000;
  // Parse a Retry-After header into a bounded millisecond delay. Supports BOTH
  // wire forms: delta-seconds ("120") and the HTTP-date form
  // ("Wed, 21 Oct <year> 07:28:00 GMT") — the old Number()-only path silently
  // dropped the date form (Number("Wed,...") === NaN) and fell back to backoff,
  // ignoring the server's hint. Returns 0 when absent/unparseable/in the past
  // so the caller falls back to exponential backoff.
  function parseRetryAfterMs(headerValue) {
    const raw = String(headerValue || "").trim();
    if (!raw) return 0;
    const seconds = Number(raw);
    if (Number.isFinite(seconds) && seconds > 0)
      return Math.min(seconds * 1000, RETRY_AFTER_MAX_MS);
    const when = Date.parse(raw);
    if (Number.isFinite(when)) {
      const delta = when - Date.now();
      if (delta > 0) return Math.min(delta, RETRY_AFTER_MAX_MS);
    }
    return 0;
  }

  async function fetchRetry(url, options = {}, retry = {}) {
    const attempts = Math.max(1, Number(retry.attempts) || 5);
    const baseDelay = Math.max(1, Number(retry.baseDelay) || 600);
    let lastError = null;
    for (let attempt = 0; attempt < attempts; attempt += 1) {
      let response;
      try {
        response = await fetch(url, options);
      } catch (networkError) {
        // DNS/reset/offline: retry unless this was the final attempt.
        lastError = networkError;
        if (attempt < attempts - 1) {
          await sleep(backoffDelay(baseDelay, attempt));
          continue;
        }
        throw networkError;
      }
      if (RETRYABLE_STATUS.has(response.status) && attempt < attempts - 1) {
        const honored = parseRetryAfterMs(response.headers.get("Retry-After"));
        const wait = honored > 0 ? honored : backoffDelay(baseDelay, attempt);
        // Drain the throttling-page body so the connection can be reused.
        try {
          await response.arrayBuffer();
        } catch (_) {}
        await sleep(wait);
        continue;
      }
      return response;
    }
    throw lastError || new Error("SharePoint request failed after retries");
  }

  async function spFetchRetry(url, options = {}, retry = {}) {
    const response = await fetchRetry(
      url,
      {
        credentials: "include",
        cache: "no-store",
        ...options,
        headers: {
          Accept: "application/json;odata=nometadata",
          ...(options.headers || {}),
        },
      },
      retry,
    );
    return readJson(response);
  }

  // FormDigest 缓存与滑动 Cookie 续期：
  // SharePoint Online 的 POST /_api/contextinfo（带 credentials: "include"）不仅返回
  // FormDigestValue（默认有效期 1800s），还会由服务端下发 Set-Cookie: FedAuth=...;
  // rtFa=...; SIMI=...，把 FedAuth 内的会话过期时间戳向后顺延 24 小时，并重置
  // spo_abt（32400s = 9h）非活动超时计时。
  let _digestCache = {
    siteUrl: "",
    value: "",
    timeout: 0,
    exp: 0,
    updatedAt: 0,
  };
  let _selfSpoToken = {
    origin: "",
    token: "",
    exp: 0,
  };

  async function getDigest(siteUrl, force = false) {
    if (
      !force &&
      _digestCache.siteUrl === siteUrl &&
      _digestCache.value &&
      _digestCache.exp - Date.now() > 5 * 60 * 1000
    ) {
      return {
        value: _digestCache.value,
        timeout: _digestCache.timeout,
        exp: _digestCache.exp,
      };
    }
    const data = await spFetchRetry(siteUrl + "/_api/contextinfo", {
      method: "POST",
    });
    const info =
      data.GetContextWebInformation ||
      data?.d?.GetContextWebInformation ||
      data;
    if (!info.FormDigestValue)
      throw new Error("SharePoint did not return FormDigestValue");
    const timeout = Number(info.FormDigestTimeoutSeconds || 1800);
    const now = Date.now();
    const exp = now + Math.max(300, timeout - 120) * 1000;
    _digestCache = {
      siteUrl,
      value: String(info.FormDigestValue),
      timeout,
      exp,
      updatedAt: now,
    };
    return {
      value: _digestCache.value,
      timeout: _digestCache.timeout,
      exp: _digestCache.exp,
    };
  }

  // 通过 SharePoint 原生 /_api/SP.OAuth.Token/Acquire 获取站点自签发 SPO Bearer Token
  // （aud=00000003-0000-0ff1-ce00-000000000000@<tid>，appid=c58637bb-e2e1-4312-8a00-04b5ffcd3403）。
  // 既可在预续期时保活 SharePoint OAuth 会话，也可在后台 M365 FOCI token 尚未捕获或失效时
  // 作为 /_api/v2.1/shares/.../driveItem 的兜底 Bearer。
  async function acquireSelfIssuedSpoToken(siteUrl, force = false) {
    const origin = new URL(siteUrl).origin;
    if (
      !force &&
      _selfSpoToken.origin === origin &&
      _selfSpoToken.token &&
      _selfSpoToken.exp - Date.now() > 10 * 60 * 1000
    ) {
      return "Bearer " + _selfSpoToken.token;
    }
    const digest = await getDigest(siteUrl, false);
    const endpoint = siteUrl + "/_api/SP.OAuth.Token/Acquire";
    const data = await spFetchRetry(endpoint, {
      method: "POST",
      headers: {
        Accept: "application/json;odata.metadata=minimal",
        "Content-Type": "application/json;odata=verbose",
        "OData-Version": "4.0",
        "X-RequestDigest": digest.value,
      },
      body: JSON.stringify({ resource: origin, tokenType: "SPO" }),
    });
    const rawToken = String(
      data?.access_token || data?.d?.Acquire?.access_token || "",
    ).trim();
    if (!rawToken)
      throw new Error(
        "SharePoint SP.OAuth.Token/Acquire did not return access_token",
      );
    let exp = Date.now() + 50 * 60 * 1000;
    try {
      const part = rawToken
        .split(".")[1]
        .replace(/-/g, "+")
        .replace(/_/g, "/");
      const claims = JSON.parse(
        atob(part.padEnd(Math.ceil(part.length / 4) * 4, "=")),
      );
      if (Number(claims.exp) > 0) exp = Number(claims.exp) * 1000 - 60000;
    } catch (_) {}
    _selfSpoToken = { origin, token: rawToken, exp };
    return "Bearer " + rawToken;
  }

  async function testConnection(siteUrl) {
    const web = await spFetch(
      siteUrl + "/_api/web?$select=Id,Title,Url,ServerRelativeUrl",
    );
    const value = web?.d || web;
    return {
      ok: true,
      title: value.Title || "",
      url: value.Url || siteUrl,
      serverRelativeUrl: value.ServerRelativeUrl || new URL(siteUrl).pathname,
      pageUrl: location.href,
    };
  }

  async function resolveLibraryRoot(siteUrl) {
    // Do not assume the localized display name. Resolve the default Documents
    // library through GetFolderByServerRelativeUrl and fall back to the known
    // server-relative path used by this tenant.
    const sitePath = new URL(siteUrl).pathname.replace(/\/$/, "");
    const candidates = [
      sitePath + "/Shared Documents",
      sitePath + "/Documents",
    ];
    let lastError = null;
    for (const candidate of candidates) {
      const endpoint =
        siteUrl +
        "/_api/web/GetFolderByServerRelativeUrl('" +
        encodeURIComponent(odataString(candidate)) +
        "')?$select=ServerRelativeUrl,Exists";
      try {
        const data = await spFetchRetry(endpoint);
        const value = data?.d || data;
        if (value.Exists !== false && value.ServerRelativeUrl)
          return String(value.ServerRelativeUrl).replace(/\/$/, "");
      } catch (error) {
        lastError = error;
      }
    }
    throw lastError || new Error("SharePoint document library was not found");
  }

  async function uploadTest(siteUrl, folder) {
    const digest = await getDigest(siteUrl);
    const libraryRoot = await resolveLibraryRoot(siteUrl);
    const relativeFolder = normalizeFolder(folder);
    const folderPath = relativeFolder
      ? libraryRoot + "/" + relativeFolder
      : libraryRoot;
    const filename = `voice-edge-test-${new Date().toISOString().replace(/[:.]/g, "-")}.txt`;
    const endpoint =
      siteUrl +
      "/_api/web/GetFolderByServerRelativeUrl('" +
      encodeURIComponent(odataString(folderPath)) +
      "')/Files/add(url='" +
      encodeURIComponent(odataString(filename)) +
      "',overwrite=true)";
    const body = new TextEncoder().encode(
      "Voice Edge SharePoint browser-session upload test.\n",
    );
    const data = await spFetch(endpoint, {
      method: "POST",
      body,
      headers: {
        "Content-Type": "application/octet-stream",
        "X-RequestDigest": digest.value,
      },
    });
    const value = data?.d || data;
    return {
      ok: true,
      name: filename,
      folder: relativeFolder,
      serverRelativeUrl: value.ServerRelativeUrl || `${folderPath}/${filename}`,
      pageUrl: location.href,
    };
  }

  function decodeBase64(value) {
    const binary = atob(String(value || ""));
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
    return bytes;
  }
  function shareIdFromUrl(url) {
    const bytes = new TextEncoder().encode(String(url));
    let binary = "";
    for (const byte of bytes) binary += String.fromCharCode(byte);
    return (
      "u!" +
      btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "")
    );
  }
  function fileUrlFromServerRelative(siteUrl, serverRelativeUrl) {
    const origin = new URL(siteUrl).origin;
    const encodedPath = String(serverRelativeUrl)
      .split("/")
      .map((part) => encodeURIComponent(decodeURIComponent(part)))
      .join("/");
    return origin + "/:u:/r" + encodedPath + "?csf=1&web=1";
  }
  async function ensureFolder(siteUrl, folderPath) {
    const endpoint =
      siteUrl +
      "/_api/web/GetFolderByServerRelativePath(decodedurl='" +
      encodeURIComponent(odataString(folderPath)) +
      "')?$select=ServerRelativeUrl,Exists";
    const data = await spFetchRetry(endpoint);
    const value = data?.d || data;
    if (value.Exists === false || !value.ServerRelativeUrl)
      throw new Error("SharePoint upload folder does not exist: " + folderPath);
  }

  async function sha256Hex(bytes) {
    const digest = await crypto.subtle.digest("SHA-256", bytes);
    return Array.from(new Uint8Array(digest), (b) =>
      b.toString(16).padStart(2, "0"),
    ).join("");
  }
  async function readFileBytes(siteUrl, serverRelativeUrl) {
    const endpoint =
      siteUrl +
      "/_api/web/GetFileByServerRelativePath(decodedurl='" +
      encodeURIComponent(odataString(serverRelativeUrl)) +
      "')/$value";
    const readOnce = () =>
      fetchRetry(endpoint, {
        credentials: "include",
        cache: "no-store",
        headers: { Accept: "application/octet-stream" },
      });
    let response = await readOnce();
    // Read-your-own-write race: addUsingPath just created this exact file, so a
    // 404 here is almost never a durable not-found — it is the read-back racing
    // SharePoint's own indexing/replication of the write. 404 is deliberately
    // OUTSIDE RETRYABLE_STATUS (a genuine missing resource must fail fast), so
    // give ONLY the read-back ONE short extra retry before surfacing. Transient
    // throttling (429/503/etc.) is already absorbed inside fetchRetry.
    if (response.status === 404) {
      await sleep(1000);
      response = await readOnce();
    }
    if (!response.ok) {
      const text = await response.text();
      throw new Error(
        "SharePoint file read-back failed (" +
          response.status +
          "): " +
          String(text || response.statusText)
            .replace(/\s+/g, " ")
            .slice(0, 500),
      );
    }
    return new Uint8Array(await response.arrayBuffer());
  }
  async function uploadFiles(siteUrl, folder, files, sharePointAccessToken) {
    const list = Array.isArray(files) ? files : [];
    if (!list.length) return { ok: true, files: [] };
    const digest = await getDigest(siteUrl);
    const libraryRoot = await resolveLibraryRoot(siteUrl);
    const relativeFolder = normalizeFolder(folder);
    const folderPath = relativeFolder
      ? libraryRoot + "/" + relativeFolder
      : libraryRoot;
    await ensureFolder(siteUrl, folderPath);
    const uploaded = [];
    // One file's hard failure (post-retry) must not abort the whole batch:
    // a single unrecoverable attachment used to throw here and lose every
    // other file's upload. Collect per-file failures instead and let the
    // caller decide. Transient 503/429/etc. are already absorbed by
    // spFetchRetry/fetchRetry below, so `failed` only ever holds genuinely
    // unrecoverable files.
    const failed = [];
    let resolvedBearer = String(sharePointAccessToken || "").trim();
    for (const raw of list) {
      // SharePoint 文件名非法字符： \ / : * ? " < > | 以及首尾点/空白。
      // 上游 content-m365.js 现已解码 encoded-word 并清洗，这里保留同一套
      // 清洗作为兜底：任何漏网的坏名字都不会再把 addUsingPath 打挂。
      const name = String(raw?.name || "")
        .replace(/[\\/:*?"<>|]/g, "_")
        .replace(/^\.+|\.+$/g, "")
        .trim();
      try {
        if (!name) throw new Error("M365 attachment filename is empty");
        const bytes = decodeBase64(raw.data);
        const expected = Number(raw.size || bytes.byteLength);
        if (bytes.byteLength !== expected)
          throw new Error("M365 attachment size mismatch: " + name);
        const endpoint =
          siteUrl +
          "/_api/web/GetFolderByServerRelativePath(decodedurl='" +
          encodeURIComponent(odataString(folderPath)) +
          "')/Files/addUsingPath(decodedurl='" +
          encodeURIComponent(odataString(name)) +
          "',overwrite=true)";
        const data = await spFetchRetry(endpoint, {
          method: "POST",
          body: bytes,
          headers: {
            "Content-Type": "application/octet-stream",
            "X-RequestDigest": digest.value,
          },
        });
        const value = data?.d || data;
        const serverRelativeUrl = String(
          value.ServerRelativeUrl || folderPath + "/" + name,
        );
        const url = fileUrlFromServerRelative(siteUrl, serverRelativeUrl);
        const expectedSha256 = await sha256Hex(bytes);
        const downloaded = await readFileBytes(siteUrl, serverRelativeUrl);
        const actualSha256 = await sha256Hex(downloaded);
        if (
          downloaded.byteLength !== bytes.byteLength ||
          actualSha256 !== expectedSha256
        ) {
          throw new Error(
            "SharePoint read-back verification failed for " +
              name +
              ": expected " +
              bytes.byteLength +
              "/" +
              expectedSha256 +
              ", got " +
              downloaded.byteLength +
              "/" +
              actualSha256,
          );
        }
        if (!/^Bearer\s+/i.test(resolvedBearer)) {
          try {
            resolvedBearer = await acquireSelfIssuedSpoToken(siteUrl, false);
          } catch (_) {}
        }
        if (!/^Bearer\s+/i.test(resolvedBearer))
          throw new Error(
            "SharePoint bearer token is unavailable for richtext URL resolution",
          );
        const driveItemUrl =
          new URL(siteUrl).origin +
          "/_api/v2.1/shares/" +
          shareIdFromUrl(url) +
          // $select the pre-authenticated direct-download link alongside the
          // fields we already validate. "@content.downloadUrl" is a short-lived
          // (≈1h) credential-less URL that streams the file bytes directly; the
          // permalink (`url`) stays the stable, login-gated share link. Selecting
          // it here is the only way to surface it (it is not returned by default).
          "/driveItem?$select=id,name,size,file,parentReference,sharepointIds,content.downloadUrl";
        let item;
        try {
          item = await spFetchRetry(driveItemUrl, {
            credentials: "omit",
            headers: {
              Accept: "application/json",
              Authorization: resolvedBearer,
              Prefer: "respond-async",
              Scenario:
                "richtexturlresolution.prefetch.getdocumentsummary.searchsuggestions",
              "Client-Request-Id": crypto.randomUUID(),
            },
          });
        } catch (driveErr) {
          // 若传入的 FOCI token 已过期/401，回退用站点自签发 SPO Bearer 重试一次
          resolvedBearer = await acquireSelfIssuedSpoToken(siteUrl, true);
          item = await spFetchRetry(driveItemUrl, {
            credentials: "omit",
            headers: {
              Accept: "application/json",
              Authorization: resolvedBearer,
              Prefer: "respond-async",
              Scenario:
                "richtexturlresolution.prefetch.getdocumentsummary.searchsuggestions",
              "Client-Request-Id": crypto.randomUUID(),
            },
          });
        }
        if (
          !item?.id ||
          String(item.name || "") !== name ||
          !item?.parentReference?.driveId
        )
          throw new Error("SharePoint driveItem validation failed for " + name);
        if (
          Number(item.size || downloaded.byteLength) !== downloaded.byteLength
        )
          throw new Error("SharePoint driveItem size mismatch for " + name);
        const spIds = item.sharepointIds || {};
        const spoScope = [spIds.siteId, spIds.webId, spIds.listId]
          .map((value) => String(value || ""))
          .filter(Boolean);
        const listItemUniqueId = String(spIds.listItemUniqueId || "");
        const encodeBase64UrlUtf8 = (value) => {
          const bytes = new TextEncoder().encode(String(value));
          let binary = "";
          for (const byte of bytes) binary += String.fromCharCode(byte);
          return btoa(binary)
            .replace(/\+/g, "-")
            .replace(/\//g, "_")
            .replace(/=+$/, "");
        };
        const agentFileId =
          spoScope.length === 3 && listItemUniqueId
            ? "SPO_" +
              encodeBase64UrlUtf8(spoScope.join(",")) +
              "_" +
              listItemUniqueId.replace(/[{}]/g, "").toUpperCase()
            : "";
        uploaded.push({
          name,
          url,
          agentFileId,
          fileType: name.includes(".")
            ? name.split(".").pop().toLowerCase()
            : "",
          sharepointIds: spIds,
          // Pre-authenticated direct-download URL (may be "" if the tenant/item
          // does not expose it); consumers must fall back to `url` when absent.
          downloadUrl: String(item["@content.downloadUrl"] || ""),
          driveItemUrl,
          itemId: item.id,
          driveId: item.parentReference.driveId,
          size: downloaded.byteLength,
          sha256: actualSha256,
          verified: true,
          mimeType:
            item.file?.mimeType ||
            String(raw.mimeType || "application/octet-stream"),
        });
      } catch (error) {
        failed.push({
          name: name || "(unnamed)",
          error: String((error && error.message) || error)
            .replace(/\s+/g, " ")
            .slice(0, 500),
        });
      }
    }
    // ok stays true as long as at least one file uploaded (or the batch was
    // empty). The caller inspects `failed` to surface partial losses; only a
    // total wipeout (every file failed) reports ok:false so the forward path
    // can raise instead of sending a turn with zero attachments.
    const ok = uploaded.length > 0 || list.length === 0;
    return { ok, folder: relativeFolder, files: uploaded, failed };
  }

  // ---- SharePoint 页面世界保活钩子（ABT + _spPageContextInfo.tokenProvider）----
  // SharePoint Online 页面在以下三个层面维护会话与活动状态：
  //   1) 页面 _spPageContextInfo.activityBasedTimeoutEnabled = true，配合
  //      spo_abt Cookie（32400s = 9h）与 SpoSuiteLinks.ashx 的 AuthAbtEnabled=true；
  //   2) 空闲期间仅 WebWorker 每 15 分钟发一次 credentials:"omit" 的 socketIo 请求，
  //      不会携带或刷新 FedAuth / rtFa Cookie，也不会刷新 FormDigest；
  //   3) 页面主世界挂载了 window._spPageContextInfo.tokenProvider（含 refreshCookies
  //      与 getAuthToken）以及 O365Shell.Abt / AbtInteraction。
  // 注入轻量页面世界脚本，配合内容脚本的周期性 POST /_api/contextinfo 与
  // SP.OAuth.Token/Acquire，在过期前同步续期 Cookie、FormDigest、OAuth Token 与 ABT。
  function SP_PAGE_HOOK() {
    "use strict";
    if (window.__veSharePointPageHook) return;
    window.__veSharePointPageHook = true;

    function touchPageAbt() {
      try {
        if (
          window.AbtInteraction &&
          typeof window.AbtInteraction.AbtRegisterActivity === "function"
        ) {
          window.AbtInteraction.AbtRegisterActivity(true);
        }
      } catch (_) {}
      try {
        if (window.AbtUI && typeof window.AbtUI.AbtHidePrompt === "function") {
          window.AbtUI.AbtHidePrompt();
        }
      } catch (_) {}
      try {
        const promptRoot =
          document.getElementById("abtPrompt") ||
          document.querySelector("#abtPrompt");
        if (promptRoot) {
          const primaryBtn = promptRoot.querySelector(
            "button.ms-Button--primary, button[class*='PrimaryButton']",
          );
          if (primaryBtn && typeof primaryBtn.click === "function")
            primaryBtn.click();
        }
      } catch (_) {}
      try {
        window.dispatchEvent(
          new MouseEvent("mousedown", { bubbles: true, cancelable: true }),
        );
      } catch (_) {}
    }

    // 拦截页面世界可能触发的 /abt/userstate?action=SHOULD_PROMPT|SHOULD_SIGNOUT（XHR 与 fetch）
    try {
      const origFetch = window.fetch;
      if (typeof origFetch === "function" && !origFetch.__veSpAbtHook) {
        const wrappedSpFetch = async function (input, init) {
          const response = await origFetch.apply(this, arguments);
          try {
            const reqHref =
              typeof input === "string" ? input : (input && input.url) || "";
            if (
              /\/abt\/userstate\b/i.test(reqHref) &&
              /[?&]action=SHOULD_(?:PROMPT|SIGNOUT)\b/i.test(reqHref)
            ) {
              try {
                const data = await response.clone().json();
                if (
                  data &&
                  (data.result === true || data.respectResult === true)
                ) {
                  const patched = JSON.stringify(
                    Object.assign({}, data, {
                      result: false,
                      respectResult: false,
                      timeSinceActivity: 0,
                    }),
                  );
                  return new Response(patched, {
                    status: response.status || 200,
                    statusText: response.statusText || "OK",
                    headers: { "Content-Type": "application/json" },
                  });
                }
              } catch (_) {
              } finally {
                setTimeout(touchPageAbt, 0);
              }
            }
          } catch (_) {}
          return response;
        };
        Object.defineProperty(wrappedSpFetch, "__veSpAbtHook", { value: true });
        window.fetch = wrappedSpFetch;
      }
    } catch (_) {}

    try {
      const O = XMLHttpRequest.prototype.open;
      XMLHttpRequest.prototype.open = function (method, url) {
        try {
          const u = String(url || "");
          if (
            /\/abt\/userstate\b/i.test(u) &&
            /[?&]action=SHOULD_(?:PROMPT|SIGNOUT)\b/i.test(u)
          ) {
            this.addEventListener("load", function () {
              try {
                const rType = String(this.responseType || "");
                let data = null;
                if (rType === "" || rType === "text") {
                  const raw = String(this.responseText || "");
                  if (raw) data = JSON.parse(raw);
                } else if (rType === "json") {
                  data = this.response;
                }
                if (
                  data &&
                  typeof data === "object" &&
                  (data.result === true || data.respectResult === true)
                ) {
                  const patchedObj = Object.assign({}, data, {
                    result: false,
                    respectResult: false,
                    timeSinceActivity: 0,
                  });
                  const patchedStr = JSON.stringify(patchedObj);
                  try {
                    if (rType === "" || rType === "text") {
                      Object.defineProperty(this, "responseText", {
                        configurable: true,
                        get: () => patchedStr,
                      });
                    }
                    Object.defineProperty(this, "response", {
                      configurable: true,
                      get: () => (rType === "json" ? patchedObj : patchedStr),
                    });
                  } catch (_) {}
                }
              } catch (_) {
              } finally {
                setTimeout(touchPageAbt, 0);
              }
            });
          }
        } catch (_) {}
        return O.apply(this, arguments);
      };
    } catch (_) {}

    window.addEventListener("message", (ev) => {
      const d = ev && ev.data;
      if (!d || d.__veSharePointPage !== true || d.type !== "SP_PAGE_KEEPALIVE")
        return;
      touchPageAbt();
      try {
        const ctx = window._spPageContextInfo;
        const tp = ctx && ctx.tokenProvider;
        const webUrl = String(
          d.siteUrl || (ctx && ctx.webAbsoluteUrl) || "",
        ).replace(/\/$/, "");
        if (tp && webUrl) {
          if (typeof tp.refreshCookies === "function") {
            Promise.resolve(
              tp.refreshCookies({
                webAbsoluteUrl: webUrl,
                url: webUrl + "/_api/contextinfo",
                noRedirect: true,
                apiQosName: "VoiceEdgeKeepAlive",
              }),
            ).catch(() => {});
          }
          if (typeof tp.getAuthToken === "function") {
            Promise.resolve(
              tp.getAuthToken({
                webAbsoluteUrl: webUrl,
                url: webUrl + "/_api/v2.0/sites/root",
                disallowCookies: true,
                noRedirect: true,
                apiQosName: "VoiceEdgeKeepAlive",
              }),
            ).catch(() => {});
          }
        }
      } catch (_) {}
    });
  }

  try {
    const s = document.createElement("script");
    const nonceEl = document.querySelector("script[nonce]");
    const nonce = nonceEl
      ? String(nonceEl.nonce || nonceEl.getAttribute("nonce") || "").trim()
      : "";
    if (nonce) {
      s.nonce = nonce;
      s.setAttribute("nonce", nonce);
    }
    s.textContent = "(" + SP_PAGE_HOOK.toString() + ")();";
    (document.head || document.documentElement).appendChild(s);
    s.remove();
  } catch (_) {}

  // 内容脚本本地兜底周期（20m = 15m 基准周期 × 4/3，仍严格低于 30m FormDigest 寿命）；
  // 正常情况下由后台每 15m 下发的 SP_KEEPALIVE（force=true）抢先主导续期。
  const SP_KEEPALIVE_BASE_MS = 15 * 60 * 1000;
  const SP_KEEPALIVE_INTERVAL_MS = Math.round((SP_KEEPALIVE_BASE_MS * 4) / 3);
  let _spRenewing = null;
  let _lastSpRenewAt = 0;
  let _spRetryTimer = null;
  let _spRetryAttempt = 0;

  function scheduleSpRenewRetry(siteUrl, reason) {
    if (_spRetryTimer) return;
    if (_spRetryAttempt >= 6) return;
    const baseDelay = Math.min(120000, 15000 * 2 ** _spRetryAttempt);
    const delay = Math.round(baseDelay * (0.85 + Math.random() * 0.3));
    _spRetryAttempt += 1;
    _spRetryTimer = setTimeout(() => {
      _spRetryTimer = null;
      renewSharePointSession(siteUrl || "", true).catch(() => {});
    }, delay);
  }

  function clearSpRenewRetry() {
    _spRetryAttempt = 0;
    if (_spRetryTimer) {
      clearTimeout(_spRetryTimer);
      _spRetryTimer = null;
    }
  }

  async function renewSharePointSession(rawSiteUrl, force = false) {
    let siteUrl = "";
    try {
      siteUrl = cleanSiteUrl(rawSiteUrl);
    } catch (_) {
      return { ok: false, error: "invalid siteUrl" };
    }
    if (location.origin !== new URL(siteUrl).origin)
      return { ok: false, error: "origin mismatch" };
    if (typeof navigator !== "undefined" && navigator.onLine === false) {
      _lastSpRenewAt = 0;
      scheduleSpRenewRetry(siteUrl, "offline");
      browser.runtime
        .sendMessage({
          __veSharePoint: true,
          type: "SP_SESSION_RENEW_FAILED",
          at: Date.now(),
          siteUrl,
          pageUrl: location.href,
          error: "offline",
        })
        .catch(() => {});
      return { ok: false, error: "offline" };
    }
    const now = Date.now();
    if (!force && now - _lastSpRenewAt < 60 * 1000 && _digestCache.value) {
      return {
        ok: true,
        siteUrl,
        at: _lastSpRenewAt,
        digestTimeout: _digestCache.timeout,
      };
    }
    if (_spRenewing) return _spRenewing;
    _spRenewing = (async () => {
      try {
        window.postMessage(
          {
            __veSharePointPage: true,
            type: "SP_PAGE_KEEPALIVE",
            siteUrl,
          },
          "*",
        );
      } catch (_) {}
      let digest;
      try {
        digest = await getDigest(siteUrl, true);
      } catch (digestError) {
        _lastSpRenewAt = 0;
        const errMsg = String(
          (digestError && digestError.message) || digestError || "digest failed",
        );
        scheduleSpRenewRetry(siteUrl, errMsg);
        browser.runtime
          .sendMessage({
            __veSharePoint: true,
            type: "SP_SESSION_RENEW_FAILED",
            at: Date.now(),
            siteUrl,
            pageUrl: location.href,
            error: errMsg,
          })
          .catch(() => {});
        throw digestError;
      }
      let spoTokenError = "";
      try {
        await acquireSelfIssuedSpoToken(siteUrl, false);
      } catch (tokErr) {
        spoTokenError = String((tokErr && tokErr.message) || tokErr || "");
      }
      clearSpRenewRetry();
      _lastSpRenewAt = Date.now();
      const result = {
        ok: true,
        siteUrl,
        pageUrl: location.href,
        at: _lastSpRenewAt,
        digestTimeout: digest.timeout,
        digestExp: digest.exp,
        spoTokenExp: _selfSpoToken.exp || 0,
        spoTokenError,
      };
      browser.runtime
        .sendMessage({
          __veSharePoint: true,
          type: "SP_SESSION_RENEWED",
          ...result,
        })
        .catch(() => {});
      return result;
    })();
    try {
      return await _spRenewing;
    } finally {
      _spRenewing = null;
    }
  }

  // 自动定时预续期（每 20 分钟刷新一次，早于 30 分钟 FormDigest 过期与 9 小时 spo_abt 超时）
  setTimeout(() => {
    renewSharePointSession("", true).catch(() => {});
  }, 15000);
  setInterval(() => {
    renewSharePointSession("", true).catch(() => {});
  }, SP_KEEPALIVE_INTERVAL_MS);
  window.addEventListener("focus", () => {
    renewSharePointSession("", false).catch(() => {});
  });
  window.addEventListener("online", () => {
    clearSpRenewRetry();
    renewSharePointSession("", true).catch(() => {});
  });
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") {
      renewSharePointSession("", false).catch(() => {});
    }
  });

  browser.runtime.onMessage.addListener((message) => {
    if (!message || message.__veSharePoint !== true) return;
    let siteUrl;
    try {
      siteUrl = cleanSiteUrl(message.siteUrl);
    } catch (error) {
      return Promise.resolve({
        ok: false,
        error: String(error.message || error),
      });
    }
    if (location.origin !== new URL(siteUrl).origin) return;
    if (message.type === "SP_KEEPALIVE")
      return renewSharePointSession(siteUrl, true).catch((error) => ({
        ok: false,
        error: String(error.message || error),
      }));
    if (message.type === "SP_TEST")
      return testConnection(siteUrl).catch((error) => ({
        ok: false,
        error: String(error.message || error),
      }));
    if (message.type === "SP_UPLOAD_TEST")
      return uploadTest(siteUrl, message.uploadFolder).catch((error) => ({
        ok: false,
        error: String(error.message || error),
      }));
    if (message.type === "SP_UPLOAD_FILES") {
      const requestId = String(message.requestId || "");
      return uploadFiles(
        siteUrl,
        message.uploadFolder,
        message.files,
        message.sharePointAccessToken,
      )
        .then(async (response) => {
          const completionMessage = {
            __veSharePoint: true,
            type: "SP_UPLOAD_COMPLETE",
            requestId,
            response,
          };
          browser.runtime
            .sendMessage(completionMessage)
            .then((reply) => {})
            .catch((error) => {});
          return { ok: true, accepted: true, requestId };
        })
        .catch(async (error) => {
          const response = { ok: false, error: String(error.message || error) };
          browser.runtime
            .sendMessage({
              __veSharePoint: true,
              type: "SP_UPLOAD_COMPLETE",
              requestId,
              response,
            })
            .then((reply) => {})
            .catch((relayError) => {});
          return { ok: false, accepted: true, requestId };
        });
    }
  });

  browser.runtime
    .sendMessage({
      __veSharePoint: true,
      type: "SP_FRAME_READY",
      pageUrl: location.href,
    })
    .catch(() => {});
})();
