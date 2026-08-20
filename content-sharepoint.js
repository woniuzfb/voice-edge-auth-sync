(() => {
  "use strict";

  function cleanSiteUrl(value) {
    const url = new URL(String(value || ""));
    if (url.protocol !== "https:" || !url.hostname.endsWith(".sharepoint.com"))
      throw new Error("SharePoint site URL is invalid");
    if (!/^\/sites\/[^/]+\/?$/i.test(url.pathname))
      throw new Error(
        "SharePoint site URL must be https://tenant.sharepoint.com/sites/site",
      );
    return url.origin + url.pathname.replace(/\/$/, "");
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
  // ("Wed, 21 Oct 2026 07:28:00 GMT") — the old Number()-only path silently
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

  async function getDigest(siteUrl) {
    const data = await spFetchRetry(siteUrl + "/_api/contextinfo", {
      method: "POST",
    });
    const info =
      data.GetContextWebInformation ||
      data?.d?.GetContextWebInformation ||
      data;
    if (!info.FormDigestValue)
      throw new Error("SharePoint did not return FormDigestValue");
    return {
      value: info.FormDigestValue,
      timeout: Number(info.FormDigestTimeoutSeconds || 0),
    };
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
        const bearer = String(sharePointAccessToken || "").trim();
        if (!/^Bearer\s+/i.test(bearer))
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
        const item = await spFetchRetry(driveItemUrl, {
          credentials: "omit",
          headers: {
            Accept: "application/json",
            Authorization: bearer,
            Prefer: "respond-async",
            Scenario:
              "richtexturlresolution.prefetch.getdocumentsummary.searchsuggestions",
            "Client-Request-Id": crypto.randomUUID(),
          },
        });
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
