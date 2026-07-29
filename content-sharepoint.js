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

  async function getDigest(siteUrl) {
    const data = await spFetch(siteUrl + "/_api/contextinfo", {
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
        const data = await spFetch(endpoint);
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
    const data = await spFetch(endpoint);
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
    const response = await fetch(endpoint, {
      credentials: "include",
      cache: "no-store",
      headers: { Accept: "application/octet-stream" },
    });
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
    for (const raw of list) {
      const name = String(raw?.name || "")
        .replace(/[\\/]/g, "_")
        .trim();
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
      const data = await spFetch(endpoint, {
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
        "/driveItem?$select=id,name,size,file,parentReference,content.downloadUrl";
      const item = await spFetch(driveItemUrl, {
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
      if (Number(item.size || downloaded.byteLength) !== downloaded.byteLength)
        throw new Error("SharePoint driveItem size mismatch for " + name);
      uploaded.push({
        name,
        url,
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
    }
    return { ok: true, folder: relativeFolder, files: uploaded };
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
