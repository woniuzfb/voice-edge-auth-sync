"use strict";

// Load the WebExtension shim first. On Firefox it is also listed in the
// manifest background.scripts array and self-detects as a no-op; the
// importScripts() call only matters inside a Chromium MV3 service worker and
// harmlessly throws (caught) in a Firefox background page.
try {
  importScripts("browser-polyfill.js");
} catch (e) {}

// MV2 exposes browser_action; MV3 exposes action. Pick whichever exists.
const actionApi = browser.action || browser.browserAction;

const NATIVE_HOST = "com.voice_edge.auth_bridge";
const COMPLETION_FILTER = {
  urls: ["https://www.doubao.com/chat/completion*"],
};
const MAX_CAPTURE_BYTES = 4 * 1024 * 1024;
const AUTO_SYNC_DELAY_MS = 350;

let nativePort = null;
let reconnectTimer = null;
let authRequired = false;
let authRequiredSince = 0;
let syncInFlight = false;
let lastCaptured = null;
let lastAppliedKey = "";
let pendingAutoSyncTimer = null;
let qwenState = null;
let qwenPollTimer = null;
const QWEN_NOTIFICATION_ID = "voice-edge-qwen-verification";
const QWEN_POLL_INTERVAL_MS = 1000;
const QWEN_POLL_TIMEOUT_MS = 10 * 60 * 1000;
const DEEPSEEK_NOTIFICATION_ID = "voice-edge-deepseek-login";
let deepseekState = null;
let deepseekPollTimer = null;
let deepseekAuthorization = "";
let deepseekHifLeim = "";
let lastStateReport = null;
let lastStateRequestAt = 0;

function compactError(error) {
  return String((error && (error.message || error)) || "unknown error")
    .replace(/\s+/g, " ")
    .slice(0, 500);
}

async function notify(title, message) {
  try {
    await browser.notifications.create({
      type: "basic",
      title,
      message,
    });
  } catch (_) {
    // Notifications are best-effort only.
  }
}

async function setBadge(text, color) {
  try {
    await actionApi.setBadgeText({ text });
    if (color) {
      await actionApi.setBadgeBackgroundColor({ color });
    }
  } catch (_) {
    // Badge updates are best-effort only.
  }
}

function scheduleReconnect() {
  if (reconnectTimer) return;
  reconnectTimer = setTimeout(() => {
    reconnectTimer = null;
    connectNative();
  }, 1500);
}

function connectNative() {
  if (nativePort) return nativePort;
  try {
    const port = browser.runtime.connectNative(NATIVE_HOST);
    nativePort = port;
    port.onMessage.addListener(handleNativeMessage);
    port.onDisconnect.addListener(async () => {
      const error = browser.runtime.lastError;
      nativePort = null;
      await setBadge("!", "#d93025");
      if (error && error.message) {
        await notify(
          "Voice Edge Native Host 已断开",
          compactError(error.message),
        );
      }
      scheduleReconnect();
    });
    port.postMessage({ type: "PING" });
    return port;
  } catch (error) {
    nativePort = null;
    scheduleReconnect();
    return null;
  }
}

function normalizeCookie(cookie, fallbackStoreId) {
  return {
    name: cookie.name,
    value: cookie.value,
    domain: cookie.domain,
    hostOnly: Boolean(cookie.hostOnly),
    path: cookie.path || "/",
    secure: Boolean(cookie.secure),
    httpOnly: Boolean(cookie.httpOnly),
    sameSite: cookie.sameSite || null,
    session: Boolean(cookie.session),
    expirationDate:
      cookie.expirationDate == null ? null : cookie.expirationDate,
    storeId: cookie.storeId || fallbackStoreId || "",
    firstPartyDomain: cookie.firstPartyDomain || null,
    partitionKey: cookie.partitionKey || null,
  };
}

function cookieIdentity(cookie) {
  return `${cookie.name}\u0000${cookie.domain}\u0000${cookie.path || "/"}`;
}

function cookieFingerprint(cookies) {
  return (cookies || [])
    .map(
      (cookie) =>
        `${cookieIdentity(cookie)}\u0000${cookie.value}\u0000${cookie.expirationDate || ""}`,
    )
    .sort()
    .join("\n");
}

function qwenDomains(state) {
  const domains = new Set(["qwen.ai", "chat.qwen.ai"]);
  try {
    const host = new URL(state.verificationUrl).hostname;
    if (host) domains.add(host);
  } catch (_) {}
  return Array.from(domains);
}

async function chooseQwenCookieStore() {
  const tabs = await browser.tabs.query({
    url: ["https://chat.qwen.ai/*", "https://*.qwen.ai/*"],
  });
  const active = tabs.find((tab) => tab.active) || tabs[0];
  return {
    tabId: active ? active.id : null,
    storeId: active ? active.cookieStoreId || "" : "",
  };
}

async function collectQwenCookies(state) {
  const output = new Map();
  for (const domain of qwenDomains(state)) {
    const query = { domain };
    if (state.cookieStoreId) query.storeId = state.cookieStoreId;
    let cookies = [];
    try {
      cookies = await browser.cookies.getAll(query);
    } catch (_) {
      continue;
    }
    for (const raw of cookies) {
      const cookie = normalizeCookie(raw, state.cookieStoreId);
      output.set(cookieIdentity(cookie), cookie);
    }
  }
  return Array.from(output.values());
}

async function openQwenVerification() {
  if (!qwenState || !qwenState.verificationUrl) return;
  const options = { url: qwenState.verificationUrl, active: true };
  if (qwenState.cookieStoreId) options.cookieStoreId = qwenState.cookieStoreId;
  const tab = await browser.tabs.create(options);
  qwenState.verificationTabId = tab.id;
}

async function sendQwenSnapshot(cookies) {
  if (!qwenState || qwenState.syncInFlight) return;
  const port = connectNative();
  if (!port) throw new Error("无法连接 Voice Edge Native Host");
  qwenState.syncInFlight = true;
  await setBadge("…", "#1a73e8");
  port.postMessage({
    type: "AUTH_SNAPSHOT",
    provider: "qwen",
    account_id: qwenState.accountId,
    captured_at: Date.now(),
    verification_url: qwenState.verificationUrl,
    cookie_store_id: qwenState.cookieStoreId || "",
    cookies,
  });
}

async function pollQwenCookies() {
  if (!qwenState) return;
  if (Date.now() - qwenState.startedAt > QWEN_POLL_TIMEOUT_MS) {
    clearInterval(qwenPollTimer);
    qwenPollTimer = null;
    await setBadge("!", "#d93025");
    await notify(
      "Voice Edge：Qwen 验证超时",
      "未检测到新的 x5sec，请重新触发验证后再试。",
    );
    return;
  }
  const cookies = await collectQwenCookies(qwenState);
  const x5secCookies = cookies.filter(
    (cookie) => cookie.name === "x5sec" && cookie.value,
  );
  const hasX5sec = x5secCookies.length > 0;
  const fingerprint = cookieFingerprint(cookies);
  const x5secFingerprint = cookieFingerprint(x5secCookies);
  if (
    !hasX5sec ||
    !fingerprint ||
    !x5secFingerprint ||
    x5secFingerprint === qwenState.baselineX5secFingerprint
  )
    return;
  clearInterval(qwenPollTimer);
  qwenPollTimer = null;
  await sendQwenSnapshot(cookies);
}

async function beginQwenVerification(message) {
  if (qwenPollTimer) {
    clearInterval(qwenPollTimer);
    qwenPollTimer = null;
  }
  const selected = await chooseQwenCookieStore();
  qwenState = {
    accountId: String(message.account_id || ""),
    verificationUrl: String(message.verification_url || ""),
    cookieStoreId: selected.storeId,
    sourceTabId: selected.tabId,
    verificationTabId: null,
    startedAt: Date.now(),
    baselineFingerprint: "",
    baselineX5secFingerprint: "",
    syncInFlight: false,
  };
  if (!qwenState.accountId || !qwenState.verificationUrl) {
    throw new Error("Qwen 验证消息缺少 account_id 或 verification_url");
  }
  const baselineCookies = await collectQwenCookies(qwenState);
  qwenState.baselineFingerprint = cookieFingerprint(baselineCookies);
  qwenState.baselineX5secFingerprint = cookieFingerprint(
    baselineCookies.filter((cookie) => cookie.name === "x5sec"),
  );
  await setBadge("!", "#d93025");
  await browser.notifications.create(QWEN_NOTIFICATION_ID, {
    type: "basic",
    title: "Voice Edge：Qwen 需要验证",
    message: "点击通知打开验证页面。验证完成后，相关 Cookie 会自动同步。",
  });
  qwenPollTimer = setInterval(() => {
    pollQwenCookies().catch((error) =>
      notify("Voice Edge Qwen 同步失败", compactError(error)),
    );
  }, QWEN_POLL_INTERVAL_MS);
}

browser.webRequest.onBeforeSendHeaders.addListener(
  (details) => {
    for (const header of details.requestHeaders || []) {
      const name = String(header.name || "").toLowerCase();
      const value = String(header.value || "").trim();
      if (name === "authorization" && value.toLowerCase().startsWith("bearer "))
        deepseekAuthorization = value;
      if (name === "x-hif-leim" && value) deepseekHifLeim = value;
    }
  },
  { urls: ["https://chat.deepseek.com/api/v0/*"] },
  ["requestHeaders"],
);
async function chooseDeepSeekCookieStore() {
  const tabs = await browser.tabs.query({
    url: ["https://chat.deepseek.com/*"],
  });
  const active = tabs.find((tab) => tab.active) || tabs[0];
  return {
    tabId: active ? active.id : null,
    storeId: active ? active.cookieStoreId || "" : "",
  };
}
async function collectDeepSeekCookies(state) {
  const query = { domain: "deepseek.com" };
  if (state.cookieStoreId) query.storeId = state.cookieStoreId;
  return (await browser.cookies.getAll(query)).map((c) =>
    normalizeCookie(c, state.cookieStoreId),
  );
}
async function sendDeepSeekSnapshot() {
  if (!deepseekState || deepseekState.syncInFlight) return;
  const cookies = await collectDeepSeekCookies(deepseekState);
  if (!cookies.length) throw new Error("未捕获到 DeepSeek Cookie");
  if (!deepseekAuthorization.toLowerCase().startsWith("bearer "))
    throw new Error("未捕获到 DeepSeek Authorization；请刷新页面或新建会话");
  const port = connectNative();
  if (!port) throw new Error("无法连接 Voice Edge Native Host");
  deepseekState.syncInFlight = true;
  await setBadge("…", "#1a73e8");
  port.postMessage({
    type: "AUTH_SNAPSHOT",
    provider: "deepseek",
    captured_at: Date.now(),
    cookie_store_id: deepseekState.cookieStoreId || "",
    cookies,
    authorization: deepseekAuthorization,
    x_hif_leim: deepseekHifLeim || "",
  });
}
async function pollDeepSeekAuth() {
  if (!deepseekState || deepseekState.syncInFlight) return;
  if (Date.now() - deepseekState.startedAt > 10 * 60 * 1000) {
    clearInterval(deepseekPollTimer);
    deepseekPollTimer = null;
    await setBadge("!", "#d93025");
    await notify("Voice Edge：DeepSeek 登录超时", "请重新触发认证后再试。");
    return;
  }
  const cookies = await collectDeepSeekCookies(deepseekState);
  const hasCookies = cookies.some(
    (cookie) => cookie && cookie.name && cookie.value,
  );
  const hasAuthorization = deepseekAuthorization
    .toLowerCase()
    .startsWith("bearer ");
  if (!hasCookies || !hasAuthorization) return;

  // AUTH_REQUIRED means the server needs a usable snapshot now. Do not wait
  // for Cookie/Authorization to differ from the baseline: Firefox may already
  // hold a valid logged-in DeepSeek session when the request arrives.
  clearInterval(deepseekPollTimer);
  deepseekPollTimer = null;
  await sendDeepSeekSnapshot();
}
async function beginDeepSeekLogin(message) {
  if (deepseekPollTimer) clearInterval(deepseekPollTimer);
  const selected = await chooseDeepSeekCookieStore();
  deepseekState = {
    cookieStoreId: selected.storeId,
    tabId: selected.tabId,
    startedAt: Date.now(),
    loginUrl: String(message.login_url || "https://chat.deepseek.com/"),
    syncInFlight: false,
    baselineAuthorization: deepseekAuthorization,
    baselineFingerprint: "",
  };
  deepseekState.baselineFingerprint = cookieFingerprint(
    await collectDeepSeekCookies(deepseekState),
  );
  await setBadge("!", "#d93025");
  await browser.notifications.create(DEEPSEEK_NOTIFICATION_ID, {
    type: "basic",
    title: "Voice Edge：DeepSeek 需要认证",
    message:
      "正在检查当前 Firefox 登录状态；若未登录，请点击通知打开 DeepSeek。",
  });

  // Try immediately. Previously the extension required credentials to change
  // after AUTH_REQUIRED, so an already logged-in browser stayed red forever.
  await pollDeepSeekAuth();
  if (deepseekState && !deepseekState.syncInFlight && !deepseekPollTimer) {
    deepseekPollTimer = setInterval(
      () =>
        pollDeepSeekAuth().catch(async (e) => {
          await setBadge("!", "#d93025");
          await notify("Voice Edge DeepSeek 同步失败", compactError(e));
        }),
      1000,
    );
  }
}
browser.runtime.onMessage.addListener((message) => {
  if (!message || message.type !== "DEEPSEEK_CONVERSATION_SNAPSHOT") return;
  const port = connectNative();
  if (!port) return;
  port.postMessage({
    type: "CONVERSATION_SNAPSHOT",
    provider: "deepseek",
    conversation_id: String(message.conversation_id || ""),
    model_type: String(message.model_type || ""),
    page_url: String(message.page_url || ""),
    captured_at: Number(message.captured_at || Date.now()),
  });
});

browser.notifications.onClicked.addListener(async (notificationId) => {
  if (notificationId === DEEPSEEK_NOTIFICATION_ID) {
    try {
      const options = {
        url: deepseekState
          ? deepseekState.loginUrl
          : "https://chat.deepseek.com/",
        active: true,
      };
      if (deepseekState && deepseekState.cookieStoreId)
        options.cookieStoreId = deepseekState.cookieStoreId;
      await browser.tabs.create(options);
    } catch (error) {
      await notify("Voice Edge：无法打开 DeepSeek", compactError(error));
    }
    return;
  }
  if (notificationId === QWEN_NOTIFICATION_ID) {
    try {
      await openQwenVerification();
    } catch (error) {
      await notify("Voice Edge：无法打开 Qwen 验证页面", compactError(error));
    }
  }
});

async function handleNativeMessage(message) {
  if (!message || typeof message !== "object") return;
  const type = String(message.type || "");
  if (type === "PONG" || type === "CONVERSATION_APPLIED") return;
  if (type === "STATE_REPORT") {
    lastStateReport = message;
    return;
  }
  if (type === "AUTH_REQUIRED" && message.provider === "deepseek") {
    try {
      await beginDeepSeekLogin(message);
    } catch (error) {
      await setBadge("!", "#d93025");
      await notify("Voice Edge：DeepSeek 登录准备失败", compactError(error));
    }
    return;
  }
  if (type === "AUTH_APPLIED" && message.provider === "deepseek") {
    if (deepseekState) deepseekState.syncInFlight = false;
    if (message.success) {
      await setBadge(
        message.validated ? "✓" : "↻",
        message.validated ? "#188038" : "#f9ab00",
      );
      await notify(
        "Voice Edge：DeepSeek 同步完成",
        message.validated
          ? "认证已验证，请重试请求。"
          : "认证已保存，请重试请求。",
      );
      deepseekState = null;
      setTimeout(() => setBadge("", null), 5000);
    } else await setBadge("!", "#d93025");
    return;
  }
  if (type === "AUTH_REQUIRED" && message.provider === "qwen") {
    try {
      await beginQwenVerification(message);
    } catch (error) {
      await setBadge("!", "#d93025");
      await notify("Voice Edge：Qwen 验证准备失败", compactError(error));
    }
    return;
  }
  if (type === "AUTH_APPLIED" && message.provider === "qwen") {
    if (message.success && message.validated) {
      if (qwenPollTimer) clearInterval(qwenPollTimer);
      qwenPollTimer = null;
      if (qwenState) qwenState.syncInFlight = false;
      await setBadge("✓", "#188038");
      await notify(
        "Voice Edge：Qwen 同步成功",
        `已同步 ${message.cookie_count || 0} 个 Cookie，后续请求将使用最新状态。`,
      );
      qwenState = null;
      setTimeout(() => setBadge("", null), 5000);
    } else {
      if (qwenState) qwenState.syncInFlight = false;
      await setBadge("!", "#d93025");
    }
    return;
  }

  if (type === "AUTH_REQUIRED" && message.provider === "doubao") {
    authRequired = true;
    authRequiredSince = Date.now();
    lastAppliedKey = "";
    await setBadge("!", "#d93025");
    await notify(
      "Voice Edge：豆包状态需要刷新",
      "请在 Firefox 豆包中继续当前对话或新建对话，发送一条消息。网页完整回复后将自动同步，无需点击扩展。",
    );
    return;
  }

  if (type === "AUTH_APPLIED" && message.provider === "doubao") {
    syncInFlight = false;
    if (message.success && message.validated) {
      authRequired = false;
      authRequiredSince = 0;
      if (lastCaptured) lastAppliedKey = captureKey(lastCaptured);
      await setBadge("✓", "#188038");
      const replay = message.question_replayed
        ? "已自动重新提交刚才的问题。"
        : "小爱已恢复。";
      await notify(
        "Voice Edge 同步成功",
        `豆包状态已验证并自动同步（${message.cookie_count || 0} 个 Cookie）。${replay}`,
      );
      setTimeout(() => setBadge("", null), 5000);
    } else {
      authRequired = true;
      await setBadge("!", "#d93025");
    }
    return;
  }

  if (type === "AUTH_ERROR") {
    if (message.provider === "deepseek") {
      if (deepseekState) deepseekState.syncInFlight = false;
      await setBadge("!", "#d93025");
      await notify(
        "Voice Edge DeepSeek 同步失败",
        compactError(message.message),
      );
      return;
    }
    if (message.provider === "qwen") {
      if (qwenState) qwenState.syncInFlight = false;
      await setBadge("!", "#d93025");
      await notify("Voice Edge Qwen 同步失败", compactError(message.message));
      return;
    }
    syncInFlight = false;
    authRequired = true;
    await setBadge("!", "#d93025");
    await notify("Voice Edge 自动同步失败", compactError(message.message));
  }
}

function captureKey(capture) {
  const c = (capture && capture.conversation) || {};
  return `${c.conversation_id || ""}:${c.section_id || ""}:${c.last_message_index || 0}`;
}

function safeInt(value) {
  const number = Number(value);
  return Number.isFinite(number) && number >= 0 ? Math.trunc(number) : null;
}

function updateMax(state, value) {
  const parsed = safeInt(value);
  if (parsed !== null)
    state.lastMessageIndex = Math.max(state.lastMessageIndex, parsed);
}

function parseSseBody(body) {
  const state = {
    conversationId: "",
    sectionId: "",
    lastMessageIndex: 0,
    endType3: false,
    streamError: null,
  };

  const events = String(body || "").split(/\r?\n\r?\n/);
  for (const block of events) {
    if (!block.trim()) continue;
    let eventName = "";
    const dataLines = [];
    for (const line of block.split(/\r?\n/)) {
      if (line.startsWith("event:")) eventName = line.slice(6).trim();
      else if (line.startsWith("data:"))
        dataLines.push(line.slice(5).trimStart());
    }
    if (!eventName || !dataLines.length) continue;
    let data;
    try {
      data = JSON.parse(dataLines.join("\n"));
    } catch (_) {
      continue;
    }

    if (eventName === "SSE_ACK") {
      const meta = (data && data.ack_client_meta) || {};
      if (meta.conversation_id)
        state.conversationId = String(meta.conversation_id);
      if (meta.section_id) state.sectionId = String(meta.section_id);
      const queries = Array.isArray(data.query_list) ? data.query_list : [];
      for (const query of queries)
        updateMax(state, query && query.message_index);
    } else if (eventName === "FULL_MSG_NOTIFY") {
      updateMax(state, data && data.message && data.message.index_in_conv);
    } else if (eventName === "STREAM_MSG_NOTIFY") {
      updateMax(state, data && data.meta && data.meta.index_in_conv);
    } else if (eventName === "SSE_REPLY_END") {
      if (safeInt(data && data.end_type) === 3) state.endType3 = true;
    } else if (eventName === "STREAM_ERROR") {
      state.streamError = {
        code: data && data.error_code,
        message: data && data.error_msg,
      };
    }
  }
  return state;
}

function scheduleAutoSync(capture) {
  if (
    !authRequired ||
    syncInFlight ||
    !capture ||
    !capture.complete ||
    capture.error
  )
    return;
  if (
    !authRequiredSince ||
    capture.startedAt < authRequiredSince ||
    capture.completedAt < authRequiredSince
  )
    return;
  const key = captureKey(capture);
  if (!key || key === lastAppliedKey) return;
  if (pendingAutoSyncTimer) clearTimeout(pendingAutoSyncTimer);
  pendingAutoSyncTimer = setTimeout(() => {
    pendingAutoSyncTimer = null;
    synchronizeCapture(capture, true).catch(async (error) => {
      syncInFlight = false;
      authRequired = true;
      await setBadge("!", "#d93025");
      await notify("Voice Edge 自动同步失败", compactError(error));
    });
  }, AUTO_SYNC_DELAY_MS);
}

async function collectCookiesForTab(tabId) {
  const tab = await browser.tabs.get(tabId);
  const storeId = tab.cookieStoreId || "";
  const query = { domain: "doubao.com" };
  if (storeId) query.storeId = storeId;
  const cookies = await browser.cookies.getAll(query);
  return cookies.map((cookie) => ({
    name: cookie.name,
    value: cookie.value,
    domain: cookie.domain,
    hostOnly: Boolean(cookie.hostOnly),
    path: cookie.path || "/",
    secure: Boolean(cookie.secure),
    httpOnly: Boolean(cookie.httpOnly),
    sameSite: cookie.sameSite || null,
    session: Boolean(cookie.session),
    expirationDate:
      cookie.expirationDate == null ? null : cookie.expirationDate,
    storeId: cookie.storeId || storeId || "",
    firstPartyDomain: cookie.firstPartyDomain || null,
    partitionKey: cookie.partitionKey || null,
  }));
}

function findTtwidInCookies(cookies) {
  for (const cookie of cookies || []) {
    if (cookie && cookie.name === "ttwid" && cookie.value) {
      return cookie;
    }
  }
  return null;
}

// cookies.getAll 可能因分区（CHIPS/Total Cookie Protection）丢失 ttwid。
// 兜底从页面 document.cookie 读取 ttwid（ttwid 通常不是 HttpOnly）。
async function fetchTtwidFromTab(tabId) {
  try {
    // MV3 forbids code strings; use scripting.executeScript({func}) when present,
    // otherwise fall back to the MV2 tabs.executeScript({code}).
    if (browser.scripting && browser.scripting.executeScript) {
      const results = await browser.scripting.executeScript({
        target: { tabId },
        func: () =>
          (document.cookie.match(/(?:^|;\s*)ttwid=([^;]+)/) || [])[1] || "",
      });
      const value = results && results[0] && results[0].result;
      return typeof value === "string" ? value.trim() : "";
    }
    const results = await browser.tabs.executeScript(tabId, {
      code: "(document.cookie.match(/(?:^|;\\s*)ttwid=([^;]+)/)||[])[1]||''",
      runAt: "document_idle",
    });
    const value = results && results[0];
    return typeof value === "string" ? value.trim() : "";
  } catch (_) {
    return "";
  }
}

// 强制确保 ttwid 在同步快照中：优先用 cookies API 的结果，缺失时从
// document.cookie 兜底取回并补回 cookies 数组。返回最终的 ttwid 值。
async function ensureTtwidCaptured(cookies, tabId, cookieStoreId) {
  const existing = findTtwidInCookies(cookies);
  if (existing) return existing.value;
  const fallback = await fetchTtwidFromTab(tabId);
  if (!fallback) return "";
  cookies.push({
    name: "ttwid",
    value: fallback,
    domain: ".doubao.com",
    hostOnly: false,
    path: "/",
    secure: true,
    httpOnly: false,
    sameSite: null,
    session: false,
    expirationDate: null,
    storeId: cookieStoreId || "",
    firstPartyDomain: null,
    partitionKey: null,
  });
  return fallback;
}

async function synchronizeCapture(capture, automatic) {
  if (!capture || !capture.complete || capture.error) {
    throw new Error("尚未捕获到完整且成功的豆包回复");
  }
  if (syncInFlight) return;
  const port = connectNative();
  if (!port) throw new Error("无法连接 Voice Edge Native Host");

  syncInFlight = true;
  await setBadge("…", "#1a73e8");
  const cookies = await collectCookiesForTab(capture.tabId);
  if (!cookies.some((item) => item.name === "sessionid")) {
    syncInFlight = false;
    throw new Error("当前 Firefox Cookie Store 中没有豆包 sessionid");
  }

  // 强制把 ttwid 纳入同步快照（cookies API 缺失时从 document.cookie 兜底），
  // 并在消息顶层显式带上 ttwid，供 Voice Edge 强制覆盖回 Camoufox。
  const ttwidValue = await ensureTtwidCaptured(
    cookies,
    capture.tabId,
    capture.cookieStoreId || "",
  );

  port.postMessage({
    type: "AUTH_SNAPSHOT",
    provider: "doubao",
    captured_at: Date.now(),
    automatic: Boolean(automatic),
    page_url: `https://www.doubao.com/chat/${capture.conversation.conversation_id}`,
    cookie_store_id: capture.cookieStoreId || "",
    conversation: capture.conversation,
    cookies,
    ttwid: ttwidValue || "",
  });
}

// Shared finalizer for a captured Doubao /chat/completion SSE body. Both the
// Firefox response-stream path and the Chromium page-relay path feed into this.
async function ingestDoubaoBody(
  body,
  tabId,
  cookieStoreId,
  startedAt,
  completedAt,
) {
  if (!body || body.length > MAX_CAPTURE_BYTES) return;
  const state = parseSseBody(body);
  const complete = Boolean(
    state.conversationId &&
    state.sectionId &&
    state.lastMessageIndex > 0 &&
    state.endType3 &&
    !state.streamError,
  );
  const capture = {
    tabId,
    cookieStoreId: cookieStoreId || "",
    startedAt: Number(startedAt || Date.now()),
    completedAt: Number(completedAt || Date.now()),
    complete,
    error: state.streamError,
    conversation: {
      conversation_id: state.conversationId,
      section_id: state.sectionId,
      last_message_index: state.lastMessageIndex,
    },
  };
  lastCaptured = capture;
  if (complete) {
    // A green badge has one unambiguous meaning: Voice Edge requested an auth
    // refresh and this completed response is about to be synced.
    if (authRequired) {
      await setBadge("\u2713", "#188038");
      scheduleAutoSync(capture);
    }
  } else if (state.streamError && authRequired) {
    await setBadge("!", "#d93025");
  }
}

// Doubao completion capture — platform-detected.
//   Firefox: read the response body directly via filterResponseData.
//   Chromium: no response-body API, so content-doubao.js hooks fetch/XHR in the
//             page world and relays the raw SSE text as a DOUBAO_CAPTURE message.
// Both call the identical ingestDoubaoBody() above.
if (
  browser.webRequest &&
  typeof browser.webRequest.filterResponseData === "function"
) {
  browser.webRequest.onBeforeRequest.addListener(
    (details) => {
      if (details.method !== "POST" || details.tabId < 0) return;
      let filter;
      try {
        filter = browser.webRequest.filterResponseData(details.requestId);
      } catch (_) {
        return;
      }

      const decoder = new TextDecoder();
      let body = "";
      let capturedBytes = 0;
      let overflow = false;

      filter.ondata = (event) => {
        filter.write(event.data);
        if (overflow) return;
        capturedBytes += event.data.byteLength;
        if (capturedBytes > MAX_CAPTURE_BYTES) {
          overflow = true;
          body = "";
          return;
        }
        body += decoder.decode(event.data, { stream: true });
      };

      filter.onstop = async () => {
        try {
          body += decoder.decode();
          filter.close();
          if (overflow) return;
          const tab = await browser.tabs.get(details.tabId);
          await ingestDoubaoBody(
            body,
            details.tabId,
            tab.cookieStoreId || "",
            Number(details.timeStamp || Date.now()),
            Date.now(),
          );
        } catch (error) {
          try {
            filter.close();
          } catch (_) {}
        }
      };

      filter.onerror = () => {
        try {
          filter.close();
        } catch (_) {}
      };
    },
    COMPLETION_FILTER,
    ["blocking"],
  );
} else {
  browser.runtime.onMessage.addListener((message, sender) => {
    if (!message || message.type !== "DOUBAO_CAPTURE") return;
    const tab = sender && sender.tab;
    const tabId = tab && typeof tab.id === "number" ? tab.id : -1;
    if (tabId < 0) return;
    ingestDoubaoBody(
      message.body,
      tabId,
      (tab && tab.cookieStoreId) || "",
      Number(message.startedAt || Date.now()),
      Number(message.completedAt || Date.now()),
    ).catch(() => {});
  });
}

// (Removed a browser.action.onClicked listener here: with default_popup set in
// both manifests the toolbar click always opens the popup and onClicked never
// fires. The same manual-sync path is exposed to the popup via the
// POPUP_SYNC_DOUBAO message, so this was unreachable dead code.)

connectNative();

/* ==== Voice Edge popup panel bridge (state round-trip to Voice Edge) ==== */
let lastDeepSeekSnapshot = null;

const POPUP_OPEN_URLS = {
  deepseek: "https://chat.deepseek.com/",
  qwen: "https://chat.qwen.ai/",
  doubao: "https://www.doubao.com/chat/",
  m365: "https://outlook.cloud.microsoft/",
  aistudio: "https://aistudio.google.com/prompts/new_chat?model=gemini-4-argon",
};

function popupSafeHost(url) {
  try {
    return new URL(String(url || "")).hostname || "";
  } catch (_) {
    return "";
  }
}

// Ask Voice Edge (via the native host) for a secret-free persistent-state
// summary. The reply arrives asynchronously as a STATE_REPORT native message
// and is cached in lastStateReport; the panel auto-refreshes to pick it up.
function requestBrowserState(force) {
  const now = Date.now();
  if (!force && now - lastStateRequestAt < 1200) return;
  lastStateRequestAt = now;
  const port = connectNative();
  if (!port) return;
  try {
    port.postMessage({ type: "STATE_QUERY" });
  } catch (_) {}
}

function popupBuildState() {
  const auth = (deepseekAuthorization || "")
    .toLowerCase()
    .startsWith("bearer ");
  return {
    ok: true,
    generatedAt: Date.now(),
    native: { connected: Boolean(nativePort), host: NATIVE_HOST },
    deepseek: {
      pending: Boolean(deepseekState),
      authorizationCaptured: auth,
      hifCaptured: Boolean(deepseekHifLeim),
      lastSnapshot: lastDeepSeekSnapshot,
    },
    qwen: {
      pending: Boolean(qwenState),
      syncInFlight: Boolean(qwenState && qwenState.syncInFlight),
      accountId: qwenState ? String(qwenState.accountId || "") : "",
      verificationHost: qwenState
        ? popupSafeHost(qwenState.verificationUrl)
        : "",
      startedAt: qwenState ? Number(qwenState.startedAt || 0) : 0,
    },
    doubao: {
      authRequired: Boolean(authRequired),
      authRequiredSince: Number(authRequiredSince || 0),
      syncInFlight: Boolean(syncInFlight),
      lastCaptureApplied: Boolean(
        lastCaptured &&
        captureKey(lastCaptured) &&
        captureKey(lastCaptured) === lastAppliedKey,
      ),
      lastCapture:
        lastCaptured && lastCaptured.conversation
          ? {
              conversation_id: String(
                lastCaptured.conversation.conversation_id || "",
              ),
              section_id: String(lastCaptured.conversation.section_id || ""),
              last_message_index: Number(
                lastCaptured.conversation.last_message_index || 0,
              ),
              complete: Boolean(lastCaptured.complete),
              error: lastCaptured.error
                ? String(
                    lastCaptured.error.message ||
                      lastCaptured.error.code ||
                      "error",
                  )
                : "",
              completedAt: Number(lastCaptured.completedAt || 0),
            }
          : null,
    },
    sharepoint: {
      configured: Boolean(sharePointHomeUrl),
      homeUrl: sharePointHomeUrl,
      uploadFolder: sharePointUploadFolder,
      readyPageUrl: sharePointReadyPageUrl,
      lastResult: sharePointLastResult,
      tokenExpiresAt: Number(_sharePoint.exp || 0),
      lastKeepAliveAt: Number(sharePointLastKeepAliveAt || 0),
      lastKeepAliveError: sharePointLastKeepAliveError,
      lastKeepAliveErrorAt: Number(sharePointLastKeepAliveErrorAt || 0),
    },
    m365: {
      bridgeConnected: Boolean(m365Ws && m365Ws.readyState === WebSocket.OPEN),
      authLoaded: Boolean(m365AuthLoaded),
      authAvailable: Boolean(_rt && _clientId && _tid),
      tenantId: _tid ? `${_tid.slice(0, 8)}…${_tid.slice(-4)}` : "",
      clientId: _clientId
        ? `${_clientId.slice(0, 8)}…${_clientId.slice(-4)}`
        : "",
      authUpdatedAt: Number(m365AuthUpdatedAt || 0),
      tokenExpiresAt: Number(_sydney.exp || 0),
      refreshTokenExpiresAt: Number(_rtExp || 0),
      lastKeepAliveAt: Number(m365LastKeepAliveAt || 0),
      lastKeepAliveError: m365LastKeepAliveError,
      lastKeepAliveErrorAt: Number(m365LastKeepAliveErrorAt || 0),
      lastRefreshError: m365LastRefreshError,
      lastRefreshErrorAt: Number(m365LastRefreshErrorAt || 0),
      reauthRequired: Boolean(m365ReauthRequired),
      reauthReason: m365ReauthReason,
      reauthRequiredAt: Number(m365ReauthRequiredAt || 0),
      autoLoginEnabled: Boolean(m365AutoLoginEnabled),
      autoLoginRemember: Boolean(m365AutoLoginRemember),
      autoLoginConfigured: Boolean(m365AutoLoginPassword),
      autoLoginUsername: m365AutoLoginUsername,
      autoLoginPaused: Boolean(m365AutoLoginPaused),
      autoLoginLastError: m365AutoLoginLastError,
      autoLoginLastErrorAt: Number(m365AutoLoginLastErrorAt || 0),
      autoLoginLastSuccessAt: Number(m365AutoLoginLastSuccessAt || 0),
      reauthPendingTabId:
        _pendingTopLevelReauth && _pendingTopLevelReauth.tabId != null
          ? _pendingTopLevelReauth.tabId
          : 0,
      reauthState: m365ReauthState,
      reauthStateReason: m365ReauthStateReason,
      reauthStateAt: Number(m365ReauthStateAt || 0),
    },
    aistudio: {
      bridgeConnected: Boolean(m365Ws && m365Ws.readyState === WebSocket.OPEN),
      entryUrl: aisEntryUrl,
      tabId: aisTargetTabId || (aisLastReady && aisLastReady.tabId) || 0,
      tabUrl: (aisLastReady && aisLastReady.tabUrl) || "",
      hookReady: Boolean(aisLastReady),
      service: (aisLastReady && aisLastReady.service) || "pending",
      timezone: (aisLastReady && aisLastReady.timezone) || "",
      authOk: Boolean(
        aisLastReady && aisLastReady.auth && aisLastReady.auth.ok,
      ),
      authUser:
        (aisLastReady && aisLastReady.auth && aisLastReady.auth.authUser) ||
        "0",
      tier: (aisLastReady && aisLastReady.auth && aisLastReady.auth.tier) || "",
      hasApiKey: Boolean(
        aisLastReady && aisLastReady.auth && aisLastReady.auth.hasApiKey,
      ),
      hasExtBin: Boolean(
        aisLastReady && aisLastReady.auth && aisLastReady.auth.hasExtBin,
      ),
      driveStatus:
        (aisLastReady && aisLastReady.drive && aisLastReady.drive.status) ||
        "unknown",
      driveError:
        (aisLastReady && aisLastReady.drive && aisLastReady.drive.error) || "",
      driveCheckedAt: Number(
        (aisLastReady && aisLastReady.drive && aisLastReady.drive.checkedAt) ||
          0,
      ),
      updatedAt: Number((aisLastReady && aisLastReady.updatedAt) || 0),
      lastError: aisLastError,
    },
    // Secret-free persistent state reported by Voice Edge (may be null until
    // the async STATE_REPORT arrives). Panel triggers a refresh each poll.
    persistent: lastStateReport,
  };
}

async function popupSyncDoubao() {
  if (typeof synchronizeCapture !== "function")
    return { ok: false, error: "background 未就绪" };
  if (!lastCaptured || !lastCaptured.complete || lastCaptured.error) {
    return { ok: false, error: "请先在豆包网页发送一条消息并等待完整回复" };
  }
  try {
    await synchronizeCapture(lastCaptured, false);
    return { ok: true };
  } catch (error) {
    try {
      syncInFlight = false;
    } catch (_) {}
    return { ok: false, error: compactError(error) };
  }
}

async function popupOpen(provider) {
  const key = String(provider || "");
  if (key === "aistudio") {
    try {
      const tabs = await browser.tabs.query({ url: AIS_TAB_URL_PATTERNS });
      if (tabs && tabs.length && tabs[0].id != null) {
        aisTargetTabId = tabs[0].id;
        await browser.tabs.update(tabs[0].id, { active: true });
        return { ok: true };
      }
    } catch (_) {}
  }
  const url =
    key === "m365"
      ? m365EntryUrl || POPUP_OPEN_URLS.m365
      : key === "aistudio"
        ? aisEntryUrl || POPUP_OPEN_URLS.aistudio
        : POPUP_OPEN_URLS[key];
  if (!url) return { ok: false, error: "未知的 provider" };
  try {
    await browser.tabs.create({ url, active: true });
    return { ok: true };
  } catch (error) {
    return { ok: false, error: String((error && error.message) || error) };
  }
}

browser.runtime.onMessage.addListener((message, sender) => {
  if (!message || typeof message !== "object") return;
  if (message.type === "M365_GET_AUTO_LOGIN_CONFIG") {
    return Promise.all([
      ensureM365AuthLoaded(),
      ensureM365AutoLoginLoaded(),
    ]).then(() => {
      const armed = isSenderArmedM365LoginTab(sender, message);
      if (armed && _pendingTopLevelReauth) {
        _pendingTopLevelReauth.pageSeen = true;
        if (_pendingTopLevelReauth.loadWatchdog) {
          clearTimeout(_pendingTopLevelReauth.loadWatchdog);
          _pendingTopLevelReauth.loadWatchdog = null;
        }
      }
      log(
        "M365 auto-login config request tab=%s armed=%s pendingTab=%s armedTab=%s armedUntil=%s enabled=%s paused=%s hasPassword=%s",
        sender && sender.tab ? sender.tab.id : "?",
        armed,
        _pendingTopLevelReauth ? _pendingTopLevelReauth.tabId : "none",
        _armedPopupTabId,
        Math.max(0, Math.round((_armedPopupWindowUntil - Date.now()) / 1000)) +
          "s",
        Boolean(m365AutoLoginEnabled),
        Boolean(m365AutoLoginPaused),
        Boolean(m365AutoLoginPassword),
      );
      if (!armed) {
        return {
          enabled: false,
          rememberOnManualLogin: false,
          paused: Boolean(m365AutoLoginPaused),
          armed: false,
        };
      }
      const allowFill = Boolean(m365AutoLoginEnabled && !m365AutoLoginPaused);
      return {
        enabled: Boolean(m365AutoLoginEnabled),
        rememberOnManualLogin: Boolean(
          m365AutoLoginEnabled && m365AutoLoginRemember,
        ),
        paused: Boolean(m365AutoLoginPaused),
        armed: true,
        tenantId: _tid,
        username: allowFill ? m365AutoLoginUsername : "",
        password: allowFill ? m365AutoLoginPassword : "",
      };
    });
  }
  if (message.type === "POPUP_SAVE_M365_AUTO_LOGIN") {
    return (async () => {
      await ensureM365AutoLoginLoaded();
      if (typeof message.enabled === "boolean") {
        m365AutoLoginEnabled = message.enabled;
        if (!m365AutoLoginEnabled) {
          _stagedAutoLoginCreds = null;
        }
      }
      if (typeof message.rememberOnManualLogin === "boolean") {
        m365AutoLoginRemember = message.rememberOnManualLogin;
        if (!m365AutoLoginRemember) {
          _stagedAutoLoginCreds = null;
        }
      }
      if (typeof message.username === "string") {
        m365AutoLoginUsername = message.username.trim();
      }
      if (typeof message.password === "string" && message.password !== "") {
        m365AutoLoginPassword = message.password;
      }
      if (message.unpause !== false) {
        m365AutoLoginPaused = false;
        m365AutoLoginLastError = "";
        m365AutoLoginLastErrorAt = 0;
      }
      await saveM365AutoLogin();
      // Only an explicit password save (unpause !== false, which also clears
      // m365AutoLoginPaused above) while auto-login is enabled resets the
      // terminal state and starts a new attempt.
      if (message.unpause !== false && m365AutoLoginEnabled) {
        await resetM365ReauthState("auto-login settings saved");
        if (m365ReauthRequired || m365ReauthState !== "none") {
          tryTopLevelSilentReauth("popup-saved-password", {
            force: true,
          }).catch((error) =>
            log(
              "M365 reauth from popup-saved-password failed:",
              compactError(error),
            ),
          );
        }
      }
      return popupBuildState();
    })().catch((error) => ({ ok: false, error: compactError(error) }));
  }
  if (message.type === "POPUP_CLEAR_M365_AUTO_LOGIN") {
    return (async () => {
      await ensureM365AutoLoginLoaded();
      m365AutoLoginPassword = "";
      m365AutoLoginPaused = false;
      m365AutoLoginLastError = "";
      m365AutoLoginLastErrorAt = 0;
      await saveM365AutoLogin();
      return popupBuildState();
    })().catch((error) => ({ ok: false, error: compactError(error) }));
  }
  if (message.type === "DEEPSEEK_CONVERSATION_SNAPSHOT") {
    // Remember it for the panel; the original relay listener still forwards it.
    lastDeepSeekSnapshot = {
      conversation_id: String(message.conversation_id || ""),
      model_type: String(message.model_type || ""),
      page_url: String(message.page_url || ""),
      captured_at: Number(message.captured_at || Date.now()),
    };
    return;
  }
  if (message.type === "POPUP_GET_STATE") {
    requestBrowserState(false); // refresh persistent state (async)
    if (!aisLastReady || Date.now() - aisLastPopupPingAt > 30000) {
      aisLastPopupPingAt = Date.now();
      pingAisTabs().catch(() => {});
    }
    return Promise.resolve(popupBuildState());
  }
  if (message.type === "POPUP_AIS_PROBE") {
    return runAisPopupProbe(Boolean(message.warmup));
  }
  if (message.type === "POPUP_SYNC_DOUBAO") return popupSyncDoubao();
  if (message.type === "POPUP_CLEAR_M365_AUTH") {
    return clearM365Auth()
      .then(() => ({ ok: true }))
      .catch((error) => ({ ok: false, error: compactError(error) }));
  }
  if (message.type === "POPUP_RECONNECT") {
    try {
      connectNative();
    } catch (_) {}
    requestBrowserState(true);
    return Promise.resolve({ ok: true, connected: Boolean(nativePort) });
  }
  if (message.type === "POPUP_MUTATE") {
    const port = connectNative();
    if (!port)
      return Promise.resolve({ ok: false, error: "无法连接 Native Host" });
    try {
      port.postMessage({
        type: "STATE_MUTATE",
        action: String(message.action || "clear"),
        provider: String(message.provider || ""),
        model_type: String(message.model_type || ""),
      });
      // The STATE_REPORT reply refreshes lastStateReport via handleNativeMessage;
      // nudge another read so the panel reflects the change immediately.
      requestBrowserState(true);
      return Promise.resolve({ ok: true });
    } catch (e) {
      return Promise.resolve({
        ok: false,
        error: String((e && e.message) || e),
      });
    }
  }
  if (message.type === "POPUP_SP_TEST") return runSharePointCommand("SP_TEST");
  if (message.type === "POPUP_SP_UPLOAD_TEST")
    return runSharePointCommand("SP_UPLOAD_TEST");
  if (message.type === "POPUP_OPEN") return popupOpen(message.provider);
});

async function ensureSharePointTab() {
  if (!sharePointHomeUrl) throw new Error("SHAREPOINT_HOME_URL 未配置");
  const home = new URL(sharePointHomeUrl);
  const sitePrefix = home.origin + home.pathname.replace(/\/$/, "");
  const allItemsUrl = sitePrefix + "/Shared%20Documents/Forms/AllItems.aspx";
  const tabs = await browser.tabs.query({ url: [sitePrefix + "/*"] });
  const usable = tabs.find((tab) => tab.id != null && !tab.discarded);
  if (usable) return usable;
  return browser.tabs.create({ url: allItemsUrl, active: false });
}

async function sendSharePointCommand(tab, payload) {
  const deadline = Date.now() + 20000;
  let lastError = null;
  while (Date.now() < deadline) {
    try {
      const response = await browser.tabs.sendMessage(tab.id, payload);
      if (response) return response;
    } catch (error) {
      lastError = error;
    }
    await new Promise((resolve) => setTimeout(resolve, 300));
  }
  throw new Error(
    "SharePoint 页面监听器未就绪。请确认已登录站点并刷新页面" +
      (lastError && lastError.message ? "：" + lastError.message : ""),
  );
}

async function runSharePointCommand(type) {
  try {
    const tab = await ensureSharePointTab();
    const payload = {
      __veSharePoint: true,
      type,
      siteUrl: sharePointHomeUrl,
      uploadFolder: sharePointUploadFolder,
    };
    const response = await sendSharePointCommand(tab, payload);
    sharePointLastResult = {
      ...response,
      type,
      at: Date.now(),
      error: response.ok
        ? ""
        : String(response.error || "SharePoint test failed"),
    };
    return sharePointLastResult;
  } catch (error) {
    sharePointLastResult = {
      ok: false,
      type,
      at: Date.now(),
      error: compactError(error),
    };
    return sharePointLastResult;
  }
}

browser.runtime.onMessage.addListener((message) => {
  if (!message || message.__veSharePoint !== true) return;
  if (message.type === "SP_FRAME_READY") {
    sharePointReadyPageUrl = String(message.pageUrl || "");
    return;
  }
  if (message.type === "SP_SESSION_RENEWED") {
    sharePointLastKeepAliveAt = Number(message.at || Date.now());
    sharePointLastKeepAliveError = "";
    sharePointLastKeepAliveErrorAt = 0;
    if (message.pageUrl) sharePointReadyPageUrl = String(message.pageUrl);
    dlog(
      "SharePoint session renewed: site=%s digestExp=%d spoTokenExp=%d",
      String(message.siteUrl || ""),
      Number(message.digestExp || 0),
      Number(message.spoTokenExp || 0),
    );
    return;
  }
  if (message.type === "SP_SESSION_RENEW_FAILED") {
    sharePointLastKeepAliveError = String(message.error || "unknown");
    sharePointLastKeepAliveErrorAt = Number(message.at || Date.now());
    dlog(
      "SharePoint session renew failed: site=%s err=%s",
      String(message.siteUrl || ""),
      sharePointLastKeepAliveError,
    );
    schedulePreExpiryRetry("sp-session-failed");
    return;
  }
  if (message.type === "SP_UPLOAD_COMPLETE") {
    const requestId = String(message.requestId || "");
    const waiter = sharePointUploadWaiters.get(requestId);
    if (waiter) {
      sharePointUploadWaiters.delete(requestId);
      waiter.resolve(
        message.response || {
          ok: false,
          error: "empty SharePoint upload response",
        },
      );
    }
  }
});
// Warm the persistent-state cache shortly after startup.
setTimeout(() => requestBrowserState(true), 800);

/* =========================================================================
 * Voice Edge · M365 background — OAuth broker, entry settings, frame routing
 * ========================================================================= */
const M365_BRIDGE_URL = "ws://127.0.0.1:5003/ws";
const M365_AUTH_KEY = "voiceEdgeM365AuthV1";
const M365_SETTINGS_KEY = "voiceEdgeM365SettingsV1";
const M365_AUTO_LOGIN_KEY = "voiceEdgeM365AutoLoginV1";
// Identity marker only: it tells us which captured /token request is the 24h
// M365 SPA refresh token. Every value we actually SEND to AAD — client_id,
// redirect_uri, brk_client_id, brk_redirect_uri, scope — comes from the page's
// own captured request. There are no defaults: if the page never sent it, the
// attempt fails with the missing field named instead of guessing a value
// (hard requirement).
const M365_SYDNEY_SCOPE = "https://substrate.office.com/sydney/.default";
// Bumped on every behavioural change that a driver must be able to recognise.
// The harness prints it on M365_READY, so "did the reload actually take?" is a
// fact instead of a guess.
const M365_BUILD_TAG = "2026-11-04T19:20";
// Audience required by the AMS object endpoint (asyncgw.teams.microsoft.com/
// v1/objects/…) that serves CodeInterpreter / present_files artifacts. The
// Sydney token (aud=substrate.office.com/sydney) is REJECTED there with 401 —
// that was the harvestArtifacts failure. This IC3/Teams scope mints a token
// whose aud is ic3.teams.office.com, which AMS accepts.
const M365_IC3_SCOPE = "https://ic3.teams.office.com/.default";
const M365_TENANT_RE = /login\.microsoftonline\.com\/([0-9a-f-]{36})\/oauth2/i;
const M365_ENTRY_RE =
  /^https:\/\/outlook\.cloud\.microsoft\/host\/[0-9a-f-]{36}\/entity1-[0-9a-f-]{36}\/?$/i;
const M365_HOST_ALLOWLIST = [
  "outlook.cloud.microsoft",
  "outlook.office.com",
  "m365.cloud.microsoft",
  "m365copilotapp.svc.cloud.microsoft",
];
const M365_TAB_URL_PATTERNS = M365_HOST_ALLOWLIST.map(
  (host) => "https://" + host + "/*",
);
function isAllowedM365Sender(sender) {
  const rawUrl = String(
    (sender && sender.url) || (sender && sender.tab && sender.tab.url) || "",
  );
  if (!rawUrl) return false;
  try {
    const parsed = new URL(rawUrl);
    return (
      parsed.protocol === "https:" &&
      M365_HOST_ALLOWLIST.includes(parsed.hostname.toLowerCase())
    );
  } catch (_) {
    return false;
  }
}

let _rt = "";
const M365_APPLIED_TOKEN_CACHE_LIMIT = 16;
const _recentAppliedM365AccessTokens = new Set();
// Start time (epoch ms) of the /token request that produced the refresh token we
// currently hold. Used to reject a *previous* RT generation that arrives late:
// the page fires one /token per resource and they can complete out of order.
let _rtNewestObservedAt = 0;
let _rtExp = 0;
let _clientId = "";
let _tid = "";
let _m365TokenContext = { version: 1, query: {}, body: {} };
let _sydney = { token: "", exp: 0 };
let _ic3 = { token: "", exp: 0 };
let _ic3Refreshing = null;
let _ic3Probed = false; // one-time aud/scp probe log (see refreshIc3Token)
let _sharePoint = { token: "", exp: 0, origin: "" };
let _sharePointRefreshing = null;
let m365LastRefreshError = "";
let m365LastRefreshErrorAt = 0;
let m365ReauthRequired = false;
let m365ReauthReason = "";
let m365ReauthRequiredAt = 0;
let m365LastKeepAliveAt = 0;
let m365LastKeepAliveError = "";
let m365LastKeepAliveErrorAt = 0;
let sharePointLastKeepAliveAt = 0;
let sharePointLastKeepAliveError = "";
let sharePointLastKeepAliveErrorAt = 0;
// Base keepalive cadence (15m): half of SharePoint's 30m FormDigest lifetime
// (1800s) and 1/36 of M365's 540m ABT timeout.
const PAGE_KEEPALIVE_BASE_MS = 15 * 60 * 1000;
const M365_PRE_EXPIRY_WINDOW_MS = PAGE_KEEPALIVE_BASE_MS;
// M365 background keepalive runs at 0.8x of the page's 15m ABT interval (12m)
// so background-driven pings preempt the page's own timer even when background
// tabs are throttled.
const M365_KEEPALIVE_INTERVAL_MS = Math.round(PAGE_KEEPALIVE_BASE_MS * 0.8);
const SP_KEEPALIVE_INTERVAL_MS = PAGE_KEEPALIVE_BASE_MS;
// Proactive top-level /authorize renewal window (6h before the 24h SPA RT cap).
const M365_RT_RENEW_WINDOW_MS = 6 * 60 * 60 * 1000;
// ---------------------------------------------------------------------------
// Single-shot reauth state machine (product decision).
// ONE automatic login attempt per NEEDS_REAUTH episode. If it fails, that is
// final: no retry, no cooldown, no budget, no backoff. The state is persisted so
// a background restart cannot quietly start another attempt. The only ways out
// of `terminal` are a real sign-in (authorization_code), a plain refresh that
// proves the token is alive, or an explicit user/harness reset.
const M365_REAUTH_STATE_KEY = "voiceEdgeM365ReauthStateV2";
const M365_PENDING_REAUTH_KEY = "voiceEdgeM365PendingReauthV1";
const M365_REAUTH_TAB_TIMEOUT_MS = 45 * 1000;
// Stage-1 watchdog: "did the page become usable at all?" (login UI seen)
const M365_REAUTH_LOAD_TIMEOUT_MS = 12 * 1000;
// How long a failed episode keeps auto-login closed before a new attempt is
// allowed again. A transient failure (offline / page not loaded) must not
// disable auto-login for the rest of the day; a hard failure still waits.
const M365_REAUTH_HOLD_MS = 6 * 60 * 60 * 1000;
const M365_REAUTH_ENV_RETRY_MS = 30 * 60 * 1000;
const M365_PENDING_REAUTH_TTL_MS = 10 * 60 * 1000;
let _pendingTopLevelReauth = null;
// Synchronous single-flight flag: set before the first await so two concurrent
// callers (pre-expiry cycle + a rejected refresh) cannot both create a tab.
let _reauthCreating = false;
// none | attempting | terminal — no "wait for the user" state: auto-login
// is unattended, so a missing password is a failure, not a hand-off.
let m365ReauthState = "none";
// When a real authorization_code sign-in completed. Two independent triggers
// (handleRefreshTokenExhausted + the pre-expiry loop) run in the same cycle; if
// the first one completes fast (SSO still valid), the second used to open a
// second, pointless login window. A fresh sign-in suppresses duplicates.
let _lastAuthCodeSignInAt = 0;
let m365ReauthStateReason = "";
let m365ReauthStateAt = 0;
let _reauthStateLoaded = false;
let _reauthStatePromise = null;
let _reauthStateLastLogAt = 0;
// Earliest time a new attempt may start after an environmental failure.
let _reauthNextAttemptAt = 0;
let _lastOrphanSweepAt = 0;
let _lastRtRejectLogAt = 0;
let _lastRtSwapLogAt = 0;
let _lastRtVerifyAt = 0;
let _rtVerifyInFlight = null;
let m365AutoLoginLoaded = false;
let m365AutoLoginLoadPromise = null;
let m365AutoLoginEnabled = true;
let m365AutoLoginRemember = false;
let m365AutoLoginUsername = "";
let m365AutoLoginPassword = "";
let m365AutoLoginPaused = false;
let m365AutoLoginLastError = "";
let m365AutoLoginLastErrorAt = 0;
let m365AutoLoginLastSuccessAt = 0;
let _stagedAutoLoginCreds = null;
let _armedPopupTabId = null;
let _armedPopupWindowUntil = 0;
// Per-tab arming registry (tabId -> expiresAt). A single global id was not
// enough: once two login tabs coexisted, the second one overwrote the first, so
// after AAD redirected tab A to /login it was no longer considered armed and
// auto-login silently stopped filling it.
const M365_ARMED_TAB_TTL_MS = 5 * 60 * 1000;
const _armedLoginTabs = new Map();
function armM365LoginTab(
  tabId,
  now = Date.now(),
  ttlMs = M365_ARMED_TAB_TTL_MS,
) {
  if (tabId == null) return;
  _armedLoginTabs.set(tabId, now + ttlMs);
  _armedPopupTabId = tabId;
  _armedPopupWindowUntil = now + ttlMs;
}
function isM365LoginTabArmed(tabId, now = Date.now()) {
  const exp = _armedLoginTabs.get(tabId);
  if (!exp) return false;
  if (exp <= now) {
    _armedLoginTabs.delete(tabId);
    return false;
  }
  _armedLoginTabs.set(tabId, now + M365_ARMED_TAB_TTL_MS);
  return true;
}
function disarmM365LoginTab(tabId) {
  if (tabId == null) return;
  _armedLoginTabs.delete(tabId);
}
let _refreshing = null;
let m365AuthLoaded = false;
let m365AuthLoadPromise = null;
let m365AuthUpdatedAt = 0;
let m365LatestTokenResponseAt = 0;
let m365SettingsLoaded = false;
let m365SettingsLoadPromise = null;
let m365EntryUrl = "";
let sharePointHomeUrl = "";
// Two distinct SharePoint folders, per the user's split:
//  - sharePointUploadFolder   : "传给模型的目录" — where the user's files that
//                                are sent TO the model live (user -> model).
//                                Used by prepareM365Attachments (unchanged).
//  - sharePointDownloadFolder : "模型传的目录" — where files the MODEL produced
//                                (CodeInterpreter / present_files artifacts) are
//                                uploaded (model -> SharePoint) by the artifact
//                                harvester. Falls back to the upload folder if
//                                voice_edge has not yet been updated to send it.
let sharePointUploadFolder = "";
let sharePointDownloadFolder = "";
let sharePointLastResult = null;
let sharePointReadyPageUrl = "";
const sharePointUploadWaiters = new Map();
let m365Ws = null;
let m365ReconnectTimer = null;
let m365ReconnectAttempt = 0;
const M365_RECONNECT_BASE_MS = 1000;
const M365_RECONNECT_MAX_MS = 30000;
let m365TargetFrame = null;
const m365FrameCandidates = new Map();
// Recent log lines kept in memory so the harness page can show exactly what the
// background decided (same text as the console, without needing devtools).
const M365_LOG_RING_MAX = 400;
const _m365LogRing = [];
function formatLogArgs(args) {
  if (!args.length) return "";
  const [head, ...rest] = args;
  let used = 0;
  const render = (v) => {
    if (typeof v === "string") return v;
    try {
      return JSON.stringify(v);
    } catch (_) {
      return String(v);
    }
  };
  const text = String(head).replace(/%[sdj%]/g, (token) => {
    if (token === "%%") return "%";
    if (used >= rest.length) return token;
    const value = rest[used++];
    if (token === "%d") return String(Number(value));
    return render(value);
  });
  return [text, ...rest.slice(used).map(render)].join(" ");
}
function pushLogRing(line) {
  try {
    _m365LogRing.push(line);
    if (_m365LogRing.length > M365_LOG_RING_MAX) {
      _m365LogRing.splice(0, _m365LogRing.length - M365_LOG_RING_MAX);
    }
  } catch (_) {}
}
const log = (...args) => {
  try {
    pushLogRing(
      new Date().toISOString().slice(11, 23) + " " + formatLogArgs(args),
    );
  } catch (_) {}
  console.log("[VE-m365-bg]", ...args);
};
// Gated verbose tracing. OFF by default; driven from Python via M365_CONFIG
// `debug` (so backend M365_DEBUG=1 turns it on end-to-end) and also togglable
// live from the extension console: __veM365SetDebug(true).
let m365Debug = false;
const dlog = (...args) => {
  if (m365Debug) console.log("[VE-m365-bg][dbg]", ...args);
};
try {
  globalThis.__veM365SetDebug = (on) => {
    m365Debug = !!on;
    log("debug tracing", m365Debug ? "ON" : "OFF");
    return m365Debug;
  };
} catch (_) {}

function normalizeM365EntryUrl(value) {
  const text = String(value || "").trim();
  if (!M365_ENTRY_RE.test(text)) return "";
  return text.replace(/\/$/, "");
}

async function loadM365Auth() {
  if (m365AuthLoaded) return;
  try {
    const stored = await browser.storage.local.get(M365_AUTH_KEY);
    const saved = stored[M365_AUTH_KEY] || {};
    _rt = String(saved.rt || saved.refreshToken || "");
    _rtExp = Number(saved.rtExp || 0);
    _clientId = String(saved.clientId || "");
    _tid = String(saved.tid || saved.tenantId || "");
    _m365TokenContext =
      saved.tokenContext && typeof saved.tokenContext === "object"
        ? {
            version: 1,
            query: { ...(saved.tokenContext.query || {}) },
            body: { ...(saved.tokenContext.body || {}) },
          }
        : { version: 1, query: {}, body: {} };
    // Migrate earlier builds that duplicated transient secrets in tokenContext.
    delete _m365TokenContext.body.refresh_token;
    delete _m365TokenContext.body.client_request_id;
    delete _m365TokenContext.body["client-request-id"];
    m365AuthUpdatedAt = Number(saved.updatedAt || 0);
    m365AuthLoaded = true;
  } catch (error) {
    m365AuthLoadPromise = null;
    throw error;
  }
}

function ensureM365AuthLoaded() {
  if (m365AuthLoaded) return Promise.resolve();
  if (!m365AuthLoadPromise) m365AuthLoadPromise = loadM365Auth();
  return m365AuthLoadPromise;
}

async function saveM365Auth() {
  if (!_rt || !_clientId || !_tid) return;
  m365AuthUpdatedAt = Date.now();
  await browser.storage.local.set({
    [M365_AUTH_KEY]: {
      rt: _rt,
      rtExp: _rtExp,
      clientId: _clientId,
      tid: _tid,
      tokenContext: _m365TokenContext,
      updatedAt: m365AuthUpdatedAt,
    },
  });
}

async function clearM365Auth() {
  m365AuthLoadPromise = null;
  _rt = "";
  _rtExp = 0;
  _clientId = "";
  _tid = "";
  _m365TokenContext = { version: 1, query: {}, body: {} };
  _sydney = { token: "", exp: 0 };
  _ic3 = { token: "", exp: 0 };
  _sharePoint = { token: "", exp: 0, origin: "" };
  _sharePointRefreshing = null;
  m365LastRefreshError = "";
  m365LastRefreshErrorAt = 0;
  m365ReauthRequired = false;
  m365ReauthReason = "";
  m365ReauthRequiredAt = 0;
  m365LastKeepAliveAt = 0;
  m365LastKeepAliveError = "";
  m365LastKeepAliveErrorAt = 0;
  m365AuthUpdatedAt = 0;
  _recentAppliedM365AccessTokens.clear();
  _rtNewestObservedAt = 0;
  m365AuthLoaded = true;
  await browser.storage.local.remove(M365_AUTH_KEY);
}

async function loadM365AutoLogin() {
  if (m365AutoLoginLoaded) return;
  try {
    const stored = await browser.storage.local.get(M365_AUTO_LOGIN_KEY);
    const saved = stored[M365_AUTO_LOGIN_KEY] || {};
    m365AutoLoginEnabled = saved.enabled !== false;
    m365AutoLoginRemember = Boolean(saved.rememberOnManualLogin);
    m365AutoLoginUsername = String(saved.username || "").trim();
    m365AutoLoginPassword = String(saved.password || "");
    m365AutoLoginPaused = Boolean(saved.paused);
    m365AutoLoginLastError = String(saved.lastError || "");
    m365AutoLoginLastErrorAt = Number(saved.lastErrorAt || 0);
    m365AutoLoginLastSuccessAt = Number(saved.lastSuccessAt || 0);
    m365AutoLoginLoaded = true;
  } catch (error) {
    m365AutoLoginLoadPromise = null;
    throw error;
  }
}

function ensureM365AutoLoginLoaded() {
  if (m365AutoLoginLoaded) return Promise.resolve();
  if (!m365AutoLoginLoadPromise) m365AutoLoginLoadPromise = loadM365AutoLogin();
  return m365AutoLoginLoadPromise;
}

async function saveM365AutoLogin() {
  await browser.storage.local.set({
    [M365_AUTO_LOGIN_KEY]: {
      enabled: Boolean(m365AutoLoginEnabled),
      rememberOnManualLogin: Boolean(m365AutoLoginRemember),
      username: m365AutoLoginUsername,
      password: m365AutoLoginPassword,
      paused: Boolean(m365AutoLoginPaused),
      lastError: m365AutoLoginLastError,
      lastErrorAt: m365AutoLoginLastErrorAt,
      lastSuccessAt: m365AutoLoginLastSuccessAt,
    },
  });
}

// Strict sender + target validation before handing auto-login credentials to
// content-m365-login.js or accepting staged manual credentials.
function isSenderArmedM365LoginTab(sender, message) {
  if (!sender || !sender.tab || typeof sender.tab.id !== "number") return false;
  if (sender.frameId != null && sender.frameId !== 0) return false;
  const rawUrl = String(sender.url || sender.tab.url || "");
  let parsed;
  try {
    parsed = new URL(rawUrl);
  } catch (_) {
    return false;
  }
  if (parsed.origin !== "https://login.microsoftonline.com") return false;
  const now = Date.now();
  const tabId = sender.tab.id;

  // 1. A tab of our current pending attempt. The window is generous (5 min)
  //    because AAD redirects the tab through /login, /SAS/ProcessAuth and the
  //    account picker before it ever shows #i0118; a 60s window silently
  //    disarmed auto-login half way through the flow.
  if (
    _pendingTopLevelReauth &&
    _pendingTopLevelReauth.tabId === tabId &&
    now - Number(_pendingTopLevelReauth.startedAt || 0) < M365_ARMED_TAB_TTL_MS
  ) {
    armM365LoginTab(tabId, now);
    return true;
  }

  // 2. Already armed for this tab. Arming is tracked PER TAB: a single global
  //    id meant that when two attempts coexisted, the older tab lost its armed
  //    state on every redirect and auto-login silently stopped filling it.
  if (isM365LoginTabArmed(tabId, now)) return true;

  // 3. Initial /oauth2/v2.0/authorize navigation matching our bound tenant +
  //    M365 FOCI client_id (or state=ve-reauth:...).
  const stateParam = String(parsed.searchParams.get("state") || "");
  if (stateParam.startsWith("ve-reauth:")) {
    armM365LoginTab(tabId, now);
    return true;
  }
  const clientIdParam = String(parsed.searchParams.get("client_id") || "");
  const brkClientIdParam = String(
    parsed.searchParams.get("brk_client_id") || "",
  );
  // Only identities the page itself handed us count. The old hardcoded GUID
  // list armed auto-login on any Microsoft page that happened to use those
  // FOCI clients — a guess, and exactly what this flow must not do.
  const capturedPageBrkClientId = String(
    (_m365TokenContext && _m365TokenContext.body
      ? _m365TokenContext.body.brk_client_id
      : "") || "",
  ).trim();
  const isM365Client =
    (Boolean(_clientId) && clientIdParam === _clientId) ||
    (Boolean(capturedPageBrkClientId) &&
      (clientIdParam === capturedPageBrkClientId ||
        brkClientIdParam === capturedPageBrkClientId));
  const pathTidMatch = parsed.pathname.match(
    /^\/([0-9a-f-]{36}|organizations|common)\/oauth2\/v2\.0\/authorize\b/i,
  );
  const pathTid = pathTidMatch ? pathTidMatch[1].toLowerCase() : "";
  const tenantMatches = Boolean(
    pathTid &&
    (_tid
      ? pathTid === _tid.toLowerCase()
      : pathTid === "organizations" ||
        pathTid === "common" ||
        /^[0-9a-f-]{36}$/.test(pathTid)),
  );
  if (pathTidMatch && isM365Client && tenantMatches) {
    const cfgUser = String((message && message.cfgUsername) || "")
      .trim()
      .toLowerCase();
    if (
      cfgUser &&
      m365AutoLoginUsername &&
      cfgUser !== m365AutoLoginUsername.toLowerCase()
    ) {
      return false;
    }
    armM365LoginTab(tabId, now);
    return true;
  }
  return false;
}

async function loadM365Settings() {
  if (m365SettingsLoaded) return;
  try {
    const stored = await browser.storage.local.get(M365_SETTINGS_KEY);
    const saved = stored[M365_SETTINGS_KEY] || {};
    m365EntryUrl = normalizeM365EntryUrl(saved.entryUrl);
    m365SettingsLoaded = true;
  } catch (error) {
    m365SettingsLoadPromise = null;
    throw error;
  }
}

function ensureM365SettingsLoaded() {
  if (m365SettingsLoaded) return Promise.resolve();
  if (!m365SettingsLoadPromise) m365SettingsLoadPromise = loadM365Settings();
  return m365SettingsLoadPromise;
}

async function saveM365Settings() {
  await browser.storage.local.set({
    [M365_SETTINGS_KEY]: { entryUrl: m365EntryUrl, updatedAt: Date.now() },
  });
}

async function setM365EntryUrl(value) {
  const normalized = normalizeM365EntryUrl(value);
  if (!normalized) return false;
  await ensureM365SettingsLoaded();
  if (normalized === m365EntryUrl) return true;
  m365EntryUrl = normalized;
  await saveM365Settings();
  return true;
}

function decodeM365JwtPart(part) {
  try {
    const normalized = String(part || "")
      .replace(/-/g, "+")
      .replace(/_/g, "/");
    return JSON.parse(
      atob(normalized.padEnd(Math.ceil(normalized.length / 4) * 4, "=")),
    );
  } catch (_) {
    return {};
  }
}

function decodeM365Jwt(token) {
  const parts = String(token || "").split(".");
  if (parts.length !== 3) return {};
  return decodeM365JwtPart(parts[1]);
}

function isM365JweToken(token) {
  const parts = String(token || "").split(".");
  if (parts.length !== 5) return false;
  const header = decodeM365JwtPart(parts[0]);
  return Boolean(header && typeof header.enc === "string" && header.enc);
}

function parseM365TokenRequest(details) {
  const tenantMatch = M365_TENANT_RE.exec(String(details.url || ""));
  if (!tenantMatch) return null;
  const requestUrl = new URL(details.url);
  const params = new URLSearchParams();
  const requestBody = details.requestBody || {};

  if (requestBody.formData && typeof requestBody.formData === "object") {
    for (const [name, values] of Object.entries(requestBody.formData)) {
      const value = Array.isArray(values) ? values[0] : values;
      if (value != null) params.set(name, String(value));
    }
  }
  if (Array.isArray(requestBody.raw)) {
    const decoder = new TextDecoder();
    let rawText = "";
    for (const part of requestBody.raw) {
      if (part && part.bytes)
        rawText += decoder.decode(part.bytes, { stream: true });
    }
    rawText += decoder.decode();
    for (const [name, value] of new URLSearchParams(rawText))
      params.set(name, value);
  }

  const scopes = String(params.get("scope") || "")
    .split(/\s+/)
    .filter(Boolean);
  const grantType = String(params.get("grant_type") || "");
  if (grantType !== "refresh_token" && grantType !== "authorization_code")
    return null;
  const isSydneyScope = scopes.includes(M365_SYDNEY_SCOPE);
  const clientId = String(
    params.get("client_id") || requestUrl.searchParams.get("client_id") || "",
  );
  const refreshToken = String(params.get("refresh_token") || "");
  if (!clientId || (grantType === "refresh_token" && !refreshToken))
    return null;
  // Accept a request from a sibling FOCI client only when it is the very
  // client_id the page handed us. The hardcoded sibling list used to widen this
  // to foreign clients whose refresh_token_expires_in says nothing about our
  // 24h SPA RT — it only moved _rtExp with other people's data.
  if (!isSydneyScope && !(Boolean(_clientId) && clientId === _clientId))
    return null;

  const tokenBody = Object.fromEntries(params.entries());

  delete tokenBody.refresh_token;
  delete tokenBody.code;
  delete tokenBody.code_verifier;
  delete tokenBody.client_request_id;
  delete tokenBody["client-request-id"];

  return {
    tenantId: tenantMatch[1],
    clientId,
    refreshToken,
    isSydneyScope,
    grantType,
    tokenContext: {
      version: 1,
      query: Object.fromEntries(requestUrl.searchParams.entries()),
      body: tokenBody,
    },
    capturedAt: Date.now(),
  };
}

async function applyCapturedM365Auth(captured) {
  await ensureM365AuthLoaded();
  if (Number(captured.capturedAt || 0) <= m365LatestTokenResponseAt) return;
  if (!captured.isSydneyScope) return;
  const capturedTid = String(captured.tenantId || "");
  // A token request from a *different* tenant/identity must never overwrite the
  // bound session or its tokenContext.
  if (_tid && capturedTid && capturedTid.toLowerCase() !== _tid.toLowerCase()) {
    log(
      "M365 capture: ignored Sydney credentials from foreign tenant %s (bound=%s client=%s)",
      shortTid(capturedTid),
      shortTid(_tid),
      shortTid(captured.clientId),
    );
    return;
  }
  _m365TokenContext = captured.tokenContext || {
    version: 1,
    query: {},
    body: {},
  };
  if (!captured.refreshToken) {
    // authorization_code request bodies carry no refresh_token; only refresh the
    // token context (the response is applied and verified separately).
    if (_rt && _clientId && _tid) await saveM365Auth();
    return;
  }
  const previousTid = _tid;
  const tenantChanged = Boolean(
    previousTid &&
    capturedTid &&
    previousTid.toLowerCase() !== capturedTid.toLowerCase(),
  );
  _tid = capturedTid || _tid;
  _clientId = captured.clientId;
  // A request body carries the RT the page is about to *consume* (an older or
  // unverified generation). Only seed _rt from a request on cold start (!_rt)
  // or when switching tenants; once _rt is populated, only validated token
  // responses and self-refresh calls may update it.
  const seededRt = !_rt || tenantChanged;
  if (seededRt) {
    _rt = captured.refreshToken;
  }
  // Only drop the cached access token when the bound identity really changed;
  // wiping it on every page rotation caused constant "no Sydney token" churn.
  if (tenantChanged) _sydney = { token: "", exp: 0 };
  // Deliberately NOT clearing m365ReauthRequired here: an unverified request
  // body is not proof that the session works again. Only a validated token
  // response may clear the flag (see applyM365TokenResponse).
  await saveM365Auth();
  if (!seededRt) return;
  log(
    "captured Sydney refresh credentials tenant=%s client=%s rt=%s context=%s",
    shortTid(_tid),
    shortTid(_clientId),
    captured.refreshToken ? "present" : "missing",
    captured.grantType,
  );
}

const m365TokenRequestExtraInfo =
  browser.webRequest &&
  typeof browser.webRequest.filterResponseData === "function"
    ? ["requestBody", "blocking"]
    : ["requestBody"];

browser.webRequest.onBeforeRequest.addListener(
  (details) => {
    // Ignore extension-originated token refresh requests (tabId === -1) so
    // background.js never self-captures its own outgoing /token calls.
    if (details.tabId < 0) return;
    // Harness probe requests are synthetic: they must never overwrite the
    // captured token context (that stripped the brk_* params and made the next
    // top-level /authorize fail with AADSTS50011).
    if (/[?&]ve_probe=1\b/.test(String(details.url || ""))) return;
    let captured;
    try {
      captured = parseM365TokenRequest(details);
    } catch (error) {
      log("failed to parse M365 token request:", compactError(error));
      return;
    }
    if (!captured) return;
    if (captured.isSydneyScope) {
      applyCapturedM365Auth(captured).catch((error) => {
        log("failed to persist M365 credentials:", compactError(error));
      });
    }
    if (
      browser.webRequest &&
      typeof browser.webRequest.filterResponseData === "function"
    ) {
      try {
        const filter = browser.webRequest.filterResponseData(details.requestId);
        const decoder = new TextDecoder();
        let body = "";
        let total = 0;
        filter.ondata = (event) => {
          filter.write(event.data);
          total += event.data.byteLength;
          if (total <= 256 * 1024) {
            body += decoder.decode(event.data, { stream: true });
          }
        };
        filter.onstop = () => {
          try {
            body += decoder.decode();
            filter.close();
            if (!body) return;
            const payload = JSON.parse(body);
            if (payload && payload.access_token) {
              applyM365TokenResponse({
                tenantId: captured.tenantId,
                clientId: captured.clientId,
                isSydneyScope: captured.isSydneyScope,
                grantType: captured.grantType,
                accessToken: String(payload.access_token),
                refreshToken: String(
                  payload.refresh_token || captured.refreshToken || "",
                ),
                expiresIn: Number(payload.expires_in || 0),
                refreshTokenExpiresIn: Number(
                  payload.refresh_token_expires_in || 0,
                ),
                // Order RT generations by when the request started, never by
                // when the response happened to arrive.
                requestStartedAt: Number(details.timeStamp || 0),
                capturedAt: Date.now(),
              }).catch((err) => {
                log(
                  "filterResponseData M365 token apply error:",
                  compactError(err),
                );
              });
            }
          } catch (_) {
            try {
              filter.close();
            } catch (_) {}
          }
        };
        filter.onerror = () => {
          try {
            filter.close();
          } catch (_) {}
        };
      } catch (_) {}
    }
  },
  { urls: ["https://login.microsoftonline.com/*/oauth2/v2.0/token*"] },
  m365TokenRequestExtraInfo,
);

// A refresh token must never move backwards. Each /token response carries a new
// RT generation, and the page fires one request per resource, so a slow response
// for an *earlier* generation can land after a later one. Adopting it would put
// the old refresh token back in use — the exact behaviour that is forbidden here
// (requirement: "拿到新 RT 就用新的，不要刷旧 RT").
// Ordering is by the request's start time, not by arrival: arrival order is what
// the responses violate.
// Empirical window (ms) grouping parallel /token requests fired in a single
// page-load burst (e.g. 4-9 requests within a few milliseconds). Not an AAD
// protocol constant.
const M365_RT_CONCURRENT_SLOP_MS = 250;

function noteRefreshTokenGeneration(observedAt, label) {
  const at = Math.floor(Number(observedAt || 0));
  if (at <= 0) return true;
  if (at + M365_RT_CONCURRENT_SLOP_MS < _rtNewestObservedAt) {
    log(
      "M365 token apply: refused an older RT generation (%dms behind the newest, src=%s) — never fall back to a previous refresh token",
      _rtNewestObservedAt - at,
      String(label || "page"),
    );
    return false;
  }
  _rtNewestObservedAt = Math.max(_rtNewestObservedAt, at);
  return true;
}

async function applyM365TokenResponse(message) {
  await ensureM365AuthLoaded();
  const accessToken = String(message.accessToken || "");
  const refreshToken = String(message.refreshToken || "");
  const clientId = String(message.clientId || "");
  const tenantId = String(message.tenantId || "");
  const capturedAt = Number(message.capturedAt || Date.now());
  const source = String(message.source || "page");
  if (!accessToken || !refreshToken || !clientId || !tenantId) {
    log(
      "M365 token apply: incomplete response (accessToken=%s refreshToken=%s clientId=%s tenantId=%s src=%s)",
      Boolean(accessToken),
      Boolean(refreshToken),
      Boolean(clientId),
      Boolean(tenantId),
      source,
    );
    throw new Error("incomplete M365 token response");
  }
  const claims = decodeM365Jwt(accessToken);
  const isJwe = !claims.oid && isM365JweToken(accessToken);
  const audience = String(claims.aud || "").replace(/\/$/, "");
  const claimTid = String(claims.tid || "");
  const wildcardTenant =
    tenantId === "organizations" || tenantId === "common" || !tenantId;
  const resolvedTid = claimTid || (!wildcardTenant ? tenantId : "");
  if ((!claims.oid && !isJwe) || (!isJwe && !resolvedTid)) {
    log(
      "M365 token apply: ignored response without identity claims (aud=%s src=%s)",
      audience,
      source,
    );
    return;
  }
  const identityMismatch = Boolean(
    claimTid &&
    !wildcardTenant &&
    claimTid.toLowerCase() !== tenantId.toLowerCase(),
  );
  if (identityMismatch) {
    const adopt = Boolean(message.adoptIdentity) || source === "pkce";
    log(
      "M365 token apply: identity mismatch request->%s token->%s aud=%s src=%s adopt=%s",
      shortTid(tenantId),
      shortTid(claimTid),
      audience,
      source,
      adopt,
    );
    if (!adopt) {
      // Ignoring is the correct behaviour here: a sibling/foreign capture must
      // neither overwrite the bound session nor be surfaced as a hard error
      // (it used to throw, and the rotated token was silently dropped).
      return;
    }
    log(
      "M365 token apply: adopting token identity %s (bound was %s)",
      shortTid(claimTid),
      shortTid(_tid),
    );
  }
  const isSydney = !isJwe && audience === "https://substrate.office.com/sydney";
  if (!isSydney) {
    // Sibling FOCI token rotation (e.g. api.spaces.skype.com / clients.config.office.net / JWE):
    // only update _rt and _rtExp if we already hold the exact same non-wildcard tenant and clientId.
    if (
      !_tid ||
      !_clientId ||
      !resolvedTid ||
      clientId !== _clientId ||
      resolvedTid.toLowerCase() !== _tid.toLowerCase()
    ) {
      log(
        "M365 token apply: ignored sibling-scope rotation aud=%s tid=%s client=%s (bound=%s/%s)",
        audience || (isJwe ? "jwe" : ""),
        shortTid(resolvedTid || tenantId),
        shortTid(clientId),
        shortTid(_tid),
        shortTid(_clientId),
      );
      return;
    }
  }
  // Deduplicate identical responses arriving via Firefox filterResponseData and
  // the page fetch hook, including interleaved A->B->A->B bursts (must run
  // synchronously after all guards pass and before awaiting).
  if (_recentAppliedM365AccessTokens.has(accessToken)) {
    return;
  }
  if (_recentAppliedM365AccessTokens.size >= M365_APPLIED_TOKEN_CACHE_LIMIT) {
    const oldest = _recentAppliedM365AccessTokens.values().next().value;
    if (oldest !== undefined) _recentAppliedM365AccessTokens.delete(oldest);
  }
  _recentAppliedM365AccessTokens.add(accessToken);
  // Decided once, synchronously: may this response refresh the stored RT?
  // (Access tokens are still stored below — only the RT may not go backwards.)
  const adoptRefreshToken = noteRefreshTokenGeneration(
    message.requestStartedAt || capturedAt,
    audience || (isJwe ? "jwe" : isSydney ? "sydney" : "sibling"),
  );
  if (!isJwe) {
    await ensureM365AutoLoginLoaded();
  }
  const tokenUpn = String(
    claims.upn || claims.preferred_username || claims.unique_name || "",
  ).trim();
  const promoteStagedAutoLogin = async () => {
    let dirty = false;
    if (tokenUpn && m365AutoLoginUsername !== tokenUpn) {
      m365AutoLoginUsername = tokenUpn;
      dirty = true;
    }
    if (
      m365AutoLoginEnabled &&
      m365AutoLoginRemember &&
      _stagedAutoLoginCreds &&
      Date.now() - _stagedAutoLoginCreds.stagedAt < 2 * 60 * 1000
    ) {
      if (_stagedAutoLoginCreds.password) {
        m365AutoLoginPassword = _stagedAutoLoginCreds.password;
        dirty = true;
      }
      if (_stagedAutoLoginCreds.username) {
        m365AutoLoginUsername = _stagedAutoLoginCreds.username;
        dirty = true;
      }
      log(
        "verified and saved M365 auto-login password from successful sign-in",
      );
    }
    _stagedAutoLoginCreds = null;
    // Only a real fresh sign-in (authorization_code) may lift the auto-login
    // pause. Clearing it on any captured rotation made the pause flag flap and
    // restarted the retry loop right after a failure.
    if (
      message.grantType === "authorization_code" &&
      (m365AutoLoginPaused || m365AutoLoginLastError)
    ) {
      log(
        "M365 auto-login pause cleared by successful sign-in (was: %s)",
        m365AutoLoginLastError || "paused",
      );
      m365AutoLoginPaused = false;
      m365AutoLoginLastError = "";
      m365AutoLoginLastErrorAt = 0;
      dirty = true;
    }
    if (message.grantType === "authorization_code" || dirty) {
      m365AutoLoginLastSuccessAt = Date.now();
      dirty = true;
    }
    if (dirty) {
      await saveM365AutoLogin();
    }
  };
  const now = Date.now();
  const rtExpIn = Number(message.refreshTokenExpiresIn || 0);
  const rtExpBefore = _rtExp;
  const updateRtExp = (label, isSydneyTarget) => {
    if (rtExpIn > 0) {
      // Guard against an anomalous initial value when seeding an unknown (0)
      // _rtExp from a non-Sydney response; not an authoritative AAD constant.
      const effectiveRtExpIn =
        !isSydneyTarget && rtExpBefore === 0
          ? Math.min(rtExpIn, 24 * 60 * 60)
          : rtExpIn;
      const candidate = now + effectiveRtExpIn * 1000;
      // authorization_code starts a brand new 24h cycle; refresh_token rotation
      // reports the shrinking remainder of the SAME absolute cap, so keep the
      // tighter estimate instead of pushing the deadline forward.
      _rtExp =
        message.grantType === "authorization_code"
          ? candidate
          : rtExpBefore > 0
            ? Math.min(rtExpBefore, candidate)
            : candidate;
    } else if (message.grantType === "authorization_code") {
      // No refresh_token_expires_in in the payload: only a fresh authorize
      // resets the countdown, so assume the documented full 24h window.
      _rtExp = now + 24 * 60 * 60 * 1000;
    }
    if (_rtExp !== rtExpBefore) {
      log(
        "M365 rtExp updated (%s): %s -> %s (rt_expires_in=%s grant=%s)",
        label,
        rtExpBefore ? m365RemainingMinutes(rtExpBefore) + "min" : "unknown",
        m365RemainingMinutes(_rtExp) + "min",
        rtExpIn,
        message.grantType,
      );
    } else if (!rtExpIn) {
      log(
        "M365 rtExp unchanged (%s): response carried no refresh_token_expires_in (grant=%s)",
        label,
        message.grantType,
      );
    }
  };
  if (!isSydney) {
    m365LatestTokenResponseAt = Math.max(m365LatestTokenResponseAt, capturedAt);
    if (adoptRefreshToken) _rt = refreshToken;
    updateRtExp(isJwe ? "sibling-jwe" : "sibling", false);
    // Non-Sydney rotation proves the family RT is alive, but is not proof for
    // Sydney itself: force the next pre-expiry tick to verify via refreshSydney
    // instead of clearing m365ReauthRequired directly.
    if (m365ReauthRequired) {
      _sydney = { token: "", exp: 0 };
    }
    // JWE tokens carry encrypted payloads (no readable oid/upn), so do not
    // touch staged auto-login credentials or account username.
    if (!isJwe) {
      await promoteStagedAutoLogin();
    }
    await saveM365Auth();
    log(
      "M365 token apply: sibling FOCI rotation stored (aud=%s rt=%s rt_expires_in=%s)",
      audience || (isJwe ? "jwe" : ""),
      adoptRefreshToken ? "updated" : "kept-newer",
      rtExpIn,
    );
    return;
  }
  m365LatestTokenResponseAt = Math.max(m365LatestTokenResponseAt, capturedAt);
  _tid = resolvedTid;
  _clientId = clientId;
  if (adoptRefreshToken) _rt = refreshToken;
  updateRtExp("sydney", true);
  const jwtExp = Number(claims.exp || 0) * 1000;
  const expiresIn = Number(message.expiresIn || 0);
  const responseExp =
    expiresIn > 0 ? now + Math.max(0, expiresIn - 60) * 1000 : 0;
  _sydney = { token: accessToken, exp: jwtExp || responseExp };
  clearM365ReauthRequired();
  await promoteStagedAutoLogin();
  m365LastRefreshError = "";
  m365LastRefreshErrorAt = 0;
  if (message.grantType === "authorization_code") {
    await resetM365ReauthState("authorization_code sign-in");
  }
  await saveM365Auth();
  log(
    "captured rotated Sydney refresh token and access token (grant=%s aud=%s tid=%s rt=%s sydneyExp=%dmin)",
    message.grantType,
    audience,
    shortTid(resolvedTid),
    adoptRefreshToken ? "updated" : "kept-newer",
    m365RemainingMinutes(_sydney.exp),
  );
}

async function refreshSydney(minRemainingMs = 90000) {
  await ensureM365AuthLoaded();
  const threshold = Math.max(30000, Number(minRemainingMs) || 90000);
  if (_sydney.token && _sydney.exp - Date.now() > threshold)
    return _sydney.token;
  if (_refreshing) return _refreshing;
  const missing = [];
  if (!_clientId) missing.push("client_id");
  if (!_rt) missing.push("refresh_token");
  if (!_tid) missing.push("tenant");
  if (missing.length)
    throw new Error("not captured yet: " + missing.join(", "));
  _refreshing = (async () => {
    const requestStartedAt = Date.now();
    const capturedQuery = _m365TokenContext.query || {};
    const capturedBody = _m365TokenContext.body || {};
    const endpoint = new URL(
      "https://login.microsoftonline.com/" + _tid + "/oauth2/v2.0/token",
    );
    for (const [name, value] of Object.entries(capturedQuery)) {
      if (value != null && value !== "")
        endpoint.searchParams.set(name, String(value));
    }
    endpoint.searchParams.set("client_id", _clientId);
    endpoint.searchParams.set(
      "client-request-id",
      typeof crypto !== "undefined" && crypto.randomUUID
        ? crypto.randomUUID()
        : "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (c) => {
            const r = (Math.random() * 16) | 0;
            return (c === "x" ? r : (r & 3) | 8).toString(16);
          }),
    );

    // Reuse the complete Outlook/MSAL form captured by webRequest. This keeps
    // redirect_uri, broker parameters, claims, X-AnchorMailbox and telemetry.
    const body = new URLSearchParams();
    for (const [name, value] of Object.entries(capturedBody)) {
      if (value != null) body.set(name, String(value));
    }
    body.set("client_id", _clientId);
    body.set("grant_type", "refresh_token");
    body.set("refresh_token", _rt);
    const capturedScope = String(capturedBody.scope || "").trim();
    if (!capturedScope) {
      throw new Error(
        "no captured scope — refusing to guess which scope to ask AAD for",
      );
    }
    body.set("scope", capturedScope);

    const response = await fetchTokenWithRetry(endpoint.href, {
      method: "POST",
      headers: {
        Accept: "*/*",
        "Content-Type": "application/x-www-form-urlencoded;charset=utf-8",
      },
      body: body.toString(),
    });
    const responseText = await response.text();
    let payload = {};
    try {
      payload = responseText ? JSON.parse(responseText) : {};
    } catch (_) {
      payload = { raw: responseText };
    }
    if (!response.ok || !payload.access_token) {
      if (isRefreshTokenExhaustedPayload(payload)) {
        handleRefreshTokenExhausted(payload);
      } else if (isInvalidGrantPayload(payload)) {
        pingM365TabsKeepAlive({ reauth: true }).catch(() => {});
      }
      const detail = String(
        payload.error_description ||
          payload.error ||
          payload.raw ||
          response.statusText ||
          response.status,
      )
        .replace(/\s+/g, " ")
        .slice(0, 500);
      throw new Error(
        "sydney refresh failed (" + response.status + "): " + detail,
      );
    }
    const claims = decodeM365Jwt(payload.access_token);
    const audience = String(claims.aud || "").replace(/\/$/, "");
    const expectedAudience = M365_SYDNEY_SCOPE.replace(/\/\.default$/i, "");
    if (audience !== expectedAudience) {
      throw new Error(
        "unexpected Sydney token audience: " + (claims.aud || "missing"),
      );
    }
    if (
      !claims.oid ||
      !claims.tid ||
      String(claims.tid).toLowerCase() !== _tid.toLowerCase()
    ) {
      throw new Error("Sydney token tenant/identity mismatch");
    }
    const rtExpiresIn = Number(payload.refresh_token_expires_in || 0);
    applyRotationRtExp(rtExpiresIn, "sydney-refresh");
    if (payload.refresh_token) {
      if (
        noteRefreshTokenGeneration(requestStartedAt, "sydney-refresh") &&
        payload.refresh_token !== _rt
      ) {
        _rt = payload.refresh_token;
      }
      await saveM365Auth();
    } else if (rtExpiresIn > 0) {
      await saveM365Auth();
    }
    const jwtExp = Number(claims.exp || 0) * 1000;
    _sydney = {
      token: payload.access_token,
      exp:
        jwtExp ||
        Date.now() +
          Math.max(0, Number(payload.expires_in || 3600) - 60) * 1000,
    };
    clearM365ReauthRequired();
    m365LastRefreshError = "";
    m365LastRefreshErrorAt = 0;
    if (m365ReauthState !== "none") {
      // A plain refresh just worked, which proves the session never needed a
      // fresh sign-in: cancel the attempt/tab and clear the terminal state.
      log(
        "M365 reauth: plain refresh succeeded while state=%s; clearing reauth episode",
        m365ReauthState,
      );
      closePendingTopLevelReauthTab();
      resetM365ReauthState("plain refresh succeeded").catch(() => {});
    }
    log(
      "M365 sydney refresh ok (sydneyExp=%dmin %s)",
      m365RemainingMinutes(_sydney.exp),
      m365StateLine(),
    );
    return _sydney.token;
  })();
  try {
    return await _refreshing;
  } catch (error) {
    m365LastRefreshError = compactError(error);
    m365LastRefreshErrorAt = Date.now();
    throw error;
  } finally {
    _refreshing = null;
  }
}

const TOKEN_RETRYABLE_STATUSES = new Set([408, 425, 429, 500, 502, 503, 504]);
async function fetchTokenWithRetry(url, init, maxAttempts = 3) {
  let lastErr = null;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      const res = await fetch(url, init);
      if (
        !TOKEN_RETRYABLE_STATUSES.has(res.status) ||
        attempt === maxAttempts
      ) {
        return res;
      }
      lastErr = new Error("HTTP " + res.status);
      try {
        if (res.body && typeof res.body.cancel === "function") {
          await res.body.cancel();
        }
      } catch (_) {}
    } catch (err) {
      lastErr = err;
      if (attempt === maxAttempts) throw err;
    }
    const delay =
      Math.min(4000, 800 * 2 ** (attempt - 1)) +
      Math.floor(Math.random() * 250);
    await new Promise((r) => setTimeout(r, delay));
  }
  throw lastErr || new Error("token fetch failed");
}

// Mint a Teams/IC3-audience token for the AMS artifact endpoint. Uses the SAME
// captured FOCI refresh token as Sydney/SharePoint (same token endpoint, only
// the scope differs); FOCI lets one family refresh token mint tokens for
// sibling resources, exactly as refreshSharePointAccessToken does for the
// SharePoint audience. Cached with a 90s skew like the others.
async function refreshIc3Token(minRemainingMs = 90000) {
  await ensureM365AuthLoaded();
  const threshold = Math.max(30000, Number(minRemainingMs) || 90000);
  if (_ic3.token && _ic3.exp - Date.now() > threshold) return _ic3.token;
  if (_ic3Refreshing) return _ic3Refreshing;
  if (!_clientId || !_rt || !_tid)
    throw new Error(
      "IC3 token unavailable: M365 refresh credentials not captured",
    );
  _ic3Refreshing = (async () => {
    const requestStartedAt = Date.now();
    const capturedQuery = _m365TokenContext.query || {};
    const capturedBody = _m365TokenContext.body || {};
    const endpoint = new URL(
      "https://login.microsoftonline.com/" + _tid + "/oauth2/v2.0/token",
    );
    for (const [name, value] of Object.entries(capturedQuery)) {
      if (value != null && value !== "")
        endpoint.searchParams.set(name, String(value));
    }
    endpoint.searchParams.set("client_id", _clientId);
    endpoint.searchParams.set("client-request-id", crypto.randomUUID());
    const body = new URLSearchParams();
    for (const [name, value] of Object.entries(capturedBody)) {
      if (value != null) body.set(name, String(value));
    }
    body.set("client_id", _clientId);
    body.set("grant_type", "refresh_token");
    body.set("refresh_token", _rt);
    body.set("scope", M365_IC3_SCOPE + " openid profile offline_access");
    const response = await fetchTokenWithRetry(endpoint.href, {
      method: "POST",
      headers: {
        Accept: "*/*",
        "Content-Type": "application/x-www-form-urlencoded;charset=utf-8",
      },
      body: body.toString(),
    });
    const text = await response.text();
    let payload = {};
    try {
      payload = text ? JSON.parse(text) : {};
    } catch (_) {
      payload = { raw: text };
    }
    if (!response.ok || !payload.access_token) {
      if (isRefreshTokenExhaustedPayload(payload)) {
        handleRefreshTokenExhausted(payload);
      } else if (isInvalidGrantPayload(payload)) {
        pingM365TabsKeepAlive({ reauth: true }).catch(() => {});
      }
      const detail = String(
        payload.error_description ||
          payload.error ||
          payload.raw ||
          response.status,
      )
        .replace(/\s+/g, " ")
        .slice(0, 500);
      throw new Error(
        "IC3 token refresh failed (" + response.status + "): " + detail,
      );
    }
    const claims = decodeM365Jwt(payload.access_token);
    const audience = String(claims.aud || "")
      .toLowerCase()
      .replace(/\/$/, "");
    // Tolerant audience check (mirrors the SharePoint one): accept the resource
    // URL form or a bare "ic3" app-id/audience.
    if (!audience.includes("ic3")) {
      throw new Error(
        "unexpected IC3 token audience: " + (claims.aud || "missing"),
      );
    }
    if (claims.tid && String(claims.tid).toLowerCase() !== _tid.toLowerCase())
      throw new Error("IC3 token tenant mismatch");
    // One-time diagnostic: decode the minted token's audience and scopes so the
    // AMS-authorization surface can be confirmed at a glance (aud must be
    // ic3.teams.office.com; scp shows whether AMS needs an object scope beyond
    // the default Teams set). Logged once per process, only under debug, and
    // never prints the token itself.
    if (m365Debug && !_ic3Probed) {
      _ic3Probed = true;
      log(
        "IC3 token probe: aud=%s appid=%s scp=%s",
        claims.aud || "<none>",
        claims.appid || claims.azp || "<none>",
        claims.scp || claims.roles || "<none>",
      );
    }
    const rtExpiresIn = Number(payload.refresh_token_expires_in || 0);
    applyRotationRtExp(rtExpiresIn, "ic3-refresh");
    if (payload.refresh_token) {
      if (
        noteRefreshTokenGeneration(requestStartedAt, "ic3-refresh") &&
        payload.refresh_token !== _rt
      ) {
        _rt = payload.refresh_token;
      }
      await saveM365Auth();
    } else if (rtExpiresIn > 0) {
      await saveM365Auth();
    }
    const jwtExp = Number(claims.exp || 0) * 1000;
    _ic3 = {
      token: payload.access_token,
      exp:
        jwtExp ||
        Date.now() +
          Math.max(0, Number(payload.expires_in || 3600) - 60) * 1000,
    };
    clearM365ReauthRequired();
    return _ic3.token;
  })();
  try {
    return await _ic3Refreshing;
  } finally {
    _ic3Refreshing = null;
  }
}

async function refreshSharePointAccessToken(siteUrl, minRemainingMs = 90000) {
  await ensureM365AuthLoaded();
  const origin = new URL(String(siteUrl || "")).origin;
  if (!origin.endsWith(".sharepoint.com"))
    throw new Error("invalid SharePoint token origin");
  const threshold = Math.max(30000, Number(minRemainingMs) || 90000);
  if (
    _sharePoint.token &&
    _sharePoint.origin === origin &&
    _sharePoint.exp - Date.now() > threshold
  ) {
    // Keep the return contract identical for both fresh and cached tokens.
    // content-sharepoint.js expects the complete Authorization header value.
    return "Bearer " + _sharePoint.token;
  }
  if (_sharePointRefreshing) return _sharePointRefreshing;
  if (!_clientId || !_rt || !_tid)
    throw new Error(
      "SharePoint bearer unavailable: M365 refresh credentials not captured",
    );
  _sharePointRefreshing = (async () => {
    const requestStartedAt = Date.now();
    const capturedQuery = _m365TokenContext.query || {};
    const capturedBody = _m365TokenContext.body || {};
    const endpoint = new URL(
      "https://login.microsoftonline.com/" + _tid + "/oauth2/v2.0/token",
    );
    // Match the successful Outlook/MSAL refresh request. The previous build
    // sent a minimal form, which dropped broker/redirect/claims context and
    // caused AADSTS70000 even though the refresh token itself was current.
    for (const [name, value] of Object.entries(capturedQuery)) {
      if (value != null && value !== "")
        endpoint.searchParams.set(name, String(value));
    }
    endpoint.searchParams.set("client_id", _clientId);
    endpoint.searchParams.set("client-request-id", crypto.randomUUID());
    const body = new URLSearchParams();
    for (const [name, value] of Object.entries(capturedBody)) {
      if (value != null) body.set(name, String(value));
    }
    body.set("client_id", _clientId);
    body.set("grant_type", "refresh_token");
    body.set("refresh_token", _rt);
    body.set("scope", origin + "/.default openid profile offline_access");
    const response = await fetchTokenWithRetry(endpoint.href, {
      method: "POST",
      headers: {
        Accept: "*/*",
        "Content-Type": "application/x-www-form-urlencoded;charset=utf-8",
      },
      body: body.toString(),
    });
    const text = await response.text();
    let payload = {};
    try {
      payload = text ? JSON.parse(text) : {};
    } catch (_) {
      payload = { raw: text };
    }
    if (!response.ok || !payload.access_token) {
      if (isRefreshTokenExhaustedPayload(payload)) {
        handleRefreshTokenExhausted(payload);
      } else if (isInvalidGrantPayload(payload)) {
        pingM365TabsKeepAlive({ reauth: true }).catch(() => {});
      }
      const detail = String(
        payload.error_description ||
          payload.error ||
          payload.raw ||
          response.status,
      )
        .replace(/\s+/g, " ")
        .slice(0, 500);
      throw new Error(
        "SharePoint token refresh failed (" + response.status + "): " + detail,
      );
    }
    const claims = decodeM365Jwt(payload.access_token);
    const host = new URL(origin).hostname.toLowerCase();
    const audience = String(claims.aud || "").toLowerCase();
    if (
      !audience.includes(host) &&
      !audience.includes("00000003-0000-0ff1-ce00-000000000000")
    )
      throw new Error(
        "unexpected SharePoint token audience: " + (claims.aud || "missing"),
      );
    if (claims.tid && String(claims.tid).toLowerCase() !== _tid.toLowerCase())
      throw new Error("SharePoint token tenant mismatch");
    const rtExpiresIn = Number(payload.refresh_token_expires_in || 0);
    applyRotationRtExp(rtExpiresIn, "sharepoint-refresh");
    if (payload.refresh_token) {
      if (
        noteRefreshTokenGeneration(requestStartedAt, "sharepoint-refresh") &&
        payload.refresh_token !== _rt
      ) {
        _rt = payload.refresh_token;
      }
      await saveM365Auth();
    } else if (rtExpiresIn > 0) {
      await saveM365Auth();
    }
    const jwtExp = Number(claims.exp || 0) * 1000;
    _sharePoint = {
      token: payload.access_token,
      exp:
        jwtExp ||
        Date.now() +
          Math.max(0, Number(payload.expires_in || 3600) - 60) * 1000,
      origin,
    };
    clearM365ReauthRequired();
    return "Bearer " + _sharePoint.token;
  })();
  try {
    return await _sharePointRefreshing;
  } finally {
    _sharePointRefreshing = null;
  }
}

function isRefreshTokenExhaustedPayload(payload) {
  if (!payload || typeof payload !== "object") return false;
  const desc = String(payload.error_description || payload.raw || "");
  const err = String(payload.error || "").toLowerCase();
  return (
    /AADSTS70008[24]\b|AADSTS50058\b/i.test(desc) ||
    err === "interaction_required" ||
    err === "login_required"
  );
}

function isInvalidGrantPayload(payload) {
  if (!payload || typeof payload !== "object") return false;
  if (isRefreshTokenExhaustedPayload(payload)) return false;
  const err = String(payload.error || "").toLowerCase();
  const desc = String(payload.error_description || "");
  return err === "invalid_grant" || /AADSTS7000[08]\b/i.test(desc);
}

function markM365ReauthRequired(reason) {
  const text = String(reason || "24h SPA 刷新令牌已到期，需重新登录");
  const wasRequired = m365ReauthRequired;
  m365ReauthRequired = true;
  m365ReauthReason = text;
  m365ReauthRequiredAt = Date.now();
  if (_preExpiryRetryTimer) {
    clearTimeout(_preExpiryRetryTimer);
    _preExpiryRetryTimer = null;
  }
  _preExpiryRetryAttempt = 0;
  setBadge("!", "#d93025").catch(() => {});
  // Page-side authorize popups are never useful (they are the same silent
  // attempt that already failed), so they stay suppressed regardless of the
  // gate state. Allowing them back on give-up is what produced the 60s popup
  // storm that was observed before the suppression existed.
  pingM365TabsKeepAlive({
    authHealthy: false,
    suppressPopups: true,
  }).catch(() => {});
  if (!wasRequired) {
    log("M365 entered NEEDS_REAUTH state:", text, "—", m365StateLine());
  } else {
    dlog("M365 still NEEDS_REAUTH:", text, "—", m365StateLine());
  }
}

function clearM365ReauthRequired() {
  if (!m365ReauthRequired && !m365ReauthReason) return;
  log(
    "M365 left NEEDS_REAUTH state (was: %s) — %s",
    m365ReauthReason || "-",
    m365StateLine(),
  );
  m365ReauthRequired = false;
  m365ReauthReason = "";
  m365ReauthRequiredAt = 0;
  setBadge("", null).catch(() => {});
  pingM365TabsKeepAlive({ authHealthy: true, suppressPopups: true }).catch(
    () => {},
  );
}

function handleRefreshTokenExhausted(payload) {
  const desc = String(
    payload && (payload.error_description || payload.error || ""),
  );
  const codeMatch = desc.match(/AADSTS\d+/i);
  const code = codeMatch ? codeMatch[0].toUpperCase() : "";
  // Only these two codes really mean "the 24h SPA refresh token is dead".
  // interaction_required / AADSTS50058 mean "silent SSO failed", which needs an
  // interactive login but must not be reported as the 24h cap.
  const isHardCap = /AADSTS70008[24]/.test(code);
  const reason = isHardCap
    ? `${code}: 24h SPA 刷新令牌已达上限，需重新登录`
    : `${code || "interaction_required"}: 静默续期失败，需交互式登录`;
  const now = Date.now();
  const repeat = now - _lastRtRejectLogAt < 60 * 1000;
  _lastRtRejectLogAt = now;
  const line = isHardCap ? "24h-cap" : "silent-sso";
  const detail = desc.replace(/\s+/g, " ").slice(0, 240) || "-";
  if (repeat) {
    dlog("M365 refresh rejected again (%s): %s", line, detail);
  } else {
    log("M365 refresh rejected (%s): %s | %s", line, reason, detail);
  }
  markM365ReauthRequired(reason);
  tryTopLevelSilentReauth("rt-exhausted").catch((error) =>
    log("M365 reauth from rt-exhausted failed:", compactError(error)),
  );
}

function base64UrlEncodeBytes(bytes) {
  let binary = "";
  for (let i = 0; i < bytes.length; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary)
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/, "");
}

async function createPkcePair() {
  const randomBytes = new Uint8Array(32);
  crypto.getRandomValues(randomBytes);
  const verifier = base64UrlEncodeBytes(randomBytes);
  const encoded = new TextEncoder().encode(verifier);
  const digest = await crypto.subtle.digest("SHA-256", encoded);
  const challenge = base64UrlEncodeBytes(new Uint8Array(digest));
  return { verifier, challenge };
}

function shortTid(value) {
  const text = String(value || "");
  return text.length > 12 ? text.slice(0, 8) + "…" + text.slice(-4) : text;
}

function m365RemainingMinutes(at) {
  const exp = Number(at || 0);
  if (!exp) return -1;
  return Math.round((exp - Date.now()) / 60000);
}

// A refresh_token rotation reports the shrinking remainder of the SAME absolute
// 24h cap, so the deadline must never move forward, and a response without
// refresh_token_expires_in must not leave a stale past-due value behind.
function applyRotationRtExp(rtExpiresIn, label) {
  const before = _rtExp;
  const value = Number(rtExpiresIn || 0);
  if (value > 0) {
    const candidate = Date.now() + value * 1000;
    _rtExp = before > 0 ? Math.min(before, candidate) : candidate;
  }
  if (_rtExp !== before) {
    log(
      "M365 rtExp updated (%s): %s -> %s (rt_expires_in=%s)",
      label,
      before ? m365RemainingMinutes(before) + "min" : "unknown",
      m365RemainingMinutes(_rtExp) + "min",
      value,
    );
  } else if (!value) {
    log(
      "M365 rtExp unchanged (%s): rotation response carried no refresh_token_expires_in",
      label,
    );
  }
  return _rtExp;
}

// ---------------------------------------------------------------------------
// Reauth command over the EXISTING Voice Edge bridge (ws://127.0.0.1:5003/ws).
//
// There is exactly one socket between this extension and the outside world, so
// the harness binds that same port instead of opening a second channel (a
// parallel port meant two protocols, two retry timers, and a client in shipped
// code that could never connect in production).
//
// Contract (opt-in — nothing is emitted unless the client asks):
//   in   {type:"M365_REAUTH_NOW", force?:true, requestId?:"..."}
//   out  {type:"M365_REAUTH_STARTED", requestId, started, ...state}
//        {type:"M365_REAUTH_STATE",   requestId, ...state}      (1/s while running)
//        {type:"M365_REAUTH_DONE",    requestId, ...state, log:[last 80 lines]}
function m365BridgeState() {
  return {
    at: Date.now(),
    build: M365_BUILD_TAG,
    state: m365ReauthState,
    reason: m365ReauthStateReason || "",
    rtRemainMin: m365RemainingMinutes(_rtExp),
    rtExpiresAt: Number(_rtExp || 0),
    sydneyRemainMin: m365RemainingMinutes(_sydney.exp),
    reauthRequired: Boolean(m365ReauthRequired),
    pendingTabId: _pendingTopLevelReauth
      ? Number(_pendingTopLevelReauth.tabId || 0)
      : 0,
  };
}

// Ask the real OWA page to fire ONE /token request carrying a deliberately bogus
// old refresh_token. If the page-world swap works, AAD answers 200 with an access
// token — that is end-to-end proof that the page used our RT. Only a status code
// and a boolean travel back; no token ever leaves the page.
async function runPageTokenProbe() {
  await ensureM365AuthLoaded();
  if (!_tid || !_clientId || !_rt) {
    log("M365 probe: skipped, credentials not captured yet");
    return 0;
  }
  const template = Object.assign(
    {},
    (_m365TokenContext && _m365TokenContext.body) || {},
  );
  for (const key of ["code", "code_verifier", "refresh_token", "grant_type"]) {
    delete template[key];
  }
  // The probe replays the page's own request body verbatim. It is NOT patched:
  // a context that lacks the broker pair cannot tell us what the page really
  // sends, so the probe stays silent instead of inventing parameters (that
  // "invented" body is what poisons the cache with AADSTS50011).
  const missing = [
    "redirect_uri",
    "brk_client_id",
    "brk_redirect_uri",
    "scope",
  ].filter((key) => !String(template[key] || "").trim());
  if (missing.length) {
    log(
      "M365 probe: skipped, the captured page request is missing %s",
      missing.join(", "),
    );
    return 0;
  }
  const payload = {
    type: "M365_PAGE_TOKEN_PROBE",
    tid: _tid,
    clientId: _clientId,
    // 自带 RT，免得探针依赖"刚才是否推过 KEEPALIVE"
    rt: _rt || "",
    rtClientId: _clientId || "",
    // 真 AAD 要求 broker 参数成对出现，探针照抄页面自己的请求体（含 scope）
    template,
  };
  let sent = 0;
  try {
    const tabs = await browser.tabs.query({
      url: M365_TAB_URL_PATTERNS,
    });
    for (const tab of tabs) {
      if (!tab || tab.id == null) continue;
      try {
        await browser.tabs.sendMessage(
          tab.id,
          { __veM365ToPage: true, payload },
          { frameId: 0 },
        );
        sent += 1;
      } catch (_) {}
    }
  } catch (error) {
    log("M365 probe: query failed:", compactError(error));
  }
  log(
    "M365 probe: asked %d OWA tab(s) to fire /token with a bogus old RT (tid=%s)",
    sent,
    shortTid(_tid),
  );
  return sent;
}

async function handleM365BridgeReauthCommand(message) {
  const requestId = String((message && message.requestId) || "");
  const force = !(message && message.force === false);
  await Promise.all([
    ensureM365AuthLoaded(),
    ensureM365AutoLoginLoaded(),
    ensureM365ReauthStateLoaded(),
  ]);
  log(
    "M365 bridge: reauth command received (force=%s request=%s) — %s",
    force,
    requestId || "-",
    m365StateLine(),
  );
  closePendingTopLevelReauthTab();
  if (force || m365ReauthState === "attempting") {
    await resetM365ReauthState("bridge command");
  }
  _reauthNextAttemptAt = 0;
  const rtBefore = Number(_rtExp || 0);
  const sydneyBefore = Number(_sydney.exp || 0);
  const started = await tryTopLevelSilentReauth("bridge", { force });
  sendM365({
    type: "M365_REAUTH_STARTED",
    requestId,
    started,
    ...m365BridgeState(),
  });
  const deadline = Date.now() + 90000;
  while (Date.now() < deadline) {
    await new Promise((resolve) => setTimeout(resolve, 1000));
    sendM365({
      type: "M365_REAUTH_STATE",
      requestId,
      ...m365BridgeState(),
    });
    const refreshed =
      Number(_rtExp || 0) !== rtBefore ||
      Number(_sydney.exp || 0) !== sydneyBefore;
    if (refreshed || m365ReauthState === "terminal" || !started) break;
  }
  sendM365({
    type: "M365_REAUTH_DONE",
    requestId,
    ...m365BridgeState(),
    log: _m365LogRing.slice(-80),
  });
  log("M365 bridge: reauth command finished — %s", m365StateLine());
}

function m365StateLine() {
  return (
    "rtRemain=" +
    m365RemainingMinutes(_rtExp) +
    "min sydneyRemain=" +
    m365RemainingMinutes(_sydney.exp) +
    "min reauth=" +
    Boolean(m365ReauthRequired) +
    " state=" +
    m365ReauthState +
    (m365ReauthStateReason
      ? "(" + String(m365ReauthStateReason).slice(0, 80) + ")"
      : "")
  );
}

async function loadM365ReauthState() {
  if (_reauthStateLoaded) return;
  try {
    const stored = await browser.storage.local.get(M365_REAUTH_STATE_KEY);
    const saved = stored[M365_REAUTH_STATE_KEY] || {};
    const status = String(saved.status || "none");
    m365ReauthState = ["none", "attempting", "terminal"].includes(status)
      ? status
      : "none";
    m365ReauthStateReason = String(saved.reason || "");
    m365ReauthStateAt = Number(saved.at || 0);
    log(
      "M365 reauth state loaded: %s (%s) at %s — %s",
      m365ReauthState,
      m365ReauthStateReason || "-",
      m365ReauthStateAt ? new Date(m365ReauthStateAt).toISOString() : "-",
      m365StateLine(),
    );
  } catch (error) {
    log("failed to load M365 reauth state:", compactError(error));
  }
  _reauthStateLoaded = true;
}

function ensureM365ReauthStateLoaded() {
  if (_reauthStateLoaded) return Promise.resolve();
  if (!_reauthStatePromise) _reauthStatePromise = loadM365ReauthState();
  return _reauthStatePromise;
}

async function saveM365ReauthState() {
  try {
    await browser.storage.local.set({
      [M365_REAUTH_STATE_KEY]: {
        status: m365ReauthState,
        reason: m365ReauthStateReason,
        at: m365ReauthStateAt,
      },
    });
  } catch (error) {
    log("failed to persist M365 reauth state:", compactError(error));
  }
}

async function setM365ReauthState(status, reason) {
  const previous = m365ReauthState;
  m365ReauthState = String(status || "none");
  m365ReauthStateReason = String(reason || "");
  m365ReauthStateAt = Date.now();
  await saveM365ReauthState();
  log(
    "M365 reauth state: %s -> %s (%s) — %s",
    previous,
    m365ReauthState,
    m365ReauthStateReason || "-",
    m365StateLine(),
  );
  return m365ReauthState;
}

// The single attempt failed. Nothing is retried, ever, until something proves
// the situation changed (fresh sign-in / plain refresh / explicit reset).
async function failM365Reauth(reason) {
  const text = String(reason || "auto-login failed");
  if (m365ReauthState === "terminal") {
    dlog("M365 reauth already terminal, ignoring extra failure: %s", text);
    return;
  }
  closePendingTopLevelReauthTab();
  await setM365ReauthState("terminal", text);
  // A failed *renewal* attempt is not the same as an expired session. While the
  // current RT still has hours left, claiming NEEDS_REAUTH is a lie that turns
  // the badge red and makes every status line contradictory (seen in the field:
  // `reauth=true state=terminal rtRemain=1385min`).
  const rtHealthy = Number(_rtExp || 0) - Date.now() > M365_RT_RENEW_WINDOW_MS;
  if (rtHealthy) {
    log(
      "M365 reauth: renewal attempt failed but the current RT is still valid (%dmin left) — session stays usable",
      m365RemainingMinutes(_rtExp),
    );
  } else {
    markM365ReauthRequired(text);
  }
  log(
    "M365 reauth: TERMINAL — the single automatic attempt failed (%s); no retry will be made",
    text.slice(0, 200),
  );
  notify(
    "Voice Edge：M365 自动登录失败",
    text.slice(0, 160) +
      "（单次尝试，不再自动重试；请在面板手动登录一次以恢复）",
  ).catch(() => {});
}

async function resetM365ReauthState(reason) {
  if (m365ReauthState === "none" && !m365ReauthStateReason) return;
  await setM365ReauthState("none", "");
  log("M365 reauth state reset (%s)", String(reason || "reset"));
}

function pendingReauthStore() {
  return (browser.storage && browser.storage.session) || browser.storage.local;
}

async function persistPendingReauth(pending) {
  if (!pending) return;
  try {
    await pendingReauthStore().set({
      [M365_PENDING_REAUTH_KEY]: {
        tabId: pending.tabId,
        windowId: pending.windowId != null ? pending.windowId : null,
        state: pending.state,
        verifier: pending.verifier,
        redirectUri: pending.redirectUri,
        clientId: pending.clientId,
        // The /token exchange is worthless without the broker pair: AAD answers
        // AADSTS90014 "required field 'brk_redirect_uri' is missing". These used
        // to live only in memory, so a service-worker restart between
        // /authorize and the code callback silently dropped them.
        scope: pending.scope,
        brkClientId: pending.brkClientId,
        brkRedirect: pending.brkRedirect,
        interactive: Boolean(pending.interactive),
        promptNone: Boolean(pending.promptNone),
        startedAt: Number(pending.startedAt || Date.now()),
      },
    });
  } catch (error) {
    log("failed to persist pending M365 reauth:", compactError(error));
  }
}

async function clearPersistedPendingReauth() {
  try {
    await pendingReauthStore().remove(M365_PENDING_REAUTH_KEY);
  } catch (_) {}
}

// A background/MV3 service-worker restart used to silently drop the pending
// PKCE verifier, so the authorization code returned by the login tab could
// never be exchanged and the tab was never closed. Restore it when possible.
async function restorePendingReauth() {
  if (_pendingTopLevelReauth) return true;
  try {
    const stored = await pendingReauthStore().get(M365_PENDING_REAUTH_KEY);
    const saved = stored[M365_PENDING_REAUTH_KEY];
    if (!saved || !saved.state || !saved.verifier) return false;
    const age = Date.now() - Number(saved.startedAt || 0);
    if (age > M365_PENDING_REAUTH_TTL_MS) {
      await clearPersistedPendingReauth();
      log("discarded stale pending M365 reauth (%dms old)", age);
      return false;
    }
    const timer = setTimeout(
      () => {
        log("restored M365 reauth tab timed out; failing tab %s", saved.tabId);
        failPendingReauth("restored login tab timed out");
      },
      Math.max(5000, M365_REAUTH_TAB_TIMEOUT_MS - age),
    );
    _pendingTopLevelReauth = {
      tabId: saved.tabId,
      windowId: saved.windowId != null ? saved.windowId : null,
      state: saved.state,
      verifier: saved.verifier,
      redirectUri: saved.redirectUri,
      clientId: saved.clientId,
      scope: saved.scope,
      brkClientId: saved.brkClientId,
      brkRedirect: saved.brkRedirect,
      interactive: Boolean(saved.interactive),
      promptNone: Boolean(saved.promptNone),
      timer,
      startedAt: Number(saved.startedAt || Date.now()),
      restored: true,
    };
    if (saved.tabId != null) {
      armM365LoginTab(saved.tabId);
    }
    log(
      "restored pending M365 reauth tab=%s state=%s age=%dms",
      saved.tabId,
      String(saved.state || "").slice(0, 20),
      age,
    );
    return true;
  } catch (error) {
    log("failed to restore pending M365 reauth:", compactError(error));
    return false;
  }
}

// Close login tabs we created (client_id / state marker match) that nobody is
// tracking any more. Without this every abandoned attempt stays open forever.
async function sweepOrphanReauthTabs(reason) {
  const sweepNow = Date.now();
  const forced = reason === "startup" || reason === "manual";
  if (!forced && sweepNow - _lastOrphanSweepAt < 60 * 1000) return;
  _lastOrphanSweepAt = sweepNow;
  try {
    const tabs = await browser.tabs.query({
      url: ["https://login.microsoftonline.com/*"],
    });
    const now = Date.now();
    for (const tab of tabs) {
      if (!tab || tab.id == null) continue;
      if (_pendingTopLevelReauth && _pendingTopLevelReauth.tabId === tab.id)
        continue;
      if (tab.active) continue;
      const url = String(tab.url || "");
      const stateMarker = /state=ve-reauth(%3A|:)/i.test(url);
      const ourClient = Boolean(_clientId) && url.includes(_clientId);
      const isAuthorize = /\/oauth2\/v2\.0\/authorize\b/i.test(url);
      const isRedirect = /\/mail\/oauthRedirect\.html([?#]|$)/i.test(url);
      const ours = stateMarker || (ourClient && (isAuthorize || isRedirect));
      if (!ours) continue;
      log(
        "M365 reauth: closing orphan login tab %s (%s) url=%s",
        tab.id,
        String(reason || "sweep"),
        url.slice(0, 160),
      );
      browser.tabs.remove(tab.id).catch(() => {});
    }
    // A tracked attempt (in flight or handed to the user) is never closed here:
    // it is the single attempt and the user may still be typing into it. If the
    // tab is closed by the user/harness, tabs.onRemoved turns that into a
    // terminal failure instead.
  } catch (error) {
    log("M365 reauth orphan sweep failed:", compactError(error));
  }
}

// Auto-login is unattended: there is deliberately NO "hand this window to the
// user" path and no state that waits for typing. If we cannot produce an
// authorization code by ourselves, the attempt fails immediately, the tab is
// closed and exactly one notification is raised. Keeping a live window around
// for a human was the source of the old multi-window mess (product
// decision: 禁止人工输入).
async function failPendingReauth(reason, options = {}) {
  const pending = _pendingTopLevelReauth;
  if (!pending) return;
  if (pending.timer) {
    clearTimeout(pending.timer);
    pending.timer = null;
  }
  if (pending.loadWatchdog) {
    clearTimeout(pending.loadWatchdog);
    pending.loadWatchdog = null;
  }
  if (pending.notifyTimer) {
    clearTimeout(pending.notifyTimer);
    pending.notifyTimer = null;
  }
  const environmental = Boolean(options && options.environmental);
  log(
    "M365 reauth: giving up on tab %s (%s) environmental=%s — closing it, no manual takeover",
    pending.tabId,
    String(reason || "-"),
    environmental,
  );
  if (environmental) {
    // Offline / page never loaded: nothing was proven about the credentials, so
    // this must NOT consume the episode. Close the tab, go back to `none`, and
    // wait before trying again (bounded, not a loop).
    closePendingTopLevelReauthTab();
    _reauthNextAttemptAt = Date.now() + M365_REAUTH_ENV_RETRY_MS;
    await setM365ReauthState("none", "environmental: " + String(reason || ""));
    log(
      "M365 reauth: environmental failure — new attempt allowed in %dmin",
      Math.round(M365_REAUTH_ENV_RETRY_MS / 60000),
    );
    return;
  }
  await failM365Reauth(String(reason || "auto-login failed"));
}

function closePendingTopLevelReauthTab() {
  const pending = _pendingTopLevelReauth;
  if (!pending) {
    clearPersistedPendingReauth().catch(() => {});
    return;
  }
  _pendingTopLevelReauth = null;
  clearPersistedPendingReauth().catch(() => {});
  if (pending.notifyTimer) {
    clearTimeout(pending.notifyTimer);
    pending.notifyTimer = null;
  }
  if (pending.timer) {
    clearTimeout(pending.timer);
    pending.timer = null;
  }
  if (pending.loadWatchdog) {
    clearTimeout(pending.loadWatchdog);
    pending.loadWatchdog = null;
  }
  if (pending.windowId != null && browser.windows && browser.windows.remove) {
    log("M365 reauth: closing login window %s", pending.windowId);
    browser.windows.remove(pending.windowId).catch(() => {});
  } else if (pending.tabId != null) {
    log("M365 reauth: closing top-level tab %s", pending.tabId);
    browser.tabs.remove(pending.tabId).catch(() => {});
  }
}

// One automatic login attempt per episode, period. `force` exists only for the
// harness and for an explicit user action in the panel (both reset the state
// first); nothing in the background retries by itself.
async function tryTopLevelSilentReauth(reason = "scheduled", options = {}) {
  await Promise.all([
    ensureM365AuthLoaded(),
    ensureM365AutoLoginLoaded(),
    ensureM365ReauthStateLoaded(),
  ]);
  const force = Boolean(options && options.force);
  if (!_tid || !_clientId) {
    log("M365 reauth: skipped, refresh credentials not captured yet");
    return false;
  }
  if (_pendingTopLevelReauth || _reauthCreating) {
    log(
      "M365 reauth: skipped, the single attempt is already in flight (tab=%s creating=%s) reason=%s",
      _pendingTopLevelReauth ? _pendingTopLevelReauth.tabId : "none",
      _reauthCreating,
      reason,
    );
    return false;
  }
  if (!force) {
    const sinceSignIn = Date.now() - _lastAuthCodeSignInAt;
    if (_lastAuthCodeSignInAt && sinceSignIn < 2 * 60 * 1000) {
      log(
        "M365 reauth: skipped, a real sign-in completed %ds ago — no second window needed",
        Math.round(sinceSignIn / 1000),
      );
      return false;
    }
    if (Date.now() < _reauthNextAttemptAt) {
      log(
        "M365 reauth: skipped, %dmin until the next allowed attempt (last failure was environmental)",
        Math.ceil((_reauthNextAttemptAt - Date.now()) / 60000),
      );
      return false;
    }
    if (m365ReauthState === "terminal") {
      log(
        "M365 reauth: skipped — this episode already failed terminally (%s). No retry by design; sign in manually in the panel/harness.",
        m365ReauthStateReason || "-",
      );
      return false;
    }
    if (m365ReauthState === "attempting") {
      log(
        "M365 reauth: skipped, state=%s (attempt in flight or handed to the user)",
        m365ReauthState,
      );
      return false;
    }
    if (m365AutoLoginPaused && m365AutoLoginLastError) {
      log(
        "M365 reauth: skipped, auto-login paused by an earlier error: %s",
        m365AutoLoginLastError,
      );
      return false;
    }
    if (!m365AutoLoginEnabled) {
      log(
        "M365 reauth: skipped, interactive auto-login is disabled in settings",
      );
      return false;
    }
  }
  if (typeof navigator !== "undefined" && navigator.onLine === false) {
    log("M365 reauth: skipped, browser offline");
    return false;
  }
  const now = Date.now();
  // Set the single-flight flag BEFORE the first await. It used to be set after
  // `await sweepOrphanReauthTabs()`, so two concurrent triggers (the refresh
  // rejection handler and the pre-expiry loop both fire in the same tick) each
  // passed the in-flight check and opened a login window: exactly the "two
  // windows at once" symptom seen in the field.
  _reauthCreating = true;
  try {
    await sweepOrphanReauthTabs(reason);
    const { verifier, challenge } = await createPkcePair();
    const stateId = crypto.randomUUID();
    const capturedBody = (_m365TokenContext && _m365TokenContext.body) || {};
    const capturedQuery = (_m365TokenContext && _m365TokenContext.query) || {};
    // AAD validates redirect_uri / brk_redirect_uri against what the app is
    // actually registered for, so a guessed value is AADSTS50011 on /authorize
    // (no password form ever renders) or AADSTS90014 on /token. There is no
    // default and no HAR-derived constant: every value below is read from the
    // page's own captured /token request, and when the page has not told us we
    // name the missing field and fail instead of borrowing a value.
    const redirectUri = String(
      capturedBody.redirect_uri || capturedQuery.redirect_uri || "",
    ).trim();
    const brkClientId = String(
      capturedBody.brk_client_id || capturedQuery.brk_client_id || "",
    ).trim();
    const webRedirect = String(
      capturedBody.brk_redirect_uri || capturedQuery.brk_redirect_uri || "",
    ).trim();
    const effectiveClientId =
      String(capturedBody.client_id || "").trim() || _clientId;
    const scope = String(capturedBody.scope || "").trim();
    const missingParams = [
      ["client_id", effectiveClientId],
      ["scope", scope],
      ["redirect_uri", redirectUri],
      ["brk_client_id", brkClientId],
      ["brk_redirect_uri", webRedirect],
    ]
      .filter(([, value]) => !value)
      .map(([name]) => name);
    if (missingParams.length) {
      const detail =
        "page token context incomplete: missing " + missingParams.join(", ");
      log(
        "M365 reauth: refusing to build /authorize — %s. The page's own /token request is the only source for these; no defaults by design.",
        detail,
      );
      await failM365Reauth(detail);
      return false;
    }
    if (!/^brk-/i.test(redirectUri)) {
      const detail =
        "captured redirect_uri is not a broker URI (" + redirectUri + ")";
      log("M365 reauth: refusing to build /authorize — %s", detail);
      await failM365Reauth(detail);
      return false;
    }
    const interactive = Boolean(options && options.interactive);
    const authUrl = new URL(
      `https://login.microsoftonline.com/${_tid}/oauth2/v2.0/authorize`,
    );
    authUrl.searchParams.set("client_id", effectiveClientId);
    authUrl.searchParams.set("scope", scope);
    authUrl.searchParams.set("redirect_uri", redirectUri);
    // The broker pair is what makes AAD hand the code back to this browser tab;
    // both halves come from the captured request, never from a constant.
    authUrl.searchParams.set("brk_client_id", brkClientId);
    authUrl.searchParams.set("brk_redirect_uri", webRedirect);
    authUrl.searchParams.set("response_mode", "fragment");
    authUrl.searchParams.set("response_type", "code");
    // Silent first (that is what OWA does and it works while the AAD session
    // cookie is alive). Only when AAD answers interaction_required do we open
    // the interactive page that can render #i0118.
    if (!interactive) authUrl.searchParams.set("prompt", "none");
    const hint = String(capturedBody.login_hint || "").trim();
    if (hint) authUrl.searchParams.set("login_hint", hint);
    else if (interactive && m365AutoLoginUsername)
      authUrl.searchParams.set("login_hint", m365AutoLoginUsername);
    // clidata / client_info are not identity values and never appear in the
    // page's /token request — they are broker-protocol flags of the interactive
    // authorize itself, which nothing else can supply. Everything AAD validates
    // (client_id, redirect_uri, brk_*, scope) is the page's, above.
    authUrl.searchParams.set("clidata", "1");
    authUrl.searchParams.set("client_info", "1");
    authUrl.searchParams.set("code_challenge", challenge);
    authUrl.searchParams.set("code_challenge_method", "S256");
    authUrl.searchParams.set("nonce", crypto.randomUUID());
    authUrl.searchParams.set("state", "ve-reauth:" + stateId);
    authUrl.searchParams.set("client-request-id", crypto.randomUUID());
    if (capturedBody.claims) {
      authUrl.searchParams.set("claims", String(capturedBody.claims));
    }
    const hasStoredPassword = Boolean(
      m365AutoLoginEnabled && !m365AutoLoginPaused && m365AutoLoginPassword,
    );
    log(
      "M365 reauth: opening the single top-level authorize tab reason=%s tenant=%s client=%s redirect=%s account=%s storedPassword=%s — %s",
      reason,
      shortTid(_tid),
      shortTid(effectiveClientId),
      redirectUri,
      m365AutoLoginUsername || "-",
      hasStoredPassword,
      m365StateLine(),
    );
    // A small popup window (like OWA's own sign-in popup) instead of a tab:
    // it is visually obvious, it does not join the tab strip, and closing it
    // cannot be confused with closing a page. Falls back to a tab when the
    // browser has no windows API.
    let tabId = null;
    let windowId = null;
    if (browser.windows && typeof browser.windows.create === "function") {
      try {
        const win = await browser.windows.create({
          url: authUrl.href,
          type: "popup",
          width: 520,
          height: 760,
        });
        windowId = win && win.id != null ? win.id : null;
        tabId =
          win && win.tabs && win.tabs[0] && win.tabs[0].id != null
            ? win.tabs[0].id
            : null;
        // Diagnostic only (the login window was reported as "never
        // opened" while the log claimed otherwise). Record what the browser
        // says about the window and about the focus request, so "no window
        // appeared" can be answered with facts instead of guesswork.
        log(
          "M365 reauth: login window created id=%s tab=%s state=%s focused=%s bounds=%s,%s %sx%s",
          windowId,
          tabId,
          win && win.state,
          win && win.focused,
          win && win.left,
          win && win.top,
          win && win.width,
          win && win.height,
        );
        if (windowId != null && browser.windows.update) {
          browser.windows
            .update(windowId, { focused: true })
            .then(async (updated) => {
              log(
                "M365 reauth: login window focus request -> id=%s state=%s focused=%s",
                windowId,
                updated && updated.state,
                updated && updated.focused,
              );
              try {
                const last = await browser.windows.getLastFocused();
                log(
                  "M365 reauth: last focused window id=%s type=%s state=%s focused=%s",
                  last && last.id,
                  last && last.type,
                  last && last.state,
                  last && last.focused,
                );
              } catch (error) {
                log(
                  "M365 reauth: getLastFocused failed: %s",
                  compactError(error),
                );
              }
            })
            .catch((error) =>
              log(
                "M365 reauth: focusing login window %s failed: %s",
                windowId,
                compactError(error),
              ),
            );
        }
        if (tabId == null && windowId != null) {
          const tabs = await browser.tabs.query({ windowId });
          tabId = tabs[0] && tabs[0].id != null ? tabs[0].id : null;
        }
      } catch (error) {
        log(
          "M365 reauth: popup window failed, falling back to a tab:",
          compactError(error),
        );
        windowId = null;
      }
    }
    if (tabId == null) {
      const tab = await browser.tabs.create({
        url: authUrl.href,
        active: true,
      });
      tabId = tab && tab.id != null ? tab.id : null;
      windowId = null;
      if (tabId != null) {
        browser.tabs.update(tabId, { active: true }).catch(() => {});
      }
    }
    // Two-stage watchdog. Stage 1 (12s) catches "the page never became usable":
    // offline, DNS failure, AADSTS50011 error page — all of which used to burn
    // the single attempt for a full 45s and then lock the episode.
    const loadWatchdog = setTimeout(() => {
      const pendingNow = _pendingTopLevelReauth;
      if (!pendingNow || pendingNow.tabId !== tabId) return;
      if (pendingNow.pageSeen) return;
      const offline =
        typeof navigator !== "undefined" && navigator.onLine === false;
      log(
        "M365 reauth: tab %s showed no login UI within %dms (offline=%s) — treating as environmental",
        tabId,
        M365_REAUTH_LOAD_TIMEOUT_MS,
        offline,
      );
      failPendingReauth(
        (offline ? "offline: " : "page-not-loaded: ") +
          M365_REAUTH_LOAD_TIMEOUT_MS +
          "ms",
        { environmental: true },
      );
    }, M365_REAUTH_LOAD_TIMEOUT_MS);
    const timer = setTimeout(() => {
      log(
        "M365 reauth: tab %s produced no code within %dms — failing (no manual takeover)",
        tabId,
        M365_REAUTH_TAB_TIMEOUT_MS,
      );
      failPendingReauth(
        "login tab timed out after " + M365_REAUTH_TAB_TIMEOUT_MS + "ms",
      );
    }, M365_REAUTH_TAB_TIMEOUT_MS);
    _pendingTopLevelReauth = {
      tabId,
      url: authUrl.href,
      state: "ve-reauth:" + stateId,
      verifier,
      redirectUri,
      clientId: effectiveClientId,
      interactive: true,
      hasStoredPassword,
      windowId,
      promptNone: !interactive,
      scope,
      brkClientId,
      brkRedirect: webRedirect,
      timer,
      loadWatchdog,
      pageSeen: false,
      startedAt: now,
    };
    armM365LoginTab(tabId, now);
    await persistPendingReauth(_pendingTopLevelReauth);
    await setM365ReauthState("attempting", reason);
    log(
      "M365 reauth: tab %s created — this is the only attempt for this episode",
      tabId,
    );
    return true;
  } catch (err) {
    log("M365 reauth: failed to open the tab:", compactError(err));
    closePendingTopLevelReauthTab();
    await failM365Reauth("open-tab: " + compactError(err));
    return false;
  } finally {
    _reauthCreating = false;
  }
}

async function handleM365AutoLoginResult(message, sender) {
  await ensureM365AutoLoginLoaded();
  const senderTabId =
    sender && sender.tab && typeof sender.tab.id === "number"
      ? sender.tab.id
      : null;
  const pending = _pendingTopLevelReauth;
  const isOurTab =
    pending && pending.tabId != null && senderTabId === pending.tabId;

  log(
    "M365 auto-login result: ok=%s code=%s mfa=%s fatal=%s missingPassword=%s tab=%s ourTab=%s msg=%s",
    Boolean(message.ok),
    String(message.errorCode || "-"),
    Boolean(message.requiresMfa),
    Boolean(message.fatal),
    Boolean(message.missingPassword),
    senderTabId,
    Boolean(isOurTab),
    String(message.errorMessage || "-").slice(0, 200),
  );

  if (!message.ok) {
    // Never promote staged manual credentials if the login attempt failed.
    _stagedAutoLoginCreds = null;
    const errMsg = String(
      message.errorMessage || "M365 自动登录失败，需人工处理",
    );
    const errCode = String(message.errorCode || "ERROR");

    if (message.requiresMfa) {
      m365AutoLoginPaused = true;
      m365AutoLoginLastError = errMsg;
      m365AutoLoginLastErrorAt = Date.now();
      await saveM365AutoLogin();
      if (isOurTab) {
        // Unattended policy: nobody is going to approve a push or type a code,
        // so the window is closed instead of parked in front of the user.
        closePendingTopLevelReauthTab();
      }
      await setM365ReauthState("terminal", "mfa: " + errMsg.slice(0, 120));
      markM365ReauthRequired(errMsg);
      await notify("Voice Edge：M365 需完成二次验证", errMsg);
      return;
    }

    if (message.fatal) {
      m365AutoLoginPaused = true;
      m365AutoLoginLastError = errMsg;
      m365AutoLoginLastErrorAt = Date.now();
      await saveM365AutoLogin();
      log(
        "M365 auto-login paused due to fatal error (%s): %s",
        errCode,
        errMsg,
      );
      await failM365Reauth("fatal " + errCode + ": " + errMsg.slice(0, 120));
      return;
    }

    if (message.missingPassword) {
      m365AutoLoginLastError = errMsg;
      m365AutoLoginLastErrorAt = Date.now();
      await saveM365AutoLogin();
      log(
        "M365 auto-login: browser password manager did not fill #i0118 (%s) — failing, no manual takeover",
        errCode,
      );
      if (isOurTab)
        await failPendingReauth("no password filled by the password manager");
      else await failM365Reauth("no-password");
      return;
    }

    // Fallback: an unclassified failure used to fall through here and change
    // nothing at all, so the loop restarted with zero diagnostics.
    await failM365Reauth(errCode + ": " + errMsg.slice(0, 120));
    log(
      "M365 auto-login failed without a classified reason (%s): %s",
      errCode,
      errMsg,
    );
    return;
  }

  log("M365 auto-login step ok (%s)", String(message.errorCode || "ok"));
}

async function handleM365OAuthRedirect(message, sender) {
  await ensureM365AutoLoginLoaded();
  const senderTabId =
    sender && sender.tab && typeof sender.tab.id === "number"
      ? sender.tab.id
      : null;
  const pending = _pendingTopLevelReauth;
  const isOurTab =
    pending && pending.tabId != null && senderTabId === pending.tabId;
  if (pending && isOurTab) {
    pending.pageSeen = true;
    if (pending.loadWatchdog) {
      clearTimeout(pending.loadWatchdog);
      pending.loadWatchdog = null;
    }
  }
  const state = String(message.state || "");
  const isOurState = pending && state && state === pending.state;

  if (!message.ok || message.error) {
    const errCode = String(message.error || "interaction_required");
    const errDesc = String(message.errorDescription || "");
    const isOurAttempt = Boolean(isOurTab || isOurState);
    log(
      "M365 reauth: authorize returned error=%s tab=%s isOurTab=%s isOurState=%s desc=%s",
      errCode,
      senderTabId,
      Boolean(isOurTab),
      Boolean(isOurState),
      errDesc.slice(0, 200),
    );
    if (isOurAttempt) {
      const silentPhase = Boolean(pending && pending.promptNone);
      const canGoInteractive = Boolean(
        silentPhase && m365AutoLoginEnabled && !m365AutoLoginPaused,
      );
      if (
        canGoInteractive &&
        /interaction_required|login_required|AADSTS50058|AADSTS50076/i.test(
          errCode + " " + errDesc,
        )
      ) {
        // Silent SSO had no session: this is the same attempt escalating to the
        // page that can render #i0118. One escalation, never a loop.
        log(
          "M365 reauth: silent authorize answered %s — escalating once to the interactive page",
          errCode,
        );
        closePendingTopLevelReauthTab();
        await resetM365ReauthState("escalating after silent");
        await tryTopLevelSilentReauth("interactive-after-silent", {
          force: true,
          interactive: true,
        });
        return;
      }
      // The one automatic attempt is spent. Terminal, no escalation, no retry.
      await failM365Reauth(
        "authorize " + errCode + ": " + (errDesc || errCode).slice(0, 140),
      );
    } else if (!m365ReauthRequired) {
      // A tab we did not open (OWA's own popup/redirect). It only tells us the
      // session is gone; the single attempt is started by the pre-expiry cycle.
      markM365ReauthRequired(`AAD 会话已失效 (${errCode})，需重新登录`);
    } else {
      dlog("M365 authorize error from a page-driven tab (%s)", errCode);
    }
    return;
  }

  const code = String(message.code || "");
  if (!code) {
    log(
      "M365 reauth: authorize returned no code (tab=%s ourTab=%s ourState=%s)",
      senderTabId,
      Boolean(isOurTab),
      Boolean(isOurState),
    );
    return;
  }
  if (!isOurTab && !isOurState) {
    // Previously this silently dropped the code, e.g. after a background
    // restart that lost the in-memory pending record.
    log(
      "M365 reauth: dropping authorization code, tab/state does not match our pending attempt (tab=%s pendingTab=%s state=%s)",
      senderTabId,
      pending ? pending.tabId : "none",
      state.slice(0, 24) || "-",
    );
    return;
  }
  const activePending = pending;
  closePendingTopLevelReauthTab();
  // A pending record persisted by an older build (or a partially restored one)
  // cannot be completed by guessing: the code exchange needs exactly the values
  // the page's /token request carried.
  const missingPending = [
    ["client_id", activePending.clientId],
    ["redirect_uri", activePending.redirectUri],
    ["scope", activePending.scope],
    ["brk_client_id", activePending.brkClientId],
    ["brk_redirect_uri", activePending.brkRedirect],
  ]
    .filter(([, value]) => !String(value || "").trim())
    .map(([name]) => name);
  if (missingPending.length) {
    const detail = "pending reauth is missing " + missingPending.join(", ");
    log("M365 reauth: cannot exchange the code — %s", detail);
    await failM365Reauth(detail);
    return;
  }
  log(
    "M365 reauth: exchanging authorization code (client=%s redirect=%s verifier=%s)",
    shortTid(activePending.clientId),
    activePending.redirectUri,
    activePending.verifier ? "present" : "missing",
  );
  try {
    const endpoint = new URL(
      `https://login.microsoftonline.com/${_tid}/oauth2/v2.0/token`,
    );
    endpoint.searchParams.set("client_id", activePending.clientId);
    endpoint.searchParams.set("client-request-id", crypto.randomUUID());
    const body = new URLSearchParams();
    body.set("client_id", activePending.clientId);
    body.set("grant_type", "authorization_code");
    body.set("code", code);
    body.set("code_verifier", activePending.verifier);
    body.set("redirect_uri", activePending.redirectUri);
    body.set("scope", activePending.scope);
    // AAD requires the broker pair on /token as well. Both halves were captured
    // from the page with the rest of the authorize parameters; a partial pair is
    // rejected with AADSTS90014 "The required field 'brk_redirect_uri' is
    // missing from the credential", so both or neither.
    body.set("brk_client_id", activePending.brkClientId);
    body.set("brk_redirect_uri", activePending.brkRedirect);
    const res = await fetchTokenWithRetry(endpoint.href, {
      method: "POST",
      headers: {
        Accept: "*/*",
        "Content-Type": "application/x-www-form-urlencoded;charset=utf-8",
      },
      body: body.toString(),
    });
    const payload = await res.json();
    if (res.ok && payload && payload.access_token && payload.refresh_token) {
      log(
        "M365 reauth: PKCE code exchange ok (expires_in=%s refresh_token_expires_in=%s)",
        payload.expires_in,
        payload.refresh_token_expires_in,
      );
      // The captured context is left untouched: it is the page's own record, and
      // a flow that has to rewrite it is a flow that guessed. Missing fields are
      // reported by name before /authorize is ever built.
      await applyM365TokenResponse({
        tenantId: _tid,
        clientId: activePending.clientId,
        grantType: "authorization_code",
        source: "pkce",
        adoptIdentity: true,
        accessToken: String(payload.access_token),
        refreshToken: String(payload.refresh_token),
        expiresIn: Number(payload.expires_in || 0),
        refreshTokenExpiresIn: Number(payload.refresh_token_expires_in || 0),
        // Our own exchange started right now, so it is the newest generation.
        requestStartedAt: Date.now(),
        capturedAt: Date.now(),
      });
      _lastAuthCodeSignInAt = Date.now();
      log("renewed 24h M365 SPA refresh token via top-level PKCE authorize");
      // The open OWA tabs are NOT reloaded. A renewal happens hours before the
      // old RT dies, and the page keeps working with it; reloading threw away
      // drafts/scroll/chat state for nothing. Popup suppression is handled by
      // pingM365TabsKeepAlive on the page side.
      pingM365TabsKeepAlive({ authHealthy: true }).catch(() => {});
    } else {
      const detail = String(
        (payload && (payload.error_description || payload.error)) || res.status,
      );
      log(
        "M365 reauth: PKCE code exchange rejected status=%s error=%s",
        res.status,
        detail.slice(0, 240),
      );
      await failM365Reauth(
        "code-exchange " + res.status + ": " + detail.slice(0, 140),
      );
    }
  } catch (err) {
    log("M365 reauth: PKCE code exchange error:", compactError(err));
    await failM365Reauth("code-exchange: " + compactError(err));
  }
}

async function pingM365TabsKeepAlive(options = {}) {
  const reauth = Boolean(options && options.reauth);
  const authHealthy =
    options && typeof options.authHealthy === "boolean"
      ? options.authHealthy
      : !m365ReauthRequired;
  // While the extension is driving a reauth itself, OWA's own MSAL fallback
  // popups are pure noise (they led to dozens of stacked login windows during
  // a past incident), so ask the page hook to suppress them.
  const suppressPopups =
    options && typeof options.suppressPopups === "boolean"
      ? options.suppressPopups
      : true;
  try {
    const tabs = await browser.tabs.query({
      url: M365_TAB_URL_PATTERNS,
    });
    for (const tab of tabs) {
      if (!tab || tab.id == null || tab.discarded) continue;
      // The page's own MSAL keeps its private copy of the RT; handing it the
      // freshest one lets the page-world fetch hook swap the refresh_token on
      // its /token calls (see content-m365.js). Shared across top frame and
      // subframes so subframe /token requests also use and report the bound RT.
      const keepAlivePayload = {
        type: "M365_KEEPALIVE",
        reauth,
        authHealthy,
        suppressPopups,
        rt: _rt || "",
        rtClientId: _clientId || "",
        rtExp: Number(_rtExp || 0),
      };
      // Notify top frame (frameId 0) where OWA's AbtInteraction & MSAL live.
      browser.tabs
        .sendMessage(
          tab.id,
          {
            __veM365ToPage: true,
            payload: keepAlivePayload,
          },
          { frameId: 0 },
        )
        .catch(() => {});
      // Also notify any registered M365 subframes for this tab.
      for (const frame of m365FrameCandidates.values()) {
        if (frame.tabId !== tab.id || frame.frameId === 0) continue;
        browser.tabs
          .sendMessage(
            tab.id,
            {
              __veM365ToPage: true,
              payload: keepAlivePayload,
            },
            { frameId: frame.frameId },
          )
          .catch(() => {});
      }
    }
  } catch (_) {}
}

async function pingSharePointTabsKeepAlive() {
  let anyFailed = false;
  try {
    const tabs = await browser.tabs.query({
      url: ["https://*.sharepoint.com/*"],
    });
    for (const tab of tabs) {
      if (!tab || tab.id == null || tab.discarded) continue;
      const res = await browser.tabs
        .sendMessage(tab.id, {
          __veSharePoint: true,
          type: "SP_KEEPALIVE",
          siteUrl: sharePointHomeUrl || tab.url || "",
        })
        .catch(() => null);
      if (
        res &&
        res.ok === false &&
        res.error &&
        res.error !== "origin mismatch" &&
        res.error !== "invalid siteUrl"
      ) {
        log(
          "SharePoint tab keepalive failed tabId=%s err=%s",
          tab.id,
          String(res.error),
        );
        anyFailed = true;
      }
    }
  } catch (_) {}
  if (anyFailed) schedulePreExpiryRetry("sp-keepalive");
}

let _preExpiryRenewalInFlight = null;
let _preExpiryRetryTimer = null;
let _preExpiryRetryAttempt = 0;

function schedulePreExpiryRetry(reason) {
  if (_preExpiryRetryTimer) return;
  const baseDelay = Math.min(
    5 * 60 * 1000,
    30000 * 2 ** Math.min(_preExpiryRetryAttempt, 4),
  );
  const delay = Math.round(baseDelay * (0.85 + Math.random() * 0.3));
  _preExpiryRetryAttempt += 1;
  log(
    "scheduling pre-expiry retry (#%d) in %dms reason=%s",
    _preExpiryRetryAttempt,
    delay,
    String(reason || "unknown"),
  );
  _preExpiryRetryTimer = setTimeout(() => {
    _preExpiryRetryTimer = null;
    runPreExpirationRenewal().catch((error) =>
      log("pre-expiry retry failed:", compactError(error)),
    );
  }, delay);
}

function runPreExpirationRenewal() {
  if (_preExpiryRenewalInFlight) return _preExpiryRenewalInFlight;
  _preExpiryRenewalInFlight = (async () => {
    if (typeof navigator !== "undefined" && navigator.onLine === false) {
      schedulePreExpiryRetry("offline");
      return;
    }
    try {
      await Promise.all([
        ensureM365AuthLoaded(),
        ensureM365SettingsLoaded(),
        ensureM365AutoLoginLoaded(),
        ensureM365ReauthStateLoaded(),
      ]);
    } catch (_) {}
    const now = Date.now();
    const rtRemaining = _rtExp > 0 ? _rtExp - now : 0;
    const rtPastDue = _rtExp > 0 && rtRemaining <= 0;
    // The local rtExp countdown is a HINT only. Declaring "24h cap reached" from
    // our own bookkeeping turned into a self-locking state machine: the flag
    // stopped every real refresh, so nothing could ever prove the token still
    // worked. A real AAD answer is now always obtained first.
    const reauthTerminal = m365ReauthState === "terminal";
    if (rtPastDue && !reauthTerminal) {
      log(
        "M365: local 24h RT countdown is past due (rtRemain=%smin) — proving it with a real refresh before declaring reauth",
        Math.round(rtRemaining / 60000),
      );
      if (
        !m365ReauthRequired &&
        _sydney.exp - now > M365_PRE_EXPIRY_WINDOW_MS
      ) {
        // Force the refresh below to actually hit AAD instead of returning the
        // still-valid cached access token.
        _sydney = { token: "", exp: 0 };
      }
    }
    let hadFailure = false;
    // 1) Prove the refresh token with a real request. Deliberately NOT gated on
    // !m365ReauthRequired: if the token is actually alive this succeeds and
    // clears the flag by itself; if it is really dead AAD answers AADSTS700084
    // and the reauth gate stops the login-tab loop.
    if (m365AuthLoaded && _clientId && _rt && _tid && !reauthTerminal) {
      if (!_sydney.token || _sydney.exp - now <= M365_PRE_EXPIRY_WINDOW_MS) {
        await refreshSydney(M365_PRE_EXPIRY_WINDOW_MS).catch((error) => {
          if (!m365ReauthRequired) hadFailure = true;
          log("pre-expiry Sydney refresh failed:", compactError(error));
        });
      }
      if (
        !m365ReauthRequired &&
        sharePointHomeUrl &&
        (!_sharePoint.token ||
          _sharePoint.exp - now <= M365_PRE_EXPIRY_WINDOW_MS)
      ) {
        await refreshSharePointAccessToken(
          sharePointHomeUrl,
          M365_PRE_EXPIRY_WINDOW_MS,
        ).catch((error) => {
          if (!m365ReauthRequired) hadFailure = true;
          log(
            "pre-expiry SharePoint token refresh failed:",
            compactError(error),
          );
        });
      }
      if (
        !m365ReauthRequired &&
        _ic3.token &&
        _ic3.exp - now <= M365_PRE_EXPIRY_WINDOW_MS
      ) {
        await refreshIc3Token(M365_PRE_EXPIRY_WINDOW_MS).catch((error) => {
          if (!m365ReauthRequired) hadFailure = true;
          log("pre-expiry IC3 token refresh failed:", compactError(error));
        });
      }
    }
    // 2) Only now decide whether a top-level (re)authorization is needed. Running
    // it before the refresh was what turned a guess into an endless login-tab
    // loop: the tab was opened even when a plain refresh would have worked.
    // 2) Single-shot sign-in: start the one and only attempt of this episode.
    const rtRemainingAfter = _rtExp > 0 ? _rtExp - Date.now() : 0;
    const needsSignIn =
      m365ReauthRequired || rtRemainingAfter <= M365_RT_RENEW_WINDOW_MS;
    // A failed episode must not disable auto-login forever: after the hold
    // window the next attempt is allowed again (the reason is kept for logs).
    if (
      m365ReauthState === "terminal" &&
      m365ReauthStateAt &&
      now - m365ReauthStateAt > M365_REAUTH_HOLD_MS
    ) {
      log(
        "M365 reauth: terminal state expired after %dh — a new attempt is allowed",
        Math.round(M365_REAUTH_HOLD_MS / 3600000),
      );
      await resetM365ReauthState("terminal expired");
    }
    if (m365AuthLoaded && _clientId && _tid && needsSignIn) {
      if (m365ReauthState === "none") {
        await tryTopLevelSilentReauth(
          m365ReauthRequired ? "reauth-required" : "rt-expiring-6h",
        ).catch((error) =>
          log("M365 reauth attempt failed:", compactError(error)),
        );
      } else if (now - _reauthStateLastLogAt > 60 * 60 * 1000) {
        _reauthStateLastLogAt = now;
        log(
          "M365 reauth: state=%s — not starting another attempt (single-shot policy) — %s",
          m365ReauthState,
          m365StateLine(),
        );
      }
    }
    if (
      m365ReauthRequired ||
      rtRemainingAfter <= M365_RT_RENEW_WINDOW_MS ||
      !_sydney.token ||
      _sydney.exp - Date.now() <= M365_PRE_EXPIRY_WINDOW_MS
    ) {
      log("M365 pre-expiry cycle: %s", m365StateLine());
    }
    if (now - m365LastKeepAliveAt >= M365_KEEPALIVE_INTERVAL_MS) {
      await pingM365TabsKeepAlive();
    }
    if (now - sharePointLastKeepAliveAt >= SP_KEEPALIVE_INTERVAL_MS) {
      await pingSharePointTabsKeepAlive();
    }
    if (hadFailure && !m365ReauthRequired) {
      schedulePreExpiryRetry("token-refresh-error");
    } else {
      _preExpiryRetryAttempt = 0;
    }
  })();
  return _preExpiryRenewalInFlight.finally(() => {
    _preExpiryRenewalInFlight = null;
  });
}

setInterval(
  () => {
    runPreExpirationRenewal().catch((error) =>
      log("pre-expiry renewal loop failed:", compactError(error)),
    );
  },
  5 * 60 * 1000,
);
setTimeout(() => {
  runPreExpirationRenewal().catch(() => {});
}, 20 * 1000);
// Startup housekeeping: restore a reauth attempt that was in flight before a
// background restart (so its authorization code can still be exchanged), log the
// persisted gate budget, and close login tabs nobody is tracking any more.
setTimeout(() => {
  (async () => {
    try {
      await Promise.all([
        ensureM365AuthLoaded(),
        ensureM365ReauthStateLoaded(),
      ]);
      log("M365 startup: %s", m365StateLine());
      await restorePendingReauth();
      await sweepOrphanReauthTabs("startup");
    } catch (error) {
      log("M365 startup housekeeping failed:", compactError(error));
    }
  })();
}, 3000);
try {
  self.addEventListener("online", () => {
    _preExpiryRetryAttempt = 0;
    if (_preExpiryRetryTimer) {
      clearTimeout(_preExpiryRetryTimer);
      _preExpiryRetryTimer = null;
    }
    runPreExpirationRenewal().catch(() => {});
  });
} catch (_) {}

function scheduleM365Reconnect() {
  if (m365ReconnectTimer) return;
  const exponential = Math.min(
    M365_RECONNECT_MAX_MS,
    M365_RECONNECT_BASE_MS * 2 ** Math.min(m365ReconnectAttempt, 5),
  );
  const delay = Math.round(exponential * (0.8 + Math.random() * 0.4));
  m365ReconnectAttempt += 1;
  m365ReconnectTimer = setTimeout(() => {
    m365ReconnectTimer = null;
    m365Connect();
  }, delay);
}

function sendM365(message) {
  if (m365Ws && m365Ws.readyState === WebSocket.OPEN) {
    try {
      m365Ws.send(JSON.stringify(message));
    } catch (_) {}
  }
}

function m365Connect() {
  if (
    m365Ws &&
    (m365Ws.readyState === WebSocket.CONNECTING ||
      m365Ws.readyState === WebSocket.OPEN)
  )
    return m365Ws;
  try {
    m365Ws = new WebSocket(M365_BRIDGE_URL);
  } catch (_) {
    scheduleM365Reconnect();
    return null;
  }
  m365Ws.addEventListener("open", () => {
    m365ReconnectAttempt = 0;
    sendM365({
      type: "M365_READY",
      build: M365_BUILD_TAG,
      // `reload` lets the harness pick up a new dist without asking the user to
      // click Reload in about:debugging; it is a local-only bridge command.
      caps: ["reauth", "probe", "reload"],
    });
  });
  m365Ws.addEventListener("message", (event) => {
    let message;
    try {
      message = JSON.parse(event.data);
    } catch (_) {
      return;
    }
    if (message.type === "M365_CONFIG") {
      if (typeof message.debug !== "undefined") {
        m365Debug = !!message.debug;
        log("debug tracing", m365Debug ? "ON (from M365_CONFIG)" : "OFF");
      }
      setM365EntryUrl(message.entryUrl).catch((error) =>
        log("invalid M365 config:", compactError(error)),
      );
      sharePointHomeUrl = String(message.sharePointHomeUrl || "").replace(
        /\/$/,
        "",
      );
      sharePointUploadFolder = String(
        message.sharePointUploadFolder || "",
      ).replace(/^\/+|\/+$/g, "");
      // Second folder for model-produced artifacts. Fall back to the upload
      // folder when voice_edge has not yet been updated to send the new field,
      // so the feature degrades to "same folder" rather than breaking.
      sharePointDownloadFolder = String(
        message.sharePointDownloadFolder ||
          message.sharePointUploadFolder ||
          "",
      ).replace(/^\/+|\/+$/g, "");
      return;
    }
    if (message.type === "M365_ASK") {
      dlog(
        "M365_ASK recv id=%s conversationId=%s tone=%s attachments=%d textLen=%d",
        String(message.id || ""),
        String(message.conversationId || "") || "<new>",
        String(message.tone || ""),
        Array.isArray(message.attachments) ? message.attachments.length : 0,
        String(message.text || "").length,
      );
      if (message.entryUrl) setM365EntryUrl(message.entryUrl).catch(() => {});
      dispatchM365(message);
    }
    if (message.type === "M365_RELOAD_SELF") {
      log("M365 bridge: reload requested by driver — reloading extension");
      sendM365({ type: "M365_RELOAD_ACK", build: M365_BUILD_TAG });
      setTimeout(() => {
        try {
          browser.runtime.reload();
        } catch (error) {
          log("M365 bridge: reload failed:", compactError(error));
        }
      }, 300);
      return { ok: true };
    }
    if (message.type === "M365_PAGE_TOKEN_PROBE") {
      runPageTokenProbe()
        .then((tabs) => sendM365({ type: "M365_PAGE_TOKEN_PROBE_SENT", tabs }))
        .catch((error) => log("M365 probe failed:", compactError(error)));
      return { ok: true };
    }
    if (message.type === "M365_REAUTH_NOW") {
      handleM365BridgeReauthCommand(message).catch((error) =>
        log("M365 bridge: reauth command failed:", compactError(error)),
      );
      return;
    }
    if (message.type === "M365_STOP") {
      dlog(
        "M365_STOP recv id=%s reason=%s",
        String(message.id || ""),
        String(message.reason || ""),
      );
      routeM365Stop(message.id);
    }
    if (message.type === "AIS_CONFIG") {
      if (message.entryUrl) {
        aisEntryUrl = String(message.entryUrl).trim() || aisEntryUrl;
      }
      aisLastForwardedReadySig = "";
      if (aisLastReady) {
        const clean = Object.assign({}, aisLastReady);
        delete clean.__veAis;
        delete clean.dir;
        sendM365(clean);
      }
      pingAisTabs().catch(() => {});
      return;
    }
    if (
      message.type === "AIS_WARMUP" ||
      message.type === "AIS_CALL" ||
      message.type === "AIS_STREAM" ||
      message.type === "AIS_UPLOAD"
    ) {
      dispatchAisToTab(message).catch((error) => {
        sendM365({
          type: "AIS_ERROR",
          id: String(message.id || ""),
          code: "dispatch_failed",
          message: compactError(error),
          httpStatus: 503,
        });
      });
      return;
    }
    if (message.type === "AIS_ABORT") {
      routeAisAbort(message.id, message.reason);
      return;
    }
  });
  m365Ws.addEventListener("close", () => {
    m365Ws = null;
    aisLastForwardedReadySig = "";
    scheduleM365Reconnect();
  });
  m365Ws.addEventListener("error", () => {
    try {
      m365Ws.close();
    } catch (_) {}
  });
  return m365Ws;
}

const AIS_TAB_URL_PATTERNS = ["https://aistudio.google.com/*"];
let aisEntryUrl =
  "https://aistudio.google.com/prompts/new_chat?model=gemini-4-argon";
let aisTargetTabId = null;
let aisLastReady = null;
let aisLastForwardedReadySig = "";
let aisLastError = null;
let aisLastPopupPingAt = 0;
const aisTabByRequestId = new Map();

async function runAisPopupProbe(warmup) {
  try {
    const tab = await ensureAisTab();
    if (!tab || tab.id == null) {
      return { ok: false, error: "无法打开或定位 AI Studio 标签页" };
    }
    const cookies = await readAisAuthCookies();
    const probeId = "probe-" + Date.now();
    const beforeTs = (aisLastReady && aisLastReady.updatedAt) || 0;
    await browser.tabs.sendMessage(
      tab.id,
      {
        __veAisToPage: true,
        payload: {
          type: "AIS_PROBE",
          id: probeId,
          warmup: Boolean(warmup),
          cookies,
        },
      },
      { frameId: 0 },
    );
    const deadline = Date.now() + (warmup ? 18000 : 4000);
    while (Date.now() < deadline) {
      await new Promise((r) => setTimeout(r, 150));
      if (
        aisLastReady &&
        (aisLastReady.id === probeId || aisLastReady.updatedAt > beforeTs)
      ) {
        break;
      }
    }
    requestBrowserState(true);
    return {
      ok: true,
      state: popupBuildState(),
    };
  } catch (error) {
    return {
      ok: false,
      error: compactError(error),
    };
  }
}

async function readAisAuthCookies() {
  const out = {};
  try {
    const cookies = await browser.cookies.getAll({
      url: "https://aistudio.google.com/",
    });
    for (const c of cookies || []) {
      if (
        c &&
        (c.name === "SAPISID" ||
          c.name === "__Secure-1PAPISID" ||
          c.name === "__Secure-3PAPISID") &&
        c.value
      ) {
        out[c.name] = c.value;
      }
    }
  } catch (_) {}
  return out;
}

function pickBestAisTab(tabs) {
  if (!Array.isArray(tabs) || !tabs.length) return null;
  const valid = tabs.filter((t) => t && t.id != null);
  if (!valid.length) return null;
  const isPromptsTab = (t) => String((t && t.url) || "").includes("/prompts/");
  const activePrompts = valid.find((t) => t.active && isPromptsTab(t));
  if (activePrompts) return activePrompts;
  if (aisTargetTabId != null) {
    const matched = valid.find((t) => t.id === aisTargetTabId);
    if (matched && (isPromptsTab(matched) || !valid.some(isPromptsTab))) {
      return matched;
    }
  }
  const anyPrompts = valid.find(isPromptsTab);
  if (anyPrompts) return anyPrompts;
  const active = valid.find((t) => t.active);
  if (active) return active;
  return valid[0] || null;
}

async function pingAisTabs() {
  try {
    const tabs = await browser.tabs.query({ url: AIS_TAB_URL_PATTERNS });
    const tab = pickBestAisTab(tabs);
    if (!tab || tab.id == null) return;
    const cookies = await readAisAuthCookies();
    aisTargetTabId = tab.id;
    browser.tabs
      .sendMessage(
        tab.id,
        {
          __veAisToPage: true,
          payload: { type: "AIS_PING", cookies },
        },
        { frameId: 0 },
      )
      .catch(() => {});
  } catch (_) {}
}

async function ensureAisTab() {
  const tabs = await browser.tabs.query({ url: AIS_TAB_URL_PATTERNS });
  const chosen = pickBestAisTab(tabs);
  if (chosen && chosen.id != null) {
    aisTargetTabId = chosen.id;
    return chosen;
  }
  const created = await browser.tabs.create({
    url: aisEntryUrl,
    active: false,
  });
  aisTargetTabId = created.id;
  const cookies = await readAisAuthCookies();
  // Wait briefly for content script and page hook to load in newly created tab
  const deadline = Date.now() + 15000;
  while (Date.now() < deadline) {
    await new Promise((r) => setTimeout(r, 250));
    try {
      const ack = await browser.tabs.sendMessage(
        created.id,
        { __veAisToPage: true, payload: { type: "AIS_PING", cookies } },
        { frameId: 0 },
      );
      if (ack && ack.ok && ack.booted) break;
    } catch (_) {}
  }
  return created;
}

async function dispatchAisToTab(message) {
  const tab = await ensureAisTab();
  if (!tab || tab.id == null) {
    throw new Error("Unable to locate or open AI Studio tab");
  }
  const cookies = await readAisAuthCookies();
  const payload = Object.assign({}, message, { cookies });
  const reqId = String(message.id || "");
  if (reqId) {
    aisTabByRequestId.set(reqId, tab.id);
  }
  const ack = await browser.tabs.sendMessage(
    tab.id,
    { __veAisToPage: true, payload },
    { frameId: 0 },
  );
  if (!ack || !ack.ok) {
    if (reqId) aisTabByRequestId.delete(reqId);
    throw new Error("AI Studio content script did not acknowledge request");
  }
}

function routeAisAbort(id, reason) {
  const reqId = String(id || "");
  if (!reqId) return;
  const tabId = aisTabByRequestId.get(reqId) || aisTargetTabId;
  aisTabByRequestId.delete(reqId);
  if (tabId == null) return;
  browser.tabs
    .sendMessage(
      tabId,
      {
        __veAisToPage: true,
        payload: {
          type: "AIS_ABORT",
          id: reqId,
          reason: String(reason || "client-abort"),
        },
      },
      { frameId: 0 },
    )
    .catch(() => {});
}

async function ensureM365Tab() {
  const tabs = await browser.tabs.query({
    url: M365_TAB_URL_PATTERNS,
  });
  if (tabs.length) return tabs[0];
  await ensureM365SettingsLoaded();
  if (!m365EntryUrl) throw new Error("M365_ENTRY_URL is not configured");
  return browser.tabs.create({ url: m365EntryUrl, active: false });
}

// Stable per-document identity for frame dedup: origin + pathname, with the
// query string and fragment removed. The M365 host rewrites volatile query
// params (notably sessionId) on every reload, so the raw frameUrl differs each
// refresh; the path is stable across refreshes of the same document. Falls back
// to a query/fragment-trimmed string if URL parsing fails, and finally to the
// raw value so a non-URL never collapses unrelated frames together.
function normalizeM365FrameKey(rawUrl) {
  const value = String(rawUrl || "");
  if (!value) return "";
  try {
    const u = new URL(value);
    return u.origin + u.pathname;
  } catch (_) {
    return value.split(/[?#]/, 1)[0] || value;
  }
}

function rememberM365Frame(message, sender, hasChats) {
  if (!sender.tab || sender.tab.id == null || sender.frameId == null)
    return null;
  const frameUrl = String(message.frameUrl || sender.url || "");
  const frame = {
    tabId: sender.tab.id,
    frameId: sender.frameId,
    frameOrigin: String(message.frameOrigin || ""),
    frameUrl,
    // Identity used for dedup: origin + pathname with the query string dropped.
    // A refresh keeps the same document/path but the host rewrites volatile
    // query params (observed: sessionId changes on every reload, e.g.
    // .../semanticoverview/Users(...)?...&sessionId=<new-guid>&...). Comparing
    // the full frameUrl therefore treated each refresh as a brand-new frame and
    // the stale entries accumulated one-per-reload (the duplicate
    // "M365 page hook ready" you saw). Normalizing to path removes the volatile
    // query so refreshes collapse onto one identity.
    frameKey: normalizeM365FrameKey(frameUrl),
    hasChats: Boolean(hasChats),
    seenAt: Date.now(),
  };
  // Evict stale duplicates from the SAME document before registering the new
  // frame. A page refresh reuses the tab (tabId unchanged) but assigns the
  // document a brand-new frameId, so the previous frameId entry is a dead
  // zombie the onRemoved handler never cleans (it only fires on tab CLOSE, not
  // reload). Left in place, these accumulate one-per-refresh and can be
  // preferred by the hasChats-first ordering, forcing every dispatch to
  // try-then-prune a dead frame (added latency) or, worse, letting a mid-reload
  // zombie ack the relay without ever running doAsk (idle timeout). Keyed by
  // tabId + normalized frameKey so a query-only change (sessionId) still counts
  // as the same document and only the newest frameId survives.
  if (frame.frameKey) {
    for (const [key, existing] of m365FrameCandidates) {
      if (
        existing.tabId === frame.tabId &&
        (existing.frameKey || normalizeM365FrameKey(existing.frameUrl)) ===
          frame.frameKey &&
        existing.frameId !== frame.frameId
      ) {
        m365FrameCandidates.delete(key);
        if (
          m365TargetFrame &&
          m365TargetFrame.tabId === existing.tabId &&
          m365TargetFrame.frameId === existing.frameId
        ) {
          m365TargetFrame = null;
        }
      }
    }
  }
  m365FrameCandidates.set(frame.tabId + ":" + frame.frameId, frame);
  if (
    frame.hasChats ||
    !m365TargetFrame ||
    m365TargetFrame.tabId !== frame.tabId ||
    (!m365TargetFrame.hasChats && frame.seenAt >= m365TargetFrame.seenAt)
  ) {
    m365TargetFrame = frame;
  }
  return frame;
}

function bestM365FrameForTab(tabId) {
  if (m365TargetFrame && m365TargetFrame.tabId === tabId)
    return m365TargetFrame;
  const candidates = Array.from(m365FrameCandidates.values())
    .filter((frame) => frame.tabId === tabId)
    .sort(
      (a, b) => Number(b.hasChats) - Number(a.hasChats) || b.seenAt - a.seenAt,
    );
  return candidates[0] || null;
}

// Ordered list of every candidate frame for a tab, best-first, with the
// sticky target promoted to the front. When a page hosts more than one frame
// that injected the page hook (observed as duplicate "M365 page hook ready"),
// exactly one of them owns a live self-built Chathub socket; the others are
// stale/duplicate entries whose frameId no longer routes. dispatchM365 walks
// this list and confirms delivery by the relay ack, so a stale frameId is
// skipped instead of silently swallowing the ASK (the "0 inbound events →
// first-event timeout" failure).
function orderedM365FramesForTab(tabId) {
  const seen = new Set();
  const ordered = [];
  const push = (frame) => {
    if (!frame || frame.tabId !== tabId) return;
    const key = frame.tabId + ":" + frame.frameId;
    if (seen.has(key)) return;
    seen.add(key);
    ordered.push(frame);
  };
  push(m365TargetFrame);
  Array.from(m365FrameCandidates.values())
    .filter((frame) => frame.tabId === tabId)
    .sort(
      (a, b) => Number(b.hasChats) - Number(a.hasChats) || b.seenAt - a.seenAt,
    )
    .forEach(push);
  return ordered;
}

// Drop a candidate frame whose frameId no longer accepts messages, so a stale
// duplicate cannot be re-selected on the next turn. Also clears the sticky
// target if it pointed at the pruned frame.
function pruneStaleM365Frame(frame) {
  if (!frame) return;
  m365FrameCandidates.delete(frame.tabId + ":" + frame.frameId);
  if (
    m365TargetFrame &&
    m365TargetFrame.tabId === frame.tabId &&
    m365TargetFrame.frameId === frame.frameId
  ) {
    m365TargetFrame = null;
  }
}

async function waitForM365Hook(tabId, timeoutMs = 15000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    const target = bestM365FrameForTab(tabId);
    if (target) return target;
    await new Promise((resolve) => setTimeout(resolve, 200));
  }
  throw new Error(
    "M365 page hook did not become ready; reload the configured M365 entry page",
  );
}

async function prepareM365Attachments(message) {
  const raw = Array.isArray(message.attachments) ? message.attachments : [];
  if (!raw.length) return [];
  // Split by kind BEFORE touching SharePoint. Images must be ingested inline
  // (content-m365.js -> owahub UploadFile -> docId -> ImageFile annotation) so
  // the vision model receives the actual pixels. Routing them through
  // SharePoint produces a FileUrl document link instead, which is exactly the
  // "pasted image arrives as a file/text link" bug. Only genuine file
  // attachments use the SharePoint upload / FileUrl path.
  const imageItems = [];
  const fileItems = [];
  for (const item of raw) {
    if (item && item.kind === "image") imageItems.push(item);
    else fileItems.push(item);
  }
  // Pass images through untouched, normalized to the field names
  // content-m365.js expects (kind + dataBase64). Python sends the base64 in
  // `data`; the extension's uploadImageToM365 reads `dataBase64`.
  const passthroughImages = imageItems.map((img) => ({
    kind: "image",
    name: String(img.name || ""),
    mimeType: String(img.mimeType || "application/octet-stream"),
    dataBase64: String(img.dataBase64 || img.data || ""),
  }));
  // No document attachments → skip SharePoint entirely (and its config check).
  // A turn with only pasted images must not require SHAREPOINT_HOME_URL.
  if (!fileItems.length) return passthroughImages;
  if (!sharePointHomeUrl)
    throw new Error("SHAREPOINT_HOME_URL 未配置，无法上传 M365 附件");
  let sharePointAccessToken = "";
  try {
    sharePointAccessToken =
      await refreshSharePointAccessToken(sharePointHomeUrl);
  } catch (tokenError) {
    dlog(
      "FOCI SharePoint token unavailable, falling back to SP page token/cookies:",
      compactError(tokenError),
    );
  }
  const tab = await ensureSharePointTab();
  const uploadRequestId =
    String(message.id || "") + ":" + Date.now() + ":" + crypto.randomUUID();
  const completion = new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      sharePointUploadWaiters.delete(uploadRequestId);
      reject(new Error("SharePoint upload completion event timed out"));
    }, sharePointUploadTimeoutMs(fileItems.length));
    sharePointUploadWaiters.set(uploadRequestId, {
      resolve: (value) => {
        clearTimeout(timer);
        resolve(value);
      },
    });
  });
  // Do not await the tabs.sendMessage response here. The SharePoint listener
  // used to await runtime.sendMessage(SP_UPLOAD_COMPLETE) while this function
  // awaited the listener response, creating a cross-context response cycle.
  // Dispatch once and use only the correlated completion event as the result.
  browser.tabs
    .sendMessage(tab.id, {
      __veSharePoint: true,
      type: "SP_UPLOAD_FILES",
      requestId: uploadRequestId,
      siteUrl: sharePointHomeUrl,
      uploadFolder: sharePointUploadFolder,
      files: fileItems,
      sharePointAccessToken,
    })
    .then((ack) => {})
    .catch((error) => {
      const waiter = sharePointUploadWaiters.get(uploadRequestId);
      if (waiter) {
        sharePointUploadWaiters.delete(uploadRequestId);
        waiter.resolve({ ok: false, error: compactError(error) });
      }
    });
  // Keep the Python relay alive for the whole (now up to 10min) upload wait.
  // dispatchM365 only pings once before and once after prepareM365Attachments,
  // so without this a slow multi-file upload produced NO frames for minutes and
  // Python could idle-tear-down the request. Empty-text M365_PROGRESS is the
  // established pure-liveness contract (flips got_any / resets the idle timer,
  // never written to the answer).
  const uploadKeepAlive = setInterval(() => {
    sendM365({ type: "M365_PROGRESS", id: String(message.id || ""), text: "" });
  }, M365_ARTIFACT_HOLD_KEEPALIVE_MS);
  let response;
  try {
    response = await completion;
  } finally {
    clearInterval(uploadKeepAlive);
  }
  sharePointUploadWaiters.delete(uploadRequestId);
  if (!response || !response.ok) {
    throw new Error(
      String(
        (response && response.error) || "SharePoint attachment upload failed",
      ),
    );
  }
  const uploaded = Array.isArray(response.files) ? response.files : [];
  // sp.js now returns partial batches: at least one file uploaded (ok:true)
  // while others may have hard-failed post-retry. Do not silently drop those —
  // surface them so a subtly incomplete turn (model never sees an attachment
  // the user added) is visible rather than mysterious.
  const failed = Array.isArray(response.failed) ? response.failed : [];
  if (failed.length) {
    log(
      "SharePoint attachment upload partial: %d ok, %d failed [%s]",
      uploaded.length,
      failed.length,
      failed
        .map(
          (f) =>
            String((f && f.name) || "?") + ": " + String((f && f.error) || ""),
        )
        .join("; "),
    );
  }
  // Uploaded document attachments (FileUrl path) + inline images (ImageFile
  // path). Both kinds flow to content-m365.js: files satisfy the FileUrl
  // filter (url/itemId/driveId/verified), images carry kind+dataBase64 so
  // doAsk uploads them for a docId and buildChatArgs emits an ImageFile.
  return [...uploaded, ...passthroughImages];
}

// Bounded per-file scaling for the SharePoint completion wait. Each file now
// costs an addUsingPath + read-back GET + driveItem GET, and every one of
// those can burn several exponential-backoff retries against a throttling
// (503) tenant. A flat 60s aborted large or throttled batches mid-flight and
// surfaced as a misleading "completion event timed out" while uploads were
// still succeeding. Floor 120s (common 1-file case), +60s per file, capped at
// 10min so a stuck batch still fails in bounded time.
//
// 上传多个大文件时旧的 5min(300s) 上限仍会触发
// "SharePoint upload completion event timed out"，按需再放宽一倍：floor 60->120s、
// 每文件 30->60s、cap 300->600s。600s 仍安全低于 Python 端 M365_IDLE_MAX_SECONDS(900s)，
// 且上传等待期间下面新增了周期性 M365_PROGRESS 保活，Python relay 不会在此期间被 idle 拆掉。
function sharePointUploadTimeoutMs(fileCount) {
  const count = Math.max(1, Number(fileCount) || 1);
  return Math.min(600000, Math.max(120000, count * 60000));
}

// Reverse direction of prepareM365Attachments: upload files the MODEL produced
// (already fetched to base64 by content-m365's artifact harvester) into the
// SharePoint DOWNLOAD folder. Reuses the exact same, proven SP_UPLOAD_FILES +
// sharePointUploadWaiters + SP_UPLOAD_COMPLETE machinery as the forward path;
// the only differences are the folder (sharePointDownloadFolder) and that the
// bytes come from the harvester rather than an outbound attachment. sp.js is
// direction-agnostic (it takes uploadFolder as a parameter), so no change is
// needed there. Returns the uploaded file descriptors (name/url/itemId/…).
async function uploadArtifactsToSharePoint(files) {
  const raw = Array.isArray(files) ? files : [];
  if (!raw.length) return [];
  if (!sharePointHomeUrl)
    throw new Error("SHAREPOINT_HOME_URL 未配置，无法上传模型产物");
  let sharePointAccessToken = "";
  try {
    sharePointAccessToken =
      await refreshSharePointAccessToken(sharePointHomeUrl);
  } catch (tokenError) {
    dlog(
      "FOCI SharePoint token unavailable for artifact, falling back to SP page token/cookies:",
      compactError(tokenError),
    );
  }
  const tab = await ensureSharePointTab();
  const uploadRequestId = "artifact:" + Date.now() + ":" + crypto.randomUUID();
  const completion = new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      sharePointUploadWaiters.delete(uploadRequestId);
      reject(new Error("SharePoint artifact upload completion timed out"));
    }, sharePointUploadTimeoutMs(raw.length));
    sharePointUploadWaiters.set(uploadRequestId, {
      resolve: (value) => {
        clearTimeout(timer);
        resolve(value);
      },
    });
  });
  browser.tabs
    .sendMessage(tab.id, {
      __veSharePoint: true,
      type: "SP_UPLOAD_FILES",
      requestId: uploadRequestId,
      siteUrl: sharePointHomeUrl,
      // The one meaningful difference from the forward path: the model's
      // artifacts land in the download folder, not the upload folder.
      uploadFolder: sharePointDownloadFolder || sharePointUploadFolder,
      files: raw,
      sharePointAccessToken,
    })
    .then((ack) => {})
    .catch((error) => {
      const waiter = sharePointUploadWaiters.get(uploadRequestId);
      if (waiter) {
        sharePointUploadWaiters.delete(uploadRequestId);
        waiter.resolve({ ok: false, error: compactError(error) });
      }
    });
  const response = await completion;
  sharePointUploadWaiters.delete(uploadRequestId);
  if (!response || !response.ok)
    throw new Error(
      String(
        (response && response.error) || "SharePoint artifact upload failed",
      ),
    );
  return Array.isArray(response.files) ? response.files : [];
}

// ---- Model-artifact → SharePoint link injection ---------------------------
// A turn that produces artifacts (present_files / CodeInterpreter output)
// uploads them to the SharePoint download folder. To surface the resulting
// links INSIDE the same answer, we hold this turn's terminal M365_DONE for a
// bounded window until its uploads settle, then append a link block to
// DONE.text. Python then emits that appended tail as one final delta via its
// existing terminal-snapshot path, so neither Python nor the streaming path
// change. Turns with artifactCount==0 skip all of this (a strict no-op), and
// the whole feature is one flag away from being disabled.
const M365_INJECT_ARTIFACT_LINKS = true; // flip to false to fully disable
// 内联图片：把模型产出的图片作为 data-URI markdown 直接嵌入答案，Continue 可内联渲染。
// data-URI 不依赖 SharePoint 登录态、不会过期、进对话历史也不裂（与短时效 downloadUrl 相反）。
const M365_INLINE_IMAGES = true; // flip to false 关闭内联，退回纯链接
// 单图内联 base64 上限。base64 会跨 extension→Python 的本地 WSS 桥（随 M365_DONE.appendText），
// 桥端 max_msg_size 已抬到 32MiB（见 voice_edge.py 补丁 B），这里再设单图 ~3MB 上限做双保险；
// 超限的图片回退为 SharePoint 链接，绝不丢。
const M365_INLINE_IMAGE_MAX_B64 = 3 * 1024 * 1024;
const _veIsImageMime = (m) => /^image\//i.test(String(m || ""));
// Size-aware ceiling on how long a DONE is held waiting for artifact uploads
// to settle. A flat 20s dropped large files: a ~1.4MB script whose
// addUsingPath + read-back GET + driveItem verification (with 503 backoff
// retries) runs well past 20s finalized at settled<expected, losing the late
// SharePoint link. Two inconsistencies drove the bug: (a) each artifact's OWN
// upload budget is sharePointUploadTimeoutMs(1)=60s, far longer than the 20s
// hold, so a legitimately-slow-but-successful upload was abandoned; (b) larger
// payloads need extra read-back/verify time. Scale accordingly:
//   floor 65s  — just above the 60s per-file upload budget, so a single slow
//                upload always finishes before finalize gives up;
//   +20s / MB  — extra budget for large-file read-back/verification;
//   cap 150s   — MUST stay below Python's M365_IDLE_TIMEOUT (180s): while a DONE
//                is held the extension forwards no relay frames, so an over-long
//                hold would let Python idle-tear-down and drop the held terminal.
//                A periodic keepalive (below) further guards this, but the cap is
//                the hard safety bound. The common case still finalizes early via
//                settled>=expected, so this only bounds the pathological hang.
const M365_ARTIFACT_LINK_TIMEOUT_FLOOR_MS = 65000;
const M365_ARTIFACT_LINK_MS_PER_MB = 20000;
const M365_ARTIFACT_LINK_TIMEOUT_CAP_MS = 150000;
// Liveness ping cadence while a DONE is held. Empty-text M365_PROGRESS is the
// codebase's established pure-liveness contract (flips got_any / resets the
// Python idle timer, never written to the answer), so pinging every few seconds
// keeps the relay alive through an arbitrarily long (but capped) upload wait.
const M365_ARTIFACT_HOLD_KEEPALIVE_MS = 5000;
function m365ArtifactHoldMs(totalBytes) {
  const mb = Math.max(0, Number(totalBytes) || 0) / (1024 * 1024);
  return Math.min(
    M365_ARTIFACT_LINK_TIMEOUT_CAP_MS,
    Math.max(
      M365_ARTIFACT_LINK_TIMEOUT_FLOOR_MS,
      Math.round(
        M365_ARTIFACT_LINK_TIMEOUT_FLOOR_MS + mb * M365_ARTIFACT_LINK_MS_PER_MB,
      ),
    ),
  );
}
const m365ArtifactTurns = new Map(); // request id -> per-turn injection state

function m365ArtifactTurn(id) {
  let state = m365ArtifactTurns.get(id);
  if (!state) {
    state = {
      expected: null, // artifactCount from DONE; null until DONE seen
      settled: 0, // uploads/errors resolved for this id so far
      descriptors: [], // successful SharePoint descriptors (name/url/downloadUrl)
      images: [], // {name, displayName, mimeType, dataUrl}，独立于 SP 上传
      artifactKeys: new Set(), // 同一产物可能被多个 WSS 消息/URL 重复 harvest；只消费一次
      bytes: 0, // 累计已观测 artifact 字节数，用于按大小动态计算持有窗口
      doneMessage: null, // the held terminal payload
      timer: null,
      keepAliveTimer: null, // 持有期间的 Python liveness 保活 interval
      finalized: false,
    };
    m365ArtifactTurns.set(id, state);
  }
  return state;
}

// Forward a relayable frame (delta / terminal / snapshot) to Python unchanged.
// Extracted so the immediate path and the delayed artifact-link path emit the
// exact same wire shape.
function m365ForwardToPy(message) {
  sendM365({
    type: message.type,
    id: message.id,
    text: message.text,
    conversationId: message.conversationId,
    chats: message.chats,
    capturedAt: message.capturedAt,
    error: message.error,
    completionSignal: message.completionSignal,
    authoritative: message.authoritative === true,
    turnState: message.turnState,
    deltaSource: message.deltaSource,
    // Out-of-band trailing block (SharePoint artifact links). Carried on its
    // OWN field instead of being spliced into `text`, so Python can deliver it
    // as a final delta unconditionally — even when this turn's authoritative
    // terminal snapshot diverged from the streamed text and is dropped as
    // non-prefix (the exact case that silently ate the links before).
    appendText: message.appendText,
    // Diagnostic-only per-turn artScan counters (see content-m365 done()).
    // Forwarded so they surface in the Python [relay-artifact] log; background
    // itself never acts on this field.
    artScanSummary: message.artScanSummary,
  });
}

// (Re)arm the size-aware finalize hold for a turn whose DONE is being held.
// Called once when the held DONE is first seen, and again whenever a later
// M365_ARTIFACT grows state.bytes — so a large payload that only becomes known
// after DONE still extends the window. Idempotent: clears any prior timer /
// keepalive first. No-op once finalized or before a DONE is held.
function m365ArmArtifactHold(id) {
  const state = m365ArtifactTurns.get(id);
  if (!state || state.finalized || !state.doneMessage) return;
  if (state.timer !== null) {
    clearTimeout(state.timer);
    state.timer = null;
  }
  if (state.keepAliveTimer !== null) {
    clearInterval(state.keepAliveTimer);
    state.keepAliveTimer = null;
  }
  const holdMs = m365ArtifactHoldMs(state.bytes);
  state.timer = setTimeout(
    () => m365FinalizeArtifactTurn(id, "timeout"),
    holdMs,
  );
  // Keep the Python relay alive for the whole (bounded) hold so an upload that
  // legitimately runs tens of seconds cannot be pre-empted by the 180s idle
  // teardown. Empty-text M365_PROGRESS is pure liveness (never written to the
  // answer). Cleared on finalize.
  state.keepAliveTimer = setInterval(() => {
    const s = m365ArtifactTurns.get(id);
    if (!s || s.finalized) return;
    sendM365({ type: "M365_PROGRESS", id, text: "" });
  }, M365_ARTIFACT_HOLD_KEEPALIVE_MS);
  dlog(
    "artifact hold (re)armed id=%s bytes=%d holdMs=%d",
    id,
    state.bytes,
    holdMs,
  );
}

function m365FinalizeArtifactTurn(id, reason) {
  const state = m365ArtifactTurns.get(id);
  if (!state || state.finalized || !state.doneMessage) return;
  state.finalized = true;
  if (state.timer !== null) {
    clearTimeout(state.timer);
    state.timer = null;
  }
  if (state.keepAliveTimer !== null) {
    clearInterval(state.keepAliveTimer);
    state.keepAliveTimer = null;
  }
  m365ArtifactTurns.delete(id);
  const done = state.doneMessage;

  // 描述符按文件名索引，便于给内联图片补上 SharePoint 链接
  const descByName = new Map();
  for (const d of state.descriptors) {
    const n = String((d && d.name) || "").trim();
    if (n && !descByName.has(n)) descByName.set(n, d);
  }
  const usedNames = new Set();
  const blocks = [];

  // 1) 内联图片（data-URI）—— 独立于 SharePoint 是否成功；下方附带 SP 链接（若有）
  for (const img of state.images || []) {
    const storedName = String(img.name || "").trim();
    const cap = String(img.displayName || storedName || "image").trim();
    const d = storedName ? descByName.get(storedName) : null;
    if (d) usedNames.add(storedName);
    const links = d
      ? [
          d.downloadUrl ? "[下载](" + d.downloadUrl + ")" : "",
          d.url ? "[永久链接](" + d.url + ")" : "",
        ]
          .filter(Boolean)
          .join(" · ")
      : "";
    blocks.push(
      "![" + cap + "](" + img.dataUrl + ")" + (links ? "\n" + links : ""),
    );
  }

  // 2) 非内联产物（非图片，或超限回退）—— 保持原纯文本链接形式
  for (const d of state.descriptors) {
    const name = String((d && (d.displayName || d.name)) || "").trim();
    const storedName = String((d && d.name) || "").trim();
    if (
      (name && usedNames.has(name)) ||
      (storedName && usedNames.has(storedName))
    )
      continue; // 已作为内联图片处理
    const dl = d && d.downloadUrl ? "[下载链接](" + d.downloadUrl + ")" : "";
    const pl = d && d.url ? "[永久链接](" + d.url + ")" : "";
    const parts = [dl, pl].filter(Boolean).join(", ");
    if (!parts) continue;
    blocks.push(name ? "- " + name + "：" + parts : "- " + parts);
  }

  if (blocks.length) {
    // 有内联图片时图片即主体，不加“已上传到 SharePoint”前缀；纯链接时保留原前缀。
    // 仍走 done.appendText 专用字段（不拼进 done.text），因此与终态快照 prefix 校验无关，
    // 不会被 Python 的 non-prefix 丢弃逻辑连带丢掉。
    const hasInline = (state.images || []).length > 0;
    done.appendText = hasInline
      ? "\n\n" + blocks.join("\n\n")
      : "\n\n已上传到 SharePoint：\n" + blocks.join("\n");
  }
  dlog(
    "artifact finalize id=%s reason=%s images=%d descriptors=%d settled=%d/%s",
    id,
    reason,
    (state.images || []).length,
    state.descriptors.length,
    state.settled,
    String(state.expected),
  );
  m365ForwardToPy(done);
}

function m365MaybeFinalizeArtifactTurn(id) {
  const state = m365ArtifactTurns.get(id);
  if (!state || state.finalized || !state.doneMessage) return;
  if (state.expected !== null && state.settled >= state.expected) {
    m365FinalizeArtifactTurn(id, "complete");
  }
}

// Exact frame that accepted each id's ASK (tabId+frameId). The socket that
// streams a turn lives in ONE specific (often cross-origin) subframe; routing
// STOP to that frameId — the SAME way M365_ASK is delivered — is reliable, where
// a tab-level broadcast was not (the observed "M365_STOP recv but no
// M365_STOPPED, drops persist" bug).
const m365AskFrameById = new Map();

// Abort an in-flight M365 turn the client (e.g. Continue) interrupted. Deliver
// M365_STOP to the exact frame that ran this id's ASK (frameId-targeted), then
// to the sticky target and every known candidate frame as a fallback, so the
// frame whose self-built Chathub socket is streaming closes it and the model
// stops. A frame not running the id no-ops (its page-world registry has no
// entry). Also release any held terminal/artifact state for the id.
function routeM365Stop(id) {
  const rid = String(id || "");
  if (!rid) return;
  const targets = [];
  const seen = new Set();
  const pushFrame = (f) => {
    if (!f || f.tabId == null || f.frameId == null) return;
    const key = f.tabId + ":" + f.frameId;
    if (seen.has(key)) return;
    seen.add(key);
    targets.push(f);
  };
  pushFrame(m365AskFrameById.get(rid)); // exact frame first
  pushFrame(m365TargetFrame);
  for (const f of m365FrameCandidates.values()) pushFrame(f);
  dlog(
    "routeM365Stop id=%s targets=%d exact=%s",
    rid,
    targets.length,
    m365AskFrameById.has(rid),
  );
  for (const f of targets) {
    browser.tabs
      .sendMessage(
        f.tabId,
        { __veM365ToPage: true, payload: { type: "M365_STOP", id: rid } },
        { frameId: f.frameId },
      )
      .catch(() => {});
  }
  m365AskFrameById.delete(rid);
  const state = m365ArtifactTurns.get(rid);
  if (state) {
    if (state.timer !== null) clearTimeout(state.timer);
    if (state.keepAliveTimer !== null) clearInterval(state.keepAliveTimer);
    m365ArtifactTurns.delete(rid);
  }
}

async function dispatchM365(message) {
  // Liveness ack: tell Python we accepted the request the instant it arrives,
  // before any slow browser-side preparation (tab cold-start, SharePoint
  // attachment upload, hook wait). None of that produces a Chathub frame, so
  // without an early signal a slow prep is counted as dead air against the
  // Python first-event timeout and surfaces as a misleading timeout while any
  // real error that follows is dropped. A single empty-text M365_PROGRESS flips
  // got_any=true on the Python side (liveness only; never written to the answer).
  const keepAlive = () =>
    sendM365({ type: "M365_PROGRESS", id: message.id, text: "" });
  keepAlive();
  try {
    const attachments = await prepareM365Attachments(message);
    keepAlive(); // attachments prepared/uploaded
    const outbound = { ...message, attachments };
    const tab = await ensureM365Tab();
    keepAlive(); // tab resolved (may have just been created and is cold-loading)

    // Ensure at least one hook is ready, then try every candidate frame in
    // priority order. The relay content script answers with {ok:true,
    // relayed:true,...}; a stale/duplicate frameId instead rejects or returns
    // no ack. Confirming the ack is what makes routing robust to the
    // duplicate-hook hazard: the ASK is delivered to the frame that actually
    // runs doAsk, rather than being silently swallowed by a dead frameId.
    if (!bestM365FrameForTab(tab.id)) await waitForM365Hook(tab.id);
    let frames = orderedM365FramesForTab(tab.id);
    if (!frames.length) {
      await waitForM365Hook(tab.id);
      frames = orderedM365FramesForTab(tab.id);
    }

    let delivered = false;
    let lastError = null;
    for (const target of frames) {
      try {
        const ack = await browser.tabs.sendMessage(
          tab.id,
          { __veM365ToPage: true, payload: outbound },
          { frameId: target.frameId },
        );
        if (ack && ack.relayed === true) {
          m365TargetFrame = target; // make the proven-live frame sticky
          // Remember the EXACT frame for this id so a later M365_STOP can be
          // delivered straight to the frame whose Chathub socket is streaming.
          if (message.id)
            m365AskFrameById.set(String(message.id), {
              tabId: tab.id,
              frameId: target.frameId,
            });
          delivered = true;
          dlog(
            "dispatch delivered id=%s frameId=%s frameUrl=%s",
            String(message.id || ""),
            target.frameId,
            target.frameUrl,
          );
          break;
        }
        // A frame with the hook but no relay ack is not usable for this ASK.
        lastError = new Error("M365 frame did not acknowledge ASK relay");
        dlog(
          "dispatch no-ack, pruning frameId=%s id=%s",
          target.frameId,
          String(message.id || ""),
        );
        pruneStaleM365Frame(target);
      } catch (sendError) {
        // frameId no longer exists / no receiver: prune and try the next one.
        lastError = sendError;
        dlog(
          "dispatch send failed, pruning frameId=%s id=%s err=%s",
          target.frameId,
          String(message.id || ""),
          compactError(sendError),
        );
        pruneStaleM365Frame(target);
      }
    }

    if (!delivered) {
      throw (
        lastError ||
        new Error(
          "No M365 frame accepted the ASK; reload the configured M365 entry page",
        )
      );
    }
  } catch (error) {
    sendM365({
      type: "M365_ERROR",
      id: message.id,
      error: compactError(error),
    });
  }
}

function sendToM365Frame(sender, payload) {
  if (!sender.tab || sender.tab.id == null) return;
  const options =
    sender.frameId == null ? undefined : { frameId: sender.frameId };
  browser.tabs
    .sendMessage(sender.tab.id, { __veM365ToPage: true, payload }, options)
    .catch(() => {});
}

browser.runtime.onMessage.addListener((message, sender) => {
  if (message && message.__veAis === true) {
    if (sender && sender.tab && sender.tab.id != null) {
      aisTargetTabId = sender.tab.id;
    }
    if (message.type === "AIS_NEED_MAIN_WORLD") {
      if (
        sender &&
        sender.tab &&
        sender.tab.id != null &&
        browser.scripting &&
        typeof browser.scripting.executeScript === "function"
      ) {
        browser.scripting
          .executeScript({
            target: { tabId: sender.tab.id },
            world: "MAIN",
            func: (code) => {
              try {
                // eslint-disable-next-line no-eval
                (0, eval)(code);
              } catch (_) {}
            },
            args: [String(message.code || "")],
          })
          .catch(() => {});
      }
      return Promise.resolve({ ok: true });
    }
    if (
      message.type === "AIS_READY" ||
      message.type === "AIS_HEAD" ||
      message.type === "AIS_CHUNK" ||
      message.type === "AIS_END" ||
      message.type === "AIS_ERROR"
    ) {
      if (message.type === "AIS_READY") {
        aisLastReady = Object.assign({}, message, {
          updatedAt: Date.now(),
          tabId: (sender && sender.tab && sender.tab.id) || aisTargetTabId || 0,
        });
        if (aisLastReady.auth && aisLastReady.auth.ok) {
          aisLastError = null;
        }
        const readySig = JSON.stringify([
          message.service || "",
          message.tabUrl || "",
          message.timezone || "",
          (message.auth && message.auth.ok) || false,
          (message.auth && message.auth.authUser) || "",
          (message.auth && message.auth.tier) || "",
          (message.auth && message.auth.hasApiKey) || false,
          (message.auth && message.auth.hasExtBin) || false,
          (message.auth && message.auth.sapisidHash) || "",
          (message.drive && message.drive.status) || "",
        ]);
        if (!message.id && readySig === aisLastForwardedReadySig) {
          return Promise.resolve({ ok: true });
        }
        aisLastForwardedReadySig = readySig;
      }
      if (message.type === "AIS_ERROR") {
        aisLastError = {
          code: String(message.code || ""),
          message: String(message.message || ""),
          httpStatus: Number(message.httpStatus || 0),
          at: Date.now(),
        };
      }
      if (message.type === "AIS_END" || message.type === "AIS_ERROR") {
        const reqId = String(message.id || "");
        if (reqId) aisTabByRequestId.delete(reqId);
      }
      const clean = Object.assign({}, message);
      delete clean.__veAis;
      delete clean.dir;
      sendM365(clean);
      return Promise.resolve({ ok: true });
    }
    return undefined;
  }
  if (!message || message.__veM365 !== true) return;
  return (async () => {
    if (message.type === "M365_FRAME_READY") {
      console.log(
        "VE-FRAME-READY",
        message.frameUrl,
        message.frameOrigin,
        sender.frameId,
      );
      const caps = message.capabilities || {};
      if (
        caps.messageListener !== true ||
        caps.chatBuilder !== true ||
        caps.doAsk !== true ||
        caps.websocket !== true
      ) {
        log("ignored incomplete M365 page hook", {
          capabilities: caps,
        });
        return;
      }
      const frame = rememberM365Frame(message, sender, false);
      if (frame) {
        sendM365({
          type: "M365_READY",
          frameOrigin: frame.frameOrigin,
          frameUrl: frame.frameUrl,
        });
      }
      return;
    }
    if (message.type === "M365_ENTRY_DISCOVERED") {
      await setM365EntryUrl(message.entryUrl);
      return;
    }
    if (message.type === "M365_TOKEN_RESPONSE") {
      if (!isAllowedM365Sender(sender)) {
        log(
          "ignored M365_TOKEN_RESPONSE from untrusted sender:",
          (sender && sender.url) ||
            (sender && sender.tab && sender.tab.url) ||
            "unknown",
        );
        return;
      }
      try {
        await applyM365TokenResponse(message);
      } catch (error) {
        log("failed to apply M365 token response:", compactError(error));
      }
      return;
    }
    if (message.type === "M365_STAGE_LOGIN_CREDS") {
      await Promise.all([ensureM365AuthLoaded(), ensureM365AutoLoginLoaded()]);
      if (
        !m365AutoLoginEnabled ||
        !m365AutoLoginRemember ||
        !isSenderArmedM365LoginTab(sender, message)
      ) {
        return;
      }
      const pwd = String(message.password || "");
      const usr = String(message.username || "").trim();
      if (pwd) {
        _stagedAutoLoginCreds = {
          username: usr,
          password: pwd,
          stagedAt: Date.now(),
        };
        dlog(
          "staged M365 manual login credentials for user=%s",
          usr || "<unknown>",
        );
      }
      return;
    }
    if (message.type === "M365_AUTO_LOGIN_SUBMITTING") {
      const armed = isSenderArmedM365LoginTab(sender, message);
      log(
        "M365 auto-login submitting password tab=%s armed=%s user=%s",
        sender && sender.tab ? sender.tab.id : "?",
        armed,
        String(message.username || "-").slice(0, 60),
      );
      if (!armed) return;
      const usr = String(message.username || "").trim();
      if (usr && !m365AutoLoginUsername) {
        m365AutoLoginUsername = usr;
        saveM365AutoLogin().catch(() => {});
      }
      if (_pendingTopLevelReauth && _pendingTopLevelReauth.timer) {
        clearTimeout(_pendingTopLevelReauth.timer);
        _pendingTopLevelReauth.timer = setTimeout(() => {
          log("M365 auto-login timed out 25s after submit; failing tab");
          failPendingReauth("login tab timed out 25s after submit");
        }, 25000);
      }
      return;
    }
    if (message.type === "M365_AUTO_LOGIN_RESULT") {
      handleM365AutoLoginResult(message, sender).catch((error) => {
        log("handleM365AutoLoginResult error:", compactError(error));
      });
      return;
    }
    if (message.type === "M365_OAUTH_REDIRECT") {
      handleM365OAuthRedirect(message, sender).catch((error) => {
        log("handleM365OAuthRedirect error:", compactError(error));
      });
      return;
    }
    if (message.type === "M365_PAGE_TOKEN_PROBE_RESULT") {
      log(
        "M365 probe: page /token after RT swap → status=%s hasAccessToken=%s swapped=%s %s",
        Number(message.status || 0),
        Boolean(message.hasToken),
        Number(message.swapped || 0),
        String(message.error || "").slice(0, 160),
      );
      sendM365({
        type: "M365_PAGE_TOKEN_PROBE_RESULT",
        status: Number(message.status || 0),
        hasToken: Boolean(message.hasToken),
        swapped: Number(message.swapped || 0),
        error: String(message.error || ""),
      });
      return { ok: true };
    }
    if (message.type === "M365_RT_SWAP_FAILED") {
      // The page's /token call failed with OUR refresh token. There is no
      // fallback to the page's own (older) RT: our RT is the single source of
      // truth. So prove its state with a real refresh; if it is really dead the
      // existing failure path (AADSTS700084/700082 -> handleRefreshTokenExhausted)
      // opens the one top-level authorize that mints a new 24h RT.
      await ensureM365AuthLoaded();
      const status = Number(message.status || 0);
      log(
        "M365 page /token was rejected with our RT (status=%s) — verifying with a real refresh, no fallback to the page's old RT",
        status,
      );
      if (!_rtVerifyInFlight && Date.now() - _lastRtVerifyAt > 60 * 1000) {
        _lastRtVerifyAt = Date.now();
        _rtVerifyInFlight = (async () => {
          _sydney = { token: "", exp: 0 };
          await refreshSydney(0);
          // Still alive: push the (rotated) RT so the page retries with it.
          await pingM365TabsKeepAlive({ authHealthy: true });
          log(
            "M365 page RT verified alive after rejection; pushed fresh RT to pages",
          );
        })()
          .catch((error) =>
            log(
              "M365 page RT verification failed (reauth path decides next):",
              compactError(error),
            ),
          )
          .finally(() => {
            _rtVerifyInFlight = null;
          });
      } else {
        dlog("M365 page RT rejection already being handled");
      }
      return { ok: true };
    }
    if (message.type === "M365_RT_SWAPPED") {
      // The page-world fetch hook replaced OWA's own refresh_token with ours on
      // its /token call. This is the proof that a renewed RT is actually being
      // used by the page (no reload needed).
      if (Date.now() - _lastRtSwapLogAt > 5 * 60 * 1000) {
        _lastRtSwapLogAt = Date.now();
        log(
          "M365 page /token now uses the extension's RT (swapped at request time) — %s",
          m365StateLine(),
        );
      }
      return { ok: true };
    }
    if (message.type === "M365_POPUP_INTERCEPTED") {
      const action = String(message.action || "");
      log(
        "M365 authorize popup intercepted (%s) tab=%s url=%s",
        action,
        sender && sender.tab ? sender.tab.id : "?",
        String(message.url || "").slice(0, 160),
      );
      if (action === "suppressed-healthy") {
        return;
      }
      const now = Date.now();
      const rtHealthy = !_rtExp || _rtExp - now > 5 * 60 * 1000;
      const sydneyHealthy = Boolean(_sydney.token && _sydney.exp - now > 60000);
      if (!m365ReauthRequired && rtHealthy && sydneyHealthy) {
        pingM365TabsKeepAlive({
          authHealthy: true,
          suppressPopups: true,
        }).catch(() => {});
        return;
      }
      // An OWA-driven authorize popup is never useful: it is the same silent
      // attempt that already failed. Suppress it from now on instead of raising
      // it ("reused-open-popup" used to focus it every ~60s, which is exactly
      // what looked like "a new login window every minute"). Our own pre-expiry
      // cycle opens a single controlled tab when it is actually needed.
      pingM365TabsKeepAlive({ authHealthy: false, suppressPopups: true }).catch(
        () => {},
      );
      if (!m365ReauthRequired) {
        markM365ReauthRequired(
          "OWA 触发登录弹窗 (24h 令牌到期)，需重新输入密码登录",
        );
      } else {
        dlog(
          "M365 popup intercepted while already in NEEDS_REAUTH (%s)",
          action,
        );
      }
      return;
    }
    if (message.type === "M365_AUTO_LOGIN_DIAG") {
      log(
        "M365 auto-login diag: %s tab=%s url=%s cfg=%s",
        String(message.state || "?"),
        sender && sender.tab ? sender.tab.id : "?",
        String(message.url || "").slice(0, 120),
        String(message.detail || "-").slice(0, 240),
      );
      return;
    }
    if (message.type === "M365_SESSION_RENEWED") {
      m365LastKeepAliveAt = Number(message.at || Date.now());
      m365LastKeepAliveError = "";
      m365LastKeepAliveErrorAt = 0;
      dlog("M365 ABT session renewed at=%d", m365LastKeepAliveAt);
      return;
    }
    if (message.type === "M365_SESSION_RENEW_FAILED") {
      m365LastKeepAliveError = String(message.error || "network/offline");
      m365LastKeepAliveErrorAt = Number(message.at || Date.now());
      log("M365 ABT session renew failed: %s", m365LastKeepAliveError);
      schedulePreExpiryRetry("m365-abt");
      return;
    }
    if (message.type === "M365_NEED_TOKEN") {
      try {
        const token = await refreshSydney();
        sendToM365Frame(sender, {
          type: "M365_TOKEN",
          token,
          exp: _sydney.exp,
        });
      } catch (error) {
        sendToM365Frame(sender, {
          type: "M365_TOKEN",
          token: "",
          error: compactError(error),
        });
      }
      return;
    }
    if (message.type === "M365_NEED_AMS_TOKEN") {
      // IC3-audience token for artifact (AMS) downloads; see refreshIc3Token.
      try {
        // A forced request (content-side 401/403 retry) invalidates the cached
        // IC3 token so refreshIc3Token mints a brand-new one instead of handing
        // back the stale token that just 401'd. (Any refresh already in flight
        // still fetches a fresh token, so this cannot return the stale one.)
        if (message.force) _ic3 = { token: "", exp: 0 };
        const token = await refreshIc3Token();
        sendToM365Frame(sender, {
          type: "M365_AMS_TOKEN",
          token,
          exp: _ic3.exp,
        });
      } catch (error) {
        log("IC3 token acquisition failed:", compactError(error));
        sendToM365Frame(sender, {
          type: "M365_AMS_TOKEN",
          token: "",
          error: compactError(error),
        });
      }
      return;
    }
    if (message.type === "M365_CHATS_SNAPSHOT") {
      rememberM365Frame(message, sender, true);
    }
    // Model-produced artifact fetched by content-m365's harvester. Upload it to
    // the SharePoint download folder. Deliberately NOT forwarded to Python (it
    // is not in the sendM365 whitelist below), so it never touches the answer
    // stream. Each artifact arrives as its own message; upload independently
    // (sp.js uses overwrite=true, so this is idempotent).
    if (message.type === "M365_ARTIFACT") {
      const id = String(message.id || "");
      const track = M365_INJECT_ARTIFACT_LINKS && id;
      const file = {
        name: String(message.name || ""),
        displayName: String(message.displayName || message.name || ""),
        data: String(message.data || ""),
        size: Number(message.size || 0),
        mimeType: String(message.mimeType || "application/octet-stream"),
      };
      // M365 有时会在多个消息分支中重复公布同一个对象 URL，harvester 随后会把
      // 完全相同的 artifact 发送多次。若每帧都上传/入 blocks，DONE.appendText 就会
      // 生成重复下载链接。以文件名、声明大小和 base64 内容的首尾指纹做 turn 内去重；
      // 不只按文件名去重，避免合法的同名不同内容产物被吞掉。
      let state = null;
      if (track) {
        state = m365ArtifactTurn(id);
        const artifactKey = [
          file.name,
          String(file.size || ""),
          String(file.mimeType || "").toLowerCase(),
          file.data.length,
          file.data.slice(0, 96),
          file.data.slice(-96),
        ].join("\u0000");
        if (state.artifactKeys.has(artifactKey)) {
          dlog("duplicate M365 artifact ignored id=%s name=%s", id, file.name);
          // 该重复产物在 content-m365 侧仍对应 artifactCount 里的一个唯一 URL，
          // harvestOne 为它 post 了这一帧。若这里直接 return 而不推进 settled，
          // settled 就永远停在 expected-1，m365MaybeFinalizeArtifactTurn 的快速
          // finalize（settled>=expected）永不触发，DONE 只能干等 hold 定时器超时
          // 才收尾——每个带重复产物的回复都被拖慢甚至掉队。照 malformed 分支的做法
          // 照常记一次 settled 并尝试 finalize，再丢弃这帧的字节/上传。
          state.settled += 1;
          m365MaybeFinalizeArtifactTurn(id);
          return;
        }
        state.artifactKeys.add(artifactKey);
      }
      // 累计本轮 artifact 字节数，并按新的总字节数延长持有窗口。artifact 常在 DONE
      // 之后才 harvest 到（harvestOne 轮询最多 ~6s），大文件的字节数此时才可知。
      if (state) {
        const fileBytes =
          file.size > 0 ? file.size : Math.floor((file.data.length * 3) / 4);
        state.bytes += Math.max(0, fileBytes);
        m365ArmArtifactHold(id);
      }
      // 先同步记录可内联的图片（独立于 SharePoint 成败/是否配置）。哪怕后面 SP 上传失败或
      // 超时，图片也已经在 state.images 里，finalize 时照样内联出图。
      if (
        track &&
        M365_INLINE_IMAGES &&
        _veIsImageMime(file.mimeType) &&
        file.data
      ) {
        if (file.data.length <= M365_INLINE_IMAGE_MAX_B64) {
          const state = m365ArtifactTurn(id);
          state.images.push({
            name: file.name,
            displayName: file.displayName || file.name,
            mimeType: file.mimeType,
            dataUrl: "data:" + file.mimeType + ";base64," + file.data,
          });
        } else {
          dlog(
            "artifact too large to inline, link-only name=%s b64=%d",
            file.name,
            file.data.length,
          );
        }
      }
      if (file.name && file.data) {
        // Always resolves (never rejects) to a descriptor array, so the
        // settled-counter below advances on both success and failure.
        const uploadPromise = uploadArtifactsToSharePoint([file])
          .then((uploaded) => {
            const list = (Array.isArray(uploaded) ? uploaded : []).map(
              (descriptor) =>
                descriptor && typeof descriptor === "object"
                  ? Object.assign({}, descriptor, {
                      displayName: file.displayName || file.name,
                    })
                  : descriptor,
            );
            const info = list[0] || {};
            log(
              "M365 artifact uploaded to SharePoint:",
              file.name,
              info.url || info.serverRelativeUrl || "(no url)",
            );
            return list;
          })
          .catch((error) => {
            log("M365 artifact upload failed:", file.name, compactError(error));
            return [];
          });
        if (track) {
          const state = m365ArtifactTurn(id);
          uploadPromise
            .then((uploaded) => {
              for (const d of uploaded) if (d) state.descriptors.push(d);
            })
            .finally(() => {
              state.settled += 1;
              m365MaybeFinalizeArtifactTurn(id);
            });
        }
      } else if (track) {
        // Malformed artifact (no name/bytes): still count it as settled so a
        // held DONE can release once every expected artifact is accounted for.
        const state = m365ArtifactTurn(id);
        state.settled += 1;
        m365MaybeFinalizeArtifactTurn(id);
      }
      return;
    }
    if (message.type === "M365_ARTIFACT_ERROR") {
      log(
        "M365 artifact fetch failed:",
        String(message.url || ""),
        String(message.error || ""),
      );
      // A failed fetch still resolves one of this turn's expected artifacts, so
      // counting it lets a held DONE finalize early (with whatever links did
      // succeed) instead of always waiting out the timeout.
      const id = String(message.id || "");
      if (M365_INJECT_ARTIFACT_LINKS && id) {
        const state = m365ArtifactTurn(id);
        state.settled += 1;
        m365MaybeFinalizeArtifactTurn(id);
      }
      return;
    }
    // --- Reasoning / progress liveness bridge --------------------------------
    // During a long "thinking" phase the model streams ONLY chain-of-thought /
    // progress frames — no answer tokens.
    // content-m365 emits these as M365_PROGRESS (plus
    // M365_REASONING for genuine CoT) explicitly "for the Python relay", but they
    // were never forwarded here. The extension's own idle watchdog resets on
    // every socket frame and DELIBERATELY delegates the content-based idle
    // timeout to the Python relay — yet the relay only ever saw M365_DELTA, so
    // during the reasoning phase it received nothing, tripped its content-idle
    // timeout, and tore down the SSE to Continue BEFORE the answer body began.
    // The entire reply was lost. Forward these
    // frames so the relay's liveness/idle timer keeps resetting until the first
    // answer delta arrives.
    //
    // M365_PROGRESS is forwarded as a PURE liveness ping (empty text) — the exact
    // contract the relay already relies on (see dispatchM365's keepAlive): it
    // flips got_any / resets the idle timer and is NEVER written to the answer,
    // so the append-only answer stream is byte-for-byte untouched. M365_REASONING
    // additionally carries the cumulative CoT text on its own dedicated channel
    // (append-only per content-m365), keeping any surfaced reasoning off the
    // answer path. The liveness ping alone fixes the loss even if the relay
    // ignores M365_REASONING.
    if (message.type === "M365_STOPPED") {
      // Diagnostic confirmation that a bridge M365_STOP reached the page world
      // and tore down that turn's Chathub socket. dlog only; deliberately NOT
      // forwarded to Python (not in the relay whitelist), so it never touches
      // the answer stream.
      dlog("M365_STOPPED recv id=%s", String(message.id || ""));
      return;
    }
    if (message.type === "M365_PROGRESS") {
      sendM365({
        type: "M365_PROGRESS",
        id: message.id,
        text: "",
        codeExecuting: message.codeExecuting === true,
        messageId: message.messageId,
        contentType: message.contentType,
        addToChainOfThought: message.addToChainOfThought === true,
      });
      return;
    }
    if (message.type === "M365_REASONING") {
      sendM365({ type: "M365_PROGRESS", id: message.id, text: "" });
      sendM365({
        type: "M365_REASONING",
        id: message.id,
        text: message.text,
        messageId: message.messageId,
      });
      return;
    }
    if (
      ["M365_DELTA", "M365_DONE", "M365_ERROR", "M365_CHATS_SNAPSHOT"].includes(
        message.type,
      )
    ) {
      if (message.type === "M365_DONE" || message.type === "M365_ERROR") {
        dlog(
          "relay->py %s id=%s conversationId=%s signal=%s authoritative=%s turnState=%s textLen=%d%s",
          message.type,
          String(message.id || ""),
          String(message.conversationId || "") || "<none>",
          String(message.completionSignal || ""),
          message.authoritative === true,
          String(message.turnState || ""),
          String(message.text || "").length,
          message.error ? " error=" + String(message.error) : "",
        );
      }
      // Hold the terminal ONLY when this turn produced artifacts and injection
      // is on; deltas, snapshots, errors, and artifact-free DONEs forward
      // immediately, so the streaming path is byte-for-byte unchanged.
      if (
        message.type === "M365_DONE" &&
        Number(message.artifactCount || 0) > 0 &&
        String(message.id || "") &&
        // 满足任一即持有：①要注入 SP 链接且已配置；②要内联图片（不依赖 SharePoint）
        ((M365_INJECT_ARTIFACT_LINKS && sharePointHomeUrl) ||
          M365_INLINE_IMAGES)
      ) {
        const id = String(message.id);
        const state = m365ArtifactTurn(id);
        state.expected = Number(message.artifactCount) || 0;
        state.doneMessage = message;
        // Safety net: never hold the answer open indefinitely. The hold is
        // sized from the artifact bytes already observed (m365ArmArtifactHold),
        // and re-armed as later/larger artifacts arrive, so a big file gets a
        // proportionally longer window instead of the old flat 20s that dropped
        // it. If uploads are slow or an expected artifact never reports, the
        // (bounded) timer still releases with whatever links succeeded.
        m365ArmArtifactHold(id);
        // Artifacts (and their uploads) may already have settled before this
        // DONE arrived, in which case finalize right away.
        m365MaybeFinalizeArtifactTurn(id);
        return;
      }
      // Any terminal we forward WITHOUT holding (an error, or a DONE we chose
      // not to hold, e.g. SharePoint unconfigured) may have created artifact-
      // turn state from early M365_ARTIFACT frames. Drop it so the map cannot
      // leak. (Held DONEs return above and clean up in finalize.)
      if (message.type === "M365_DONE" || message.type === "M365_ERROR") {
        const id = String(message.id || "");
        const state = id && m365ArtifactTurns.get(id);
        if (state) {
          if (state.timer !== null) clearTimeout(state.timer);
          if (state.keepAliveTimer !== null)
            clearInterval(state.keepAliveTimer);
          m365ArtifactTurns.delete(id);
        }
      }
      m365ForwardToPy(message);
    }
  })();
});

browser.tabs.onRemoved.addListener((tabId) => {
  if (aisTargetTabId === tabId) {
    aisTargetTabId = null;
    aisLastReady = null;
    aisLastForwardedReadySig = "";
  }
  for (const [reqId, tId] of aisTabByRequestId) {
    if (tId === tabId) {
      aisTabByRequestId.delete(reqId);
      sendM365({
        type: "AIS_ERROR",
        id: reqId,
        code: "tab_closed",
        message: "AI Studio tab was closed while request was in flight",
        httpStatus: 503,
      });
    }
  }
  for (const [key, frame] of m365FrameCandidates) {
    if (frame.tabId === tabId) m365FrameCandidates.delete(key);
  }
  if (m365TargetFrame && m365TargetFrame.tabId === tabId)
    m365TargetFrame = null;
  if (_pendingTopLevelReauth && _pendingTopLevelReauth.tabId === tabId) {
    const wasLive = m365ReauthState === "attempting";
    log(
      "M365 reauth: login tab %s was closed without an authorization code (state=%s)",
      tabId,
      m365ReauthState,
    );
    if (_pendingTopLevelReauth.timer)
      clearTimeout(_pendingTopLevelReauth.timer);
    if (_pendingTopLevelReauth.notifyTimer)
      clearTimeout(_pendingTopLevelReauth.notifyTimer);
    _pendingTopLevelReauth = null;
    clearPersistedPendingReauth().catch(() => {});
    if (wasLive) {
      failM365Reauth("login tab closed before completing").catch(() => {});
    }
  }
});

Promise.all([
  ensureM365AuthLoaded(),
  ensureM365SettingsLoaded(),
  ensureM365AutoLoginLoaded(),
])
  .catch((error) => log("M365 initialization failed:", compactError(error)))
  .finally(m365Connect);
setInterval(() => {
  if (!m365Ws) m365Connect();
}, 5000);
