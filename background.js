"use strict";

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
    await browser.browserAction.setBadgeText({ text });
    if (color) {
      await browser.browserAction.setBadgeBackgroundColor({ color });
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
        const state = parseSseBody(body);
        const complete = Boolean(
          state.conversationId &&
          state.sectionId &&
          state.lastMessageIndex > 0 &&
          state.endType3 &&
          !state.streamError,
        );
        const tab = await browser.tabs.get(details.tabId);
        const capture = {
          tabId: details.tabId,
          cookieStoreId: tab.cookieStoreId || "",
          startedAt: Number(details.timeStamp || Date.now()),
          completedAt: Date.now(),
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
          // A green badge has one unambiguous meaning: Voice Edge requested an
          // auth refresh and this completed response is about to be synced.
          if (authRequired) {
            await setBadge("✓", "#188038");
            scheduleAutoSync(capture);
          }
        } else if (state.streamError && authRequired) {
          await setBadge("!", "#d93025");
        }
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

browser.browserAction.onClicked.addListener(async () => {
  try {
    if (!lastCaptured || !lastCaptured.complete || lastCaptured.error) {
      throw new Error("请先在豆包网页发送一条消息并等待完整回复");
    }
    await synchronizeCapture(lastCaptured, false);
  } catch (error) {
    syncInFlight = false;
    await setBadge("!", "#d93025");
    await notify("Voice Edge 手动同步失败", compactError(error));
  }
});

connectNative();

/* ==== Voice Edge popup panel bridge (state round-trip to Voice Edge) ==== */
let lastDeepSeekSnapshot = null;

const POPUP_OPEN_URLS = {
  deepseek: "https://chat.deepseek.com/",
  qwen: "https://chat.qwen.ai/",
  doubao: "https://www.doubao.com/chat/",
  m365: "https://outlook.cloud.microsoft/",
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
      lastRefreshError: m365LastRefreshError,
      lastRefreshErrorAt: Number(m365LastRefreshErrorAt || 0),
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
  const url =
    key === "m365"
      ? m365EntryUrl || POPUP_OPEN_URLS.m365
      : POPUP_OPEN_URLS[key];
  if (!url) return { ok: false, error: "未知的 provider" };
  try {
    await browser.tabs.create({ url, active: true });
    return { ok: true };
  } catch (error) {
    return { ok: false, error: String((error && error.message) || error) };
  }
}

browser.runtime.onMessage.addListener((message) => {
  if (!message || typeof message !== "object") return;
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
    return Promise.resolve(popupBuildState());
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
const M365_BRIDGE_URL = "ws://127.0.0.1:5002/ws";
const M365_AUTH_KEY = "voiceEdgeM365AuthV1";
const M365_SETTINGS_KEY = "voiceEdgeM365SettingsV1";
const M365_SYDNEY_SCOPE = "https://substrate.office.com/sydney/.default";
// Audience required by the AMS object endpoint (asyncgw.teams.microsoft.com/
// v1/objects/…) that serves CodeInterpreter / present_files artifacts. The
// Sydney token (aud=substrate.office.com/sydney) is REJECTED there with 401 —
// that was the harvestArtifacts failure. This IC3/Teams scope mints a token
// whose aud is ic3.teams.office.com, which AMS accepts.
const M365_IC3_SCOPE = "https://ic3.teams.office.com/.default";
const M365_TENANT_RE = /login\.microsoftonline\.com\/([0-9a-f-]{36})\/oauth2/i;
const M365_ENTRY_RE =
  /^https:\/\/outlook\.cloud\.microsoft\/host\/[0-9a-f-]{36}\/entity1-[0-9a-f-]{36}\/?$/i;

let _rt = "";
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
const log = (...args) => console.log("[VE-m365-bg]", ...args);
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
browser.runtime.onMessage.addListener((message, sender) => {
  if (!message || message.__veSharePoint !== true) return;
});

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
  _clientId = "";
  _tid = "";
  _m365TokenContext = { version: 1, query: {}, body: {} };
  _sydney = { token: "", exp: 0 };
  _sharePoint = { token: "", exp: 0, origin: "" };
  _sharePointRefreshing = null;
  m365LastRefreshError = "";
  m365LastRefreshErrorAt = 0;
  m365AuthUpdatedAt = 0;
  m365AuthLoaded = true;
  await browser.storage.local.remove(M365_AUTH_KEY);
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

function decodeM365Jwt(token) {
  try {
    const part = String(token || "")
      .split(".")[1]
      .replace(/-/g, "+")
      .replace(/_/g, "/");
    return JSON.parse(atob(part.padEnd(Math.ceil(part.length / 4) * 4, "=")));
  } catch (_) {
    return {};
  }
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
  if (
    params.get("grant_type") !== "refresh_token" ||
    !scopes.includes(M365_SYDNEY_SCOPE)
  )
    return null;
  const clientId = String(
    params.get("client_id") || requestUrl.searchParams.get("client_id") || "",
  );
  const refreshToken = String(params.get("refresh_token") || "");
  if (!clientId || !refreshToken) return null;

  const tokenBody = Object.fromEntries(params.entries());

  delete tokenBody.refresh_token;
  delete tokenBody.client_request_id;
  delete tokenBody["client-request-id"];

  return {
    tenantId: tenantMatch[1],
    clientId,
    refreshToken,
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
  _tid = captured.tenantId;
  _clientId = captured.clientId;
  _rt = captured.refreshToken;
  _m365TokenContext = captured.tokenContext || {
    version: 1,
    query: {},
    body: {},
  };
  _sydney = { token: "", exp: 0 };
  await saveM365Auth();
  log("captured Sydney refresh credentials");
}

browser.webRequest.onBeforeRequest.addListener(
  (details) => {
    let captured;
    try {
      captured = parseM365TokenRequest(details);
    } catch (error) {
      log("failed to parse M365 token request:", compactError(error));
      return;
    }
    if (!captured) return;
    applyCapturedM365Auth(captured).catch((error) => {
      log("failed to persist M365 credentials:", compactError(error));
    });
  },
  { urls: ["https://login.microsoftonline.com/*/oauth2/v2.0/token*"] },
  ["requestBody"],
);

async function applyM365TokenResponse(message) {
  await ensureM365AuthLoaded();
  const accessToken = String(message.accessToken || "");
  const refreshToken = String(message.refreshToken || "");
  const clientId = String(message.clientId || "");
  const tenantId = String(message.tenantId || "");
  const capturedAt = Number(message.capturedAt || Date.now());
  if (!accessToken || !refreshToken || !clientId || !tenantId) {
    throw new Error("incomplete M365 token response");
  }
  const claims = decodeM365Jwt(accessToken);
  const audience = String(claims.aud || "").replace(/\/$/, "");
  if (audience !== "https://substrate.office.com/sydney") {
    throw new Error(
      "unexpected Sydney token audience: " + (claims.aud || "missing"),
    );
  }
  if (
    !claims.oid ||
    !claims.tid ||
    String(claims.tid).toLowerCase() !== tenantId.toLowerCase()
  ) {
    throw new Error("Sydney token tenant/identity mismatch");
  }
  m365LatestTokenResponseAt = Math.max(m365LatestTokenResponseAt, capturedAt);
  _tid = tenantId;
  _clientId = clientId;
  _rt = refreshToken;
  const jwtExp = Number(claims.exp || 0) * 1000;
  const expiresIn = Number(message.expiresIn || 0);
  const responseExp =
    expiresIn > 0 ? Date.now() + Math.max(0, expiresIn - 60) * 1000 : 0;
  _sydney = { token: accessToken, exp: jwtExp || responseExp };
  await saveM365Auth();
  log("captured rotated Sydney refresh token and access token");
}

async function refreshSydney() {
  await ensureM365AuthLoaded();
  if (_sydney.token && _sydney.exp - Date.now() > 90000) return _sydney.token;
  if (_refreshing) return _refreshing;
  const missing = [];
  if (!_clientId) missing.push("client_id");
  if (!_rt) missing.push("refresh_token");
  if (!_tid) missing.push("tenant");
  if (missing.length)
    throw new Error("not captured yet: " + missing.join(", "));
  _refreshing = (async () => {
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
    body.set(
      "scope",
      String(capturedBody.scope || "").trim() ||
        M365_SYDNEY_SCOPE + " openid profile offline_access",
    );

    const response = await fetch(endpoint.href, {
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
    if (audience !== "https://substrate.office.com/sydney") {
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
    if (payload.refresh_token) {
      _rt = payload.refresh_token;
      await saveM365Auth();
    }
    _sydney = {
      token: payload.access_token,
      exp:
        Date.now() +
        Math.max(0, Number(payload.expires_in || 3600) - 60) * 1000,
    };
    m365LastRefreshError = "";
    m365LastRefreshErrorAt = 0;
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

// Mint a Teams/IC3-audience token for the AMS artifact endpoint. Uses the SAME
// captured FOCI refresh token as Sydney/SharePoint (same token endpoint, only
// the scope differs); FOCI lets one family refresh token mint tokens for
// sibling resources, exactly as refreshSharePointAccessToken does for the
// SharePoint audience. Cached with a 90s skew like the others.
async function refreshIc3Token() {
  await ensureM365AuthLoaded();
  if (_ic3.token && _ic3.exp - Date.now() > 90000) return _ic3.token;
  if (_ic3Refreshing) return _ic3Refreshing;
  if (!_clientId || !_rt || !_tid)
    throw new Error(
      "IC3 token unavailable: M365 refresh credentials not captured",
    );
  _ic3Refreshing = (async () => {
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
    const response = await fetch(endpoint.href, {
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
    if (payload.refresh_token) {
      _rt = payload.refresh_token;
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
    return _ic3.token;
  })();
  try {
    return await _ic3Refreshing;
  } finally {
    _ic3Refreshing = null;
  }
}

async function refreshSharePointAccessToken(siteUrl) {
  await ensureM365AuthLoaded();
  const origin = new URL(String(siteUrl || "")).origin;
  if (!origin.endsWith(".sharepoint.com"))
    throw new Error("invalid SharePoint token origin");
  if (
    _sharePoint.token &&
    _sharePoint.origin === origin &&
    _sharePoint.exp - Date.now() > 90000
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
    const response = await fetch(endpoint.href, {
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
    if (payload.refresh_token) {
      _rt = payload.refresh_token;
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
    return "Bearer " + _sharePoint.token;
  })();
  try {
    return await _sharePointRefreshing;
  } finally {
    _sharePointRefreshing = null;
  }
}
setInterval(
  () => {
    if (m365AuthLoaded && _clientId && _rt && _tid)
      refreshSydney().catch((error) =>
        log("scheduled Sydney refresh failed:", compactError(error)),
      );
  },
  25 * 60 * 1000,
);

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
    sendM365({ type: "M365_READY" });
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
    if (message.type === "M365_STOP") {
      dlog(
        "M365_STOP recv id=%s reason=%s",
        String(message.id || ""),
        String(message.reason || ""),
      );
      routeM365Stop(message.id);
    }
  });
  m365Ws.addEventListener("close", () => {
    m365Ws = null;
    scheduleM365Reconnect();
  });
  m365Ws.addEventListener("error", () => {
    try {
      m365Ws.close();
    } catch (_) {}
  });
  return m365Ws;
}

async function ensureM365Tab() {
  const tabs = await browser.tabs.query({
    url: [
      "https://outlook.cloud.microsoft/*",
      "https://outlook.office.com/*",
      "https://m365copilotapp.svc.cloud.microsoft/*",
      "https://m365.cloud.microsoft/*",
    ],
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
  const sharePointAccessToken =
    await refreshSharePointAccessToken(sharePointHomeUrl);
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
  const response = await completion;
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
// still succeeding. Floor 60s (unchanged for the common 1-file case), +30s per
// file, capped at 5min so a stuck batch still fails in bounded time.
function sharePointUploadTimeoutMs(fileCount) {
  const count = Math.max(1, Number(fileCount) || 1);
  return Math.min(300000, Math.max(60000, count * 30000));
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
  const sharePointAccessToken =
    await refreshSharePointAccessToken(sharePointHomeUrl);
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
      images: [], // {name, mimeType, dataUrl} 于 harvest 时同步记录，独立于 SP 上传
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
    const name = String(img.name || "").trim();
    const cap = name || "image";
    const d = name ? descByName.get(name) : null;
    if (d) usedNames.add(name);
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
    const name = String((d && d.name) || "").trim();
    if (name && usedNames.has(name)) continue; // 已作为内联图片处理
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
      try {
        await applyM365TokenResponse(message);
      } catch (error) {
        log("failed to apply M365 token response:", compactError(error));
      }
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
        dlog("IC3 token acquisition failed:", compactError(error));
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
            const list = Array.isArray(uploaded) ? uploaded : [];
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
      sendM365({ type: "M365_PROGRESS", id: message.id, text: "" });
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
  for (const [key, frame] of m365FrameCandidates) {
    if (frame.tabId === tabId) m365FrameCandidates.delete(key);
  }
  if (m365TargetFrame && m365TargetFrame.tabId === tabId)
    m365TargetFrame = null;
});

Promise.all([ensureM365AuthLoaded(), ensureM365SettingsLoaded()])
  .catch((error) => log("M365 initialization failed:", compactError(error)))
  .finally(m365Connect);
setInterval(() => {
  if (!m365Ws) m365Connect();
}, 5000);
