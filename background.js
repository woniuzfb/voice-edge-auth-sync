"use strict";

const NATIVE_HOST = "com.voice_edge.auth_bridge";
const COMPLETION_FILTER = {
  urls: ["https://www.doubao.com/chat/completion*"]
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

function compactError(error) {
  return String(error && (error.message || error) || "unknown error")
    .replace(/\s+/g, " ")
    .slice(0, 500);
}

async function notify(title, message) {
  try {
    await browser.notifications.create({
      type: "basic",
      title,
      message
    });
  } catch (_) {
    // Notifications are best-effort only.
  }
}

async function setBadge(text, color) {
  try {
    await browser.browserAction.setBadgeText({text});
    if (color) {
      await browser.browserAction.setBadgeBackgroundColor({color});
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
    port.onDisconnect.addListener(() => {
      nativePort = null;
      scheduleReconnect();
    });
    port.postMessage({type: "PING"});
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
    expirationDate: cookie.expirationDate == null ? null : cookie.expirationDate,
    storeId: cookie.storeId || fallbackStoreId || "",
    firstPartyDomain: cookie.firstPartyDomain || null,
    partitionKey: cookie.partitionKey || null
  };
}

function cookieIdentity(cookie) {
  return `${cookie.name}\u0000${cookie.domain}\u0000${cookie.path || "/"}`;
}

function cookieFingerprint(cookies) {
  return (cookies || [])
    .map(cookie => `${cookieIdentity(cookie)}\u0000${cookie.value}\u0000${cookie.expirationDate || ""}`)
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
    url: ["https://chat.qwen.ai/*", "https://*.qwen.ai/*"]
  });
  const active = tabs.find(tab => tab.active) || tabs[0];
  return {
    tabId: active ? active.id : null,
    storeId: active ? (active.cookieStoreId || "") : ""
  };
}

async function collectQwenCookies(state) {
  const output = new Map();
  for (const domain of qwenDomains(state)) {
    const query = {domain};
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
  const options = {url: qwenState.verificationUrl, active: true};
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
    cookies
  });
}

async function pollQwenCookies() {
  if (!qwenState) return;
  if (Date.now() - qwenState.startedAt > QWEN_POLL_TIMEOUT_MS) {
    clearInterval(qwenPollTimer);
    qwenPollTimer = null;
    await setBadge("!", "#d93025");
    await notify("Voice Edge：Qwen 验证超时", "未检测到新的 x5sec，请重新触发验证后再试。");
    return;
  }
  const cookies = await collectQwenCookies(qwenState);
  const x5secCookies = cookies.filter(cookie => cookie.name === "x5sec" && cookie.value);
  const hasX5sec = x5secCookies.length > 0;
  const fingerprint = cookieFingerprint(cookies);
  const x5secFingerprint = cookieFingerprint(x5secCookies);
  if (
    !hasX5sec ||
    !fingerprint ||
    !x5secFingerprint ||
    x5secFingerprint === qwenState.baselineX5secFingerprint
  ) return;
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
    syncInFlight: false
  };
  if (!qwenState.accountId || !qwenState.verificationUrl) {
    throw new Error("Qwen 验证消息缺少 account_id 或 verification_url");
  }
  const baselineCookies = await collectQwenCookies(qwenState);
  qwenState.baselineFingerprint = cookieFingerprint(baselineCookies);
  qwenState.baselineX5secFingerprint = cookieFingerprint(
    baselineCookies.filter(cookie => cookie.name === "x5sec")
  );
  await setBadge("!", "#d93025");
  await browser.notifications.create(QWEN_NOTIFICATION_ID, {
    type: "basic",
    title: "Voice Edge：Qwen 需要验证",
    message: "点击通知打开验证页面。验证完成后，相关 Cookie 会自动同步。"
  });
  qwenPollTimer = setInterval(() => {
    pollQwenCookies().catch(error => notify("Voice Edge Qwen 同步失败", compactError(error)));
  }, QWEN_POLL_INTERVAL_MS);
}

browser.notifications.onClicked.addListener(async notificationId => {
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
        `已同步 ${message.cookie_count || 0} 个 Cookie，后续请求将使用最新状态。`
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
      "请在 Firefox 豆包中继续当前对话或新建对话，发送一条消息。网页完整回复后将自动同步，无需点击扩展。"
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
        `豆包状态已验证并自动同步（${message.cookie_count || 0} 个 Cookie）。${replay}`
      );
      setTimeout(() => setBadge("", null), 5000);
    } else {
      authRequired = true;
      await setBadge("!", "#d93025");
    }
    return;
  }

  if (type === "AUTH_ERROR") {
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
  const c = capture && capture.conversation || {};
  return `${c.conversation_id || ""}:${c.section_id || ""}:${c.last_message_index || 0}`;
}

function safeInt(value) {
  const number = Number(value);
  return Number.isFinite(number) && number >= 0 ? Math.trunc(number) : null;
}

function updateMax(state, value) {
  const parsed = safeInt(value);
  if (parsed !== null) state.lastMessageIndex = Math.max(state.lastMessageIndex, parsed);
}

function parseSseBody(body) {
  const state = {
    conversationId: "",
    sectionId: "",
    lastMessageIndex: 0,
    endType3: false,
    streamError: null
  };

  const events = String(body || "").split(/\r?\n\r?\n/);
  for (const block of events) {
    if (!block.trim()) continue;
    let eventName = "";
    const dataLines = [];
    for (const line of block.split(/\r?\n/)) {
      if (line.startsWith("event:")) eventName = line.slice(6).trim();
      else if (line.startsWith("data:")) dataLines.push(line.slice(5).trimStart());
    }
    if (!eventName || !dataLines.length) continue;
    let data;
    try {
      data = JSON.parse(dataLines.join("\n"));
    } catch (_) {
      continue;
    }

    if (eventName === "SSE_ACK") {
      const meta = data && data.ack_client_meta || {};
      if (meta.conversation_id) state.conversationId = String(meta.conversation_id);
      if (meta.section_id) state.sectionId = String(meta.section_id);
      const queries = Array.isArray(data.query_list) ? data.query_list : [];
      for (const query of queries) updateMax(state, query && query.message_index);
    } else if (eventName === "FULL_MSG_NOTIFY") {
      updateMax(state, data && data.message && data.message.index_in_conv);
    } else if (eventName === "STREAM_MSG_NOTIFY") {
      updateMax(state, data && data.meta && data.meta.index_in_conv);
    } else if (eventName === "SSE_REPLY_END") {
      if (safeInt(data && data.end_type) === 3) state.endType3 = true;
    } else if (eventName === "STREAM_ERROR") {
      state.streamError = {
        code: data && data.error_code,
        message: data && data.error_msg
      };
    }
  }
  return state;
}

function scheduleAutoSync(capture) {
  if (!authRequired || syncInFlight || !capture || !capture.complete || capture.error) return;
  if (
    !authRequiredSince ||
    capture.startedAt < authRequiredSince ||
    capture.completedAt < authRequiredSince
  ) return;
  const key = captureKey(capture);
  if (!key || key === lastAppliedKey) return;
  if (pendingAutoSyncTimer) clearTimeout(pendingAutoSyncTimer);
  pendingAutoSyncTimer = setTimeout(() => {
    pendingAutoSyncTimer = null;
    synchronizeCapture(capture, true).catch(async error => {
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
  const query = {domain: "doubao.com"};
  if (storeId) query.storeId = storeId;
  const cookies = await browser.cookies.getAll(query);
  return cookies.map(cookie => ({
    name: cookie.name,
    value: cookie.value,
    domain: cookie.domain,
    hostOnly: Boolean(cookie.hostOnly),
    path: cookie.path || "/",
    secure: Boolean(cookie.secure),
    httpOnly: Boolean(cookie.httpOnly),
    sameSite: cookie.sameSite || null,
    session: Boolean(cookie.session),
    expirationDate: cookie.expirationDate == null ? null : cookie.expirationDate,
    storeId: cookie.storeId || storeId || "",
    firstPartyDomain: cookie.firstPartyDomain || null,
    partitionKey: cookie.partitionKey || null
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
    partitionKey: null
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
  if (!cookies.some(item => item.name === "sessionid")) {
    syncInFlight = false;
    throw new Error("当前 Firefox Cookie Store 中没有豆包 sessionid");
  }

  // 强制把 ttwid 纳入同步快照（cookies API 缺失时从 document.cookie 兜底），
  // 并在消息顶层显式带上 ttwid，供 Voice Edge 强制覆盖回 Camoufox。
  const ttwidValue = await ensureTtwidCaptured(
    cookies,
    capture.tabId,
    capture.cookieStoreId || ""
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
    ttwid: ttwidValue || ""
  });
}

browser.webRequest.onBeforeRequest.addListener(
  details => {
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

    filter.ondata = event => {
      filter.write(event.data);
      if (overflow) return;
      capturedBytes += event.data.byteLength;
      if (capturedBytes > MAX_CAPTURE_BYTES) {
        overflow = true;
        body = "";
        return;
      }
      body += decoder.decode(event.data, {stream: true});
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
          !state.streamError
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
            last_message_index: state.lastMessageIndex
          }
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
        try { filter.close(); } catch (_) {}
      }
    };

    filter.onerror = () => {
      try { filter.close(); } catch (_) {}
    };
  },
  COMPLETION_FILTER,
  ["blocking"]
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
