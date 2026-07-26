"use strict";

const OPEN_URLS = {
  deepseek: "https://chat.deepseek.com/",
  qwen: "https://chat.qwen.ai/",
  doubao: "https://www.doubao.com/chat/"
};

const $ = (id) => document.getElementById(id);

function shortId(value, keep = 8) {
  const s = String(value || "");
  if (!s) return "";
  return s.length <= keep * 2 + 1 ? s : `${s.slice(0, keep)}…${s.slice(-4)}`;
}

function agoText(ts) {
  const t = Number(ts || 0);
  if (!t) return "";
  const delta = Date.now() - t;
  if (delta < 0) return "刚刚";
  const s = Math.round(delta / 1000);
  if (s < 60) return `${s} 秒前`;
  const m = Math.round(s / 60);
  if (m < 60) return `${m} 分钟前`;
  const h = Math.round(m / 60);
  if (h < 24) return `${h} 小时前`;
  return new Date(t).toLocaleString();
}

function setStatus(el, text, cls) {
  el.textContent = text;
  el.className = "ve-status" + (cls ? " " + cls : "");
}

// Set a <dd> value and hide its whole row when the value is empty.
function setRow(rowId, ddId, value) {
  const dd = $(ddId);
  const row = $(rowId);
  const text = String(value == null ? "" : value).trim();
  if (dd) dd.textContent = text || "—";
  if (row) row.hidden = !text;
}

function authText(auth) {
  if (!auth) return "";
  const parts = [];
  const count = Number(auth.cookie_count || 0);
  if (count) parts.push(`${count} Cookie`);
  if (auth.has_authorization) parts.push("Bearer ✓");
  if (auth.has_sessionid) parts.push("sessionid ✓");
  if (Array.isArray(auth.cookie_names) && auth.cookie_names.includes("x5sec")) {
    parts.push("x5sec ✓");
  }
  return parts.join(" · ");
}

// Enable a "新建对话" button only when there is a stored conversation to clear.
function setNewButton(provider, hasConversation) {
  const btn = document.querySelector(`.ve-new[data-new="${provider}"]`);
  if (btn) btn.disabled = !hasConversation;
}

function render(state) {
  if (!state || !state.ok) {
    $("native-text").textContent = "无法读取后台状态";
    $("native-dot").className = "ve-dot err";
    return;
  }

  const p = state.persistent || null;
  const pp = (p && p.providers) || {};

  // Native host
  const connected = state.native && state.native.connected;
  $("native-dot").className = "ve-dot " + (connected ? "ok" : "err");
  $("native-text").textContent = connected
    ? `Native Host 已连接 · ${state.native.host}`
    : "Native Host 未连接（无法读取持久化状态）";
  $("foot-path").textContent = p && p.state_path
    ? `状态文件：${p.state_path}` +
      (p.updated_at ? ` · 更新于 ${agoText(p.updated_at * 1000)}` : "")
    : "";

  // ---------------- DeepSeek ----------------
  const ds = state.deepseek || {};
  setStatus(
    $("ds-status"),
    ds.pending ? "等待登录/同步" : ds.authorizationCaptured ? "凭证已捕获" : "空闲",
    ds.pending ? "warn" : ds.authorizationCaptured ? "ok" : ""
  );
  $("ds-auth").textContent = ds.authorizationCaptured ? "已捕获 (Bearer)" : "未捕获";
  $("ds-hif").textContent = ds.hifCaptured ? "已捕获" : "未捕获";
  const dss = ds.lastSnapshot;
  setRow(
    "ds-conv-row",
    "ds-conv",
    dss ? `${shortId(dss.conversation_id)}｜${dss.model_type || "?"}｜${agoText(dss.captured_at)}` : ""
  );
  const dsp = pp.deepseek || {};
  const sbm = dsp.sessions_by_mode || {};
  setRow("ds-p-default-row", "ds-p-default", shortId(sbm.default));
  setRow("ds-p-expert-row", "ds-p-expert", shortId(sbm.expert));
  setRow("ds-p-vision-row", "ds-p-vision", shortId(sbm.vision));
  setRow("ds-p-count-row", "ds-p-count", p ? String(dsp.known_conversation_count || 0) : "");
  setRow("ds-p-auth-row", "ds-p-auth", p ? authText(dsp.auth) : "");
  setNewButton("deepseek", Boolean(sbm.default || sbm.expert || sbm.vision));

  // ---------------- Qwen ----------------
  const qw = state.qwen || {};
  setStatus(
    $("qw-status"),
    qw.syncInFlight ? "同步中" : qw.pending ? "等待验证" : "空闲",
    qw.syncInFlight ? "warn" : qw.pending ? "warn" : ""
  );
  setRow("qw-host-row", "qw-host", qw.verificationHost);
  const qwp = pp.qwen || {};
  setRow("qw-p-conv-row", "qw-p-conv", p ? shortId(qwp.conversation_id) : "");
  setRow("qw-p-acct-row", "qw-p-acct", p && qwp.auth ? shortId(qwp.auth.account_id) : "");
  setRow("qw-p-auth-row", "qw-p-auth", p ? authText(qwp.auth) : "");
  setNewButton("qwen", Boolean(qwp.conversation_id));

  // ---------------- Doubao ----------------
  const db = state.doubao || {};
  let dbCls = "";
  let dbText = "空闲";
  if (db.authRequired) { dbText = "需要刷新状态"; dbCls = "err"; }
  if (db.syncInFlight) { dbText = "同步中"; dbCls = "warn"; }
  setStatus($("db-status"), dbText, dbCls);

  const cap = db.lastCapture;
  $("db-live-block").hidden = !cap;
  if (cap) {
    $("db-conv").textContent = `${shortId(cap.conversation_id)}｜${agoText(cap.completedAt)}`;
    $("db-index").textContent = String(cap.last_message_index);
    $("db-complete").textContent = cap.error
      ? `错误：${cap.error}`
      : cap.complete ? "是" : "否";
  }
  const dbp = pp.doubao || {};
  setRow("db-p-conv-row", "db-p-conv", p ? shortId(dbp.conversation_id) : "");
  setRow("db-p-auth-row", "db-p-auth", p ? authText(dbp.auth) : "");
  setNewButton("doubao", Boolean(dbp.conversation_id));

  const canSync = Boolean(cap && cap.complete && !cap.error) && !db.syncInFlight;
  const btn = $("sync-doubao");
  btn.hidden = !cap;
  btn.disabled = !canSync;
  btn.textContent = db.syncInFlight ? "同步中…" : "立即同步豆包状态";
}

async function refresh() {
  try {
    const state = await browser.runtime.sendMessage({ type: "POPUP_GET_STATE" });
    render(state);
  } catch (e) {
    $("native-text").textContent = "读取后台失败：" + ((e && e.message) || e);
    $("native-dot").className = "ve-dot err";
  }
}

function setHint(text, cls) {
  const el = $("db-hint");
  el.textContent = text || "";
  el.className = "ve-hint" + (cls ? " " + cls : "");
}

async function syncDoubao() {
  setHint("正在同步…");
  $("sync-doubao").disabled = true;
  try {
    const res = await browser.runtime.sendMessage({ type: "POPUP_SYNC_DOUBAO" });
    if (res && res.ok) setHint("已提交同步请求。", "ok");
    else setHint(res && res.error ? res.error : "同步失败", "err");
  } catch (e) {
    setHint("同步失败：" + ((e && e.message) || e), "err");
  }
  setTimeout(refresh, 400);
}

async function newConversation(provider) {
  const btn = document.querySelector(`.ve-new[data-new="${provider}"]`);
  const original = btn ? btn.textContent : "";
  if (btn) { btn.disabled = true; btn.textContent = "处理中…"; }
  try {
    const res = await browser.runtime.sendMessage({
      type: "POPUP_MUTATE",
      action: "new",
      provider
    });
    if (btn) btn.textContent = res && res.ok ? "已清除" : "失败";
  } catch (_) {
    if (btn) btn.textContent = "失败";
  }
  // Refresh twice: the STATE_REPORT reply is async over native messaging.
  setTimeout(refresh, 300);
  setTimeout(() => {
    if (btn) btn.textContent = original || "新建对话";
    refresh();
  }, 1200);
}

async function reconnect() {
  try { await browser.runtime.sendMessage({ type: "POPUP_RECONNECT" }); } catch (_) {}
  setTimeout(refresh, 500);
}

async function openProvider(provider) {
  try {
    await browser.runtime.sendMessage({ type: "POPUP_OPEN", provider });
  } catch (_) {
    const url = OPEN_URLS[provider];
    if (url) browser.tabs.create({ url });
  }
}

document.addEventListener("DOMContentLoaded", () => {
  $("refresh").addEventListener("click", refresh);
  $("reconnect").addEventListener("click", reconnect);
  $("sync-doubao").addEventListener("click", syncDoubao);
  for (const b of document.querySelectorAll(".ve-open")) {
    b.addEventListener("click", () => openProvider(b.getAttribute("data-open")));
  }
  for (const b of document.querySelectorAll(".ve-new")) {
    b.addEventListener("click", () => newConversation(b.getAttribute("data-new")));
  }
  refresh();
  const timer = setInterval(refresh, 2000);
  window.addEventListener("unload", () => clearInterval(timer));
});
