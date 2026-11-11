"use strict";

const OPEN_URLS = {
  deepseek: "https://chat.deepseek.com/",
  qwen: "https://chat.qwen.ai/",
  doubao: "https://www.doubao.com/chat/",
  m365: "https://m365.cloud.microsoft/",
  aistudio: "https://aistudio.google.com/prompts/new_chat?model=gemini-4-argon",
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
  $("foot-path").textContent =
    p && p.state_path
      ? `状态文件：${p.state_path}` +
        (p.updated_at ? ` · 更新于 ${agoText(p.updated_at * 1000)}` : "")
      : "";

  // ---------------- DeepSeek ----------------
  const ds = state.deepseek || {};
  setStatus(
    $("ds-status"),
    ds.pending
      ? "等待登录/同步"
      : ds.authorizationCaptured
        ? "凭证已捕获"
        : "空闲",
    ds.pending ? "warn" : ds.authorizationCaptured ? "ok" : "",
  );
  $("ds-auth").textContent = ds.authorizationCaptured
    ? "已捕获 (Bearer)"
    : "未捕获";
  $("ds-hif").textContent = ds.hifCaptured ? "已捕获" : "未捕获";
  const dss = ds.lastSnapshot;
  setRow(
    "ds-conv-row",
    "ds-conv",
    dss
      ? `${shortId(dss.conversation_id)}｜${dss.model_type || "?"}｜${agoText(dss.captured_at)}`
      : "",
  );
  const dsp = pp.deepseek || {};
  const sbm = dsp.sessions_by_mode || {};
  setRow("ds-p-default-row", "ds-p-default", shortId(sbm.default));
  setRow("ds-p-expert-row", "ds-p-expert", shortId(sbm.expert));
  setRow("ds-p-vision-row", "ds-p-vision", shortId(sbm.vision));
  setRow(
    "ds-p-count-row",
    "ds-p-count",
    p ? String(dsp.known_conversation_count || 0) : "",
  );
  setRow("ds-p-auth-row", "ds-p-auth", p ? authText(dsp.auth) : "");
  setNewButton("deepseek", Boolean(sbm.default || sbm.expert || sbm.vision));

  // ---------------- Qwen ----------------
  const qw = state.qwen || {};
  setStatus(
    $("qw-status"),
    qw.syncInFlight ? "同步中" : qw.pending ? "等待验证" : "空闲",
    qw.syncInFlight ? "warn" : qw.pending ? "warn" : "",
  );
  setRow("qw-host-row", "qw-host", qw.verificationHost);
  const qwp = pp.qwen || {};
  setRow("qw-p-conv-row", "qw-p-conv", p ? shortId(qwp.conversation_id) : "");
  setRow(
    "qw-p-acct-row",
    "qw-p-acct",
    p && qwp.auth ? shortId(qwp.auth.account_id) : "",
  );
  setRow("qw-p-auth-row", "qw-p-auth", p ? authText(qwp.auth) : "");
  setNewButton("qwen", Boolean(qwp.conversation_id));

  // ---------------- Doubao ----------------
  const db = state.doubao || {};
  let dbCls = "";
  let dbText = "空闲";
  if (db.authRequired) {
    dbText = "需要刷新状态";
    dbCls = "err";
  }
  if (db.syncInFlight) {
    dbText = "同步中";
    dbCls = "warn";
  }
  setStatus($("db-status"), dbText, dbCls);

  const cap = db.lastCapture;
  $("db-live-block").hidden = !cap;
  if (cap) {
    $("db-conv").textContent =
      `${shortId(cap.conversation_id)}｜${agoText(cap.completedAt)}`;
    $("db-index").textContent = String(cap.last_message_index);
    $("db-complete").textContent = cap.error
      ? `错误：${cap.error}`
      : cap.complete
        ? "是"
        : "否";
  }
  const dbp = pp.doubao || {};
  setRow("db-p-conv-row", "db-p-conv", p ? shortId(dbp.conversation_id) : "");
  setRow("db-p-auth-row", "db-p-auth", p ? authText(dbp.auth) : "");
  setNewButton("doubao", Boolean(dbp.conversation_id));

  // ---------------- SharePoint Online ----------------
  const sp = state.sharepoint || {};
  const spr = sp.lastResult || null;
  // "页面已就绪" is a healthy state, so it uses the same green as M365's
  // "已就绪": the SP tab is open and keeping the session alive, even before the
  // first upload proves it with a result.
  setStatus(
    $("sp-status"),
    !sp.configured
      ? "未配置"
      : spr && spr.ok
        ? "已连接"
        : sp.readyPageUrl
          ? "页面已就绪"
          : "等待页面",
    !sp.configured ? "err" : (spr && spr.ok) || sp.readyPageUrl ? "ok" : "warn",
  );
  setRow("sp-home-row", "sp-home", sp.homeUrl || "");
  setRow("sp-folder-row", "sp-folder", sp.uploadFolder || "");
  setRow(
    "sp-result-row",
    "sp-result",
    spr
      ? spr.error || spr.serverRelativeUrl || spr.title || "成功"
      : sp.lastKeepAliveError
        ? `保活异常 (${sp.lastKeepAliveError})｜${agoText(sp.lastKeepAliveErrorAt)}`
        : sp.lastKeepAliveAt
          ? `自动续期于 ${agoText(sp.lastKeepAliveAt)}`
          : "",
  );
  $("sp-test").disabled = !sp.configured;
  $("sp-upload-test").disabled = !sp.configured;

  // ---------------- M365 Copilot ----------------
  const m365 = state.m365 || {};
  const m365p = pp.m365 || {};
  const m365Ids = m365p.conversations_by_tone || {};
  const m365Ready = Boolean(
    m365.bridgeConnected && m365.authAvailable && !m365.reauthRequired,
  );
  let m365Text = "等待 M365 页面与认证";
  let m365Class = "warn";
  if (m365.reauthRequired) {
    // Single-shot policy: exactly one automatic attempt per episode, then stop.
    m365Text =
      m365.reauthState === "attempting"
        ? "正在自动登录（单次尝试）"
        : m365.reauthState === "terminal"
          ? "自动登录已失败（不重试）"
          : "需重新登录 (24h 令牌已到期)";
    m365Class = "err";
  } else if (m365Ready) {
    m365Text = "已就绪";
    m365Class = "ok";
  } else if (!m365.bridgeConnected) {
    m365Text = "Bridge 未连接";
    m365Class = "err";
  } else if (!m365.authAvailable) {
    m365Text = "等待登录凭证";
    m365Class = "warn";
  }
  setStatus($("m365-status"), m365Text, m365Class);
  setRow("m365-tenant-row", "m365-tenant", m365.tenantId);
  setRow("m365-client-row", "m365-client", m365.clientId);
  const m365TimeParts = [];
  if (m365.authUpdatedAt) m365TimeParts.push(agoText(m365.authUpdatedAt));
  if (m365.lastKeepAliveAt)
    m365TimeParts.push(`保活 ${agoText(m365.lastKeepAliveAt)}`);
  setRow("m365-auth-time-row", "m365-auth-time", m365TimeParts.join(" · "));
  const m365ErrText = m365.autoLoginLastError
    ? `${m365.autoLoginLastError}｜${agoText(m365.autoLoginLastErrorAt)}`
    : m365.reauthState === "terminal"
      ? `自动登录失败：${m365.reauthStateReason || "-"}（单次尝试，不再自动重试）`
      : m365.reauthRequired
        ? `${m365.reauthReason || "24h SPA 刷新令牌已到期，需重新登录"}｜${agoText(m365.reauthRequiredAt)}`
        : m365.lastRefreshError
          ? `${m365.lastRefreshError}｜${agoText(m365.lastRefreshErrorAt)}`
          : m365.lastKeepAliveError
            ? `保活异常 (${m365.lastKeepAliveError})｜${agoText(m365.lastKeepAliveErrorAt)}`
            : "";
  setRow("m365-refresh-error-row", "m365-refresh-error", m365ErrText);
  const enabledChk = $("m365-autologin-enabled");
  const rememberChk = $("m365-autologin-remember");
  if (enabledChk) enabledChk.checked = m365.autoLoginEnabled !== false;
  if (rememberChk) {
    rememberChk.checked = Boolean(m365.autoLoginRemember);
    rememberChk.disabled = m365.autoLoginEnabled === false;
  }
  let autoLoginText = "未保存密码";
  if (m365.autoLoginEnabled === false) {
    autoLoginText = "已关闭";
  } else if (m365.autoLoginPaused) {
    autoLoginText = "已暂停（登录失败或需验证，输入新密码可解除）";
  } else if (m365.autoLoginConfigured) {
    autoLoginText =
      "已启用" +
      (m365.autoLoginUsername ? ` (${m365.autoLoginUsername})` : "") +
      (m365.autoLoginLastSuccessAt
        ? ` · 上次成功 ${agoText(m365.autoLoginLastSuccessAt)}`
        : "");
  } else if (m365.autoLoginRemember) {
    autoLoginText = "已开启网页登录自动记忆（等待下次登录）";
  }
  setRow("m365-autologin-row", "m365-autologin-status", autoLoginText);
  const clearAutoBtn = $("m365-autologin-clear");
  if (clearAutoBtn) {
    clearAutoBtn.hidden = !m365.autoLoginConfigured && !m365.autoLoginPaused;
  }
  const passInput = $("m365-autologin-pass");
  if (passInput && document.activeElement !== passInput) {
    passInput.placeholder = m365.autoLoginPaused
      ? "更新 M365 密码并保存以解除暂停"
      : m365.autoLoginConfigured
        ? "已保存自动重登密码（输入可覆盖更新）"
        : "输入 M365 密码";
  }
  setRow("m365-opus-row", "m365-opus", shortId(m365Ids.Claude_Opus));
  setRow("m365-sonnet-row", "m365-sonnet", shortId(m365Ids.Claude_Sonnet));
  setRow("m365-gpt61-row", "m365-gpt61", shortId(m365Ids.Gpt_61_Sol_Reasoning));
  setRow(
    "m365-gpt6-chat-row",
    "m365-gpt6-chat",
    shortId(m365Ids.Gpt_6_Sol_Chat),
  );
  setRow("m365-gpt56-row", "m365-gpt56", shortId(m365Ids.Gpt_5_6_Reasoning));
  setRow(
    "m365-gpt56-chat-row",
    "m365-gpt56-chat",
    shortId(m365Ids.Gpt_5_6_Chat),
  );
  const anyM365Conversation = Boolean(
    m365Ids.Claude_Opus ||
    m365Ids.Claude_Sonnet ||
    m365Ids.Gpt_61_Sol_Reasoning ||
    m365Ids.Gpt_6_Sol_Chat ||
    m365Ids.Gpt_5_6_Reasoning ||
    m365Ids.Gpt_5_6_Chat,
  );
  setNewButton("m365", anyM365Conversation);
  $("m365-clear-auth").disabled = !m365.authAvailable;

  // ---------------- Google AI Studio ----------------
  const ais = state.aistudio || {};
  const aisp = pp.aistudio || {};
  const aisBridgeOk = Boolean(ais.bridgeConnected || aisp.bridge_connected);
  const aisAuthOk = Boolean(ais.authOk || aisp.auth_ok);
  const aisService = String(ais.service || aisp.service || "pending");
  const aisDriveStatus =
    ais.driveStatus && ais.driveStatus !== "unknown"
      ? ais.driveStatus
      : String(aisp.drive_status || "unknown");
  const aisDriveErr = String(ais.driveError || aisp.drive_error || "");

  let aisStatusText = "等待页面与登录";
  let aisStatusCls = "warn";
  if (!aisBridgeOk) {
    aisStatusText = "Bridge 未连接";
    aisStatusCls = "err";
  } else if (!aisAuthOk) {
    aisStatusText = ais.hookReady ? "需登录 Google 账号" : "等待打开页面";
    aisStatusCls = "warn";
  } else if (aisDriveStatus === "unauthorized") {
    aisStatusText =
      aisService === "captured" ? "已就绪 (Drive 未授权)" : "Drive 未授权";
    aisStatusCls = "warn";
  } else if (aisService === "captured") {
    aisStatusText = "已就绪";
    aisStatusCls = "ok";
  } else {
    aisStatusText = "已连接 (待首轮自举)";
    aisStatusCls = "ok";
  }
  setStatus($("ais-status"), aisStatusText, aisStatusCls);

  const aisAcctParts = [];
  const aisUser = String(ais.authUser || aisp.auth_user || "");
  if (aisAuthOk && aisUser !== "") {
    aisAcctParts.push(`/u/${aisUser}`);
  }
  const aisTier = String(ais.tier || aisp.tier || "");
  if (aisTier) {
    const tierLabel =
      aisTier === "TIER1"
        ? "TIER1 (Pro)"
        : aisTier === "TIER2"
          ? "TIER2 (Ultra)"
          : aisTier === "TIER0"
            ? "TIER0 (Plus)"
            : aisTier;
    aisAcctParts.push(tierLabel);
  }
  const aisTz = String(ais.timezone || aisp.timezone || "");
  if (aisTz) aisAcctParts.push(aisTz);
  setRow("ais-account-row", "ais-account", aisAcctParts.join(" · "));

  const serviceEl = $("ais-service");
  if (serviceEl) {
    serviceEl.textContent =
      aisService === "captured"
        ? "已捕获 (BotGuard Ready ✓)"
        : aisAuthOk
          ? "待捕获 (首轮请求自动静默自举)"
          : "未就绪";
  }

  const driveEl = $("ais-drive");
  if (driveEl) {
    if (aisDriveStatus === "authorized") {
      driveEl.textContent = "已授权 ✓ (可同步 Playground History)";
      driveEl.style.color = "var(--ok)";
    } else if (aisDriveStatus === "unauthorized") {
      driveEl.textContent =
        "未授权 ✗ (请在 AI Studio 页面开启 Drive 保存)" +
        (aisDriveErr ? ` · ${aisDriveErr}` : "");
      driveEl.style.color = "var(--warn)";
    } else {
      driveEl.textContent = aisAuthOk
        ? "未检测 (点击下方按钮检测)"
        : "等待登录";
      driveEl.style.color = "";
    }
  }

  const rawTabUrl = String(ais.tabUrl || aisp.tab_url || "");
  let shortTab = "";
  if (rawTabUrl) {
    try {
      const uObj = new URL(rawTabUrl);
      shortTab = uObj.pathname + (uObj.search ? uObj.search.slice(0, 36) : "");
    } catch (_) {
      shortTab = rawTabUrl.slice(0, 48);
    }
  }
  setRow("ais-tab-row", "ais-tab", shortTab);

  const aisErrObj = ais.lastError;
  setRow(
    "ais-error-row",
    "ais-error",
    aisErrObj && aisErrObj.message
      ? `${aisErrObj.message}${aisErrObj.at ? "｜" + agoText(aisErrObj.at) : ""}`
      : "",
  );

  const lastPid = String(aisp.last_prompt_id || "");
  const lastTitle = String(aisp.last_title || "");
  const sessCount = Number(aisp.session_count || 0);
  const dynModels = Number(aisp.dynamic_models_count || 0);
  setRow(
    "ais-p-prompt-row",
    "ais-p-prompt",
    lastPid ? shortId(lastPid.replace(/^prompts\//, ""), 10) : "",
  );
  setRow("ais-p-title-row", "ais-p-title", lastTitle);
  const metaParts = [];
  if (sessCount > 0) metaParts.push(`${sessCount} 活跃会话`);
  if (dynModels > 0) metaParts.push(`${dynModels} 动态模型`);
  setRow("ais-p-meta-row", "ais-p-meta", metaParts.join(" · "));
  const aisNewPlaygroundBtn = $("ais-new-playground");
  if (aisNewPlaygroundBtn) {
    aisNewPlaygroundBtn.disabled = !Boolean(
      sessCount > 0 || lastPid || lastTitle,
    );
  }

  const lastBuildApplet = String(
    aisp.last_build_applet_id || aisp.last_applet_id || "",
  );
  const lastBuildTitle = String(aisp.last_build_title || "");
  const buildSessCount = Number(aisp.build_session_count || 0);
  const buildAppletCount = Number(aisp.build_applet_count || 0);
  setRow(
    "ais-b-applet-row",
    "ais-b-applet",
    lastBuildApplet
      ? shortId(lastBuildApplet.replace(/^applets\//, ""), 10)
      : "",
  );
  setRow("ais-b-title-row", "ais-b-title", lastBuildTitle);
  const buildMetaParts = [];
  if (buildSessCount > 0) buildMetaParts.push(`${buildSessCount} 活跃会话`);
  if (buildAppletCount > 0)
    buildMetaParts.push(`${buildAppletCount} 驻留 Applet`);
  setRow("ais-b-meta-row", "ais-b-meta", buildMetaParts.join(" · "));
  const aisNewBuildBtn = $("ais-new-build");
  if (aisNewBuildBtn) {
    aisNewBuildBtn.disabled = !Boolean(
      buildSessCount > 0 ||
      buildAppletCount > 0 ||
      lastBuildApplet ||
      lastBuildTitle,
    );
  }

  const hasUsableCapture = Boolean(cap && cap.complete && !cap.error);
  const captureApplied = hasUsableCapture && db.lastCaptureApplied;
  const btn = $("sync-doubao");
  btn.hidden = !cap;
  btn.disabled = !hasUsableCapture || db.syncInFlight || captureApplied;
  if (db.syncInFlight) {
    btn.textContent = "同步中…";
  } else if (captureApplied) {
    btn.textContent = "豆包状态已同步";
  } else if (db.authRequired) {
    btn.textContent = "立即同步豆包状态";
  } else {
    btn.textContent = "重新发送最近豆包快照";
  }
}

async function refresh() {
  try {
    const state = await browser.runtime.sendMessage({
      type: "POPUP_GET_STATE",
    });
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
    const res = await browser.runtime.sendMessage({
      type: "POPUP_SYNC_DOUBAO",
    });
    if (res && res.ok) setHint("已提交同步请求。", "ok");
    else setHint(res && res.error ? res.error : "同步失败", "err");
  } catch (e) {
    setHint("同步失败：" + ((e && e.message) || e), "err");
  }
  setTimeout(refresh, 400);
}

async function newConversation(provider, modelType, button) {
  const btn =
    button || document.querySelector(`.ve-new[data-new="${provider}"]`);
  const original = btn ? btn.textContent : "";
  if (btn) {
    btn.disabled = true;
    btn.textContent = "处理中…";
  }
  try {
    const res = await browser.runtime.sendMessage({
      type: "POPUP_MUTATE",
      action: "new",
      provider,
      model_type: modelType || "",
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

async function clearM365Auth() {
  const btn = $("m365-clear-auth");
  const original = btn.textContent;
  btn.disabled = true;
  btn.textContent = "处理中…";
  try {
    const res = await browser.runtime.sendMessage({
      type: "POPUP_CLEAR_M365_AUTH",
    });
    btn.textContent = res && res.ok ? "已清除" : "失败";
  } catch (_) {
    btn.textContent = "失败";
  }
  setTimeout(() => {
    btn.textContent = original;
    refresh();
  }, 900);
}

async function saveM365AutoLogin() {
  const input = $("m365-autologin-pass");
  const btn = $("m365-autologin-save");
  const enabledChk = $("m365-autologin-enabled");
  const rememberChk = $("m365-autologin-remember");
  const pwd = String((input && input.value) || "");
  if (!pwd) return;
  const original = btn.textContent;
  btn.disabled = true;
  btn.textContent = "保存中…";
  try {
    const res = await browser.runtime.sendMessage({
      type: "POPUP_SAVE_M365_AUTO_LOGIN",
      enabled: enabledChk ? Boolean(enabledChk.checked) : true,
      rememberOnManualLogin: rememberChk ? Boolean(rememberChk.checked) : false,
      password: pwd,
      unpause: true,
    });
    if (res && res.ok !== false) {
      input.value = "";
      btn.textContent = "已保存";
      render(res);
    } else {
      btn.textContent = "失败";
    }
  } catch (_) {
    btn.textContent = "失败";
  }
  setTimeout(() => {
    btn.disabled = false;
    btn.textContent = original;
    refresh();
  }, 900);
}

async function updateM365AutoLoginToggles() {
  const enabledChk = $("m365-autologin-enabled");
  const rememberChk = $("m365-autologin-remember");
  try {
    const res = await browser.runtime.sendMessage({
      type: "POPUP_SAVE_M365_AUTO_LOGIN",
      enabled: enabledChk ? Boolean(enabledChk.checked) : true,
      rememberOnManualLogin: rememberChk ? Boolean(rememberChk.checked) : false,
      unpause: false,
    });
    if (res && res.ok !== false) {
      render(res);
    }
  } catch (_) {}
}

async function clearM365AutoLogin() {
  const btn = $("m365-autologin-clear");
  const original = btn.textContent;
  btn.disabled = true;
  btn.textContent = "处理中…";
  try {
    const res = await browser.runtime.sendMessage({
      type: "POPUP_CLEAR_M365_AUTO_LOGIN",
    });
    if (res && res.ok !== false) {
      render(res);
    }
  } catch (_) {}
  setTimeout(() => {
    btn.disabled = false;
    btn.textContent = original;
    refresh();
  }, 700);
}

async function sharePointAction(type) {
  const hint = $("sp-hint");
  hint.textContent = "处理中…";
  try {
    const result = await browser.runtime.sendMessage({ type });
    hint.textContent =
      result && result.ok
        ? result.serverRelativeUrl || result.title || "成功"
        : (result && result.error) || "失败";
    hint.className = "ve-hint " + (result && result.ok ? "ok" : "err");
  } catch (error) {
    hint.textContent = String((error && error.message) || error);
    hint.className = "ve-hint err";
  }
  setTimeout(refresh, 300);
}

async function probeAisStatus() {
  const btn = $("ais-probe");
  const hint = $("ais-hint");
  const original = btn ? btn.textContent : "";
  if (btn) {
    btn.disabled = true;
    btn.textContent = "正在检测…";
  }
  if (hint) {
    hint.textContent = "正在检测 AI Studio 页面与 Google Drive OAuth 状态…";
    hint.className = "ve-hint";
  }
  try {
    const res = await browser.runtime.sendMessage({
      type: "POPUP_AIS_PROBE",
      warmup: false,
    });
    if (res && res.ok && res.state) {
      render(res.state);
      const dSt =
        (res.state.aistudio && res.state.aistudio.driveStatus) || "unknown";
      if (hint) {
        if (dSt === "authorized") {
          hint.textContent =
            "Google Drive 已授权 ✓，Playground History 同步正常";
          hint.className = "ve-hint ok";
        } else if (dSt === "unauthorized") {
          hint.textContent =
            "当前 Google 账号未授权 Google Drive，请在 AI Studio 页面开启 Drive 保存";
          hint.className = "ve-hint err";
        } else {
          hint.textContent = "已刷新 AI Studio 页面状态";
          hint.className = "ve-hint ok";
        }
      }
    } else if (hint) {
      hint.textContent = (res && res.error) || "检测失败";
      hint.className = "ve-hint err";
    }
  } catch (err) {
    if (hint) {
      hint.textContent = String((err && err.message) || err);
      hint.className = "ve-hint err";
    }
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.textContent = original || "检测 Drive 授权与 WAA 状态";
    }
    setTimeout(refresh, 400);
  }
}

async function reconnect() {
  try {
    await browser.runtime.sendMessage({ type: "POPUP_RECONNECT" });
  } catch (_) {}
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
  $("m365-clear-auth").addEventListener("click", clearM365Auth);
  const saveBtn = $("m365-autologin-save");
  const clearBtn = $("m365-autologin-clear");
  const passInput = $("m365-autologin-pass");
  const enabledChk = $("m365-autologin-enabled");
  const rememberChk = $("m365-autologin-remember");
  if (saveBtn) saveBtn.addEventListener("click", saveM365AutoLogin);
  if (clearBtn) clearBtn.addEventListener("click", clearM365AutoLogin);
  if (enabledChk)
    enabledChk.addEventListener("change", updateM365AutoLoginToggles);
  if (rememberChk)
    rememberChk.addEventListener("change", updateM365AutoLoginToggles);
  if (passInput) {
    passInput.addEventListener("keydown", (e) => {
      if (e.key === "Enter") saveM365AutoLogin();
    });
  }
  $("sp-test").addEventListener("click", () =>
    sharePointAction("POPUP_SP_TEST"),
  );
  $("sp-upload-test").addEventListener("click", () =>
    sharePointAction("POPUP_SP_UPLOAD_TEST"),
  );
  const aisProbeBtn = $("ais-probe");
  if (aisProbeBtn) {
    aisProbeBtn.addEventListener("click", probeAisStatus);
  }
  for (const b of document.querySelectorAll(".ve-open")) {
    b.addEventListener("click", () =>
      openProvider(b.getAttribute("data-open")),
    );
  }
  for (const b of document.querySelectorAll(".ve-new")) {
    b.addEventListener("click", () =>
      newConversation(
        b.getAttribute("data-new"),
        b.getAttribute("data-model") || "",
        b,
      ),
    );
  }
  refresh();
  const timer = setInterval(refresh, 2000);
  window.addEventListener("unload", () => clearInterval(timer));
});
