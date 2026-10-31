#!/usr/bin/env node
/*
 * M365 自动登录 harness —— 单文件，跑 mock 登录窗口 + 真实发布代码
 *
 *   npm run test:reauth        (node test/reauth.mjs)
 *   npm run test:reauth -- -v  (额外打印后台日志)
 *
 * 被测代码 = 构建产物本身（不是 src 的另一份拷贝）：
 *   - dist/chrome/background.js         在 node:vm 里原样执行（内存版 browser.*）
 *   - dist/chrome/content-m365-login.js 在登录窗口里原样执行（jsdom 或真浏览器）
 *   启动时先比对 src 与 dist 的 sha256，不一致直接报错退出（免得测到旧产物）。
 *
 * 模拟的部分（本文件下半部分）：
 *   - AAD 登录窗口：/authorize、账号选择器、密码页 #i0118、错误页、MFA、KMSI
 *   - 浏览器密码库：立即填 / 晚填 / 不填（用户手输）
 *   - /oauth2/v2.0/token 端点
 *
 * 每个场景都会打印登录窗口的生命周期（[窗口 #1000] 开头的行），
 * 让你直接看到“窗口怎么开的、页面渲染成什么样、密码什么时候被填、谁点了提交”。
 */
import vm from "node:vm";
import http from "node:http";
import { spawn } from "node:child_process";
import { createHash } from "node:crypto";
import { readFileSync, existsSync, readdirSync } from "node:fs";
import { homedir, tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { JSDOM } from "jsdom";
import { WebSocketServer } from "ws";

const TENANT = "024ecf63-00bd-4137-a47c-82c2d93ddb62";
const CLIENT_ID = "c0ab8ce9-e9a0-42e7-b064-33d422df41f1";
const UPN = "harness@contoso.com";
const PASSWORD = "CorrectHorseBatteryStaple!";
const ROOT = fileURLToPath(new URL("../", import.meta.url));
const DIST = join(ROOT, "dist/chrome");
const BACKGROUND_FILE =
  process.env.VE_BACKGROUND || join(DIST, "background.js");
const CONTENT_FILE = join(DIST, "content-m365-login.js");
const CONTENT_SCRIPT = existsSync(CONTENT_FILE)
  ? readFileSync(CONTENT_FILE, "utf8")
  : (() => {
      throw new Error(
        "找不到 " +
          CONTENT_FILE +
          "，先跑 npm run build:chrome（harness 测的是构建产物）",
      );
    })();

function sha256(file) {
  return createHash("sha256")
    .update(readFileSync(file))
    .digest("hex")
    .slice(0, 12);
}

// 确保"测的产物"就是"现在要发的产物"：src 与 dist 必须逐字节一致。
function assertProductionBuild() {
  const pairs = [
    ["src/background.js", "dist/chrome/background.js"],
    ["src/content-m365-login.js", "dist/chrome/content-m365-login.js"],
  ];
  const problems = [];
  for (const [srcRel, distRel] of pairs) {
    const srcFile = join(ROOT, srcRel);
    const distFile = join(ROOT, distRel);
    if (!existsSync(distFile)) {
      problems.push(distRel + " 不存在（先 npm run build:chrome）");
      continue;
    }
    if (!existsSync(srcFile)) continue;
    if (sha256(srcFile) !== sha256(distFile)) {
      problems.push(
        distRel + " 与 " + srcRel + " 不一致（dist 是旧的，先重新构建）",
      );
    }
  }
  if (problems.length) {
    console.error("✗ 产物校验失败：");
    for (const p of problems) console.error("  - " + p);
    process.exit(3);
  }
  console.log(
    `产物校验: dist/chrome/background.js (${sha256(join(ROOT, "dist/chrome/background.js"))})  ` +
      `dist/chrome/content-m365-login.js (${sha256(CONTENT_FILE)})`,
  );
}
const REDIRECT_URI = "https://outlook.cloud.microsoft/mail/oauthRedirect.html";
// What OWA's own token request carries (registered value). The authorize request
// MUST reuse this, otherwise AAD answers AADSTS50011 and no password form shows.
const BROKER_REDIRECT = "brk-multihub://outlook.office.com";
const WEB_REDIRECT = "https://outlook.cloud.microsoft/mail/oauthRedirect.html";
const CAPTURED_REDIRECT = BROKER_REDIRECT;
const BRK_CLIENT_ID = "9199bf20-a13f-4107-85dc-02114787ef48";
const SEED_CONTEXT = {
  version: 1,
  query: {},
  body: {
    client_id: CLIENT_ID,
    redirect_uri: BROKER_REDIRECT,
    scope: "https://substrate.office.com/sydney/.default openid profile offline_access",
    brk_client_id: BRK_CLIENT_ID,
    brk_redirect_uri: WEB_REDIRECT,
    login_hint: "user@contoso.com",
  },
};
const VERBOSE = process.argv.includes("-v") || process.env.VE_VERBOSE === "1";

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const fmtArgs = (args) => {
  const [head, ...rest] = args;
  let i = 0;
  const render = (v) => (typeof v === "string" ? v : safeJson(v));
  return (
    String(head).replace(/%[sdj%]/g, (t) => {
      if (t === "%%") return "%";
      if (i >= rest.length) return t;
      const v = rest[i++];
      return t === "%d" ? String(Number(v)) : render(v);
    }) + (rest.length > i ? " " + rest.slice(i).map(render).join(" ") : "")
  );
};
const safeJson = (value) => {
  try {
    return JSON.stringify(value);
  } catch (_) {
    return String(value);
  }
};

// ---------------------------------------------------------------------------
// 1. Token endpoint mock (/oauth2/v2.0/token)
// ---------------------------------------------------------------------------
const b64url = (value) =>
  Buffer.from(JSON.stringify(value))
    .toString("base64")
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/, "");

function accessToken() {
  return [
    b64url({ alg: "none", typ: "JWT" }),
    b64url({
      aud: "https://substrate.office.com/sydney",
      tid: TENANT,
      oid: "harness-oid",
      upn: UPN,
      preferred_username: UPN,
      unique_name: UPN,
      exp: Math.floor(Date.now() / 1000) + 3600,
    }),
    "harness-signature",
  ].join(".");
}

const jsonResponse = (status, body) => ({
  ok: status >= 200 && status < 300,
  status,
  async json() {
    return body;
  },
  async text() {
    return JSON.stringify(body);
  },
});

function createNetwork(flow) {
  const calls = [];
  const fetchImpl = async (url, init = {}) => {
    const form = new URLSearchParams(String(init.body || ""));
    const grant = form.get("grant_type") || "";
    const code = form.get("code") || "";
    calls.push({ url: String(url), grant, code });
    if (!/\/oauth2\/v2\.0\/token\b/.test(String(url))) {
      return jsonResponse(404, { error: "not_found" });
    }
    if (grant === "authorization_code") {
      if (!/^mock-code-/.test(code)) {
        return jsonResponse(400, {
          error: "invalid_grant",
          error_description:
            "AADSTS9002313: Invalid request. Request is malformed or invalid.",
        });
      }
      return jsonResponse(200, {
        access_token: accessToken(),
        refresh_token: "rotated-rt",
        expires_in: 3600,
        refresh_token_expires_in: 86400,
      });
    }
    if (grant === "refresh_token") {
      if (flow.refreshOk === false) {
        return jsonResponse(400, {
          error: "invalid_grant",
          error_description:
            "AADSTS700084: The refresh token was issued to a single page app (SPA), and therefore has a fixed, limited lifetime of 1.00:00:00, which cannot be extended.",
        });
      }
      return jsonResponse(200, {
        access_token: accessToken(),
        refresh_token: "rotated-rt-refresh",
        expires_in: 3600,
        refresh_token_expires_in: 80000,
      });
    }
    return jsonResponse(400, { error: "unsupported_grant_type" });
  };
  return { fetchImpl, calls };
}

// ---------------------------------------------------------------------------
// 2. 内存版 WebExtension 运行时（只实现 background.js 用到的接口）
// ---------------------------------------------------------------------------
function createBrowserStub(hooks = {}, options = {}) {
  const local = new Map();
  const session = new Map();
  const tabs = [];
  const listeners = { message: [], tabRemoved: [] };
  const record = { notifications: [], badge: { text: "" }, tabReloads: [] };
  let nextTabId = Number(options.startTabId || 1000);

  const area = (map) => ({
    async get(keys) {
      if (keys == null) return Object.fromEntries(map);
      const list =
        typeof keys === "string"
          ? [keys]
          : Array.isArray(keys)
            ? keys
            : Object.keys(keys);
      const out = {};
      for (const k of list) if (map.has(k)) out[k] = map.get(k);
      return out;
    },
    async set(values) {
      for (const [k, v] of Object.entries(values || {})) map.set(k, v);
    },
    async remove(keys) {
      for (const k of typeof keys === "string" ? [keys] : keys || [])
        map.delete(k);
    },
    async clear() {
      map.clear();
    },
  });

  const matchUrl = (pattern, url) => {
    if (!pattern || pattern === "<all_urls>") return true;
    if (pattern.endsWith("/*"))
      return String(url).startsWith(pattern.slice(0, -1));
    return String(url) === pattern;
  };

  const badge = {
    async setBadgeText(a) {
      record.badge.text = String((a && a.text) || "");
    },
    async setBadgeBackgroundColor() {},
    async setTitle() {},
    async setIcon() {},
  };

  const api = {
    runtime: {
      id: "harness-extension-id",
      onMessage: {
        addListener: (fn) => listeners.message.push(fn),
        removeListener() {},
        hasListener: () => false,
      },
      onInstalled: { addListener() {} },
      onStartup: { addListener() {} },
      getURL: (p) => `chrome-extension://harness/${String(p || "")}`,
      getManifest: () => ({ name: "Voice Edge Auth Sync", version: "harness" }),
      sendMessage: async () => undefined,
      lastError: undefined,
    },
    storage: { local: area(local), session: area(session) },
    tabs: {
      async create(options) {
        const tab = {
          id: nextTabId++,
          url: String((options && options.url) || ""),
          active: Boolean(options && options.active),
        };
        tabs.push(tab);
        if (hooks.onTabCreated) {
          try {
            hooks.onTabCreated({ ...tab });
          } catch (_) {}
        }
        return { ...tab };
      },
      async remove(tabId) {
        const index = tabs.findIndex((t) => t.id === tabId);
        if (index < 0) throw new Error("no such tab");
        tabs.splice(index, 1);
        if (hooks.onTabRemoved) {
          try {
            hooks.onTabRemoved(tabId);
          } catch (_) {}
        }
        queueMicrotask(() => {
          for (const fn of listeners.tabRemoved) {
            try {
              fn(tabId, { isWindowClosing: false });
            } catch (_) {}
          }
        });
      },
      async query(query) {
        const wanted = query && query.url;
        const patterns =
          wanted == null ? null : Array.isArray(wanted) ? wanted : [wanted];
        return tabs
          .filter((t) => !patterns || patterns.some((p) => matchUrl(p, t.url)))
          .map((t) => ({ ...t }));
      },
      async update(tabId, props) {
        const tab = tabs.find((t) => t.id === tabId);
        if (tab && props && typeof props.active === "boolean")
          tab.active = props.active;
        return tab ? { ...tab } : {};
      },
      async reload(tabId) {
        // 续期成功后绝不允许刷新页面（会丢草稿/滚动/对话状态）
        record.tabReloads.push(tabId);
      },
      async sendMessage() {
        return undefined;
      },
      onRemoved: { addListener: (fn) => listeners.tabRemoved.push(fn) },
    },
    // 不提供 filterResponseData：走 Chromium 分支
    webRequest: new Proxy(
      {},
      {
        get(_t, key) {
          if (key === "filterResponseData") return undefined;
          if (typeof key === "string" && key.startsWith("on")) {
            return {
              addListener() {},
              removeListener() {},
              hasListener: () => false,
            };
          }
          return undefined;
        },
      },
    ),
    notifications: {
      async create(arg, maybe) {
        const options = typeof arg === "string" ? maybe : arg;
        record.notifications.push({
          title: String((options && options.title) || ""),
          message: String((options && options.message) || ""),
        });
        return "notification-id";
      },
      onClicked: { addListener() {} },
    },
    cookies: {
      async getAll() {
        return [];
      },
    },
    action: badge,
    browserAction: badge,
  };

  return {
    api,
    listeners,
    record,
    tabs,
    loginTabs: () =>
      tabs.filter((t) => /login\.microsoftonline\.com/.test(t.url)),
    navigate: (tabId, url) => {
      const tab = tabs.find((t) => t.id === tabId);
      if (tab) tab.url = String(url);
    },
    removeTab: (tabId) => api.tabs.remove(tabId),
  };
}

// ---------------------------------------------------------------------------
// 3. 在 VM 里加载真实 src/background.js
// ---------------------------------------------------------------------------
function loadBackground({ browser, fetchImpl, logSink }) {
  const wsSockets = [];
  const sandbox = {
    browser,
    console: {
      log: (...args) => logSink.push(fmtArgs(args)),
      warn() {},
      error: (...args) => logSink.push("ERROR " + fmtArgs(args)),
    },
    // 记录所有 WebSocket 客户端实例，测试可以手动 open()/receive()，
    // 用来验证扩展唯一的对外通道（5003 bridge）上的 M365_REAUTH_* 协议。
    WebSocket: class {
      constructor(url) {
        this.url = String(url);
        this.readyState = 0;
        this.sent = [];
        this.handlers = { open: [], message: [], close: [], error: [] };
        wsSockets.push(this);
      }
      addEventListener(type, fn) {
        (this.handlers[type] || (this.handlers[type] = [])).push(fn);
      }
      removeEventListener() {}
      send(data) {
        this.sent.push(String(data));
      }
      close() {
        this.readyState = 3;
        for (const fn of this.handlers.close || []) fn({});
      }
      open() {
        this.readyState = 1;
        for (const fn of this.handlers.open || []) fn({});
      }
      receive(payload) {
        for (const fn of this.handlers.message || [])
          fn({ data: JSON.stringify(payload) });
      }
      // background.js 用 WebSocket.OPEN 判断能否发送；缺了它 sendM365 会静默丢弃。
      static CONNECTING = 0;
      static OPEN = 1;
      static CLOSING = 2;
      static CLOSED = 3;
    },
    TextDecoder,
    TextEncoder,
    URL,
    URLSearchParams,
    crypto: globalThis.crypto,
    btoa,
    atob,
    setTimeout,
    clearTimeout,
    setInterval,
    clearInterval,
    queueMicrotask,
    navigator: { onLine: true, userAgent: "node-harness" },
    fetch: fetchImpl,
    importScripts: () => {},
    self: { addEventListener: () => {} },
    location: { href: "https://outlook.cloud.microsoft/" },
    document: {
      addEventListener() {},
      documentElement: null,
      querySelector: () => null,
      querySelectorAll: () => [],
    },
  };
  sandbox.globalThis = sandbox;
  sandbox.window = sandbox;
  const context = vm.createContext(sandbox);
  const file = BACKGROUND_FILE.startsWith("/")
    ? BACKGROUND_FILE
    : new URL("../../" + BACKGROUND_FILE, import.meta.url).pathname;
  vm.runInContext(readFileSync(file, "utf8"), context, { filename: file });
  return {
    sockets: wsSockets,
    run: (body) => vm.runInContext(`(async () => { ${body} })()`, context),
    async sendMessage(message, sender = {}) {
      for (const listener of browser.__listeners.message) {
        const result = await listener(message, sender);
        if (result !== undefined) return result;
      }
      return undefined;
    },
  };
}

// ---------------------------------------------------------------------------
// 4. Mock AAD 登录窗口（jsdom + 真实 content-m365-login.js）
// ---------------------------------------------------------------------------
function patchVisibility(win) {
  // jsdom 没有布局：rect 全是 0、opacity 计算值是 ""（Number("") === 0），
  // 内容脚本的 isVisible() 会把所有元素判成不可见。给 mock 一个真实尺寸。
  win.Element.prototype.getBoundingClientRect = () => ({
    width: 120,
    height: 32,
    top: 10,
    left: 10,
    right: 130,
    bottom: 42,
    x: 10,
    y: 10,
    toJSON: () => ({}),
  });
  win.getComputedStyle = (el) => {
    const style = (el && el.style) || {};
    const hidden =
      Boolean(el && el.hidden) ||
      Boolean(
        el && el.getAttribute && el.getAttribute("aria-hidden") === "true",
      );
    return {
      display: hidden ? "none" : style.display || "block",
      visibility: style.visibility || "visible",
      opacity: style.opacity || "1",
    };
  };
}

const page = (body, cfg) =>
  `<!doctype html><html><head><meta charset="utf-8">${
    cfg ? `<script>window.$Config=${JSON.stringify(cfg)};</script>` : ""
  }</head><body>${body}</body></html>`;

class MockLoginWindow {
  constructor({ tabId, scenario, onEvent, onTabUrl, sendMessage }) {
    this.tabId = tabId;
    this.scenario = scenario;
    this.onEvent = onEvent;
    this.onTabUrl = onTabUrl;
    this.sendMessage = sendMessage;
    this.win = null;
    this.closed = false;
    this.submitted = 0;
    this.kmsiDone = false;
    this.url = "";
    this._timers = new Set();
    this._session = new Map();
    this.events = [];
  }

  start(authorizeUrl) {
    const url = new URL(authorizeUrl);
    // 真实 AAD 只认应用注册过的 redirect_uri；不一致时回 50011 错误页，
    // 页面里根本没有密码框。这样一旦代码退回硬编码 url，测试立刻红。
    this.redirectUri = url.searchParams.get("redirect_uri") || REDIRECT_URI;
    const registered = this.scenario.registeredRedirect || CAPTURED_REDIRECT;
    const forcedError = this.scenario.forceAadError;
    if (forcedError || this.redirectUri !== registered) {
      this._later(() => this._renderAadErrorPage(forcedError || "50011"), 60);
      return;
    }
    this.stateParam = url.searchParams.get("state") || "";
    this._emit(`打开授权页 (state=${this.stateParam.slice(0, 18)}…)`);
    const sso = this.scenario.sso || "expired";
    if (sso === "pass") {
      this._later(() => this._redirect(), 80);
      return;
    }
    // prompt=none = 静默 SSO：会话还活着就直接给 code；会话死了才回
    // interaction_required（由后台升级一次交互式页面）。
    this.finalRedirect =
      url.searchParams.get("brk_redirect_uri") || REDIRECT_URI;
    if (url.searchParams.get("prompt") === "none" && !this.scenario.silentFails) {
      this._later(() => this._redirect(), 60);
      return;
    }
    if (
      sso === "interaction_required" ||
      url.searchParams.get("prompt") === "none"
    ) {
      this._later(() => {
        this._emit("AAD 拒绝：interaction_required（该页不会出现密码框）");
        this.url = `${REDIRECT_URI}#error=interaction_required&state=${this.stateParam}`;
        this.onTabUrl(this.url);
        this._send({
          ok: false,
          error: "interaction_required",
          errorDescription:
            "AADSTS50058: A silent sign-in request was sent but no user is signed in.",
        });
        this.close();
      }, 80);
      return;
    }
    this._later(
      () =>
        this.scenario.picker ? this._renderTiles() : this._renderPassword(),
      60,
    );
  }

  close() {
    if (this.closed) return;
    this.closed = true;
    for (const t of this._timers) clearTimeout(t);
    this._timers.clear();
    try {
      if (this.win) this.win.close();
    } catch (_) {}
  }

  _emit(text) {
    this.events.push(text);
    this.onEvent(this.tabId, text);
  }

  _later(fn, delay) {
    const timer = setTimeout(() => {
      this._timers.delete(timer);
      if (this.closed) return;
      try {
        fn();
      } catch (error) {
        this._emit("脚本异常: " + (error && error.message));
      }
    }, delay);
    this._timers.add(timer);
  }

  _send(payload) {
    return this.sendMessage(
      { __veM365: true, type: "M365_OAUTH_REDIRECT", ...payload },
      { tab: { id: this.tabId, url: this.url }, frameId: 0, url: this.url },
    );
  }

  _open(body, cfg) {
    if (this.closed) return null;
    if (this.win) {
      try {
        for (let i = 0; i < this.win.sessionStorage.length; i++) {
          const k = this.win.sessionStorage.key(i);
          this._session.set(k, this.win.sessionStorage.getItem(k));
        }
      } catch (_) {}
      try {
        this.win.close();
      } catch (_) {}
      this.win = null;
    }
    const url = `https://login.microsoftonline.com/${TENANT}/login`;
    this.url = url;
    this.onTabUrl(url);
    const dom = new JSDOM(page(body, cfg), {
      url,
      runScripts: "outside-only",
      pretendToBeVisual: true,
    });
    const win = dom.window;
    this.win = win;
    patchVisibility(win);
    for (const [k, v] of this._session) {
      try {
        win.sessionStorage.setItem(k, v);
      } catch (_) {}
    }
    win.browser = {
      runtime: {
        sendMessage: (message) =>
          Promise.resolve(
            this.sendMessage(message, {
              tab: { id: this.tabId, url: this.url },
              frameId: 0,
              url: this.url,
            }),
          ),
      },
    };
    win.eval(CONTENT_SCRIPT);
    return win;
  }

  _renderTiles() {
    this._emit("渲染账号选择器 (#tilesHolder)");
    const win = this._open(
      `<div id="tilesHolder"><div class="tile-container"><div class="table" role="button" data-test-id="${UPN}">${UPN}</div></div>
       <div class="tile-container"><div class="table" role="button" id="otherTile">使用其他账号</div></div></div>`,
      { pgid: "ConvergedSignIn", sPOST_Username: "" },
    );
    if (!win) return;
    for (const tile of win.document.querySelectorAll("[data-test-id]")) {
      tile.addEventListener("click", () => {
        this._emit(`脚本点击账号 ${tile.getAttribute("data-test-id")}`);
        this._later(() => this._renderPassword(), 120);
      });
    }
  }

  _renderPassword() {
    this._emit("渲染密码页 (#i0118 可见)");
    const win = this._open(
      `<form><input id="i0118" name="passwd" type="password" autocomplete="current-password">
       <button id="idSIButton9" type="button">登录</button></form>`,
      { pgid: "ConvergedSignIn", sPOST_Username: UPN },
    );
    if (!win) return;
    win.document
      .querySelector("#idSIButton9")
      .addEventListener("click", () => this._submit());
    this._scheduleAutofill(win);
  }

  _scheduleAutofill(win) {
    const mode = this.scenario.autofill || { mode: "immediate" };
    const fill = () => {
      if (this.closed || this.win !== win) return;
      const input = win.document.querySelector("#i0118");
      if (!input) return;
      input.value = this.scenario.fillValue || PASSWORD;
      input.dispatchEvent(new win.Event("input", { bubbles: true }));
      input.dispatchEvent(new win.Event("change", { bubbles: true }));
      this._emit(
        `浏览器密码库填充 #i0118${this.scenario.fillValue ? "（故意错的密码）" : ""}`,
      );
    };
    if (mode.mode === "immediate") this._later(fill, 120);
    else if (mode.mode === "delay")
      this._later(fill, Number(mode.delayMs || 1500));
    else if (mode.mode === "manual")
      this._later(fill, Number(mode.atMs || 6000));
    else this._emit("浏览器密码库不填充（等待用户手输）");
  }

  _submit() {
    if (this.closed || !this.win) return;
    const input = this.win.document.querySelector("#i0118");
    const value = String((input && input.value) || "");
    const ok = value === PASSWORD;
    this.submitted += 1;
    this._emit(`提交 #${this.submitted}（密码${ok ? "正确" : "错误"}）`);
    this._later(() => {
      if (!ok) return this._renderError();
      if (this.scenario.mfa) return this._renderMfa();
      if (this.scenario.kmsi && !this.kmsiDone) return this._renderKmsi();
      return this._redirect();
    }, 60);
  }

  _renderAadErrorPage(code) {
    this._emit(`AAD 错误页：AADSTS${code}（redirect_uri 未注册）`);
    this._open(
      `<div id="errorText">AADSTS${code}: The redirect URI specified in the request does not match the redirect URIs configured for the application.</div>`,
      { pgid: "ConvergedError", sErrorCode: String(code), sPOST_Username: "" },
    );
  }

  _renderError() {
    this._emit("渲染错误页：AADSTS50126 账号或密码错误");
    this._open(
      `<div id="passwordError">AADSTS50126: Error validating credentials due to invalid username or password.</div>
       <form><input id="i0118" name="passwd" type="password"><button id="idSIButton9" type="button">登录</button></form>`,
      { pgid: "ConvergedSignIn", sPOST_Username: UPN, sErrorCode: "50126" },
    );
  }

  _renderMfa() {
    this._emit("渲染二次验证页 (#idDiv_SAOTCS_Proofs)");
    this._open(
      `<div id="idDiv_SAOTCS_Proofs">请在手机上批准登录</div><input id="idTxtBx_SAOTCC_OTC" name="otc" type="text">`,
      { pgid: "ConvergedSignIn", sPOST_Username: UPN },
    );
  }

  _renderKmsi() {
    this._emit("渲染 KMSI 页（是否保持登录）");
    const win = this._open(
      `<div id="KmsiInterrupt">保持登录？</div>
       <input id="KmsiCheckboxField" name="DontShowAgain" type="checkbox">
       <button id="idSIButton9" type="button">是</button>`,
      { pgid: "KmsiInterrupt", sPOST_Username: UPN },
    );
    if (!win) return;
    win.document.querySelector("#idSIButton9").addEventListener("click", () => {
      this.kmsiDone = true;
      this._emit("脚本处理 KMSI 并提交");
      this._later(() => this._redirect(), 60);
    });
  }

  _redirect() {
    if (this.closed) return;
    const code = `mock-code-${this.tabId}-${this.submitted}`;
    this.url = `${this.finalRedirect || this.redirectUri}#code=${code}&state=${encodeURIComponent(this.stateParam)}`;
    this.onTabUrl(this.url);
    this._emit(`跳回 oauthRedirect.html#code=${code}`);
    this._send({ ok: true, code, state: this.stateParam });
    this._later(() => this.close(), 20);
  }
}

// ---------------------------------------------------------------------------
// 5. 场景与断言
// ---------------------------------------------------------------------------
let pass = 0;
let fail = 0;

function check(name, condition, detail) {
  if (condition) {
    pass += 1;
    console.log(`  PASS  ${name}`);
    return true;
  }
  fail += 1;
  console.log(`  FAIL  ${name}`);
  if (detail)
    console.log(`        ${String(detail).split("\n").join("\n        ")}`);
  return false;
}

async function scenario(title, config, body) {
  console.log(`\n=== ${title}`);
  const logSink = [];
  const flow = { refreshOk: config.refreshOk !== false };
  const state = { createdTabs: [], windows: new Map(), closed: 0, traces: [] };
  let bg = null;
  const t0 = Date.now();

  const stub = createBrowserStub({
    onTabCreated: (tab) => {
      state.createdTabs.push(tab);
      if (!/login\.microsoftonline\.com/.test(tab.url)) return;
      // 模拟“断网 / 页面根本没加载出来”：只有标签，没有任何页面。
      if (config.noPage) return;
      const win = new MockLoginWindow({
        tabId: tab.id,
        scenario: config,
        onEvent: (tabId, text) =>
          console.log(
            `  [窗口 #${tabId}] +${String(Date.now() - t0).padStart(4)}ms ${text}`,
          ),
        onTabUrl: (url) => stub.navigate(tab.id, url),
        sendMessage: (message, sender) => bg.sendMessage(message, sender),
      });
      state.windows.set(tab.id, win);
      win.start(tab.url);
    },
    onTabRemoved: (tabId) => {
      const win = state.windows.get(tabId);
      if (!win) return;
      state.traces.push(win.events);
      state.closed += 1;
      state.windows.delete(tabId);
      win.close();
    },
  });

  const browser = Object.assign(stub.api, { __listeners: stub.listeners });
  const net = createNetwork(flow);
  bg = loadBackground({ browser, fetchImpl: net.fetchImpl, logSink });

  await bg.run(`
    await Promise.all([ensureM365AuthLoaded(), ensureM365AutoLoginLoaded(), ensureM365ReauthStateLoaded()]);
    return true;
  `);
  await bg.run(`
    _tid = ${JSON.stringify(TENANT)};
    _clientId = ${JSON.stringify(CLIENT_ID)};
    _rt = "seed-refresh-token";
    _rtExp = Date.now() - 60 * 1000;
    _sydney = { token: "", exp: 0 };
    _m365TokenContext = ${JSON.stringify(SEED_CONTEXT)};
    m365AuthLoaded = true;
    // 用户真实配置：面板里开了自动登录，但扩展内不存密码，靠浏览器密码库
    m365AutoLoginEnabled = true;
    m365AutoLoginRemember = false;
    m365AutoLoginPassword = "";
    m365AutoLoginUsername = ${JSON.stringify(UPN)};
    m365AutoLoginPaused = false;
    await saveM365AutoLogin();
    await saveM365Auth();
    return true;
  `);

  const snapshot = () =>
    bg.run(`
      const tabs = await browser.tabs.query({ url: "https://login.microsoftonline.com/*" });
      return {
        state: m365ReauthState,
        reason: m365ReauthStateReason,
        pendingUrl: _pendingTopLevelReauth ? String(_pendingTopLevelReauth.url || "") : "",
        rt: _rt,
        rtExp: _rtExp,
        paused: Boolean(m365AutoLoginPaused),
        loginTabs: tabs.length,
      };
    `);

  const bridge = {
    socket() {
      return bg.sockets[bg.sockets.length - 1] || null;
    },
    open() {
      const ws = this.socket();
      if (ws) ws.open();
      return Boolean(ws);
    },
    command(payload) {
      const ws = this.socket();
      if (!ws) throw new Error("扩展还没有建立 bridge 连接");
      ws.receive(payload);
    },
    sent(type) {
      const ws = this.socket();
      if (!ws) return [];
      return ws.sent
        .map((raw) => {
          try {
            return JSON.parse(raw);
          } catch (_) {
            return null;
          }
        })
        .filter((m) => m && (!type || m.type === type));
    },
  };

  const harness = {
    bg,
    stub,
    net,
    logSink,
    state,
    snapshot,
    bridge,
    async waitFor(predicate, { timeout = 15000 } = {}) {
      const startedAt = Date.now();
      let last = null;
      while (Date.now() - startedAt < timeout) {
        last = await snapshot();
        last.openWindows = state.windows.size;
        last.createdTabs = state.createdTabs.length;
        if (predicate(last)) return last;
        await sleep(40);
      }
      throw new Error(`超时 ${timeout}ms；最后状态=${safeJson(last)}`);
    },
  };

  try {
    await body(harness);
  } catch (error) {
    check(
      `${title} — 流程跑完未抛异常`,
      false,
      (error && (error.stack || error.message)) || String(error),
    );
    console.log("  --- 后台日志（最后 20 行）---");
    for (const line of logSink.slice(-20)) console.log("    " + line);
  } finally {
    for (const win of state.windows.values()) win.close();
    state.windows.clear();
    if (VERBOSE) {
      console.log("  --- 后台日志 ---");
      for (const line of logSink) console.log("    " + line);
    }
  }
}

// 一个“真实的 OWA 页面”：jsdom + 真实的 src/content-m365.js（页面世界钩子），
// 它的 browser.runtime.sendMessage 直接通到被测后台，fetch 由桩接管。
function openOwaPage(h) {
  const dom = new JSDOM(
    `<!doctype html><html><head></head><body><div id="app"></div></body></html>`,
    {
      url: "https://outlook.cloud.microsoft/mail/",
      runScripts: "dangerously",
      pretendToBeVisual: true,
    },
  );
  const w = dom.window;
  w.Element.prototype.getBoundingClientRect = () => ({
    width: 120,
    height: 32,
    top: 0,
    left: 0,
    right: 120,
    bottom: 32,
  });
  const posted = [];
  const sender = {
    tab: { id: 900, url: "https://outlook.cloud.microsoft/mail/" },
    frameId: 0,
    url: "https://outlook.cloud.microsoft/mail/",
  };
  w.browser = {
    runtime: {
      id: "test",
      sendMessage: (m) => {
        posted.push(m);
        try {
          return h.bg.sendMessage(m, sender);
        } catch (_) {
          return Promise.resolve(undefined);
        }
      },
      getURL: (p) => "moz-extension://test/" + p,
      onMessage: { addListener: () => {} },
    },
    storage: {
      local: { get: () => Promise.resolve({}), set: () => Promise.resolve() },
    },
  };
  const calls = [];
  w.fetch = async (input, init) => {
    calls.push({
      url: typeof input === "string" ? input : (input && input.url) || "",
      body: init && init.body,
    });
    return {
      ok: true,
      status: 200,
      json: async () => ({}),
      text: async () => "",
      clone() {
        return this;
      },
    };
  };
  w.eval(readFileSync(join(ROOT, "src/content-m365.js"), "utf8"));
  return {
    w,
    posted,
    post(payload) {
      w.postMessage(
        Object.assign({ __veM365: true, dir: "toPage" }, payload),
        "*",
      );
    },
    fetchToken(refreshToken) {
      return w.fetch(
        `https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`,
        {
          method: "POST",
          body: `client_id=${CLIENT_ID}&grant_type=refresh_token&refresh_token=${refreshToken}&scope=x`,
        },
      );
    },
    lastBody() {
      return String((calls.at(-1) || {}).body || "");
    },
    close() {
      w.close();
    },
  };
}

async function main() {
  assertProductionBuild();
  console.log("M365 自动登录 harness — 构建产物 + mock 登录窗口");
  console.log(
    `tenant=${TENANT}  upn=${UPN}  扩展内密码=未保存（模拟你的真实配置）`,
  );
  console.log(`被测代码: ${BACKGROUND_FILE}\n`);

  // 1. 生产真实触发路径
  await scenario(
    "1. 预续期收到 AADSTS700084 → 顶层登录窗口完成自动登录",
    { sso: "expired", autofill: { mode: "immediate" }, refreshOk: false },
    async (h) => {
      await h.bg.run(`await runPreExpirationRenewal(); return true;`);
      const snap = await h.waitFor(
        (s) => s.state === "none" && s.rt === "rotated-rt",
      );
      check(
        "refresh 被拒后转为顶层登录",
        h.net.calls.some((c) => c.grant === "refresh_token"),
      );
      check(
        "授权码换取成功，令牌落库",
        snap.rt === "rotated-rt",
        `rt=${snap.rt}`,
      );
      check("新的 24h RT 寿命已刷新", (snap.rtExp - Date.now()) / 3600000 > 23);
      check(
        "只开过 1 个登录窗口",
        h.state.createdTabs.length === 1,
        `tabs=${h.state.createdTabs.length}`,
      );
    },
  );

  // 2. SSO 直通
  await scenario(
    "2. SSO 仍有效：授权页直接回跳，不走密码页",
    { sso: "pass" },
    async (h) => {
      await h.bg.run(`return await tryTopLevelSilentReauth("e2e");`);
      const snap = await h.waitFor(
        (s) => s.state === "none" && s.rt === "rotated-rt",
      );
      const url = (h.state.createdTabs[0] && h.state.createdTabs[0].url) || "";
      check("授权码换取成功", snap.rt === "rotated-rt", `rt=${snap.rt}`);
      check(
        "授权 URL 含 PKCE",
        /code_challenge=/.test(url) && /code_challenge_method=S256/.test(url),
        url.slice(0, 120),
      );
      check(
        "静默阶段带 prompt=none（和 OWA 自己一致）",
        url.includes("prompt=none"),
        url.slice(0, 220),
      );
      check(
        "redirect_uri / brk_redirect_uri 按抓取值配对",
        url.includes("redirect_uri=" + encodeURIComponent(BROKER_REDIRECT)) &&
          url.includes("brk_redirect_uri=" + encodeURIComponent(WEB_REDIRECT)),
        url.slice(0, 260),
      );
      check(
        "带上 brk_client_id",
        url.includes("brk_client_id=" + BRK_CLIENT_ID),
        url.slice(0, 260),
      );
      check(
        "登录窗口开在前台（用户看得到）",
        h.state.createdTabs[0] && h.state.createdTabs[0].active === true,
        `active=${h.state.createdTabs[0] && h.state.createdTabs[0].active}`,
      );
      check("只开过 1 个登录窗口", h.state.createdTabs.length === 1);
      const ctx = await h.bg.run(
        `return (_m365TokenContext && _m365TokenContext.body) || {};`,
      );
      check(
        "token 上下文保持页面原值，没有被后台“修好”（不接受非页面来源的值）",
        ctx.redirect_uri === SEED_CONTEXT.body.redirect_uri &&
          ctx.brk_client_id === SEED_CONTEXT.body.brk_client_id &&
          ctx.brk_redirect_uri === SEED_CONTEXT.body.brk_redirect_uri,
        JSON.stringify({
          redirect_uri: ctx.redirect_uri,
          brk_client_id: ctx.brk_client_id,
          brk_redirect_uri: ctx.brk_redirect_uri,
        }),
      );
      check(
        "续期成功后不刷新任何 M365 页面",
        h.stub.record.tabReloads.length === 0,
        `reloads=${JSON.stringify(h.stub.record.tabReloads)}`,
      );
    },
  );

  // 静默阶段会话已死 → 升级一次交互式页面（只升级一次，不是循环）
  await scenario(
    "静默 prompt=none 被判 interaction_required → 升级一次交互式页面并成功",
    { sso: "expired", silentFails: true, autofill: { mode: "immediate" } },
    async (h) => {
      await h.bg.run(`return await tryTopLevelSilentReauth("e2e");`);
      const snap = await h.waitFor((s) => s.state === "none" && s.rt === "rotated-rt");
      const tabs = h.state.createdTabs;
      check("先静默、再交互，共 2 个标签页", tabs.length === 2, `tabs=${tabs.length}`);
      check(
        "第一个是 prompt=none",
        tabs[0] && /prompt=none/.test(tabs[0].url),
        tabs[0] && tabs[0].url.slice(0, 160),
      );
      check(
        "第二个转成交互式（没有 prompt=none）",
        tabs[1] && !/prompt=none/.test(tabs[1].url),
        tabs[1] && tabs[1].url.slice(0, 160),
      );
      check("最终换到新令牌", snap.rt === "rotated-rt", `rt=${snap.rt}`);
    },
  );

  // AAD 只认注册过的 redirect_uri：不对就回 50011 错误页（页面上没有密码框）
  await scenario(
    "AAD 拒绝未注册的 redirect_uri (50011) → 快速失败，不空等 45 秒",
    { sso: "expired", forceAadError: "50011" },
    async (h) => {
      await h.bg.run(`return await tryTopLevelSilentReauth("e2e", { interactive: true });`);
      const snap = await h.waitFor((s) => s.state === "terminal", { timeout: 8000 });
      check("进入 terminal", snap.state === "terminal", snap.reason);
      check("原因里带 AADSTS50011", /50011/.test(snap.reason), snap.reason);
      check("窗口已关闭", snap.openWindows === 0);
    },
  );

  // 断网 / 页面压根没加载：这不算“凭证失败”，不能把当天锁死
  await scenario(
    "登录页始终不加载（模拟断网）→ 环境性失败，不消耗唯一尝试",
    { sso: "expired", noPage: true },
    async (h) => {
      await h.bg.run(`return await tryTopLevelSilentReauth("e2e", { interactive: true });`);
      const snap = await h.waitFor(
        (s) => s.state === "none" && s.openWindows === 0,
        { timeout: 20000 },
      );
      check(
        "最终回到 none（不是 terminal）",
        snap.state === "none",
        snap.state + " / " + snap.reason,
      );
      check(
        "原因标记为环境性失败",
        /environmental|page-not-loaded|offline/.test(snap.reason),
        snap.reason,
      );
      check(
        "日志写明可重试时间",
        h.logSink.some((l) => /new attempt allowed in/.test(l)),
      );
      check("窗口已关闭", snap.openWindows === 0);
    },
  );

  // 浏览器离线：连窗口都不该开
  await scenario(
    "浏览器离线 → 直接跳过，不开窗",
    { sso: "expired" },
    async (h) => {
      await h.bg.run(`navigator.onLine = false; return true;`);
      const started = await h.bg.run(`return await tryTopLevelSilentReauth("e2e", { interactive: true });`);
      await sleep(200);
      const snap = await h.snapshot();
      check("没有发起尝试", started === false);
      check(
        "没有开任何窗口",
        h.state.createdTabs.length === 0,
        `tabs=${h.state.createdTabs.length}`,
      );
      check("状态没有被污染", snap.state === "none", snap.state);
    },
  );

  // 3. 密码库立即填充
  await scenario(
    "3. 会话失效 + 密码库立即填充：窗口内自动填密并提交",
    { sso: "expired", autofill: { mode: "immediate" } },
    async (h) => {
      await h.bg.run(`return await tryTopLevelSilentReauth("e2e", { interactive: true });`);
      const snap = await h.waitFor(
        (s) => s.state === "none" && s.rt === "rotated-rt",
      );
      const trace = h.state.traces.flat();
      check(
        "密码页被渲染",
        trace.includes("渲染密码页 (#i0118 可见)"),
        trace.join(" | "),
      );
      check(
        "密码被填充且提交成功",
        trace.some((t) => t === "提交 #1（密码正确）"),
        trace.join(" | "),
      );
      check("令牌已更新", snap.rt === "rotated-rt", `rt=${snap.rt}`);
    },
  );

  // 4. 晚 1.2 秒填充
  await scenario(
    "4. 密码库晚 1.2 秒填充：仍在等待窗口内",
    { sso: "expired", autofill: { mode: "delay", delayMs: 1200 } },
    async (h) => {
      await h.bg.run(`return await tryTopLevelSilentReauth("e2e", { interactive: true });`);
      const snap = await h.waitFor(
        (s) => s.state === "none" && s.rt === "rotated-rt",
      );
      const trace = h.state.traces.flat();
      check(
        "晚填后脚本自动提交",
        trace.some((t) => t === "提交 #1（密码正确）"),
        trace.join(" | "),
      );
      check("令牌已更新", snap.rt === "rotated-rt", `rt=${snap.rt}`);
    },
  );

  // 5. 密码库始终不填充：自动登录失败，关窗，绝不等人输入
  await scenario(
    "5. 密码库始终不填充 → 失败、关窗、通知（禁止人工输入）",
    { sso: "expired", autofill: { mode: "never" } },
    async (h) => {
      await h.bg.run(`return await tryTopLevelSilentReauth("e2e", { interactive: true });`);
      const snap = await h.waitFor((s) => s.state === "terminal", {
        timeout: 12000,
      });
      check(
        "进入 terminal（不进入等待人工的状态）",
        snap.state === "terminal",
        snap.reason,
      );
      check(
        "原因写明密码库没填",
        /password manager/i.test(snap.reason),
        snap.reason,
      );
      check(
        "登录窗口已关闭",
        snap.openWindows === 0,
        `open=${snap.openWindows}`,
      );
      check(
        "弹出了失败通知",
        h.stub.record.notifications.some((n) => /自动登录失败/.test(n.title)),
      );
      check("全程只开过 1 个窗口", h.state.createdTabs.length === 1);
    },
  );

  // 6. 密码错误
  await scenario(
    "6. 密码错误 (AADSTS50126)：终态 + 暂停自动登录 + 关窗 + 通知",
    {
      sso: "expired",
      autofill: { mode: "immediate" },
      fillValue: "wrong-password",
    },
    async (h) => {
      await h.bg.run(`return await tryTopLevelSilentReauth("e2e", { interactive: true });`);
      const snap = await h.waitFor((s) => s.state === "terminal");
      check("进入 terminal", snap.state === "terminal", snap.reason);
      check("原因含 AADSTS50126", /50126/.test(snap.reason), snap.reason);
      check("自动登录被暂停（防锁号）", snap.paused === true);
      check("窗口已关闭", snap.openWindows === 0);
      check(
        "弹出失败通知",
        h.stub.record.notifications.some((n) => /自动登录失败/.test(n.title)),
      );
    },
  );

  // 7. MFA
  await scenario(
    "7. 需要 MFA：终态 + 关窗 + 通知（没人会去点批准）",
    { sso: "expired", autofill: { mode: "immediate" }, mfa: true },
    async (h) => {
      await h.bg.run(`return await tryTopLevelSilentReauth("e2e", { interactive: true });`);
      const snap = await h.waitFor((s) => s.state === "terminal");
      check("进入 terminal", snap.state === "terminal", snap.reason);
      check(
        "窗口被关闭（没人会去点批准）",
        snap.openWindows === 0,
        `open=${snap.openWindows}`,
      );
      check(
        "弹出 MFA 通知",
        h.stub.record.notifications.some((n) => /二次验证/.test(n.title)),
      );
    },
  );

  // 8. AAD 直接拒绝
  await scenario(
    "8. AAD 直接回 interaction_required：一次即终态",
    { sso: "interaction_required" },
    async (h) => {
      await h.bg.run(`return await tryTopLevelSilentReauth("e2e", { interactive: true });`);
      const snap = await h.waitFor((s) => s.state === "terminal", {
        timeout: 8000,
      });
      check("进入 terminal", snap.state === "terminal", snap.reason);
      check(
        "原因含 interaction_required",
        /interaction_required/.test(snap.reason),
        snap.reason,
      );
      check("窗口已关闭", snap.openWindows === 0);
    },
  );

  // 9. 账号选择器
  await scenario(
    "9. 账号选择器 → 密码页：整条链路走通",
    { sso: "expired", picker: true, autofill: { mode: "immediate" } },
    async (h) => {
      await h.bg.run(`return await tryTopLevelSilentReauth("e2e", { interactive: true });`);
      const snap = await h.waitFor(
        (s) => s.state === "none" && s.rt === "rotated-rt",
      );
      const trace = h.state.traces.flat();
      check(
        "先渲染账号选择器",
        trace.some((t) => t.includes("账号选择器")),
        trace.join(" | "),
      );
      check("选择账号后完成换取", snap.rt === "rotated-rt", `rt=${snap.rt}`);
    },
  );

  // 10. 登录页被关闭
  await scenario(
    "10. 登录页被关闭：终态，之后不再自动开窗",
    { sso: "expired", autofill: { mode: "never" } },
    async (h) => {
      await h.bg.run(`return await tryTopLevelSilentReauth("e2e", { interactive: true });`);
      const opened = await h.waitFor((s) => s.openWindows === 1);
      await h.stub.removeTab(
        await h.bg.run(`return _pendingTopLevelReauth.tabId;`),
      );
      const snap = await h.waitFor((s) => s.state === "terminal", {
        timeout: 8000,
      });
      check("关窗即终态", snap.state === "terminal", snap.reason);
      check("原因记录为窗口被关闭", /closed/.test(snap.reason), snap.reason);
      const before = h.state.createdTabs.length;
      await h.bg.run(`return await tryTopLevelSilentReauth("e2e-auto");`);
      await h.bg.run(`await runPreExpirationRenewal(); return true;`);
      await h.bg.run(`await runPreExpirationRenewal(); return true;`);
      await sleep(120);
      check(
        "终态后自动路径被拒绝，未新建窗口",
        h.state.createdTabs.length === before && h.state.windows.size === 0,
        `tabs ${before} -> ${h.state.createdTabs.length}`,
      );
      void opened;
    },
  );

  // 11. 复用既有 5003 通道：M365_REAUTH_NOW → 新 RT，且握手前不发任何新增流量
  await scenario(
    "11. 5003 bridge 协议：M365_REAUTH_NOW 触发续期并回传新 RT",
    { sso: "expired", autofill: { mode: "immediate" } },
    async (h) => {
      check(
        "扩展连的是唯一通道 ws://127.0.0.1:5003/ws",
        h.bridge.socket() && h.bridge.socket().url === "ws://127.0.0.1:5003/ws",
        h.bridge.socket() && h.bridge.socket().url,
      );
      check(
        "没有任何 5004 之类的第二条连接",
        h.bg.sockets.length === 1,
        `sockets=${h.bg.sockets.length}`,
      );
      h.bridge.open();
      await sleep(30);
      check(
        "连接后先发 M365_READY（既有握手）",
        h.bridge.sent("M365_READY").length === 1,
        JSON.stringify(h.bridge.sent().map((m) => m.type)),
      );
      check(
        "客户端没发命令前，不主动推任何 M365_REAUTH_* 消息",
        h.bridge.sent("M365_REAUTH_STARTED").length === 0 &&
          h.bridge.sent("M365_REAUTH_STATE").length === 0,
        JSON.stringify(h.bridge.sent().map((m) => m.type)),
      );

      h.bridge.command({
        type: "M365_REAUTH_NOW",
        force: true,
        requestId: "harness-test-1",
      });
      const snap = await h.waitFor((s) => s.rt === "rotated-rt", { timeout: 15000 });
      check("收到命令后真的换了新 RT", snap.rt === "rotated-rt", `rt=${snap.rt}`);

      const startedAt = Date.now();
      while (Date.now() - startedAt < 8000 && !h.bridge.sent("M365_REAUTH_DONE").length) {
        await sleep(100);
      }
      const started = h.bridge.sent("M365_REAUTH_STARTED")[0];
      const done = h.bridge.sent("M365_REAUTH_DONE")[0];
      check(
        "回了 M365_REAUTH_STARTED（带 requestId）",
        started && started.requestId === "harness-test-1" && started.started === true,
        JSON.stringify(started),
      );
      check(
        "回了 M365_REAUTH_DONE，且 rtRemain 是新的 24h",
        done && done.requestId === "harness-test-1" && done.rtRemainMin > 1000,
        done ? `rtRemain=${done.rtRemainMin}` : "no DONE",
      );
      check(
        "DONE 里带后台日志（含 renewed 24h）",
        Array.isArray(done && done.log) &&
          done.log.some((line) => /renewed 24h M365 SPA refresh token/.test(line)),
        done && Array.isArray(done.log) ? String(done.log.length) + " lines" : "no log",
      );
    },
  );

  // 12. 页面世界的 /token 改写：拿到新 RT 后，页面自己刷新也用新的
  {
    console.log("\n=== 12. 页面 /token 的 refresh_token 被换成扩展的新 RT");
    const dom = new JSDOM(
      `<!doctype html><html><head></head><body><div id="app"></div></body></html>`,
      {
        url: "https://outlook.cloud.microsoft/mail/",
        runScripts: "dangerously",
        pretendToBeVisual: true,
      },
    );
    const w = dom.window;
    w.Element.prototype.getBoundingClientRect = () => ({
      width: 120,
      height: 32,
      top: 0,
      left: 0,
      right: 120,
      bottom: 32,
    });
    const posted = [];
    w.browser = {
      runtime: {
        id: "test",
        sendMessage: (m) => {
          posted.push(m);
          return Promise.resolve({ ok: true });
        },
        getURL: (p) => "moz-extension://test/" + p,
        onMessage: { addListener: () => {} },
      },
      storage: {
        local: { get: () => Promise.resolve({}), set: () => Promise.resolve() },
      },
    };
    const calls = [];
    let stubStatus = 200;
    w.fetch = async (input, init) => {
      calls.push({
        url: typeof input === "string" ? input : (input && input.url) || "",
        body: init && init.body,
      });
      return {
        ok: stubStatus < 400,
        status: stubStatus,
        json: async () => ({}),
        text: async () => "",
        clone() {
          return this;
        },
      };
    };
    // 模拟旧扩展留下的包装器：只有标记、没有版本号。
    // 若 hookTokenFetch 仍按"有标记就跳过"，新闭包里的 RT 永远用不上。
    w.fetch.__veM365TokenHook = true;
    w.eval(readFileSync(join(ROOT, "src/content-m365.js"), "utf8"));
    await sleep(250);
    check("页面 fetch 钩子已安装", Boolean(w.fetch.__veM365TokenHook));
    check(
      "旧包装器存在时，新钩子仍会叠一层（带版本号）",
      Number(w.fetch.__veM365TokenHookVersion || 0) >= 2,
      `version=${w.fetch.__veM365TokenHookVersion}`,
    );

    w.postMessage(
      {
        __veM365: true,
        dir: "toPage",
        type: "M365_KEEPALIVE",
        authHealthy: true,
        rt: "NEW-RT-24H",
        rtClientId: CLIENT_ID,
      },
      "*",
    );
    await sleep(120);
    await w.fetch(
      `https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`,
      {
        method: "POST",
        body: `client_id=${CLIENT_ID}&grant_type=refresh_token&refresh_token=OLD-RT&scope=x%20openid`,
      },
    );
    const sent = calls.at(-1) || {};
    const body = String(sent.body || "");
    check(
      "页面 /token 里的 refresh_token 被替换成新 RT",
      body.includes("refresh_token=NEW-RT-24H"),
      body.slice(0, 160),
    );
    check(
      "其它字段原样保留（grant_type/scope/client_id）",
      body.includes("grant_type=refresh_token") &&
        body.includes("client_id=" + CLIENT_ID) &&
        body.includes("scope="),
      body.slice(0, 160),
    );
    await sleep(120);
    check(
      "页面把替换事件报给后台",
      posted.some((m) => m && m.type === "M365_RT_SWAPPED"),
      JSON.stringify(posted.slice(-3)),
    );

    // 不是 token 端点 / 不是 refresh_token 授权 → 一个字都不能动
    await w.fetch("https://outlook.cloud.microsoft/mail/startup", {
      method: "POST",
      body: "refresh_token=OLD-RT",
    });
    check(
      "非 /token 请求不被改写",
      String((calls.at(-1) || {}).body || "").includes("refresh_token=OLD-RT"),
    );
    await w.fetch(
      `https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`,
      { method: "POST", body: `client_id=other&grant_type=refresh_token&refresh_token=OLD-RT` },
    );
    check(
      "client_id 不匹配时不改写",
      String((calls.at(-1) || {}).body || "").includes("refresh_token=OLD-RT"),
    );

    // 换上去的 RT 被 AAD 拒绝 → 立刻退避，后续不再改写（页面用自己的 RT 活下去）
    stubStatus = 400;
    w.postMessage(
      {
        __veM365: true,
        dir: "toPage",
        type: "M365_KEEPALIVE",
        authHealthy: true,
        rt: "REJECTED-RT",
        rtClientId: CLIENT_ID,
      },
      "*",
    );
    await sleep(120);
    const tokenUrl = `https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`;
    await w.fetch(tokenUrl, {
      method: "POST",
      body: `client_id=${CLIENT_ID}&grant_type=refresh_token&refresh_token=GOOD-RT`,
    });
    check(
      "第一次仍会换（就是它被拒的）",
      String((calls.at(-1) || {}).body || "").includes("refresh_token=REJECTED-RT"),
      String((calls.at(-1) || {}).body || "").slice(0, 120),
    );
    // 被拒之后也不回退：下一次照样换成我们这条（旧 RT 一律不用）
    await w.fetch(tokenUrl, {
      method: "POST",
      body: `client_id=${CLIENT_ID}&grant_type=refresh_token&refresh_token=OLD-RT-2`,
    });
    check(
      "被拒后仍然只用自己的 RT（不回退旧 RT）",
      String((calls.at(-1) || {}).body || "").includes("refresh_token=REJECTED-RT"),
      String((calls.at(-1) || {}).body || "").slice(0, 120),
    );
    await sleep(80);
    check(
      "上报了换用失败事件",
      posted.some((m) => m && m.type === "M365_RT_SWAP_FAILED"),
      JSON.stringify(posted.slice(-3)),
    );
    dom.window.close();
  }

  // 13. 页面报告“我们的 RT 被拒” → 后台用真实 /token 验证，不回退旧 RT
  await scenario(
    "13. 页面报告 RT 被拒 → 后台真刷新验证（不回退页面旧 RT）",
    { sso: "expired" },
    async (h) => {
      // access token 还没过期，正常情况下 refreshSydney 会直接返回缓存；
      // 这里必须证明“被拒”会强制做一次真实刷新。
      await h.bg.run(
        `_sydney = { token: "still-valid-at", exp: Date.now() + 3600 * 1000 }; return true;`,
      );
      const before = h.net.calls.length;
      await h.bg.sendMessage(
        { __veM365: true, type: "M365_RT_SWAP_FAILED", status: 400 },
        {
          tab: { id: 7, url: "https://outlook.cloud.microsoft/mail/" },
          frameId: 0,
          url: "https://outlook.cloud.microsoft/mail/",
        },
      );
      await sleep(400);
      const refreshCalls = h.net.calls
        .slice(before)
        .filter((c) => c && c.grant === "refresh_token");
      check(
        "立刻发起一次真实 /token 验证（不吃 access token 缓存）",
        refreshCalls.length >= 1,
        JSON.stringify(h.net.calls.slice(before)),
      );
      check(
        "日志写明不回退页面旧 RT",
        h.logSink.some((l) => /no fallback to the page's old RT/.test(l)),
      );
      check(
        "验证用的是我们自己的 refresh_token",
        h.logSink.some((l) =>
          /M365 page \/token was rejected with our RT/.test(l),
        ),
      );
    },
  );

  // 14. 端到端：真实的续期 → 页面真用上这条新 RT
  await scenario(
    "14. 端到端：续期换到新 RT 后，OWA 页面的 /token 立刻改用新 RT",
    { sso: "expired", autofill: { mode: "immediate" }, refreshOk: false },
    async (h) => {
      const page = openOwaPage(h);
      await sleep(250);
      check(
        "OWA 页面加载了真实的页面世界钩子",
        Boolean(page.w.fetch.__veM365TokenHook),
      );

      // 真实触发：refresh_token 被 700084 拒 → 顶层登录 → 换到全新 24h RT
      await h.bg.run(`await runPreExpirationRenewal(); return true;`);
      const snap = await h.waitFor((s) => s.rt === "rotated-rt", {
        timeout: 15000,
      });
      check("续期真的换到了新 RT", snap.rt === "rotated-rt", `rt=${snap.rt}`);

      // 后台把新 RT 推给页面（生产路径是 pingM365TabsKeepAlive，桩里没有真实 tab）
      const rt = await h.bg.run(`return _rt;`);
      const rtClientId = await h.bg.run(`return _clientId;`);
      page.post({ type: "M365_KEEPALIVE", authHealthy: true, rt, rtClientId });
      await sleep(150);

      await page.fetchToken("PAGE-OLD-RT");
      const body = page.lastBody();
      check(
        "页面 /token 用的是刚换到的新 RT（不是它自己那条）",
        body.includes("refresh_token=" + rt),
        body.slice(0, 150),
      );
      check(
        "页面没有把旧 RT 发出去",
        !body.includes("PAGE-OLD-RT"),
        body.slice(0, 150),
      );
      await sleep(120);
      check(
        "页面上报了替换事件",
        page.posted.some((m) => m && m.type === "M365_RT_SWAPPED"),
        JSON.stringify(page.posted.slice(-2)),
      );
      page.close();
    },
  );

  // 15. 抓包不全 → 绝不拿常量兜底
  await scenario(
    "15. 抓包缺 brk_* → 拒绝拼 /authorize，失败原因点名缺失字段",
    { sso: "pass" },
    async (h) => {
      await h.bg.run(`
        _m365TokenContext = { version: 1, query: {}, body: {
          client_id: ${JSON.stringify(CLIENT_ID)},
          scope: ${JSON.stringify("https://substrate.office.com/sydney/.default openid profile offline_access")},
          redirect_uri: ${JSON.stringify(BROKER_REDIRECT)}
        } };
        return true;
      `);
      const started = await h.bg.run(
        `return await tryTopLevelSilentReauth("missing-capture");`,
      );
      const snap = await h.waitFor((s) => s.state === "terminal");
      check("没有为此打开任何登录窗口", h.state.createdTabs.length === 0,
        `tabs=${h.state.createdTabs.length}`);
      check("调用方拿到 false（没有假装开始）", started === false, `started=${started}`);
      check(
        "失败原因点名缺的两个 broker 参数",
        /brk_client_id/.test(snap.reason) && /brk_redirect_uri/.test(snap.reason),
        snap.reason,
      );
      check(
        "日志里没有借用任何硬编码值",
        !h.net.calls.some((c) => String(c.url || "").includes("brk-multihub://")),
        JSON.stringify(h.net.calls.map((c) => c.url).slice(0, 3)),
      );
    },
  );

  // 16. 同 clientId 的非 Sydney 轮换：只喂 _rt/_rtExp，清标留给真刷 sydney
  await scenario(
    "16. 同 clientId 的非 Sydney 轮换：采纳 RT、夹紧初值、不直接清标",
    { sso: "pass" },
    async (h) => {
      const token = (aud) =>
        [
          b64url({ alg: "none", typ: "JWT" }),
          b64url({
            aud,
            tid: TENANT,
            oid: "harness-oid",
            upn: UPN,
            exp: Math.floor(Date.now() / 1000) + 3600,
          }),
          "harness-signature",
        ].join(".");
      const sibling = token("https://api.spaces.skype.com");
      const apply = (clientId, accessToken, refreshToken) =>
        h.bg.run(`
          await applyM365TokenResponse({
            tenantId: ${JSON.stringify(TENANT)},
            clientId: ${JSON.stringify(clientId)},
            grantType: "refresh_token",
            source: "page",
            accessToken: ${JSON.stringify(accessToken)},
            refreshToken: ${JSON.stringify(refreshToken)},
            expiresIn: 3600,
            refreshTokenExpiresIn: 30 * 24 * 3600,
            capturedAt: Date.now()
          });
          return {
            rt: _rt, rtExp: _rtExp, sydney: _sydney.token,
            reauth: m365ReauthRequired, clientId: _clientId
          };
        `);
      await h.bg.run(`
        _rtExp = 0;
        _rt = "before-sibling";
        _sydney = { token: "cached-at", exp: Date.now() + 45 * 60 * 1000 };
        m365ReauthRequired = true;
        m365ReauthReason = "harness";
        return true;
      `);
      const r1 = await apply(CLIENT_ID, sibling, "sibling-rt");
      check("同 clientId 的兄弟轮换被采纳（_rt 更新）", r1.rt === "sibling-rt", `rt=${r1.rt}`);
      check(
        "未知初值被夹到 24h 以内（30 天策略不会污染倒计时）",
        r1.rtExp - Date.now() <= 24 * 3600 * 1000 + 5000,
        `remainMin=${Math.round((r1.rtExp - Date.now()) / 60000)}`,
      );
      check("没有直接清 reauth 标（清标只认真刷出来的 sydney 令牌）", r1.reauth === true, `reauth=${r1.reauth}`);
      check("把 _sydney 置空，逼下一轮真刷", r1.sydney === "", `sydney=${r1.sydney}`);
      const r2 = await apply("00000000-0000-0000-0000-000000000000", sibling, "foreign-rt");
      check("不同 clientId 的兄弟轮换被忽略", r2.rt === "sibling-rt", `rt=${r2.rt}`);
      const r3 = await apply(CLIENT_ID, sibling, "duplicate-rt");
      check("同一条 accessToken 第二次到达被去重（双通道不重复写）", r3.rt === "sibling-rt", `rt=${r3.rt}`);
      const sibB = token("https://clients.config.office.net");
      await apply(CLIENT_ID, sibB, "sibling-b-rt");
      const rInterleavedA = await apply(CLIENT_ID, sibling, "interleaved-dup-a");
      const rInterleavedB = await apply(CLIENT_ID, sibB, "interleaved-dup-b");
      check(
        "A -> B -> A -> B 双通道交错到达时，A 和 B 的重复副本均被有界 Set 去重",
        rInterleavedA.rt === "sibling-b-rt" && rInterleavedB.rt === "sibling-b-rt",
        `rtA=${rInterleavedA.rt} rtB=${rInterleavedB.rt}`,
      );
      // 有界缓存的淘汰行为：灌满 16 条后，最旧的一条被挤出，再次到达要当新响应处理。
      const evicted = await h.bg.run(`
        _recentAppliedM365AccessTokens.clear();
        const base = ${JSON.stringify(token("https://api.spaces.skype.com"))};
        const t0 = Date.now();
        const push = (suffix, rt) =>
          applyM365TokenResponse({
            tenantId: ${JSON.stringify(TENANT)},
            clientId: ${JSON.stringify(CLIENT_ID)},
            grantType: "refresh_token",
            source: "page",
            accessToken: base + "-" + suffix,
            refreshToken: rt,
            refreshTokenExpiresIn: 70000,
            requestStartedAt: t0,
            capturedAt: t0
          });
        for (let i = 0; i < 16; i++) await push("fill-" + i, "fill-rt-" + i);
        const sizeAfterFill = _recentAppliedM365AccessTokens.size;
        await push("fill-16", "fill-rt-16");
        const sizeAfterEvict = _recentAppliedM365AccessTokens.size;
        const oldestStillCached = _recentAppliedM365AccessTokens.has(base + "-fill-0");
        const rtBeforeReplay = _rt;
        await push("fill-0", "fill-rt-0-replay");
        return {
          sizeAfterFill,
          sizeAfterEvict,
          oldestStillCached,
          rtBeforeReplay,
          rtAfterReplay: _rt,
        };
      `);
      check(
        "去重缓存严格有界（16 条，加满后仍是 16 条）",
        evicted.sizeAfterFill === 16 && evicted.sizeAfterEvict === 16,
        `fill=${evicted.sizeAfterFill} evict=${evicted.sizeAfterEvict}`,
      );
      check(
        "最旧一条被淘汰出缓存",
        evicted.oldestStillCached === false,
        `oldestStillCached=${evicted.oldestStillCached}`,
      );
      check(
        "被淘汰的 token 再次到达会被重新应用（不是永久去重）",
        evicted.rtBeforeReplay === "fill-rt-16" &&
          evicted.rtAfterReplay === "fill-rt-0-replay",
        `before=${evicted.rtBeforeReplay} after=${evicted.rtAfterReplay}`,
      );
      const jweToken = [
        b64url({ alg: "RSA-OAEP", enc: "A128CBC-HS256", zip: "DEF" }),
        "encrypted-cek",
        "iv-bytes",
        "ciphertext-payload",
        "auth-tag",
      ].join(".");
      await h.bg.run(`
        m365AutoLoginEnabled = true;
        m365AutoLoginRemember = true;
        m365AutoLoginPassword = "";
        _stagedAutoLoginCreds = {
          username: "staged@contoso.com",
          password: "staged-secret-password",
          stagedAt: Date.now(),
        };
        return true;
      `);
      const rJwe = await h.bg.run(`
        await applyM365TokenResponse({
          tenantId: ${JSON.stringify(TENANT)},
          clientId: ${JSON.stringify(CLIENT_ID)},
          grantType: "refresh_token",
          source: "page",
          accessToken: ${JSON.stringify(jweToken)},
          refreshToken: "jwe-rotated-rt",
          expiresIn: 3600,
          refreshTokenExpiresIn: 75017,
          capturedAt: Date.now()
        });
        return {
          rt: _rt,
          savedPwd: m365AutoLoginPassword,
          stagedStillPresent: Boolean(_stagedAutoLoginCreds && _stagedAutoLoginCreds.password === "staged-secret-password"),
        };
      `);
      check("同 clientId + 非通配符租户的 5 段 JWE 轮换被采纳", rJwe.rt === "jwe-rotated-rt", `rt=${rJwe.rt}`);
      check(
        "JWE 轮换不落库、也不清空暂存的自动登录密码（身份不可读时不碰身份状态）",
        rJwe.savedPwd === "" && rJwe.stagedStillPresent === true,
        `savedPwd=${rJwe.savedPwd} staged=${rJwe.stagedStillPresent}`,
      );
      const rJweWildcard = await h.bg.run(`
        await applyM365TokenResponse({
          tenantId: "organizations",
          clientId: ${JSON.stringify(CLIENT_ID)},
          grantType: "refresh_token",
          source: "page",
          accessToken: ${JSON.stringify(jweToken + "x")},
          refreshToken: "jwe-wildcard-rt",
          expiresIn: 3600,
          refreshTokenExpiresIn: 75017,
          capturedAt: Date.now()
        });
        return { rt: _rt };
      `);
      check("通配符租户（organizations/common）下的 JWE 轮换被拒绝", rJweWildcard.rt === "jwe-rotated-rt", `rt=${rJweWildcard.rt}`);
      const r4 = await apply(CLIENT_ID, token("https://substrate.office.com/sydney"), "sydney-rt");
      check("真 Sydney 响应仍然直接落库并清标", r4.rt === "sydney-rt" && r4.reauth === false, `rt=${r4.rt} reauth=${r4.reauth}`);

      // 乱序到达：新一代先落地，旧一代迟到 —— 旧 RT 绝不允许覆盖新一代；
      // 同批并发（250ms 窗口内）放行但水位使用 Math.max 不倒退；
      // 请求体捕获（applyCapturedM365Auth）已有 _rt 时只更新上下文、不回退 _rt。
      const ordered = await h.bg.run(`
        _rtNewestObservedAt = 0;
        _recentAppliedM365AccessTokens.clear();
        const sib = ${JSON.stringify(token("https://api.spaces.skype.com"))};
        const newer = sib + "new";
        const sameBatch = sib + "batch";
        const older = sib + "old";
        const t0 = Date.now();
        await applyM365TokenResponse({
          tenantId: ${JSON.stringify(TENANT)},
          clientId: ${JSON.stringify(CLIENT_ID)},
          grantType: "refresh_token",
          source: "page",
          accessToken: newer,
          refreshToken: "gen-new",
          refreshTokenExpiresIn: 70000,
          requestStartedAt: t0 + 0.8,
          capturedAt: t0
        });
        const watermarkAfterNew = _rtNewestObservedAt;
        await applyM365TokenResponse({
          tenantId: ${JSON.stringify(TENANT)},
          clientId: ${JSON.stringify(CLIENT_ID)},
          grantType: "refresh_token",
          source: "page",
          accessToken: sameBatch,
          refreshToken: "gen-same-batch",
          refreshTokenExpiresIn: 70000,
          requestStartedAt: t0 - 20,
          capturedAt: t0 + 10
        });
        const watermarkAfterBatch = _rtNewestObservedAt;
        await applyM365TokenResponse({
          tenantId: ${JSON.stringify(TENANT)},
          clientId: ${JSON.stringify(CLIENT_ID)},
          grantType: "refresh_token",
          source: "page",
          accessToken: older,
          refreshToken: "gen-old",
          refreshTokenExpiresIn: 70005,
          requestStartedAt: t0 - 5000,
          capturedAt: Date.now()
        });
        await applyCapturedM365Auth({
          tenantId: ${JSON.stringify(TENANT)},
          clientId: ${JSON.stringify(CLIENT_ID)},
          refreshToken: "stale-request-body-rt",
          isSydneyScope: true,
          grantType: "refresh_token",
          tokenContext: ${JSON.stringify(SEED_CONTEXT)},
          capturedAt: Date.now() + 10000
        });
        return {
          rt: _rt,
          watermarkAfterNew,
          watermarkAfterBatch,
        };
      `);
      check(
        "同批并发请求（250ms 内）正常采纳，且迟到的旧一代 RT（-5000ms）被拒绝，请求体捕获也不覆盖已有 _rt",
        ordered.rt === "gen-same-batch",
        `rt=${ordered.rt}`,
      );
      check(
        "同批稍早请求的响应放行后，_rtNewestObservedAt 水位保持单调不退（Math.max + Math.floor）",
        ordered.watermarkAfterBatch === ordered.watermarkAfterNew &&
          Number.isInteger(ordered.watermarkAfterBatch),
        `before=${ordered.watermarkAfterNew} after=${ordered.watermarkAfterBatch}`,
      );
      check(
        "日志写明拒绝旧一代 RT",
        h.logSink.some((l) => /refused an older RT generation/.test(l)),
        JSON.stringify(h.logSink.slice(-3)),
      );
    },
  );

  // 17. 回归：面板开关负/正向对照、恢复超时真定时器、提交 25s 超时、子/顶层 frame 保活对照
  await scenario(
    "17. 面板开关保存负/正向对照 + 恢复/提交超时真定时器推进 terminal + 子/顶层 frame 保活对照",
    { sso: "expired", autofill: { mode: "immediate" } },
    async (h) => {
      // ① 面板仅切换开关（unpause: false）或关闭自动登录：不重置 terminal、不弹窗
      await h.bg.run(`
        m365ReauthRequired = true;
        m365ReauthReason = "expired";
        await setM365ReauthState("terminal", "prior-failure");
        return true;
      `);
      const tabsBeforeToggle = h.state.createdTabs.length;
      await h.bg.sendMessage({
        type: "POPUP_SAVE_M365_AUTO_LOGIN",
        enabled: true,
        rememberOnManualLogin: true,
        unpause: false,
      });
      await h.bg.sendMessage({
        type: "POPUP_SAVE_M365_AUTO_LOGIN",
        enabled: false,
        rememberOnManualLogin: true,
        unpause: true,
      });
      await sleep(80);
      const afterToggle = await h.snapshot();
      check(
        "仅切换面板开关（unpause:false）或关闭自动重登时不重置 terminal、不弹窗",
        afterToggle.state === "terminal" &&
          h.state.createdTabs.length === tabsBeforeToggle,
        `state=${afterToggle.state} tabs=${h.state.createdTabs.length}`,
      );

      // ④ 正向对照：显式保存密码并启用（unpause: true, enabled: true）→ 重置 terminal 并弹窗完成续期
      await h.bg.sendMessage({
        type: "POPUP_SAVE_M365_AUTO_LOGIN",
        enabled: true,
        rememberOnManualLogin: false,
        password: PASSWORD,
        unpause: true,
      });
      const afterSave = await h.waitFor(
        (s) => s.state === "none" && s.rt === "rotated-rt",
        { timeout: 8000 },
      );
      check(
        "面板显式保存密码（unpause:true + enabled:true）重置 terminal 并开窗完成续期",
        afterSave.state === "none" &&
          afterSave.rt === "rotated-rt" &&
          h.state.createdTabs.length === tabsBeforeToggle + 1,
        `state=${afterSave.state} rt=${afterSave.rt} tabs=${h.state.createdTabs.length}`,
      );

      // ② restorePendingReauth 恢复 windowId / promptNone，且不 clearTimeout，等 5s 真定时器自然触发进入 terminal
      const restored = await h.bg.run(`
        m365AutoLoginEnabled = true;
        await setM365ReauthState("attempting", "before-sw-restart");
        await persistPendingReauth({
          tabId: 4321,
          windowId: 8765,
          state: "ve-reauth:restored-test",
          verifier: "verifier-123",
          redirectUri: ${JSON.stringify(BROKER_REDIRECT)},
          clientId: ${JSON.stringify(CLIENT_ID)},
          scope: ${JSON.stringify(SEED_CONTEXT.body.scope)},
          brkClientId: ${JSON.stringify(BRK_CLIENT_ID)},
          brkRedirect: ${JSON.stringify(WEB_REDIRECT)},
          interactive: true,
          promptNone: true,
          startedAt: Date.now() - 44000,
        });
        _pendingTopLevelReauth = null;
        await restorePendingReauth();
        return {
          windowId: _pendingTopLevelReauth && _pendingTopLevelReauth.windowId,
          promptNone: _pendingTopLevelReauth && _pendingTopLevelReauth.promptNone,
        };
      `);
      check(
        "restorePendingReauth 完整恢复 windowId 与 promptNone",
        restored.windowId === 8765 && restored.promptNone === true,
        JSON.stringify(restored),
      );
      await sleep(5200);
      const afterRestoreTimeout = await h.bg.run(`
        return {
          state: m365ReauthState,
          reason: m365ReauthStateReason,
          pending: _pendingTopLevelReauth,
        };
      `);
      check(
        "恢复的登录页经真实定时器超时后自然落入 terminal（不卡在 attempting）",
        afterRestoreTimeout.state === "terminal" &&
          /restored login tab timed out/.test(afterRestoreTimeout.reason) &&
          afterRestoreTimeout.pending === null,
        JSON.stringify(afterRestoreTimeout),
      );

      // ②-b M365_AUTO_LOGIN_SUBMITTING 重置 25s 定时器，超时回调自然触发后落入 terminal
      const submitTimeoutRes = await h.bg.run(`
        await setM365ReauthState("attempting", "submitting-test");
        const dummyTimer = setTimeout(() => {}, 60000);
        _pendingTopLevelReauth = {
          tabId: 5432,
          windowId: null,
          state: "ve-reauth:submit-timeout",
          verifier: "v",
          redirectUri: ${JSON.stringify(BROKER_REDIRECT)},
          clientId: ${JSON.stringify(CLIENT_ID)},
          scope: ${JSON.stringify(SEED_CONTEXT.body.scope)},
          brkClientId: ${JSON.stringify(BRK_CLIENT_ID)},
          brkRedirect: ${JSON.stringify(WEB_REDIRECT)},
          interactive: true,
          promptNone: false,
          timer: dummyTimer,
          startedAt: Date.now(),
        };
        armM365LoginTab(5432, Date.now());
        const origSetTimeout = globalThis.setTimeout;
        let capturedSubmitDelay = 0;
        globalThis.setTimeout = (fn, ms, ...rest) => {
          if (ms === 25000) {
            capturedSubmitDelay = ms;
            return origSetTimeout(fn, 60, ...rest);
          }
          return origSetTimeout(fn, ms, ...rest);
        };
        globalThis.__restoreSetTimeout = () => {
          globalThis.setTimeout = origSetTimeout;
          return capturedSubmitDelay;
        };
        return true;
      `);
      void submitTimeoutRes;
      let capturedDelay = 0;
      try {
        await h.bg.sendMessage(
          {
            __veM365: true,
            type: "M365_AUTO_LOGIN_SUBMITTING",
            username: UPN,
          },
          {
            tab: {
              id: 5432,
              url: `https://login.microsoftonline.com/${TENANT}/login`,
            },
            frameId: 0,
            url: `https://login.microsoftonline.com/${TENANT}/login`,
          },
        );
      } finally {
        capturedDelay = await h.bg.run(`return globalThis.__restoreSetTimeout();`);
      }
      await sleep(160);
      const afterSubmitTimeout = await h.bg.run(`
        return {
          state: m365ReauthState,
          reason: m365ReauthStateReason,
          pending: _pendingTopLevelReauth,
        };
      `);
      check(
        "M365_AUTO_LOGIN_SUBMITTING 挂载 25s 定时器且超时后真实回调落入 terminal",
        capturedDelay === 25000 &&
          afterSubmitTimeout.state === "terminal" &&
          /login tab timed out 25s after submit/.test(afterSubmitTimeout.reason) &&
          afterSubmitTimeout.pending === null,
        JSON.stringify({ capturedDelay, ...afterSubmitTimeout }),
      );

      // bridge 显式传 force:false 时关窗也不会把状态卡在 attempting
      const bridgeNonForce = await h.bg.run(`
        await setM365ReauthState("attempting", "stale-attempt");
        m365AutoLoginEnabled = false;
        await handleM365BridgeReauthCommand({ force: false, requestId: "nf-1" });
        return { state: m365ReauthState };
      `);
      check(
        "bridge force:false 关窗后不会把状态滞留在 attempting",
        bridgeNonForce.state === "none",
        JSON.stringify(bridgeNonForce),
      );

      // ③ 子 frame vs 顶层 frame 收到同一 M365_KEEPALIVE 的负/正向对照
      const parentDom = new JSDOM(
        `<!doctype html><html><head></head><body><iframe id="sub"></iframe></body></html>`,
        {
          url: "https://outlook.cloud.microsoft/mail/",
          runScripts: "dangerously",
          pretendToBeVisual: true,
        },
      );
      const topWin = parentDom.window;
      const subWin = topWin.document.getElementById("sub").contentWindow;
      const wireFrame = (win, bucket) => {
        win.Element.prototype.getBoundingClientRect = () => ({
          width: 120,
          height: 32,
          top: 0,
          left: 0,
          right: 120,
          bottom: 32,
        });
        win.fetch = async () => ({
          ok: true,
          status: 200,
          json: async () => ({}),
          text: async () => "",
          clone() {
            return this;
          },
        });
        win.browser = {
          runtime: {
            id: "test",
            sendMessage: (m) => {
              bucket.push(m);
              return Promise.resolve({ ok: true });
            },
            getURL: (p) => "moz-extension://test/" + p,
            onMessage: { addListener: () => {} },
          },
          storage: {
            local: { get: () => Promise.resolve({}), set: () => Promise.resolve() },
          },
        };
        win.eval(readFileSync(join(ROOT, "src/content-m365.js"), "utf8"));
      };
      const topPosted = [];
      const subPosted = [];
      wireFrame(topWin, topPosted);
      wireFrame(subWin, subPosted);
      await sleep(120);
      const keepAliveMsg = {
        __veM365: true,
        dir: "toPage",
        type: "M365_KEEPALIVE",
        authHealthy: true,
      };
      topWin.postMessage(keepAliveMsg, "*");
      subWin.postMessage(keepAliveMsg, "*");
      await sleep(120);
      check(
        "顶层 frame 收到 M365_KEEPALIVE 正常上报 M365_SESSION_RENEWED（正向对照）",
        topWin === topWin.top &&
          topPosted.some((m) => m && m.type === "M365_SESSION_RENEWED"),
        JSON.stringify(topPosted.map((m) => m && m.type)),
      );
      check(
        "子 frame 已启动 PAGE_HOOK 且无 AbtInteraction 时不虚报 M365_SESSION_RENEWED（负向对照）",
        subWin !== subWin.top &&
          Boolean(subWin.__veM365Hook && subWin.__veM365Hook.ready) &&
          !subPosted.some((m) => m && m.type === "M365_SESSION_RENEWED"),
        JSON.stringify(subPosted.map((m) => m && m.type)),
      );
      parentDom.window.close();
    },
  );

  console.log(
    `\n${fail ? "FAILED" : "ALL PASSED"} — ${pass} passed, ${fail} failed`,
  );
  process.exit(fail ? 1 : 0);
}

// ---------------------------------------------------------------------------
// 6. 可视化模式：node test/reauth.mjs --visual
//    真的在屏幕上打开一个登录窗口（真实 DOM、真实 src/content-m365-login.js），
//    后台仍是真实 src/background.js。你看到的是：密码被填 → 自动点登录 →
//    回跳 #code → 后台换到新的 24h 令牌。
// ---------------------------------------------------------------------------
function findBrowser() {
  const list = [];
  const root = join(homedir(), "Library/Caches/ms-playwright");
  if (existsSync(root)) {
    for (const entry of readdirSync(root)) {
      if (!entry.startsWith("chromium-")) continue;
      for (const arch of ["chrome-mac-arm64", "chrome-mac-x64", "chrome-mac"]) {
        list.push(
          join(
            root,
            entry,
            arch,
            "Google Chrome for Testing.app/Contents/MacOS/Google Chrome for Testing",
          ),
        );
        list.push(
          join(root, entry, arch, "Chromium.app/Contents/MacOS/Chromium"),
        );
      }
    }
  }
  list.push("/Applications/Chromium.app/Contents/MacOS/Chromium");
  list.push("/Applications/Google Chrome.app/Contents/MacOS/Google Chrome");
  return list.find((candidate) => existsSync(candidate)) || "";
}

function authorizeHtml(cfg) {
  return `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8">
<title>登录 - Microsoft 在线服务</title>
<style>
  *{box-sizing:border-box}
  body{margin:0;font-family:-apple-system,"Segoe UI",sans-serif;background:#f2f2f2;color:#1b1b1b}
  .bar{height:36px;background:#e6e6e6;border-bottom:1px solid #d0d0d0;display:flex;align-items:center;gap:8px;padding:0 12px;font-size:12px;color:#444}
  .bar .url{flex:1;background:#fff;border:1px solid #cfcfcf;border-radius:10px;padding:4px 10px;overflow:hidden;white-space:nowrap;text-overflow:ellipsis;color:#555}
  .card{max-width:420px;margin:40px auto;background:#fff;padding:34px 38px 26px;box-shadow:0 2px 12px rgba(0,0,0,.12);min-height:230px}
  .ms{color:#5b5b5b;font-weight:600;font-size:13px;letter-spacing:.2px}
  h1{font-size:24px;font-weight:600;margin:14px 0 6px}
  .sub{color:#555;font-size:13px;margin-bottom:20px}
  .hint{font-size:12px;color:#0f6cbd;background:#eef6fd;border:1px solid #cfe7fb;padding:8px 10px;margin:12px 0}
  input[type=password]{width:100%;height:38px;border:1px solid #8a8a8a;border-radius:2px;padding:0 10px;font-size:15px}
  input[type=password]:focus{outline:2px solid #0f6cbd;border-color:#0f6cbd}
  button{width:100%;height:38px;margin-top:16px;background:#0f6cbd;color:#fff;border:0;font-size:15px;cursor:pointer}
  button:active{background:#0b5a9e}
  #passwordError{display:none;margin-top:14px;color:#a4262c;font-size:13px}
  .mfa{margin-top:16px;padding:14px;background:#f3f2f1;border:1px dashed #8a8886;font-size:14px}
  #tilesHolder .table{margin-top:12px;padding:10px 12px;border:1px solid #d0d0d0;cursor:pointer;font-size:14px}
  #KmsiCheckboxField{margin-top:14px}
  .trace{margin:20px auto 36px;max-width:520px;font:12px/1.5 ui-monospace,Menlo,monospace;color:#333;background:#fff;border:1px solid #ddd;padding:12px 14px;max-height:200px;overflow:auto}
</style>
<script>window.$Config=${JSON.stringify({ pgid: "ConvergedSignIn", sPOST_Username: cfg.upn })};</script>
<script>window.__VEV=${JSON.stringify(cfg)};</script>
</head><body>
<div class="bar"><span>🔒</span><span class="url">login.microsoftonline.com/${cfg.tenant}/oauth2/v2.0/authorize?client_id=${cfg.clientId}&amp;redirect_uri=${cfg.redirectUri || ""}&amp;state=ve-reauth…</span></div>
<div class="card" id="card"></div>
<div class="trace" id="trace"></div>
<script>
(function () {
  var V = window.__VEV;
  var card = document.getElementById("card");
  var traceEl = document.getElementById("trace");
  var done = false;
  var submitting = false;

  function trace(text) {
    var line = document.createElement("div");
    line.textContent = new Date().toISOString().slice(11, 23) + "  " + text;
    traceEl.appendChild(line);
    traceEl.scrollTop = traceEl.scrollHeight;
  }
  function hint(text) {
    var h = document.getElementById("hint");
    if (h) h.textContent = text;
  }
  window.__veTrace = trace;

  window.browser = {
    runtime: {
      sendMessage: function (message) {
        trace("内容脚本 → 后台: " + message.type);
        return fetch("/api/message", {
          method: "POST",
          headers: { "content-type": "application/json" },
          body: JSON.stringify({
            message: message,
            sender: { tab: { id: V.tabId, url: V.senderUrl }, frameId: 0, url: V.senderUrl }
          })
        })
          .then(function (r) { return r.json(); })
          .then(function (body) { return body.result === null ? undefined : body.result; });
      }
    }
  };

  trace("登录窗口已加载 mode=" + V.mode);

  function passwordPage() {
    card.innerHTML =
      '<div class="ms">Microsoft</div><h1>登录</h1>' +
      '<div class="sub">继续前往 Microsoft 365 Copilot</div>' +
      '<div id="hint" class="hint">等待浏览器密码库填充…</div>' +
      '<form onsubmit="return false"><input id="i0118" name="passwd" type="password" autocomplete="current-password" placeholder="密码">' +
      '<button id="idSIButton9" type="button">登录</button></form>' +
      '<div id="passwordError"></div>';
    document.getElementById("idSIButton9").addEventListener("click", submit);
    if (V.fillMs < 600000) setTimeout(fill, V.fillMs);
  }

  function fill() {
    var input = document.getElementById("i0118");
    if (!input) return;
    input.value = V.wrong ? "definitely-wrong-password" : V.password;
    input.dispatchEvent(new Event("input", { bubbles: true }));
    input.dispatchEvent(new Event("change", { bubbles: true }));
    trace("浏览器密码库填充 #i0118" + (V.wrong ? "（故意错的密码）" : ""));
    hint("密码库已填充，等待内容脚本自动提交…");
  }

  function submit() {
    if (submitting || done) return;
    submitting = true;
    var input = document.getElementById("i0118");
    var value = input ? input.value : "";
    trace("AAD 收到提交，校验密码…");
    setTimeout(function () {
      submitting = false;
      if (value !== V.password) return showError();
      if (V.mode === "mfa") return showMfa();
      if (V.mode === "kmsi") return showKmsi();
      redirect();
    }, 400);
  }

  function showError() {
    var box = document.getElementById("passwordError");
    if (box) {
      box.style.display = "block";
      box.textContent = "AADSTS50126: Error validating credentials due to invalid username or password.";
    }
    trace("密码错误 → AADSTS50126");
    hint("登录失败：密码错误");
  }

  function showMfa() {
    done = true;
    card.innerHTML =
      '<div class="ms">Microsoft</div><h1>验证你的身份</h1>' +
      '<div class="sub">请在 Microsoft Authenticator 中批准请求</div>' +
      '<div id="idDiv_SAOTCS_Proofs" class="mfa">等待批准…</div>';
    trace("渲染二次验证页 (#idDiv_SAOTCS_Proofs)");
  }

  function showKmsi() {
    card.innerHTML =
      '<div class="ms">Microsoft</div><h1>是否保持登录？</h1>' +
      '<div id="KmsiInterrupt" class="sub">减少登录次数</div>' +
      '<input id="KmsiCheckboxField" name="DontShowAgain" type="checkbox">' +
      '<button id="idSIButton9" type="button">是</button>';
    document.getElementById("idSIButton9").addEventListener("click", function () {
      trace("KMSI 已勾选并提交");
      redirect();
    });
    trace("渲染 KMSI 页");
  }

  function showPicker() {
    card.innerHTML =
      '<div class="ms">Microsoft</div><h1>选择账户</h1>' +
      '<div id="tilesHolder"><div class="tile-container"><div class="table" role="button" data-test-id="' +
      V.upn + '">' + V.upn + '</div></div></div>';
    trace("渲染账号选择器 (#tilesHolder)");
    var tile = card.querySelector("[data-test-id]");
    if (tile) {
      tile.addEventListener("click", function () {
        trace("脚本点击账号 " + V.upn);
        passwordPage();
      });
    }
  }

  function showInteraction() {
    card.innerHTML =
      '<div class="ms">Microsoft</div><h1>无法静默登录</h1>' +
      '<div id="passwordError" style="display:block">interaction_required (AADSTS50058)</div>';
    trace("AAD 拒绝：interaction_required");
    fetch("/api/message", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        message: {
          __veM365: true,
          type: "M365_OAUTH_REDIRECT",
          ok: false,
          error: "interaction_required",
          errorDescription: "AADSTS50058: A silent sign-in request was sent but no user is signed in."
        },
        sender: {
          tab: { id: V.tabId, url: "https://outlook.cloud.microsoft/mail/oauthRedirect.html" },
          frameId: 0,
          url: "https://outlook.cloud.microsoft/mail/oauthRedirect.html"
        }
      })
    });
  }

  function redirect() {
    if (done) return;
    done = true;
    trace("密码正确 → 302 跳回 redirect_uri");
    hint("登录成功，正在跳回应用…");
    location.href = "/oauthRedirect?h=" + V.tabId + "#code=mock-code-" + V.tabId + "-1&state=" + encodeURIComponent(V.state);
  }

  if (V.mode === "pass") {
    trace("SSO 仍有效，直接回跳");
    setTimeout(redirect, 80);
  } else if (V.mode === "interaction") {
    showInteraction();
  } else if (V.mode === "picker") {
    showPicker();
  } else {
    passwordPage();
  }
})();
</script>
<script src="/content-script.js"></script>
</body></html>`;
}

function redirectHtml() {
  return `<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><title>正在完成登录…</title>
<style>
  body{font-family:-apple-system,"Segoe UI",sans-serif;display:flex;align-items:center;justify-content:center;height:100vh;margin:0;background:#f2f2f2}
  .box{background:#fff;padding:28px 34px;box-shadow:0 2px 12px rgba(0,0,0,.12);min-width:360px}
  h1{font-size:18px;margin:0 0 12px}
  .row{font:12px/1.7 ui-monospace,Menlo,monospace;color:#444}
  .ok{color:#107c10;font-weight:600}
  .err{color:#a4262c;font-weight:600}
</style></head><body>
<div class="box"><h1>oauthRedirect.html</h1><div class="row" id="row">收到授权码，正在交给后台…</div></div>
<script>
(function () {
  var q = new URLSearchParams(location.search);
  var tabId = Number(q.get("h") || 0);
  var hash = new URLSearchParams(location.hash.slice(1));
  var code = hash.get("code") || "";
  var state = hash.get("state") || "";
  var row = document.getElementById("row");
  var url = "https://outlook.cloud.microsoft/mail/oauthRedirect.html";
  row.textContent = "authorization code = " + code;

  fetch("/api/message", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({
      message: { __veM365: true, type: "M365_OAUTH_REDIRECT", ok: true, code: code, state: state },
      sender: { tab: { id: tabId, url: url }, frameId: 0, url: url }
    })
  })
    .then(function (r) { return r.json(); })
    .then(function () { return poll(0); });

  function poll(n) {
    fetch("/api/state").then(function (r) { return r.json(); }).then(function (s) {
      if (s.state === "none" && s.rt === "rotated-rt") {
        row.innerHTML = '<span class="ok">✅ 后台已用授权码换到新的 24h 刷新令牌</span><br>RT 剩余 ' + Math.round(s.rtRemainMin) + " 分钟";
        document.title = "登录完成";
        return;
      }
      if (s.state === "terminal") {
        row.innerHTML = '<span class="err">❌ 后台拒绝：' + (s.reason || "unknown") + "</span>";
        return;
      }
      if (n < 40) setTimeout(function () { poll(n + 1); }, 300);
      else row.textContent += " · 等待后台换取令牌超时";
    });
  }
})();
</script></body></html>`;
}

function startVisualServer(link) {
  const server = http.createServer((req, res) => {
    const url = new URL(req.url, "http://127.0.0.1");
    if (url.pathname !== "/api/state")
      console.log(`  [http] ${req.method} ${url.pathname}`);
    const send = (code, body, type) => {
      res.writeHead(code, {
        "content-type": type + "; charset=utf-8",
        "cache-control": "no-store",
      });
      res.end(body);
    };
    if (url.pathname === "/content-script.js") {
      return send(200, CONTENT_SCRIPT, "application/javascript");
    }
    if (url.pathname === "/authorize") {
      return send(
        200,
        authorizeHtml({
          tabId: Number(url.searchParams.get("h") || 0),
          state: url.searchParams.get("state") || "",
          tenant: TENANT,
          upn: UPN,
          password: PASSWORD,
          // 页面模式必须从 URL 带下来，否则每种情况都渲染成默认密码页
          mode: url.searchParams.get("mode") || "expired",
          wrong: url.searchParams.get("wrong") === "1",
          redirectUri: url.searchParams.get("redirect") || CAPTURED_REDIRECT,
          fillMs: Number(url.searchParams.get("fill") || 1200),
          clientId: url.searchParams.get("client_id") || CLIENT_ID,
          senderUrl:
            "https://login.microsoftonline.com/" +
            TENANT +
            "/oauth2/v2.0/authorize?state=" +
            encodeURIComponent(url.searchParams.get("state") || ""),
        }),
        "text/html",
      );
    }
    if (url.pathname === "/oauthRedirect")
      return send(200, redirectHtml(), "text/html");
    if (url.pathname === "/api/log" && req.method === "POST") {
      let raw = "";
      req.on("data", (c) => (raw += c));
      req.on("end", () => {
        try {
          console.log("  [页面日志] " + (JSON.parse(raw || "{}").text || ""));
        } catch (_) {}
        send(200, "{}", "application/json");
      });
      return;
    }
    if (url.pathname === "/api/state") {
      return link
        .snapshot()
        .then((snap) => send(200, JSON.stringify(snap), "application/json"));
    }
    if (url.pathname === "/api/message" && req.method === "POST") {
      let raw = "";
      req.on("data", (chunk) => (raw += chunk));
      req.on("end", async () => {
        let payload = {};
        try {
          payload = JSON.parse(raw || "{}");
        } catch (_) {}
        const message = payload.message || {};
        console.log(`  [页面] ${message.type || "?"}`);
        const result = await link.send(message, payload.sender || {});
        send(
          200,
          JSON.stringify({ result: result === undefined ? null : result }),
          "application/json",
        );
      });
      return;
    }
    send(404, "not found", "text/plain");
  });
  return new Promise((resolve) => {
    server.listen(0, "127.0.0.1", () =>
      resolve({ server, port: server.address().port }),
    );
  });
}

const VISUAL_CASES = [
  {
    mode: "pass",
    expect: "success",
    title: "SSO 仍有效 → 直接回跳，不渲染密码页",
  },
  {
    mode: "expired",
    fill: 300,
    expect: "success",
    title: "密码库立即填充 → 自动提交",
  },
  {
    mode: "expired",
    fill: 1500,
    expect: "success",
    title: "密码库晚 1.5 秒填充 → 仍在等待窗口内，自动提交",
  },
  {
    mode: "expired",
    fill: 999999,
    expect: "failure",
    title: "密码库始终不填充 → 失败、关窗、通知（不等人）",
  },
  {
    mode: "expired",
    fill: 300,
    wrong: true,
    expect: "failure",
    title: "密码错误 → AADSTS50126，终态 + 暂停",
  },
  {
    mode: "mfa",
    fill: 300,
    expect: "failure",
    title: "需要 MFA → 终态、关窗、通知",
  },
  {
    mode: "kmsi",
    fill: 300,
    expect: "success",
    title: "KMSI 页 → 自动勾选并继续",
  },
  {
    mode: "picker",
    fill: 300,
    expect: "success",
    title: "账号选择器 → 点击账号 → 密码页",
  },
  {
    mode: "interaction",
    expect: "failure",
    title: "AAD 直接 interaction_required → 终态",
  },
];

function findBrowserOrDie() {
  const bin = findBrowser();
  if (!bin) {
    console.error("找不到 Chromium/Chrome for Testing，无法做可视化演示");
    process.exit(4);
  }
  return bin;
}

async function runVisualCase(testCase, index, total, link, port) {
  console.log(`\n=== [${index}/${total}] ${testCase.title}`);
  const logSink = [];
  const flow = { refreshOk: true };
  const net = createNetwork(flow);
  let bg = null;
  let browserProc = null;

  const stub = createBrowserStub(
    {
      onTabCreated(tab) {
        if (!/login\.microsoftonline\.com/.test(tab.url)) return;
        const parsed = new URL(tab.url);
        const pageUrl =
          "http://127.0.0.1:" +
          port +
          "/authorize" +
          "?h=" +
          tab.id +
          "&state=" +
          encodeURIComponent(parsed.searchParams.get("state") || "") +
          "&client_id=" +
          encodeURIComponent(
            parsed.searchParams.get("client_id") || CLIENT_ID,
          ) +
          "&mode=" +
          encodeURIComponent(testCase.mode) +
          "&fill=" +
          Number(testCase.fill || 300) +
          (testCase.wrong ? "&wrong=1" : "");
        const bin = findBrowserOrDie();
        browserProc = spawn(
          bin,
          [
            "--app=" + pageUrl,
            "--user-data-dir=" + join(tmpdir(), "ve-visual-" + Date.now()),
            "--no-first-run",
            "--no-default-browser-check",
            "--window-size=560,780",
            "--no-sandbox",
            "--disable-gpu",
            "--disable-dev-shm-usage",
          ],
          { stdio: "ignore", detached: true },
        );
        console.log(
          "  窗口已打开: " + pageUrl.replace(/state=[^&]+/, "state=…"),
        );
      },
      onTabRemoved() {},
    },
    { startTabId: 1000 + index * 100 },
  );
  const browser = Object.assign(stub.api, { __listeners: stub.listeners });
  bg = loadBackground({ browser, fetchImpl: net.fetchImpl, logSink });

  link.send = (message, sender) => bg.sendMessage(message, sender);
  link.snapshot = () =>
    bg.run(`
      return {
        state: m365ReauthState,
        reason: m365ReauthStateReason,
        rt: _rt,
        rtRemainMin: Math.round((_rtExp - Date.now()) / 60000),
      };
    `);

  await bg.run(`
    await Promise.all([ensureM365AuthLoaded(), ensureM365AutoLoginLoaded(), ensureM365ReauthStateLoaded()]);
    _tid = ${JSON.stringify(TENANT)};
    _clientId = ${JSON.stringify(CLIENT_ID)};
    _rt = "seed-refresh-token";
    _rtExp = Date.now() - 60 * 1000;
    _sydney = { token: "", exp: 0 };
    _m365TokenContext = ${JSON.stringify(SEED_CONTEXT)};
    m365AuthLoaded = true;
    m365AutoLoginEnabled = true;
    m365AutoLoginRemember = false;
    m365AutoLoginPassword = "";
    m365AutoLoginUsername = ${JSON.stringify(UPN)};
    m365AutoLoginPaused = false;
    await saveM365AutoLogin();
    await saveM365Auth();
    return true;
  `);
  const started = await bg.run(
    `return await tryTopLevelSilentReauth("visual");`,
  );
  if (!started) console.log("  （后台没有发起尝试，见上方日志）");

  const deadline = Date.now() + 12000;
  let snap = null;
  while (Date.now() < deadline) {
    snap = await link.snapshot();
    if (snap.state === "none" || snap.state === "terminal") break;
    await sleep(150);
  }

  const success = Boolean(
    snap && snap.state === "none" && snap.rt === "rotated-rt",
  );
  const matched = success === (testCase.expect !== "failure");
  if (success) {
    console.log(
      `  ✅ 成功：换到新的 24h 令牌（剩余 ${snap.rtRemainMin} 分钟）`,
    );
  } else if (snap && snap.state === "terminal") {
    console.log(`  ⛔ 终态（预期内）：${snap.reason}`);
  } else {
    console.log(`  ⚠ 超时未结束：state=${snap && snap.state}`);
  }

  if (browserProc) {
    const pid = browserProc.pid;
    const alive = () => {
      try {
        process.kill(pid, 0);
        return true;
      } catch (_) {
        return false;
      }
    };
    try {
      process.kill(-pid, "SIGTERM");
    } catch (_) {
      try {
        browserProc.kill("SIGTERM");
      } catch (_) {}
    }
    for (let i = 0; i < 12 && alive(); i++) await sleep(200);
    if (alive()) {
      try {
        process.kill(-pid, "SIGKILL");
      } catch (_) {}
      for (let i = 0; i < 8 && alive(); i++) await sleep(150);
    }
    browserProc = null;
  }
  console.log(
    `  窗口已关闭 · 结果${matched ? "符合预期" : "不符合预期（期望 " + testCase.expect + "）"}`,
  );
  await sleep(250);
  return { success, matched, snap };
}

async function visualDemo() {
  assertProductionBuild();
  const hold = process.argv.includes("--hold");
  console.log(
    `可视化演示：${VISUAL_CASES.length} 种情况逐个开窗跑一遍（每次跑完自动关窗）\n`,
  );

  const link = {};
  const started = await startVisualServer(link);
  const server = started.server;
  const port = started.port;

  const results = [];
  for (let i = 0; i < VISUAL_CASES.length; i++) {
    results.push(
      await runVisualCase(
        VISUAL_CASES[i],
        i + 1,
        VISUAL_CASES.length,
        link,
        port,
      ),
    );
  }

  const okCount = results.filter((r) => r.success).length;
  const terminalCount = results.filter(
    (r) => r.snap && r.snap.state === "terminal",
  ).length;
  const mismatched = results.filter((r) => !r.matched);
  console.log(
    `\n汇总：${VISUAL_CASES.length} 种情况 — 成功 ${okCount}，按设计失败(终态) ${terminalCount}，超时 ${VISUAL_CASES.length - okCount - terminalCount}，与预期不符 ${mismatched.length}`,
  );

  server.close();
  const shutdown = (code = 0) => process.exit(code);
  process.on("SIGINT", () => shutdown(0));
  if (hold) {
    console.log("--hold：保留最后一个窗口？此模式已逐案关窗，直接退出。");
  }
  shutdown(0);
}

// 构建指纹：从 dist/firefox 里读，用来判断浏览器里那份是不是最新。
function expectedBuildTag() {
  try {
    const src = readFileSync(join(ROOT, "dist/firefox/background.js"), "utf8");
    const m = src.match(/const M365_BUILD_TAG = "([^"]+)"/);
    return m ? m[1] : "";
  } catch (_) {
    return "";
  }
}

async function liveReauth() {
  assertProductionBuild();
  // 复用扩展唯一的对外通道（Voice Edge 的 ws://127.0.0.1:5003/ws），不再另开端口。
  // 5003 被 voice_edge 的 python 服务占着时必须先停它——同一个端口只能有一个
  // 监听者，而扩展只会连这一个地址。
  const PORT = 5003;
  const PATH = "/ws";
  const requestId = "harness-" + Date.now().toString(36);
  let wss;
  try {
    wss = new WebSocketServer({ host: "127.0.0.1", port: PORT, path: PATH });
    await new Promise((resolve, reject) => {
      wss.once("listening", resolve);
      wss.once("error", reject);
    });
  } catch (error) {
    const code = error && error.code;
    console.error(
      code === "EADDRINUSE"
        ? `端口 ${PORT} 已被占用（voice_edge 的 python 服务在跑）。先停掉它，扩展会在 30 秒内自动重连到本 harness。`
        : `无法监听 ws://127.0.0.1:${PORT}${PATH}: ${(error && error.message) || error}`,
    );
    try {
      wss && wss.close();
    } catch (_) {}
    process.exit(3);
  }
  console.log(`已监听 ${"ws://127.0.0.1:" + PORT + PATH}（复用扩展的既有通道）`);
  console.log("等待扩展连上来…（它自带 1s→30s 退避重连）\n");

  let commanded = false;
  let rtBefore = 0;
  let probeSent = false;
  let doneMsg = null;
  let activeWs = null;
  const expectedBuild = expectedBuildTag();
  let reloadTried = false;
  const result = await new Promise((resolve, reject) => {
    // probe-only 需要先等浏览器把扩展装好并连上来，不能 90 秒就放弃。
    const firstWaitMs = process.argv.includes("--probe-only") ? 12 * 60 * 1000 : 90000;
    let timer = setTimeout(
      () =>
        reject(
          new Error(
            `${Math.round(firstWaitMs / 1000)} 秒内没有等到扩展连接 / M365_REAUTH_DONE`,
          ),
        ),
      firstWaitMs,
    );
    const overallDeadline = Date.now() + 12 * 60 * 1000;
    const armProbeWait = () => {
      clearTimeout(timer);
      timer = setTimeout(() => {
        if (Date.now() > overallDeadline) {
          reject(
            new Error("5 分钟内页面仍没有回报探针（扩展没加载到新构建？）"),
          );
          return;
        }
        console.log(
          "   …15 秒仍无回报（浏览器里还是旧构建）；等你重载扩展，重连后会自动重试",
        );
        commanded = false;
        probeSent = false;
        // 扩展可能刚重连上（新构建），直接重发一次探针，不必等下一次 READY。
        try {
          if (activeWs && activeWs.readyState === 1) {
            console.log("   重发探针（扩展已重连）");
            probeSent = true;
            activeWs.send(JSON.stringify({ type: "M365_PAGE_TOKEN_PROBE" }));
          }
        } catch (_) {}
        armProbeWait();
      }, 15000);
    };
    wss.on("connection", (ws, req) => {
      activeWs = ws;
      console.log(`✅ 扩展已连接（${req && req.socket ? req.socket.remotePort : "?"}）`);
      ws.on("message", (data) => {
        let msg = null;
        try {
          msg = JSON.parse(String(data));
        } catch (_) {
          return;
        }
        if (msg.type === "M365_READY") {
          // M365_READY 是复用的：bridge 握手带 build（后台 4175），页面 frame
          // 报到带 frameOrigin/frameUrl（后台 5054）。后者没有 build，以前会被
          // 当成"旧构建"打印出来，白吓一跳。
          if (!msg.build && msg.frameOrigin !== undefined) {
            console.log(`   页面 frame 报到: ${msg.frameOrigin}`);
            return;
          }
          console.log(
            `   扩展构建: ${msg.build || "unknown（旧构建）"}  期望: ${expectedBuild || "-"}`,
          );
          if (
            !reloadTried &&
            expectedBuild &&
            msg.build &&
            msg.build !== expectedBuild &&
            (msg.caps || []).includes("reload")
          ) {
            reloadTried = true;
            commanded = false;
            probeSent = false;
            console.log("   → 构建不一致，发送 M365_RELOAD_SELF（扩展会自己重启加载新 dist）");
            ws.send(JSON.stringify({ type: "M365_RELOAD_SELF" }));
            return;
          }
          if (commanded) return;
          commanded = true;
          if (process.argv.includes("--probe-only")) {
            probeSent = true;
            console.log("   收到 M365_READY → 跳过续期，直接让页面发一次 /token 探针");
            ws.send(JSON.stringify({ type: "M365_PAGE_TOKEN_PROBE" }));
            armProbeWait();
            return;
          }
          // 给用户 3 秒准备时间，不然静默 SSO 下这个窗口只存在 2~4 秒，容易错过。
          console.log("   收到 M365_READY → 3 秒后打开登录窗口，请看屏幕…");
          setTimeout(() => {
            try {
              ws.send(
                JSON.stringify({ type: "M365_REAUTH_NOW", force: true, requestId }),
              );
            } catch (_) {}
          }, 3000);
          return;
        }
        if (msg.requestId && msg.requestId !== requestId) return;
        if (msg.type === "M365_REAUTH_STARTED") {
          rtBefore = Number(msg.rtExpiresAt || 0);
          console.log(
            `   已发起: started=${msg.started} state=${msg.state} rtRemain=${msg.rtRemainMin}min rtExpiresAt=${rtBefore}`,
          );
        } else if (msg.type === "M365_REAUTH_STATE") {
          console.log(
            `   … state=${msg.state} rtRemain=${msg.rtRemainMin}min pendingTab=${msg.pendingTabId}`,
          );
        } else if (msg.type === "M365_REAUTH_DONE") {
          console.log(
            `\n续期完成: state=${msg.state} rtRemain=${msg.rtRemainMin}min reason=${msg.reason || "-"}`,
          );
          if (Array.isArray(msg.log) && msg.log.length) {
            console.log("--- 扩展后台日志（续期）---");
            for (const line of msg.log) console.log("  " + line);
          }
          doneMsg = msg;
          if (!probeSent && msg.state === "none") {
            probeSent = true;
            console.log(
              "\n→ 让页面自己发一次 /token（旧 RT 故意写错），看后台有没有把它改成新 RT",
            );
            ws.send(JSON.stringify({ type: "M365_PAGE_TOKEN_PROBE" }));
            armProbeWait();
            return;
          }
          clearTimeout(timer);
          msg.rtBefore = rtBefore;
          resolve(msg);
          try {
            ws.close();
          } catch (_) {}
        } else if (msg.type === "M365_RELOAD_ACK") {
          console.log(`   扩展确认重启（旧构建 ${msg.build || "?"}），等它重连…`);
        } else if (msg.type === "M365_PAGE_TOKEN_PROBE_SENT") {
          console.log(`   已投递到 ${msg.tabs} 个 OWA 标签页，等待页面回报…`);
        } else if (msg.type === "M365_PAGE_TOKEN_PROBE_RESULT") {
          const probeOk =
            msg.ok === true ||
            (Number(msg.status) >= 200 &&
              Number(msg.status) < 300 &&
              msg.hasToken === true);
          console.log(
            `页面 /token（改写后）→ HTTP ${msg.status} · 拿到 access_token=${msg.hasToken} · 本次替换=${msg.swapped} 次` +
              (msg.error ? ` · ${String(msg.error).slice(0, 160)}` : ""),
          );
          if (!probeOk) {
            console.log(
              "   未通过 — 多半是浏览器里还是旧构建；重载后会自动重试",
            );
            commanded = false;
            probeSent = false;
            armProbeWait();
            return;
          }
          clearTimeout(timer);
          msg.ok = true;
          if (doneMsg) {
            doneMsg.rtBefore = rtBefore;
            doneMsg.probe = msg;
          }
          resolve(doneMsg || msg);
          try {
            ws.close();
          } catch (_) {}
        }
      });
    });
  });

  await new Promise((resolve) => wss.close(resolve));
  // Success means the 24h clock was actually RESET by this run. A large
  // remaining value proves nothing — the old RT may still have 23h left
  // (a false ✅ used to be printed while the exchange was failing).
  const rtAfter = Number(result.rtExpiresAt || 0);
  const refreshed = result.state === "none" && rtBefore > 0 && rtAfter - rtBefore > 60000;
  if (refreshed) {
    console.log(
      `\n✅ 已换取新的 24h RT（剩余 ${result.rtRemainMin} 分钟，到期时间 +${Math.round((rtAfter - rtBefore) / 60000)} 分钟）`,
    );
    process.exit(0);
  }
  if (result.probe && !result.probe.ok) {
    console.log(
      `\n❌ 页面 /token 没有被正确改写（status=${result.probe.status}, swapped=${result.probe.swapped}）` +
        (result.probe.error ? `\n   ${String(result.probe.error).slice(0, 200)}` : ""),
    );
    process.exit(1);
  }
  if (result.probe && result.probe.ok) {
    console.log(
      "\n✅ 页面确实用的是扩展的新 RT（AAD 用改写后的请求返回了 access_token）",
    );
  }
  if (result.state === "terminal") {
    console.log(`\n❌ 自动登录失败（state=terminal）：${result.reason || "-"}`);
  } else {
    console.log(
      `\n⚠ 没等到新的 24h RT（state=${result.state}, rtRemain=${result.rtRemainMin}, rtExpiresAt ${rtBefore} -> ${rtAfter}）`,
    );
  }
  process.exit(1);
}

if (process.argv.includes("--live")) {
  liveReauth().catch((error) => {
    console.error("LIVE ERROR:", (error && (error.stack || error.message)) || error);
    process.exit(2);
  });
} else if (process.argv.includes("--visual")) {
  visualDemo().catch((error) => {
    console.error(
      "VISUAL ERROR:",
      (error && (error.stack || error.message)) || error,
    );
    process.exit(2);
  });
} else {
  main().catch((error) => {
    console.error(
      "HARNESS ERROR:",
      (error && (error.stack || error.message)) || error,
    );
    process.exit(2);
  });
}
