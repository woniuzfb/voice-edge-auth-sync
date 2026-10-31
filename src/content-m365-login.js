"use strict";

// Automatic Entra ID (login.microsoftonline.com) password sign-in helper for
// M365 Copilot 24h SPA refresh token renewal, with strict armed-tab scope
// checks, 1-attempt circuit breaker, and error/MFA detection.
(function () {
  if (window !== window.top) return;
  if (
    /\/oauth2\/v2\.0\/logout\b|\/common\/uxlogout\b|\/forgetuser\b/i.test(
      location.pathname,
    )
  ) {
    return;
  }

  const SUBMIT_GUARD_KEY = "__veM365AutoLoginSubmittedAt";
  const USER_GUARD_KEY = "__veM365AutoLoginUserSubmittedAt";
  const TILE_GUARD_KEY = "__veM365AutoLoginTileClickedAt";
  const KMS_GUARD_KEY = "__veM365AutoLoginKmsiClickedAt";
  const ARMED_FLOW_KEY = "__veM365ArmedFlowAt";
  const GUARD_TTL_MS = 2 * 60 * 1000;

  // Benign AAD codes that appear on initial silent-SSO fallback before the user
  // submits a password (50058 = silent sign-in cookie missing).
  const BENIGN_INITIAL_CODES = new Set(["50058", "0", ""]);

  // Known fatal/interactive AAD error codes mapped to clear human-readable reasons.
  const AAD_ERROR_DESCRIPTIONS = {
    "50126": "AADSTS50126: 账号或密码错误，已暂停自动登录以防锁号",
    "50053": "AADSTS50053: 登录尝试次数过多，账户触发锁定保护",
    "50055": "AADSTS50055: 密码已过期，请手动登录并修改密码",
    "50057": "AADSTS50057: 用户账户已被禁用",
    "50072": "AADSTS50072: 需要注册多因素认证 (MFA)",
    "50074": "AADSTS50074: 强身份验证策略要求二次验证 (MFA)",
    "50076": "AADSTS50076: 需要完成二次验证 (MFA)",
    "50079": "AADSTS50079: 需要配置或完成多因素认证 (MFA)",
    "50105": "AADSTS50105: 当前账户未获授权访问该应用",
    "53003": "AADSTS53003: 登录被条件访问策略 (Conditional Access) 阻止",
  };

  function isVisible(el) {
    if (!el || !(el instanceof Element)) return false;
    if (el.hidden || el.getAttribute("aria-hidden") === "true") return false;
    const style = window.getComputedStyle(el);
    if (
      !style ||
      style.display === "none" ||
      style.visibility === "hidden" ||
      Number(style.opacity) === 0
    ) {
      return false;
    }
    const rect = el.getBoundingClientRect();
    return rect.width > 0 && rect.height > 0;
  }

  function setNativeInputValue(input, value) {
    if (!input) return;
    try {
      const proto = Object.getPrototypeOf(input);
      const desc =
        (proto && Object.getOwnPropertyDescriptor(proto, "value")) ||
        Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value");
      if (desc && typeof desc.set === "function") {
        desc.set.call(input, value);
      } else {
        input.value = value;
      }
    } catch (_) {
      input.value = value;
    }
    try {
      input.dispatchEvent(new Event("input", { bubbles: true }));
      input.dispatchEvent(new Event("change", { bubbles: true }));
    } catch (_) {}
  }

  function parseAadConfig() {
    try {
      const scripts = document.querySelectorAll("script");
      for (const s of scripts) {
        const txt = s.textContent || "";
        if (!txt.includes("$Config=")) continue;
        const m = txt.match(/\$Config=(\{[\s\S]*?\});/);
        if (m && m[1]) {
          return JSON.parse(m[1]);
        }
      }
    } catch (_) {}
    return null;
  }

  function extractCfgUsername(cfg) {
    if (!cfg || typeof cfg !== "object") return "";
    return String(
      cfg.sPOST_Username ||
        (cfg.oGetCredTypeResult && cfg.oGetCredTypeResult.Username) ||
        "",
    ).trim();
  }

  function readRecentGuard(key) {
    try {
      const raw = Number(sessionStorage.getItem(key) || 0);
      if (raw > 0 && Date.now() - raw < GUARD_TTL_MS) return raw;
    } catch (_) {}
    return 0;
  }

  function writeGuard(key) {
    try {
      sessionStorage.setItem(key, String(Date.now()));
    } catch (_) {}
  }

  function clearAllGuards() {
    try {
      sessionStorage.removeItem(SUBMIT_GUARD_KEY);
      sessionStorage.removeItem(USER_GUARD_KEY);
      sessionStorage.removeItem(TILE_GUARD_KEY);
      sessionStorage.removeItem(KMS_GUARD_KEY);
      sessionStorage.removeItem(ARMED_FLOW_KEY);
    } catch (_) {}
  }

  function detectLoginFailureOrMfa(cfg) {
    // 1. Check visible DOM error banners first.
    const errorSelectors = [
      "#passwordError",
      "#usernameError",
      "#idSpan_SAOTCC_Error_OTC",
      "#errorText",
      ".alert-error",
    ];
    for (const sel of errorSelectors) {
      const el = document.querySelector(sel);
      if (el && isVisible(el)) {
        const text = String(el.textContent || "").replace(/\s+/g, " ").trim();
        if (text) {
          const codeMatch = text.match(/AADSTS(\d+)/i);
          const code = codeMatch ? codeMatch[1] : "DOM_ERROR";
          return {
            failed: true,
            fatal: true,
            errorCode: code,
            errorMessage: AAD_ERROR_DESCRIPTIONS[code] || text.slice(0, 240),
          };
        }
      }
    }

    // 2. Check $Config error codes.
    if (cfg && typeof cfg === "object") {
      const rawCodes = [];
      if (cfg.sErrorCode != null) rawCodes.push(String(cfg.sErrorCode).trim());
      if (cfg.stsErrorCode != null)
        rawCodes.push(String(cfg.stsErrorCode).trim());
      if (Array.isArray(cfg.arrValErrs)) {
        for (const c of cfg.arrValErrs) rawCodes.push(String(c || "").trim());
      }
      for (const code of rawCodes) {
        if (!code || BENIGN_INITIAL_CODES.has(code)) continue;
        const msg =
          AAD_ERROR_DESCRIPTIONS[code] ||
          String(cfg.sErrTxt || "").replace(/\s+/g, " ").trim() ||
          `AAD 登录返回错误码 AADSTS${code}`;
        return {
          failed: true,
          fatal: true,
          errorCode: code,
          errorMessage: msg,
        };
      }
    }

    // 3. Check MFA / OTC / Proof / Captcha challenge elements.
    const mfaSelectors = [
      "#idDiv_SAOTCS_Proofs",
      "#idTxtBx_SAOTCC_OTC",
      "#idRichContext_DisplaySign",
      "#idDiv_SAOTCAS_Description",
      "#hipTemplateContainer",
      "input[name='otc']",
    ];
    for (const sel of mfaSelectors) {
      const el = document.querySelector(sel);
      if (el && isVisible(el)) {
        return {
          failed: true,
          fatal: true,
          requiresMfa: true,
          errorCode: "MFA_REQUIRED",
          errorMessage: "检测到二次验证 (MFA/动态码/验证码)，需人工完成登录",
        };
      }
    }

    // 4. Check if we already submitted a password in this flow and landed back
    // on the password input screen again without an explicit error code.
    const passInput = document.querySelector(
      "input[name='passwd']#i0118, input[type='password'][name='passwd']",
    );
    const submittedAt = readRecentGuard(SUBMIT_GUARD_KEY);
    if (
      passInput &&
      isVisible(passInput) &&
      submittedAt > 0 &&
      Date.now() - submittedAt > 1200
    ) {
      return {
        failed: true,
        fatal: true,
        errorCode: "PASSWORD_REJECTED",
        errorMessage: "自动提交密码后仍停留在密码页，已停止重试以防锁号",
      };
    }

    return null;
  }

  // Passive capture: only attached when BOTH autoLoginEnabled and
  // rememberOnManualLogin are explicitly enabled AND this tab is an armed
  // M365 Copilot sign-in tab.
  let _passiveHookAttached = false;
  function attachPassiveManualLoginCapture(cfg) {
    if (_passiveHookAttached) return;
    const passInput = document.querySelector(
      "input[name='passwd']#i0118, input[type='password'][name='passwd']",
    );
    if (!passInput) return;
    _passiveHookAttached = true;

    const stageCurrentInput = () => {
      try {
        const pwd = String(passInput.value || "");
        if (!pwd) return;
        const userInput = document.querySelector("input[name='loginfmt']#i0116");
        const username = String(
          (userInput && userInput.value) || extractCfgUsername(cfg) || "",
        ).trim();
        browser.runtime
          .sendMessage({
            __veM365: true,
            type: "M365_STAGE_LOGIN_CREDS",
            cfgUsername: extractCfgUsername(cfg),
            username,
            password: pwd,
          })
          .catch(() => {});
      } catch (_) {}
    };

    const submitBtn = document.querySelector(
      "#idSIButton9, input[type='submit'], button[type='submit']",
    );
    if (submitBtn) {
      submitBtn.addEventListener("click", stageCurrentInput, true);
    }
    if (passInput.form) {
      passInput.form.addEventListener("submit", stageCurrentInput, true);
    }
  }

  let _autoLoginFinished = false;

  // Diagnostics: every "nothing to do here" exit used to be a silent return, so
  // the background log could never explain why auto-login never advanced.
  let _lastDiagKey = "";
  function reportDiag(state, detail) {
    const key = String(state) + "|" + String(detail || "");
    if (key === _lastDiagKey) return;
    _lastDiagKey = key;
    try {
      browser.runtime
        .sendMessage({
          __veM365: true,
          type: "M365_AUTO_LOGIN_DIAG",
          state: String(state),
          detail: String(detail || ""),
          url: location.href.split(/[?#]/, 1)[0],
        })
        .catch(() => {});
    } catch (_) {}
  }

  function describePage(cfg) {
    const vis = (sel) => {
      const el = document.querySelector(sel);
      return el && isVisible(el) ? "1" : "0";
    };
    return [
      "pgid=" + String((cfg && cfg.pgid) || "-"),
      "err=" +
        String(
          (cfg && (cfg.sErrorCode || cfg.stsErrorCode)) || "-",
        ),
      "pass=" + vis("input[name='passwd']#i0118, input[type='password'][name='passwd']"),
      "user=" + vis("input[name='loginfmt']#i0116"),
      "tiles=" + vis("#tilesHolder"),
      "kmsi=" + vis("#KmsiCheckboxField, input[name='DontShowAgain']"),
      "submit=" + vis("#idSIButton9, input[type='submit'], button[type='submit']"),
    ].join(" ");
  }

  // Auto-login is unattended: we never wait for a human to type. The only
  // password source is the browser's own Password Manager, and it gets one
  // bounded window (the 3s loop above). If it stays empty, the attempt fails —
  // no "watch for late input", no hand-off, nothing to click.
  // 重入锁：runAutoLoginStepInner 内部有 await（消息往返），MutationObserver
  // 会在等待期间再次触发；没有它，同一次填充会并发发出多份 CONFIG/SUBMITTING。
  let _stepInFlight = false;
  async function runAutoLoginStep() {
    if (_autoLoginFinished || _stepInFlight) return;
    _stepInFlight = true;
    try {
      await runAutoLoginStepInner();
    } finally {
      _stepInFlight = false;
    }
  }

  async function runAutoLoginStepInner() {
    if (_autoLoginFinished) return;
    const cfg = parseAadConfig();
    const cfgUsername = extractCfgUsername(cfg);

    // Query background first to verify this tab is an armed M365 Copilot login
    // flow (never touch unrelated Microsoft sign-in pages in the same browser).
    let autoConfig = null;
    try {
      autoConfig = await browser.runtime.sendMessage({
        __veM365: true,
        type: "M365_GET_AUTO_LOGIN_CONFIG",
        cfgUsername,
      });
    } catch (_) {}

    const isArmed = Boolean(
      (autoConfig && autoConfig.armed) || readRecentGuard(ARMED_FLOW_KEY),
    );
    if (!isArmed) {
      // Only report for flows that look like our own authorization request, so
      // unrelated Microsoft sign-in pages stay out of the extension log.
      if (/[?&](state|code_challenge)=/.test(location.search)) {
        reportDiag("not-armed", describePage(cfg));
      }
      return;
    }
    if (autoConfig && autoConfig.armed) {
      writeGuard(ARMED_FLOW_KEY);
    }

    if (
      autoConfig &&
      autoConfig.enabled === true &&
      autoConfig.rememberOnManualLogin === true
    ) {
      attachPassiveManualLoginCapture(cfg);
    }

    // Always check for failures or MFA first on an armed tab, before touching any inputs.
    const failure = detectLoginFailureOrMfa(cfg);
    if (failure) {
      _autoLoginFinished = true;
      clearAllGuards();
      browser.runtime
        .sendMessage({
          __veM365: true,
          type: "M365_AUTO_LOGIN_RESULT",
          ok: false,
          fatal: Boolean(failure.fatal),
          requiresMfa: Boolean(failure.requiresMfa),
          errorCode: String(failure.errorCode || "ERROR"),
          errorMessage: String(failure.errorMessage || "自动登录失败"),
          url: location.href.split(/[?#]/, 1)[0],
        })
        .catch(() => {});
      return;
    }

    if (!autoConfig || autoConfig.enabled === false || autoConfig.paused) {
      reportDiag(
        "disabled",
        (autoConfig && autoConfig.paused ? "paused " : "disabled ") +
          describePage(cfg),
      );
      return;
    }

    const configuredUser = String(autoConfig.username || "")
      .trim()
      .toLowerCase();
    const configuredPass = String(autoConfig.password || "");

    // Guard against filling credentials into a different account's password form.
    if (
      cfgUsername &&
      configuredUser &&
      cfgUsername.toLowerCase() !== configuredUser
    ) {
      reportDiag(
        "username-mismatch",
        "page=" + cfgUsername + " saved=" + configuredUser,
      );
      return;
    }

    // Step 1: KMSI ("是否保持登录状态?") interrupt page.
    const kmsiCheckbox = document.querySelector(
      "#KmsiCheckboxField, input[name='DontShowAgain']",
    );
    const kmsiSubmit = document.querySelector("#idSIButton9");
    if (
      ((cfg && cfg.pgid === "KmsiInterrupt") ||
        (kmsiCheckbox && isVisible(kmsiCheckbox))) &&
      kmsiSubmit &&
      isVisible(kmsiSubmit) &&
      !readRecentGuard(KMS_GUARD_KEY)
    ) {
      writeGuard(KMS_GUARD_KEY);
      try {
        if (kmsiCheckbox && !kmsiCheckbox.checked) kmsiCheckbox.click();
      } catch (_) {}
      setTimeout(() => {
        try {
          kmsiSubmit.click();
        } catch (_) {}
      }, 150);
      return;
    }

    // Step 2: Password input screen (#i0118).
    const passInput = document.querySelector(
      "input[name='passwd']#i0118, input[type='password'][name='passwd']",
    );
    const submitBtn = document.querySelector(
      "#idSIButton9, input[type='submit'], button[type='submit']",
    );
    if (passInput && isVisible(passInput) && submitBtn && isVisible(submitBtn)) {
      if (readRecentGuard(SUBMIT_GUARD_KEY)) {
        return;
      }
      // No password inside the extension is the normal case: give the browser's
      // own Password Manager up to 3s to autofill #i0118, then keep watching for
      // a late fill instead of giving up (Chrome/Firefox may fill after the
      // page settles, or the user may type it into the handed-off tab).
      let passwordToUse = configuredPass || String(passInput.value || "");
      if (!passwordToUse) {
        for (let i = 0; i < 12; i++) {
          await new Promise((r) => setTimeout(r, 250));
          if (detectLoginFailureOrMfa(cfg)) return;
          passwordToUse = String(passInput.value || "");
          if (passwordToUse) break;
        }
      }
      if (!passwordToUse) {
        _autoLoginFinished = true;
        browser.runtime
          .sendMessage({
            __veM365: true,
            type: "M365_AUTO_LOGIN_RESULT",
            ok: false,
            fatal: false,
            missingPassword: true,
            errorCode: "NO_PASSWORD",
            errorMessage:
              "浏览器密码库未填充密码（扩展不保存密码），自动登录失败",
            url: location.href.split(/[?#]/, 1)[0],
          })
          .catch(() => {});
        reportDiag("no-password", describePage(cfg));
        return;
      }

      // 不要在这里把流程标记为“结束”：真实 AAD 的错误页/MFA/KMSI 是在同一个
      // 文档里就地重渲染的，一旦关掉观察者就再也看不到它们（密码错误会一直挂到
      // 超时）。SUBMIT_GUARD_KEY 已经保证不会重复提交，继续观察即可。
      writeGuard(SUBMIT_GUARD_KEY);
      if (configuredPass && passInput.value !== configuredPass) {
        setNativeInputValue(passInput, configuredPass);
      } else {
        setNativeInputValue(passInput, passwordToUse);
      }

      browser.runtime
        .sendMessage({
          __veM365: true,
          type: "M365_AUTO_LOGIN_SUBMITTING",
          cfgUsername,
          username: cfgUsername || autoConfig.username || "",
        })
        .catch(() => {});

      setTimeout(() => {
        if (detectLoginFailureOrMfa(parseAadConfig())) return;
        try {
          submitBtn.click();
        } catch (_) {}
      }, 200);
      return;
    }

    // Step 3: Account Picker tile (#tilesHolder).
    const tilesHolder = document.querySelector("#tilesHolder");
    if (tilesHolder && isVisible(tilesHolder) && !readRecentGuard(TILE_GUARD_KEY)) {
      const tiles = Array.from(
        tilesHolder.querySelectorAll(
          "[data-test-id], .table[role='button'], .tile-container .table",
        ),
      ).filter((el) => isVisible(el) && el.id !== "otherTile");
      let targetTile = null;
      if (configuredUser) {
        targetTile = tiles.find((el) =>
          String(
            el.getAttribute("data-test-id") || el.textContent || "",
          )
            .toLowerCase()
            .includes(configuredUser),
        );
      }
      if (!targetTile && tiles.length === 1) {
        targetTile = tiles[0];
      }
      if (targetTile) {
        writeGuard(TILE_GUARD_KEY);
        setTimeout(() => {
          try {
            targetTile.click();
          } catch (_) {}
        }, 200);
        return;
      }
    }

    // Step 4: Username input screen (input[name="loginfmt"]#i0116).
    const userInput = document.querySelector("input[name='loginfmt']#i0116");
    if (
      userInput &&
      isVisible(userInput) &&
      submitBtn &&
      isVisible(submitBtn) &&
      !readRecentGuard(USER_GUARD_KEY)
    ) {
      const existingVal = String(userInput.value || "").trim().toLowerCase();
      if (configuredUser && existingVal && existingVal !== configuredUser) {
        return;
      }
      const userToUse =
        String(autoConfig.username || "").trim() ||
        String(userInput.value || "").trim();
      if (userToUse) {
        writeGuard(USER_GUARD_KEY);
        if (userInput.value !== userToUse) {
          setNativeInputValue(userInput, userToUse);
        }
        setTimeout(() => {
          try {
            submitBtn.click();
          } catch (_) {}
        }, 200);
      }
    }

    // Armed, enabled, but nothing on this page is actionable: report the page
    // state once so the background log shows what the login tab was showing.
    reportDiag("stuck", describePage(cfg));
  }

  // Run on initial idle and watch for KnockoutJS dynamic view transitions
  // (AAD often swaps username -> password or renders #passwordError in-place).
  setTimeout(() => {
    runAutoLoginStep().catch(() => {});
  }, 300);

  if (typeof MutationObserver === "function" && document.documentElement) {
    let debounceTimer = null;
    const obs = new MutationObserver(() => {
      if (_autoLoginFinished) {
        obs.disconnect();
        return;
      }
      if (debounceTimer) clearTimeout(debounceTimer);
      debounceTimer = setTimeout(() => {
        debounceTimer = null;
        runAutoLoginStep().catch(() => {});
      }, 250);
    });
    obs.observe(document.documentElement, {
      childList: true,
      subtree: true,
      attributes: true,
      attributeFilter: ["class", "style", "hidden", "aria-hidden"],
    });
    setTimeout(() => obs.disconnect(), 30000);
  }
})();
