/* =========================================================================
 * Voice Edge · M365 Copilot content script — self-built Chathub WSS
 * -------------------------------------------------------------------------
 * 页面停在 /host/{uuid}/entity1-{uuid}。插件【自己建】Chathub WSS,不依赖页面打字。
 *
 * 链路(全部在浏览器内,Python 零碰 token):
 *   background: webRequest 抓 login.microsoftonline.com token 响应里的 refresh_token
 *               → 需要时用 RT 自刷 sydney access_token(FOCI: client_id=...,
 *                 scope=https://substrate.office.com/sydney/.default)→ 滚动续期
 *   content(page world):
 *     M365_ASK 到来 → 使用 Python 指定的 ConversationId，空值则新建
 *       → 在正确的 m365copilotapp.svc.cloud.microsoft frame 建立 page-world Chathub WebSocket
 *       → 握手 {"protocol":"json","version":1}\x1e → 收 {} → 发 chat+Metrics 帧
 *       → 按 cursor messageId 读取正文快照/writeAtCursor 增量→ type2/type3 收尾
 * ========================================================================= */
(() => {
  "use strict";

  function PAGE_HOOK() {
    "use strict";
    // Origin self-gate. The manifest injects this script into every frame of
    // three match origins (outlook.cloud.microsoft, m365.cloud.microsoft,
    // m365copilotapp.svc.cloud.microsoft) with all_frames:true, so every host-shell subframe
    // also advertises itself as a Chathub-socket candidate. Only the
    // m365copilotapp.svc.cloud.microsoft frame ever builds the working socket (observed
    // frameOrigin in every VE-FRAME-READY). We therefore BLOCK the two proven
    // host-shell origins from advertising as socket candidates, and FAIL OPEN
    // for m365copilotapp.svc.cloud.microsoft and any unforeseen origin — gating can only ever
    // remove known-noise candidates, never suppress a frame that might be the
    // real one, so it cannot break connectivity. Entry discovery
    // (M365_ENTRY_DISCOVERED) is deliberately NOT gated: it must keep reporting
    // from the shell origin.
    const M365_SHELL_ORIGINS = {
      "https://outlook.cloud.microsoft": true,
      "https://m365.cloud.microsoft": true,
    };
    const frameMayHostSocket = () =>
      M365_SHELL_ORIGINS[location.origin] !== true;
    const priorHook = window.__veM365Hook;
    if (priorHook && priorHook.ready === true) {
      if (frameMayHostSocket()) {
        window.postMessage(
          {
            __veM365: true,
            dir: "fromPage",
            type: "M365_FRAME_READY",
            capabilities: priorHook.capabilities,
            frameOrigin: location.origin,
            frameUrl: location.href,
          },
          "*",
        );
      }
      return;
    }
    // Do not advertise readiness until every handler below has been installed.
    const hookState = {
      ready: false,
      capabilities: null,
    };
    window.__veM365Hook = hookState;

    const RS = "\u001e";
    // 固定 feature-flag 串(照抄实测抓包;服务端如要求更新再改)
    const VARIANTS = [
      "EnableMcpServerWidgets",
      "feature.EnableMcpServerWidgets",
      "feature.EnableImageGenInsufficientTokensThrottled",
      "feature.EnableImageGenSystemCapacityThrottled",
      "feature.EnableLuForChatCIQ",
      "feature.enableChatCIQPlugin",
      "EnableRequestPlugins",
      "feature.EnableSensitivityLabels",
      "EnableUnsupportedUrlDetector",
      "feature.IsCustomEngineCopilotEnabled",
      "feature.bizchatfluxv3",
      "feature.enablechatpages",
      "feature.turnOnWorkTabRecommendation",
      "feature.turnOnDARecommendation",
      "feature.IsStreamingModeInChatRequestEnabled",
      "IncludeSourceAttributionsConcise",
      "SkipPublishEmptyMessage",
      "feature.EnableDeduplicatingSourceAttributions",
      "feature.IsCitationsReferencesOutputEnabled",
      "feature.enableDeltaStreamingForReferences",
      "feature.enableIncludeReferencesInDeltaResponse",
      "feature.enablereferencesforagents",
      "Enable3PActionProgressMessages",
      "feature.enableClientWebRtc",
      "feature.EnableMeetingRecapOfSeriesMeetingWithCiq",
      "feature.EnableReferencesListCompleteSignal",
      "feature.StorageMessageSplitDisabled",
      "feature.EnableCuaTakeControlApi",
      "SingletonEnvOn",
      "EnableComposeWidget",
      "feature.EnableMergingPureDeltas",
      "feature.isExternalEmailEnabled",
      "feature.isExcludedEmailEnabled",
      "feature.disabledisallowedmsgs",
      "feature.enableCitationsForSynthesisData",
      "feature.EnableConversationShareApis",
      "feature.enableGenerateGraphicArtOptionsSet",
      "cdximagen",
      "feature.EnableContentApiandDocTypeHtmlInRichAnswers",
      "cdxgrounding_api_v2_rich_web_answers_reference_bottom_force",
      "cdxenablerenderforisocomp",
      "feature.EnableDesignEditorImageGrounding",
      "feature.EnableDesignerEditor",
      "feature.EnableSkipRehydrationForSpeCIdImages",
      "feature.sourcescontrolmainline",
      "feature.sourcescontrolmainlineal",
      "feature.EnableConnectorExecutionControlsAllowlist",
      "feature.EnableBizchatMainlineExecutionControlsResolution",
      "feature.EnablePersonalization",
      "cdxentrecapvifluxv3",
      "rich_responses",
      "feature.EnableBase64DataInMessageAnnotations",
      "feature.EnableStarterLicenseCheckBypass",
      "feature.DisableMimir3sFlow",
      "feature.EnablePersonalWorkingSetFor3s",
      "feature.EnableSkipEmittingMessageOnFlush",
      "feature.EnableRemoveEmptySourceAttributions",
      "feature.EnableRemoveStreamingMode",
      "feature.OfficeWebToHelix",
      "feature.OfficeDesktopToHelix",
      "feature.M365TeamsHubToHelix",
      "feature.OwaHubToHelix",
      "feature.MonarchHubToHelix",
      "feature.Win32OutlookHubToHelix",
      "feature.MacOutlookHubToHelix",
      "Agt_bizchat_enableGpt5ForHelix",
    ].join(",");

    const uuid = () =>
      (crypto.randomUUID && crypto.randomUUID()) ||
      "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (c) => {
        const r = (Math.random() * 16) | 0;
        return (c === "x" ? r : (r & 0x3) | 0x8).toString(16);
      });
    const parseFrames = (p) => {
      const out = [];
      for (const c of String(p).split(RS)) {
        const t = c.trim();
        if (t) {
          try {
            out.push(JSON.parse(t));
          } catch (_) {}
        }
      }
      return out;
    };
    const jwtClaims = (tok) => {
      try {
        return JSON.parse(
          atob(String(tok).split(".")[1].replace(/-/g, "+").replace(/_/g, "/")),
        );
      } catch (_) {
        return {};
      }
    };
    const post = (m) =>
      window.postMessage(
        Object.assign({ __veM365: true, dir: "fromPage" }, m),
        "*",
      );
    // Registry of in-flight doAsk turns, keyed by the Python request id, so an
    // M365_STOP from the bridge can abort the exact turn the client (Continue)
    // interrupted. Without this the self-built Chathub socket keeps streaming
    // after the client is gone; Python then logs [relay-inbound-drop] no pending
    // queue for that id and the browser silently finishes an answer nobody sees.
    const _veActiveAsks = new Map();
    // Gated page-world tracing. OFF by default; enable live from the page
    // console with `window.__veM365Debug = true` (no reload). Focused on the
    // conversation-id decision and terminal signal.
    const dbg = (...a) => {
      try {
        if (window.__veM365Debug) console.log("[VE-m365][dbg]", ...a);
      } catch (_) {}
    };

    // ---- Artifact auth capture ------------------------------------------
    // Downloading a CodeInterpreter artifact from the AMS object endpoint
    // (asyncgw.teams.microsoft.com/v1/objects/…) needs the SAME bearer the page
    // itself uses — NOT the sydney token (wrong audience -> 401) and NOT cookies
    // (also 401). Rather than guess the audience, we passively reuse the page's
    // own successful request in two ways, both with the page's real auth:
    //   1) _amsBytes : if the page fetches the object itself (to build its blob
    //      download), we tee those exact bytes — zero auth guessing.
    //   2) _amsPageAuth : the Authorization header the page attaches to any AMS
    //      request, captured from both fetch and XHR, reused for our own GET of
    //      artifact URLs the page did not fetch.
    const AMS_OBJECT_URL_RE = /^https:\/\/[^/]*asyncgw[^/]*\/v1\/objects\//i;
    const normalizeAmsUrl = (u) =>
      String(u || "").replace(/(\/views\/original)\/[^/?#]*(?=$|[?#])/i, "$1");
    let _amsPageAuth = ""; // most recent "Bearer …" the page used for AMS
    const _amsBytes = new Map(); // normalized url -> base64 (teed from page)
    const bufToB64 = (buf) => {
      const b = new Uint8Array(buf);
      let bin = "";
      const CH = 0x8000;
      for (let i = 0; i < b.length; i += CH)
        bin += String.fromCharCode.apply(null, b.subarray(i, i + CH));
      return btoa(bin);
    };
    const rememberAmsAuth = (auth) => {
      const s = String(auth || "");
      if (/^Bearer\s+\S/i.test(s)) _amsPageAuth = s;
    };
    // ---------------------------------------------------------------------

    // ---- 捕获页面 token fetch 的成功响应，保存服务端轮换后的 refresh_token ----
    (function hookTokenFetch() {
      const originalFetch = window.fetch;
      if (originalFetch.__veM365TokenHook) return;

      async function wrappedFetch(input, init) {
        const response = await originalFetch.apply(this, arguments);
        // Passively reuse the page's own AMS auth/bytes (see _amsBytes /
        // _amsPageAuth above). Never blocks or alters the page's response.
        try {
          const reqHref =
            typeof input === "string" ? input : (input && input.url) || "";
          if (AMS_OBJECT_URL_RE.test(reqHref)) {
            // capture the page's bearer for this AMS request
            let auth = "";
            if (init && init.headers) {
              const h = init.headers;
              auth =
                typeof h.get === "function"
                  ? h.get("Authorization") || h.get("authorization")
                  : h.Authorization || h.authorization || "";
            }
            if (
              !auth &&
              typeof Request !== "undefined" &&
              input instanceof Request
            )
              auth = input.headers.get("Authorization") || "";
            rememberAmsAuth(auth);
            // tee the exact bytes the page just downloaded (best case)
            if (response && response.ok) {
              response
                .clone()
                .arrayBuffer()
                .then((buf) =>
                  _amsBytes.set(normalizeAmsUrl(reqHref), bufToB64(buf)),
                )
                .catch(() => {});
            }
          }
        } catch (_) {}
        try {
          const requestUrl = new URL(
            typeof input === "string" ? input : (input && input.url) || "",
            location.href,
          );
          const match = requestUrl.href.match(
            /^https:\/\/login\.microsoftonline\.com\/([0-9a-f-]{36})\/oauth2\/v2\.0\/token/i,
          );
          if (!match) return response;

          let bodyText = "";
          if (init && typeof init.body === "string") bodyText = init.body;
          else if (typeof Request !== "undefined" && input instanceof Request) {
            try {
              bodyText = await input.clone().text();
            } catch (_) {}
          }
          const params = new URLSearchParams(bodyText);
          const scopes = String(params.get("scope") || "")
            .split(/\s+/)
            .filter(Boolean);
          if (
            params.get("grant_type") !== "refresh_token" ||
            !scopes.includes("https://substrate.office.com/sydney/.default")
          )
            return response;

          const clientId = String(
            params.get("client_id") ||
              requestUrl.searchParams.get("client_id") ||
              "",
          );
          if (!clientId || !response.ok) return response;
          const payload = await response.clone().json();
          if (!payload || !payload.access_token || !payload.refresh_token)
            return response;
          post({
            type: "M365_TOKEN_RESPONSE",
            tenantId: match[1],
            clientId,
            accessToken: String(payload.access_token),
            refreshToken: String(payload.refresh_token),
            expiresIn: Number(payload.expires_in || 0),
            capturedAt: Date.now(),
          });
        } catch (_) {}
        return response;
      }
      Object.defineProperty(wrappedFetch, "__veM365TokenHook", { value: true });
      try {
        Object.defineProperty(wrappedFetch, "name", {
          value: originalFetch.name,
        });
      } catch (_) {}
      window.fetch = wrappedFetch;
    })();

    // ---- 被动截获页面【自己发】的 GetChats(XHR),缓存;不主动 fetch、不碰 token ----
    let _chatsCache = { at: 0, chats: [] };
    (function hookXHR() {
      const O = XMLHttpRequest.prototype.open;
      const S = XMLHttpRequest.prototype.send;
      const SRH = XMLHttpRequest.prototype.setRequestHeader;
      XMLHttpRequest.prototype.open = function (method, url) {
        try {
          this.__veUrl = String(url || "");
        } catch (_) {}
        return O.apply(this, arguments);
      };
      // Capture the page's bearer if it downloads the AMS object over XHR
      // (covers the case where the download path is XHR rather than fetch).
      XMLHttpRequest.prototype.setRequestHeader = function (name, value) {
        try {
          if (
            String(name || "").toLowerCase() === "authorization" &&
            AMS_OBJECT_URL_RE.test(this.__veUrl || "")
          )
            rememberAmsAuth(value);
        } catch (_) {}
        return SRH.apply(this, arguments);
      };
      XMLHttpRequest.prototype.send = function () {
        try {
          if (/m365Copilot\/GetChats/i.test(this.__veUrl || "")) {
            this.addEventListener("load", function () {
              try {
                const data = JSON.parse(this.responseText || "{}");
                if (Array.isArray(data.chats)) {
                  _chatsCache = {
                    at: Date.now(),
                    chats: data.chats.map((c) => ({
                      conversationId: c.conversationId,
                      chatName: c.chatName,
                      tone: c.tone,
                      updateTimeUtc: c.updateTimeUtc,
                      turnState: c.turnState,
                    })),
                  };
                  post({
                    type: "M365_CHATS_SNAPSHOT",
                    chats: _chatsCache.chats,
                    capturedAt: _chatsCache.at,
                  });
                }
              } catch (_) {}
            });
          }
        } catch (_) {}
        return S.apply(this, arguments);
      };
    })();

    // ---- token 供给:向 background 要 sydney token(background 用 RT 自刷) ----
    let _tokenCache = { token: "", exp: 0 };
    const _tokenWaiters = [];
    function requestSydneyToken() {
      return new Promise((resolve, reject) => {
        // 90s 内有缓存直接用
        if (_tokenCache.token && _tokenCache.exp - Date.now() > 90000) {
          return resolve(_tokenCache.token);
        }
        _tokenWaiters.push({ resolve, reject });
        post({ type: "M365_NEED_TOKEN" });
        setTimeout(() => {
          const i = _tokenWaiters.findIndex((w) => w.resolve === resolve);
          if (i >= 0) {
            _tokenWaiters.splice(i, 1);
            reject(new Error("token timeout"));
          }
        }, 15000);
      });
    }
    // AMS/artifact token bridge. The AMS object endpoint requires an IC3-
    // audience token (aud=ic3.teams.office.com); the Sydney token 401s there.
    // Mirrors requestSydneyToken but over the M365_NEED_AMS_TOKEN channel.
    let _amsTokenCache = { token: "", exp: 0 };
    const _amsTokenWaiters = [];
    function requestAmsToken(forceRefresh) {
      return new Promise((resolve, reject) => {
        // forceRefresh (used on a 401/403 retry) drops the cached token so a
        // fresh IC3-audience token is minted; the flag is forwarded so
        // background invalidates ITS cache too — otherwise both sides would
        // hand back the same stale token and the retry could never recover.
        if (forceRefresh) _amsTokenCache = { token: "", exp: 0 };
        if (_amsTokenCache.token && _amsTokenCache.exp - Date.now() > 90000) {
          return resolve(_amsTokenCache.token);
        }
        _amsTokenWaiters.push({ resolve, reject });
        post({ type: "M365_NEED_AMS_TOKEN", force: !!forceRefresh });
        setTimeout(() => {
          const i = _amsTokenWaiters.findIndex((w) => w.resolve === resolve);
          if (i >= 0) {
            _amsTokenWaiters.splice(i, 1);
            reject(new Error("ams token timeout"));
          }
        }, 15000);
      });
    }
    // 上传单张图片到 BizChat，返回可用于 messageAnnotations 的 docId。
    // 端点与 Chathub 同属 substrate.office.com/m365Copilot/*，因此复用 Sydney token。
    // file: { name, mimeType, dataBase64 } —— dataBase64 可带或不带 data: 前缀
    async function uploadImageToM365(file, conversationId, oid, tid) {
      const token = await requestSydneyToken();
      // 补齐 data URL 前缀（UploadFile 的 FileBase64 期望完整 data:*;base64, 串）
      let b64 = String(file.dataBase64 || "");
      if (!/^data:/i.test(b64)) {
        b64 = "data:" + (file.mimeType || "image/png") + ";base64," + b64;
      }
      const fd = new FormData();
      fd.append("scenario", "UploadImage");
      fd.append("conversationId", String(conversationId || ""));
      fd.append("FileBase64", b64);
      // 抓包里这三个 optionsSets 各占一段，FormData 允许重复 key
      fd.append("optionsSets", "cwcgptvsan");
      fd.append(
        "optionsSets",
        "flux_v3_gptv_enable_upload_multi_image_in_turn_wo_ch",
      );
      fd.append("optionsSets", "gptvnorm2048");

      const res = await fetch(
        "https://substrate.office.com/m365Copilot/UploadFile",
        {
          method: "POST",
          // 不要手动设 Content-Type：让浏览器自带 multipart boundary
          headers: {
            Accept: "*/*",
            Authorization: "Bearer " + token,
            "x-anchormailbox": "Oid:" + oid + "@" + tid,
            "x-scenario": "owahub",
            "x-variants": "feature.EnableImageSupportInUploadFile",
          },
          body: fd,
        },
      );
      const text = await res.text();
      let json = {};
      try {
        json = text ? JSON.parse(text) : {};
      } catch (_) {}
      if (!res.ok || !json.docId) {
        throw new Error(
          "UploadFile failed (" +
            res.status +
            "): " +
            String(
              json.error || (json.result && json.result.message) || text,
            ).slice(0, 300),
        );
      }
      return json.docId; // e.g. "0-eus-d5-ab087f10c64e9689517a5eab0185791c"
    }
    // background 回传 token
    window.addEventListener("message", (ev) => {
      const d = ev.data;
      if (!d || d.__veM365 !== true || d.dir !== "toPage") return;
      if (d.type === "M365_PROBE") {
        if (frameMayHostSocket()) {
          post({
            type: "M365_FRAME_READY",
            frameOrigin: location.origin,
            frameUrl: location.href,
          });
        }
      } else if (d.type === "M365_TOKEN") {
        _tokenCache = { token: String(d.token || ""), exp: Number(d.exp || 0) };
        while (_tokenWaiters.length) {
          const w = _tokenWaiters.shift();
          _tokenCache.token
            ? w.resolve(_tokenCache.token)
            : w.reject(new Error(d.error || "no token"));
        }
      } else if (d.type === "M365_AMS_TOKEN") {
        _amsTokenCache = {
          token: String(d.token || ""),
          exp: Number(d.exp || 0),
        };
        while (_amsTokenWaiters.length) {
          const w = _amsTokenWaiters.shift();
          _amsTokenCache.token
            ? w.resolve(_amsTokenCache.token)
            : w.reject(new Error(d.error || "no ams token"));
        }
      } else if (d.type === "M365_ASK") {
        doAsk(
          d.id,
          String(d.text || ""),
          String(d.tone || "Claude_Opus"),
          String(d.conversationId || ""),
          Array.isArray(d.attachments) ? d.attachments : [],
          Number(d.idleTimeoutMs || 0),
        );
      } else if (d.type === "M365_STOP") {
        // Abort the in-flight turn for this id (client interrupted the request).
        // dbg makes it unambiguous whether the STOP reached THIS frame and
        // whether the id was live here (routing vs registry miss).
        const _veStopId = String(d.id || "");
        const stop = _veActiveAsks.get(_veStopId);
        dbg("M365_STOP page recv id=%s known=%s", _veStopId, !!stop);
        if (stop) stop();
      }
    });

    // ---- 自建 Chathub WSS,发一轮,读流 ----
    function buildChatArgs(text, tone, conversationId, attachments = []) {
      const rid = uuid();
      // FileUrl（ODB 文档）—— 这条会触发 officeweb 分支
      const fileUrlAnnotations = attachments
        .filter(
          (file) =>
            file &&
            file.name &&
            file.url &&
            file.verified === true &&
            file.itemId &&
            file.driveId,
        )
        .map((file) => ({
          id: String(file.url),
          text: String(file.name),
          url: String(file.url),
          messageAnnotationType: "FileUrl",
        }));
      // ImageFile（图片 blob）—— 只需已拿到 docId；保持 owahub 默认帧
      const imageAnnotations = attachments
        .filter((file) => file && file.docId && file.kind === "image")
        .map((file) => {
          const ext =
            String(file.name || "")
              .split(".")
              .pop()
              .toLowerCase() || "png";
          return {
            id: String(file.docId),
            messageAnnotationMetadata: {
              "@type": "File",
              annotationType: "File",
              fileType: ext, // 裸扩展名
              fileName: String(file.name || "image." + ext),
            },
            messageAnnotationType: "ImageFile",
          };
        });
      const messageAnnotations = [...fileUrlAnnotations, ...imageAnnotations];
      // 关键：只有 FileUrl 才切 officeweb；纯图片留在 owahub 默认分支。
      const hasAttachments = fileUrlAnnotations.length > 0;
      const capturedAttachmentOptionsSets = [
        "search_result_progress_messages_with_search_queries",
        "update_textdoc_response_after_streaming",
        "deepleo_networking_timeout_10minutes_canmore",
        "cwc_flux_image",
        "cwc_code_interpreter",
        "cwc_code_interpreter_amsfix",
        "cwcfluxgptv",
        "flux_v3_gptv_enable_upload_multi_image_in_turn_wo_ch",
        "gptvnorm2048",
        "cwc_code_interpreter_citation_fix",
        "code_interpreter_interactive_charts",
        "cwc_code_interpreter_interactive_charts_inline_image",
        "code_interpreter_matplotlib_patching",
        "cwc_fileupload_odb",
        "update_memory_plugin",
        "add_custom_instructions",
        "cwc_flux_v3",
        "flux_v3_progress_messages",
        "enable_batch_token_processing",
        "enable_gg_gpt",
        "flux_v3_references",
        "flux_v3_references_entities",
        "flux_v3_image_gen_enable_dimensions",
        "flux_v3_image_gen_enable_non_watermarked_storage",
        "flux_v3_image_gen_enable_icon_dimensions",
        "flux_v3_image_gen_enable_system_text_with_params",
        "flux_v3_image_gen_enable_designer_dimensions_meta_prompting_in_system_prompts",
        "flux_v3_image_gen_enable_story",
        "rich_responses",
      ];
      const officeClientInfo = {
        clientPlatform: "mcmcopilot-web",
        clientAppName: "Office",
        clientEntrypoint: "mcmcopilot-officeweb",
        clientSessionId: uuid(),
        ProductCategory: "Chat",
        clientAppType: "Web",
        productEntryPoint: "ChatPanel",
        deviceOS: "macOS",
        deviceType: "Desktop",
        clientPlatformVersion: "10.15",
      };
      const owaClientInfo = {
        clientPlatform: "OwaHub-web",
        clientAppName: "OwaHub",
        clientEntrypoint: "owahub",
        clientSessionId: uuid(),
        clientAppType: "Web",
        deviceOS: "macOS",
        deviceType: "Desktop",
        clientPlatformVersion: "10.15",
      };
      return {
        args: {
          source: hasAttachments ? "officeweb" : "owahub",
          clientCorrelationId: rid,
          sessionId: uuid(),
          optionsSets: hasAttachments
            ? capturedAttachmentOptionsSets
            : [
                "enterprise_flux_web",
                "enterprise_flux_work",
                "enable_request_response_interstitials",
                "enterprise_flux_image_v1",
                "enterprise_toolbox_with_skdsstore_search_message_extensions",
                "enable_ME_auth_interstitial",
                "enable_confirmation_interstitial",
                "enable_plugin_auth_interstitial",
                "enable_response_action_processing",
                "enterprise_pagination_support",
                "search_result_progress_messages_with_search_queries",
                "flux_v3_gptv_enable_upload_multi_image_in_turn_wo_ch",
                "rich_responses",
                "gptvnorm2048",
                "enterprise_flux_work_code_interpreter",
                "cwc_code_interpreter_citation_fix",
                "code_interpreter_interactive_charts",
                "enterprise_code_interpreter_citation_fix",
                "cwc_code_interpreter_interactive_charts_inline_image",
                "code_interpreter_matplotlib_patching",
                "enable_batch_token_processing",
                "disable_cea_message_listener",
                "enable_selective_url_redaction",
                "update_memory_plugin",
                "add_custom_instructions",
                "agent_recommendations",
                "enable_gg_gpt",
                "enable_inferred_memory_read",
                "update_textdoc_response_after_streaming",
                "deepleo_networking_timeout_10minutes_canmore",
                "flux_v3_references",
                "flux_v3_references_entities",
                "flux_v3_image_gen_enable_dimensions",
                "flux_v3_image_gen_enable_non_watermarked_storage",
                "flux_v3_image_gen_enable_icon_dimensions",
                "flux_v3_image_gen_enable_system_text_with_params",
                "flux_v3_image_gen_enable_designer_dimensions_meta_prompting_in_system_prompts",
                "flux_v3_image_gen_enable_story",
                ...(fileUrlAnnotations.length ? ["cwc_fileupload_odb"] : []),
              ],
          streamingMode: "ConciseWithPadding",
          options: {},
          extraExtensionParameters: {},
          allowedMessageTypes: [
            "Chat",
            "Suggestion",
            "InternalSearchQuery",
            "Disengaged",
            "InternalLoaderMessage",
            "Progress",
            "GeneratedCode",
            "RenderCardRequest",
            "AdsQuery",
            "SemanticSerp",
            "GenerateContentQuery",
            "GenerateGraphicArt",
            "SearchQuery",
            "ConfirmationCard",
            "AuthError",
            "DeveloperLogs",
            "TriggerPlugin",
            "HintInvocation",
            "MemoryUpdate",
            "EndOfRequest",
            "TriggerConfirmation",
            "ResumeInvokeAction",
            "ResumeUserInputRequest",
            "TriggerUserInputRequest",
            "EscapeHatch",
            "TriggerPluginAuth",
            "ResumePluginAuth",
            ...(hasAttachments ? ["SideBySide"] : []),
            "ReferencesListComplete",
            ...(hasAttachments
              ? []
              : ["CompleteExtension", "TriggerExtension"]),
            "SwitchRespondingEndpoint",
          ],
          sliceIds: [],
          threadLevelGptId: {},
          traceId: rid,
          isStartOfSession: false,
          clientInfo: hasAttachments ? officeClientInfo : owaClientInfo,
          message: {
            author: "user",
            inputMethod: "Keyboard",
            text: text,
            entityAnnotationTypes: [
              "People",
              "File",
              "Event",
              "Email",
              "TeamsMessage",
            ],
            requestId: rid,
            locale: "zh-cn",
            messageType: "Chat",
            experienceType: "Default",
            adaptiveCards: [],
            messageAnnotations,
            clientPreferences: hasAttachments
              ? {}
              : { executionControls: { web: {}, work: {} } },
            ...(hasAttachments
              ? {
                  connectedFederatedConnections: ["dummyId"],
                  clientInfo: officeClientInfo,
                }
              : {}),
          },
          ...(hasAttachments
            ? {}
            : {
                gpts: [
                  {
                    id: "bizchat-as-gpt-scenario",
                    source: "BuiltInAgents",
                    clientOverrides: {
                      capabilities: [
                        { name: "WebSearch" },
                        { name: "WorkSearch" },
                      ],
                      "deepResearchModels@odata.type": "Collection(String)",
                    },
                  },
                ],
              }),
          plugins: [{ Id: "BingWebSearch", Source: "BuiltIn" }],
          ...(hasAttachments ? { isSbsSupported: true } : {}),
          tone: tone || "Claude_Opus",
          renderReferencesBehindEOS: true,
          disconnectBehavior: "continue",
        },
        requestId: rid,
      };
    }

    async function doAsk(
      id,
      text,
      tone,
      conversationId,
      attachments = [],
      requestedIdleTimeoutMs = 0,
    ) {
      let ws = null;
      // Register this turn so a bridge M365_STOP can abort it. A stop can arrive
      // during the pre-socket awaits below (token fetch / image upload) before
      // the socket-level teardown exists, so record intent in _veAborted and
      // install the real teardown (_veAbortLocal) once done() is defined.
      let _veAborted = false;
      let _veAbortLocal = null;
      const _veOnStop = () => {
        _veAborted = true;
        if (_veAbortLocal) {
          try {
            _veAbortLocal();
          } catch (_) {}
        }
      };
      _veActiveAsks.set(id, _veOnStop);
      try {
        const token = await requestSydneyToken();
        const claims = jwtClaims(token);
        const oid = claims.oid,
          tid = claims.tid;
        if (!oid || !tid) throw new Error("token has no oid/tid");

        // Python selects an existing ID; empty means create a new conversation.
        const convId = conversationId || uuid();
        dbg(
          "doAsk id=%s tone=%s convId=%s (%s) attachments=%d textLen=%d",
          id,
          tone,
          convId,
          conversationId ? "reused-from-python" : "new-uuid",
          Array.isArray(attachments) ? attachments.length : 0,
          String(text || "").length,
        );

        const sid = uuid(),
          reqSess = uuid();
        // 图片附件：send 前先上传拿 docId，回填到 attachment，供 buildChatArgs 生成
        // ImageFile 注解。上传失败只丢弃该图、不阻断整轮。
        for (const f of Array.isArray(attachments) ? attachments : []) {
          if (f && f.kind === "image" && f.dataBase64 && !f.docId) {
            try {
              f.docId = await uploadImageToM365(f, convId, oid, tid);
              dbg("image uploaded name=%s docId=%s", f.name, f.docId);
            } catch (e) {
              post({
                type: "M365_ATTACH_ERROR",
                id,
                name: f.name,
                error: String((e && e.message) || e),
              });
            }
          }
        }
        // transport/officeweb 只由 FileUrl 决定；纯图片轮保持 owahub。
        const hasAttachments =
          Array.isArray(attachments) &&
          attachments.some(
            (a) => a && a.url && a.verified === true && a.itemId && a.driveId,
          );
        const transportProfile = hasAttachments
          ? "&source=%22officeweb%22&product=Office&agentHost=Bizchat.ChatPanel" +
            "&licenseType=Starter&isEdu=true&agent=work&scenario=officeweb"
          : "&source=%22owahub%22&product=OwaHub&agentHost=Bizchat.FullScreen" +
            "&licenseType=Starter&isEdu=true&agent=work&scenario=owahub";
        const url =
          "wss://substrate.office.com/m365Copilot/Chathub/" +
          encodeURIComponent(oid) +
          "@" +
          encodeURIComponent(tid) +
          "?chatsessionid=" +
          reqSess +
          "&XRoutingParameterSessionKey=" +
          reqSess +
          "&clientrequestid=" +
          reqSess +
          "&X-SessionId=" +
          sid +
          "&ConversationId=" +
          encodeURIComponent(convId) +
          "&access_token=" +
          encodeURIComponent(token) +
          "&variants=" +
          VARIANTS +
          transportProfile;

        // Honor a stop that arrived during the pre-socket awaits: never open
        // the Chathub socket for a turn the client already abandoned.
        if (_veAborted) {
          _veActiveAsks.delete(id);
          dbg("doAsk aborted before socket open id=%s", id);
          return;
        }
        ws = new WebSocket(url); // browser supplies this frame's Origin
        const { args } = buildChatArgs(text, tone, convId, attachments);
        const invId = "0"; // 全新 socket,首个调用用 "0"(每 socket 只发一轮)
        // The answer may be delivered across MULTIPLE messageIds within a
        // single turn: when the model takes a tool-call / reasoning break, the
        // server closes one answer message and continues under a NEW messageId.
        // Each message's `text` is independently append-only, but the second
        // message does NOT start with the first — so a single cumulative `best`
        // guarded by startsWith() would reject the entire second segment and
        // freeze the answer. We therefore accumulate PER SEGMENT:
        //   committed = concatenation of all finished (superseded) segments
        //   curId/curBest = the messageId and cumulative text of the segment
        //                   currently streaming
        // and publish committed + curBest. This total is still strictly
        // append-only across the whole turn (committed only grows, curBest is
        // prefix-monotonic within its segment), so the Python relay's
        // incremental text[len(best):] contract is preserved.
        let committed = "";
        let curId = "";
        let curBest = "";
        const totalText = () => committed + curBest;
        // Chain-of-thought accumulator — MIRRORS the answer accumulator above,
        // but on the reasoning channel. The model's real reasoning is delivered
        // as Progress frames flagged addToChainOfThought:true, cumulative per
        // messageId, and spanning MULTIPLE messageIds per turn (observed:
        // 10 CoT messageIds interleaved with 5 answer messageIds). Forwarding
        // each frame's raw text would hit the SAME island-loss bug as C4 on the
        // Python side (its len(text) > len(best_reasoning) guard drops a shorter
        // new-segment snapshot), so we accumulate committed + current exactly
        // like the answer and emit the cumulative TOTAL as M365_REASONING.
        // Distinct thought bursts (messageId switches / non-prefix rewrites) are
        // joined with a blank line for readability; the total stays strictly
        // append-only so the Python reasoning slice text[len(best_reasoning):]
        // remains correct. This never reads or writes the answer channel.
        let cotCommitted = "";
        let cotCurId = "";
        let cotCurBest = "";
        const cotTotal = () => cotCommitted + cotCurBest;
        const commitCotSegment = () => {
          if (cotCurBest) cotCommitted += cotCurBest + "\n\n";
          cotCurBest = "";
        };
        const publishReasoning = (candidate, messageId) => {
          const value = String(candidate || "");
          if (!value) return;
          const mid = String(messageId || cotCurId || "");
          if (!cotCurId) {
            cotCurId = mid;
          } else if (mid && mid !== cotCurId) {
            commitCotSegment(); // new messageId -> new thought burst
            cotCurId = mid;
          }
          if (value === cotCurBest) return;
          if (value.startsWith(cotCurBest)) {
            cotCurBest = value; // in-segment growth (prefix-monotonic)
          } else {
            commitCotSegment(); // non-prefix rewrite -> start a fresh burst
            cotCurBest = value;
          }
          post({
            type: "M365_REASONING",
            id,
            text: cotTotal(),
            messageId: mid,
          });
        };
        // True once we have accepted at least one authoritative cumulative
        // snapshot for the CURRENT segment. While snapshots are flowing they are
        // the single source of truth; writeAtCursor deltas are a redundant
        // *provisional* view that can diverge (e.g. unresolved citation
        // placeholders like 【1-xxxx】 the server later rewrites to the canonical
        // \ue200cite\ue202…\ue201 form). Mixing them poisons the segment and
        // makes later authoritative snapshots fail the prefix guard.
        let sawSnapshot = false;
        let activeAnswerMessageId = "";
        // Whether the CURRENT answer segment issued a tool call (carried an
        // `invocation`). Capture-verified pattern (3/3 restarts): a
        // segment that ends in a tool call is superseded by a NEW messageId whose
        // text shares no prefix with it — either the model narrates a ReAct step
        // then continues or abandons a draft and re-answers. Either
        // way the finished segment was already streamed live to Continue over an
        // append-only SSE and CANNOT be retracted, so we keep it and just separate
        // it from the post-tool output with a blank line.
        let curSegHadInvocation = false;
        // A plain blank line only. The answer is already broken up by interleaved
        // thinking/tool blocks, so a paragraph break matches that rhythm; no visual
        // rule ("---") is added since we have no evidence for how the page renders
        // these segments and the same signal covers both ReAct narration and
        // re-answers.
        const TOOL_STEP_SEPARATOR = "\n\n";
        // Switch the active segment when the server moves to a new answer
        // messageId. The finished segment is committed verbatim; NO prefix check
        // is applied across segments (that is exactly the bug that dropped the
        // continuation). Within a segment, prefix-monotonicity still holds.
        const ensureSegment = (messageId) => {
          const mid = String(messageId || activeAnswerMessageId || curId || "");
          if (!mid) return;
          if (!curId) {
            curId = mid;
          } else if (mid !== curId) {
            committed += curBest;
            if (curSegHadInvocation) committed += TOOL_STEP_SEPARATOR;
            curBest = "";
            curId = mid;
            sawSnapshot = false; // snapshot-vs-writeAtCursor is per segment
            curSegHadInvocation = false; // tool-call flag is per segment
          }
        };
        let handshook = false;
        let terminal = false;
        // ---- Artifact harvesting ---------------------------------------
        // Capture the files THIS turn produced (CodeInterpreter / present_files
        // output) and forward their bytes so the background can upload them to
        // the SharePoint "download" folder. The real download URLs arrive inside
        // the answer message as sourceAttributions[*].seeMoreUrl and/or
        // references[*].targetLink, pointing at
        // asyncgw.teams.microsoft.com/v1/objects/… (verified in captures
        // references.targetLink, sourceAttributions.seeMoreUrl;
        // the field varies per turn, so BOTH are scanned). Those URLs are
        // fetchable with a bare, credential-less GET from this m365copilotapp.svc.cloud.microsoft
        // page origin (probe verified: default GET -> 200 + bytes;
        // credentials:include -> 401). URLs are collected while parsing frames
        // and fetched once at terminal, off the answer path.
        const artifactUrls = new Set();
        // 跨轮旧下载链接污染 + 挤掉新链接的根因隔离：M365 Chathub 的 type=2 终帧
        // item.messages 携带【整段会话历史】(含前几轮 bot 消息及其 sourceAttributions /
        // references 里的旧 asyncgw 对象 URL)。原先 collectArtifactsDeep 对整帧 JSON.stringify
        // 正则抓取，会把所有历史链接一并收进 artifactUrls；轮次越多旧链接越多，harvest 后
        // 既冒出大量旧链接，又会因同名产物在 background 侧被指纹去重而把当轮新链接挤掉。
        // 对策：只信任【本轮】实时流经的 messageId 与权威 result，历史消息一律不采集。
        const currentTurnMsgIds = new Set();
        let artifactsHarvested = false;
        // OBSERVATION-ONLY per-turn artifact-detection counters (gated by
        // window.__veM365Debug via dbg()). These NEVER touch artifactUrls or
        // any behavior; they exist only to classify, on the NEXT missing-link
        // occurrence, WHY artifactCount ended at 0 (read alongside the Python
        // [relay-artifact] line):
        //   loose_asyncgw=0            -> no AMS URL appeared in ANY frame this
        //     turn (upstream M365 did not emit it) == NOT a plugin bug.
        //   loose_asyncgw>0 & strict_urls=0 -> a URL WAS present but the strict
        //     AMS_OBJECT_RE_G (or cleanAmsUrl) rejected it (a detection gap);
        //     the per-frame "MISSED" dbg below shows its real shape to fix.
        const artScan = {
          frames: 0, // frames passed through collectArtifactsDeep
          looseAsyncgw: 0, // frames whose text contained "asyncgw"
          looseObjects: 0, // frames whose text contained "/v1/objects/"
          missed: 0, // frames with a loose marker but strict added nothing
          // Up to ARTSCAN_MISSED_CAP short, redacted windows around the loose
          // marker of a MISSED frame — rides the DONE payload to the PYTHON
          // [relay-artifact] log so the real URL SHAPE is diagnosable without
          // the page console. Bounded in count AND per-snippet length so a
          // pathological turn cannot bloat the terminal payload.
          missedSnippets: [],
          // Loose-URL fingerprints already accounted for this turn. A single
          // artifact URL is echoed across MANY type=1 frames; strict uses a
          // Set so only the FIRST frame grows artifactUrls, and every later
          // echo used to be counted as a fresh "missed" (9 identical MISSED
          // lines for one already-harvested file). Dedup on the normalized URL
          // so "missed" reflects only a GENUINELY-undetected url, not repeats.
          seenLoose: new Set(),
        };
        const ARTSCAN_MISSED_CAP = 6;
        const ARTSCAN_SNIPPET_LEN = 160;
        // Sanitize a captured AMS object URL before it enters artifactUrls.
        // Two failure modes were observed when a URL appeared in the ANSWER
        // PROSE rather than a citation field (e.g. the model literally writing
        // an asyncgw/objects URL in its answer text): (a) trailing markdown /
        // sentence punctuation captured from "[name](URL)" or "URL." → a
        // malformed target that 404s (".../voice_edge.py)"); (b) the bare
        // ".../v1/objects/" prefix with no object id → 405. Strip trailing
        // delimiters and REQUIRE a non-empty object-id path segment; return ""
        // for anything that is not a fetchable AMS object URL.
        const cleanAmsUrl = (raw) => {
          let u = String(raw || "").trim();
          u = u.replace(/[)\]}>,.;:!?'"\u3001\uff0c\u3002\uff09\u3011]+$/u, "");
          const m = /\/v1\/objects\/([^/?#\s]+)/i.exec(u);
          if (!/asyncgw/i.test(u) || !m || !m[1]) return "";
          return u;
        };
        const AMS_OBJECT_RE =
          /https:\/\/[^"'\s]*asyncgw[^"'\s]*\/v1\/objects\/[^"'\s]*/i;
        const collectArtifacts = (message) => {
          if (!message || typeof message !== "object") return;
          const scan = (url) => {
            const s = String(url || "");
            if (AMS_OBJECT_RE.test(s)) {
              const c = cleanAmsUrl(s);
              if (c) artifactUrls.add(c);
            }
          };
          const sa = message.sourceAttributions;
          if (Array.isArray(sa)) for (const s of sa) if (s) scan(s.seeMoreUrl);
          const refs = message.references;
          if (refs && typeof refs === "object")
            for (const k of Object.keys(refs)) {
              const r = refs[k];
              if (r) scan(r.targetLink);
            }
        };
        // 无死角兜底：整帧扫描 asyncgw 对象 URL。collectArtifacts 只看两个固定
        // 字段（seeMoreUrl / targetLink），而 URL 在不同轮次会落到不同字段（长回答 +
        // result.message 分叉时尤甚），漏掉就会 artifactUrls.size===0 -> artifactCount=0
        // -> background 不 hold 终态 -> appendText 从不生成 -> 链接与内联图全丢（本次现象）。
        // 对每一帧 JSON.stringify 后全局正则抓取，捕获 URL 无论嵌在哪个字段；Set 幂等，
        // 与上面的字段扫描叠加无副作用，也绝不触碰答案文本流。反斜杠排除在字符类外，
        // 避免把 JSON 转义或结尾的 \" 卷进 URL。
        const AMS_OBJECT_RE_G =
          /https:\/\/[^"'\s\\)\]]*asyncgw[^"'\s\\)\]]*\/v1\/objects\/[^"'\s\\)\]]*/gi;
        const collectArtifactsDeep = (frame) => {
          if (frame == null) return;
          let s;
          try {
            s = typeof frame === "string" ? frame : JSON.stringify(frame);
          } catch (_) {
            return;
          }
          if (!s) return;
          const before = artifactUrls.size;
          const hits = s.match(AMS_OBJECT_RE_G);
          if (hits)
            for (const u of hits) {
              const c = cleanAmsUrl(u);
              if (c) artifactUrls.add(c);
            }
          // OBSERVATION-ONLY (no behavior change): a loose substring probe that
          // runs regardless of whether the strict regex+cleanAmsUrl accepted a
          // url. If a frame clearly references an AMS artifact ("asyncgw" or
          // "/v1/objects/") yet NO new url entered artifactUrls, detection has a
          // gap; dbg a short redacted window around the marker so its real shape
          // is visible next time. Purely diagnostic; artifactUrls is untouched.
          try {
            artScan.frames += 1;
            const lo = s.toLowerCase();
            const hasAsyncgw = lo.indexOf("asyncgw") >= 0;
            const hasObjects = lo.indexOf("/v1/objects/") >= 0;
            if (hasAsyncgw) artScan.looseAsyncgw += 1;
            if (hasObjects) artScan.looseObjects += 1;
            if ((hasAsyncgw || hasObjects) && artifactUrls.size === before) {
              // This frame carried an AMS marker but added no NEW strict url.
              // Decide whether that is a GENUINE miss or just a repeat of an
              // already-known url. Extract this frame's loose AMS urls, cleanUp
              // each to the same key strict would use, and only count/report a
              // url that is neither already in artifactUrls (strict-detected)
              // nor already tallied this turn. If every loose url here is a
              // known repeat, this frame is benign and is NOT counted.
              const looseHits = s.match(AMS_OBJECT_RE_G) || [];
              const genuinelyNew = [];
              for (const h of looseHits) {
                const key = cleanAmsUrl(h) || h; // fall back to raw if unclean
                if (artifactUrls.has(key)) continue; // strict already has it
                if (artScan.seenLoose.has(key)) continue; // already tallied
                artScan.seenLoose.add(key);
                genuinelyNew.push(key);
              }
              if (genuinelyNew.length) {
                artScan.missed += 1;
                const at = lo.indexOf("asyncgw");
                const idx = at >= 0 ? at : lo.indexOf("/v1/objects/");
                const snippet = s
                  .slice(Math.max(0, idx - 24), idx + ARTSCAN_SNIPPET_LEN)
                  .replace(/\s+/g, " ");
                if (artScan.missedSnippets.length < ARTSCAN_MISSED_CAP) {
                  artScan.missedSnippets.push(
                    "t" + (frame && frame.type) + ":" + snippet,
                  );
                }
                if (window.__veM365Debug) {
                  dbg(
                    "[artifact-scan] loose marker but strict MISSED:",
                    "frameType=" + (frame && frame.type),
                    "snippet=" + JSON.stringify(snippet),
                  );
                }
              }
            }
          } catch (_) {}
        };
        // RFC 2047 encoded-word 解码（=?utf-8?B?..?= / =?utf-8?Q?..?=）。
        // AMS 的 Content-Disposition 常把中文名编成 encoded-word，
        // decodeURIComponent 不认这种编码，直接返回会残留 ? = 等字符，
        // 而这些正是 SharePoint 文件名的非法字符，导致 addUsingPath 被拒。
        const decodeMimeWord = (s) =>
          String(s).replace(
            /=\?(utf-8)\?([bq])\?([^?]*)\?=/gi,
            (_, _cs, enc, data) => {
              try {
                if (enc.toLowerCase() === "b") {
                  const bin = atob(data);
                  const bytes = Uint8Array.from(bin, (c) => c.charCodeAt(0));
                  return new TextDecoder("utf-8").decode(bytes);
                }
                // Q-encoding: _ => space, =XX => byte
                const q = data
                  .replace(/_/g, " ")
                  .replace(/=([0-9A-Fa-f]{2})/g, (_m, h) =>
                    String.fromCharCode(parseInt(h, 16)),
                  );
                const bytes = Uint8Array.from(q, (c) => c.charCodeAt(0));
                return new TextDecoder("utf-8").decode(bytes);
              } catch (_) {
                return "";
              }
            },
          );
        // SharePoint 文件名非法字符： \ / : * ? " < > | 以及首尾空白/点。
        const sanitizeName = (s) =>
          String(s || "")
            .trim()
            .replace(/[\\/:*?"<>|]/g, "_")
            .replace(/^\.+|\.+$/g, "")
            .trim();
        const artifactNameFromUrl = (url, disposition) => {
          const dispo = String(disposition || "");
          // 优先 RFC 5987: filename*=UTF-8''%E8%AF%BE...
          let m = /filename\*\s*=\s*(?:UTF-8'')?([^;]+)/i.exec(dispo);
          if (m && m[1]) {
            const decoded = sanitizeName(
              decodeURIComponent(m[1].trim().replace(/^"|"$/g, "")),
            );
            if (decoded) return decoded;
          }
          // 再 filename="..."，可能是 RFC 2047 encoded-word
          m = /filename\s*=\s*"?([^";]+)"?/i.exec(dispo);
          if (m && m[1]) {
            let raw = m[1].trim();
            if (/=\?[^?]+\?[bq]\?/i.test(raw)) raw = decodeMimeWord(raw);
            const decoded = sanitizeName(raw);
            if (decoded) return decoded;
          }
          const parts = String(url).split("?")[0].split("/").filter(Boolean);
          const last = parts[parts.length - 1] || "";
          // ".../views/original[/<name>]" — "original" means no trailing name.
          return last && last.toLowerCase() !== "original"
            ? sanitizeName(decodeURIComponent(last))
            : "artifact-" + Date.now();
        };
        const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
        // 由扩展名推断图片 mime。仅覆盖常见图片类型，其余返回 ""（保持 octet-stream）。
        const guessMimeFromName = (name) => {
          const ext = String(name || "")
            .toLowerCase()
            .split(".")
            .pop();
          return (
            {
              png: "image/png",
              jpg: "image/jpeg",
              jpeg: "image/jpeg",
              gif: "image/gif",
              webp: "image/webp",
              bmp: "image/bmp",
              svg: "image/svg+xml",
            }[ext] || ""
          );
        };
        // 由文件头 magic bytes 嗅探图片 mime（比扩展名更可靠，即便无扩展名也能识别）。
        const sniffImageMime = (bytes) => {
          const b = bytes || [];
          if (
            b.length >= 8 &&
            b[0] === 0x89 &&
            b[1] === 0x50 &&
            b[2] === 0x4e &&
            b[3] === 0x47
          )
            return "image/png";
          if (b.length >= 3 && b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff)
            return "image/jpeg";
          if (
            b.length >= 4 &&
            b[0] === 0x47 &&
            b[1] === 0x49 &&
            b[2] === 0x46 &&
            b[3] === 0x38
          )
            return "image/gif";
          if (
            b.length >= 12 &&
            b[0] === 0x52 &&
            b[1] === 0x49 &&
            b[2] === 0x46 &&
            b[3] === 0x46 &&
            b[8] === 0x57 &&
            b[9] === 0x45 &&
            b[10] === 0x42 &&
            b[11] === 0x50
          )
            return "image/webp";
          if (b.length >= 2 && b[0] === 0x42 && b[1] === 0x4d)
            return "image/bmp";
          return "";
        };
        // 从 base64 头部解出前若干字节用于嗅探（不解整包，够识别 magic 即可）。
        const headBytesFromB64 = (b64) => {
          try {
            const bin = atob(String(b64 || "").slice(0, 32));
            const out = new Uint8Array(bin.length);
            for (let i = 0; i < bin.length; i += 1) out[i] = bin.charCodeAt(i);
            return out;
          } catch (_) {
            return new Uint8Array(0);
          }
        };
        // 综合决策：magic 嗅探优先，其次扩展名，再次响应头，最后 octet-stream。
        // 关键点：无论走 page-tee 还是 fallback，图片都能拿到 image/* 类型，
        // 这样 background.js 的 _veIsImageMime 才会命中并内联为 data-URI；
        // 同时 data:image/png;base64 头正确，浏览器/Continue 才能直接渲染。
        const resolveMime = (bytes, name, headerMime) => {
          const sniffed = sniffImageMime(bytes);
          if (sniffed) return sniffed;
          const byName = guessMimeFromName(name);
          if (byName) return byName;
          const hm = String(headerMime || "").trim();
          if (hm && !/octet-stream/i.test(hm)) return hm;
          return "application/octet-stream";
        };
        const harvestOne = async (url) => {
          // The AMS endpoint serves bytes at ".../views/original"; the captured
          // link often appends the display filename, and that longer path 404s.
          // Normalize the request target but keep `url` for name/logging.
          const fetchUrl = normalizeAmsUrl(url);
          // BEST CASE: the page downloads the object itself (to build its blob
          // link); wrappedFetch tees those exact bytes with the page's real
          // auth. Poll briefly since that download may land shortly AFTER the
          // answer terminates.
          for (let i = 0; i < 12; i++) {
            const b64 = _amsBytes.get(fetchUrl);
            if (b64) {
              const nm = artifactNameFromUrl(url, null);
              post({
                type: "M365_ARTIFACT",
                id,
                url,
                name: nm,
                size: (function () {
                  try {
                    return atob(b64).length;
                  } catch (_) {
                    return 0;
                  }
                })(),
                // 曾硬编码 octet-stream，导致 background 的 _veIsImageMime 判定失败、
                // 图片被当成普通文件走纯链接，从不内联。改为按 magic bytes + 扩展名推断。
                mimeType: resolveMime(headBytesFromB64(b64), nm, null),
                data: b64,
                source: "page-tee",
              });
              return;
            }
            await sleep(500);
          }
          // FALLBACK: fetch it ourselves. Prefer the exact bearer the page used
          // for AMS (captured from its fetch/XHR — guaranteed-correct audience).
          // If the page never issued an AMS request (common for artifacts the
          // page does not auto-download), acquire an IC3-audience token
          // (aud=ic3.teams.office.com), the audience the AMS object endpoint
          // requires. The Sydney token is deliberately NOT used here: its
          // audience (substrate.office.com/sydney) is rejected by AMS with 401 —
          // that was the harvestArtifacts failure.
          let auth = _amsPageAuth;
          if (!auth) {
            try {
              const t = await requestAmsToken();
              if (t) auth = "Bearer " + t;
            } catch (_) {}
          }
          // Bounded retry with exponential backoff + jitter for the AMS fetch.
          // Error classification (retry must not waste attempts on permanent
          // failures):
          //   * transient  → retry: network/timeout (no status), 408/425/429,
          //     500/502/503/504, and any other unclassified status.
          //   * auth 401/403 → FORCE-refresh the IC3 token (bypassing both the
          //     content and background caches), then retry.
          //   * permanent  → do NOT retry: 400/404/405/410/411/413/414/415/422.
          //     A 404/405 here almost always means the URL is bogus or the
          //     object is gone (e.g. a placeholder URL that leaked into the
          //     answer prose), which no number of retries can fix.
          const AMS_PERMANENT = new Set([
            400, 404, 405, 410, 411, 413, 414, 415, 422,
          ]);
          const AMS_MAX_ATTEMPTS = 4;
          let lastError = "unknown";
          let lastStatus = 0;
          for (let attempt = 1; attempt <= AMS_MAX_ATTEMPTS; attempt++) {
            const headers = { Accept: "*/*", "MS-IC3-Product": "Copilot" };
            if (auth) headers["Authorization"] = auth;
            try {
              const res = await fetch(fetchUrl, {
                method: "GET",
                credentials: "omit",
                headers,
              });
              if (res.ok) {
                const buf = await res.arrayBuffer();
                const nm = artifactNameFromUrl(
                  url,
                  res.headers.get("content-disposition"),
                );
                const head = new Uint8Array(buf.slice(0, 16));
                post({
                  type: "M365_ARTIFACT",
                  id,
                  url,
                  name: nm,
                  size: buf.byteLength,
                  mimeType: resolveMime(
                    head,
                    nm,
                    res.headers.get("content-type"),
                  ),
                  data: bufToB64(buf),
                  source: auth ? "authed-fetch" : "bare-fetch",
                });
                return;
              }
              lastStatus = res.status;
              lastError = "fetch " + res.status;
              if (AMS_PERMANENT.has(res.status)) break;
              if (res.status === 401 || res.status === 403) {
                try {
                  const t = await requestAmsToken(true);
                  if (t) auth = "Bearer " + t;
                } catch (_) {}
              }
            } catch (e) {
              lastStatus = 0; // network/timeout → transient
              lastError = String((e && e.message) || e);
            }
            if (attempt < AMS_MAX_ATTEMPTS) {
              const backoff =
                Math.min(4000, 400 * 2 ** (attempt - 1)) +
                Math.floor(Math.random() * 200);
              await sleep(backoff);
            }
          }
          post({
            type: "M365_ARTIFACT_ERROR",
            id,
            url,
            fetchUrl,
            error: lastError,
            status: lastStatus,
            usedPageAuth: auth === _amsPageAuth && !!_amsPageAuth,
          });
        };
        const harvestArtifacts = () => {
          if (artifactsHarvested) return;
          artifactsHarvested = true;
          if (!artifactUrls.size) return;
          for (const url of artifactUrls) harvestOne(url);
        };
        // ----------------------------------------------------------------
        let timeoutTimer = null;
        let completionFallbackTimer = null;
        let type2ConversationId = "";
        let type2TurnState = "";
        // Code Interpreter can remain legitimately busy while Chathub emits no
        // user-visible answer text. Capture-verified execution markers are
        // Progress messages with contentType="Code" and a non-empty hiddenText;
        // GeneratedCode is the corresponding result. Keep this state strictly
        // per turn and bounded so ordinary requests retain their normal idle
        // protection.
        const CODE_EXECUTING_PROGRESS_MS = 30000;
        const CODE_EXECUTING_MAX_MS = 10 * 60 * 1000;
        let codeExecutingTimer = null;
        let codeExecutingDeadline = 0;
        const leaveCodeExecuting = (reason) => {
          if (codeExecutingTimer !== null) clearInterval(codeExecutingTimer);
          if (codeExecutingDeadline)
            dbg(
              "CODE_EXECUTING leave id=%s reason=%s",
              id,
              reason || "unknown",
            );
          codeExecutingTimer = null;
          codeExecutingDeadline = 0;
        };
        const enterCodeExecuting = (message) => {
          if (terminal || codeExecutingTimer !== null) return;
          codeExecutingDeadline = Date.now() + CODE_EXECUTING_MAX_MS;
          dbg(
            "CODE_EXECUTING enter id=%s messageId=%s",
            id,
            String((message && message.messageId) || ""),
          );
          // The triggering Progress frame is forwarded by the normal Progress
          // branch below. Only synthesize later empty liveness signals.
          codeExecutingTimer = setInterval(() => {
            if (terminal || Date.now() >= codeExecutingDeadline) {
              leaveCodeExecuting(terminal ? "terminal" : "bounded-timeout");
              return;
            }
            post({ type: "M365_PROGRESS", id, text: "", codeExecuting: true });
          }, CODE_EXECUTING_PROGRESS_MS);
        };
        const cursorMessageId = (cursor) => {
          const path = String((cursor && cursor.j) || "");
          // Example: $['66f8...'].adaptiveCards[0].body[0].text
          const match = /^\$\[['"]([^'"]+)['"]\]/.exec(path);
          return match ? match[1] : "";
        };
        // Citation-form normalizer. The server first streams an UNRESOLVED
        // placeholder (e.g. 【1-turn1file1】) and later REWRITES it IN PLACE to
        // the resolved private-use form \ue200cite\ue202…\ue201. Both encode the
        // same citation, so we collapse either span to one neutral token before
        // any prefix comparison. This lets an in-place citation rewrite be seen
        // as forward progress instead of a token retraction.
        const CITE_TOKEN = "\ue200\ue201";
        const stripCites = (s) =>
          String(s)
            .replace(/\ue200[\s\S]*?\ue201/g, CITE_TOKEN) // resolved cite span
            .replace(/【\d+-[^】]*】/g, CITE_TOKEN); // unresolved placeholder
        const publishSnapshot = (candidate, source, messageId) => {
          const value = String(candidate || "");
          if (!value) return;
          leaveCodeExecuting("answer-delta");
          // Route the snapshot to its segment first; a new messageId commits the
          // previous segment instead of being rejected as "non-prefix".
          ensureSegment(messageId);
          if (value === curBest) return;
          // Within the current segment the answer is append-only. A stale/short
          // or non-prefix rewrite of the SAME segment cannot be represented
          // without retracting tokens, so it is ignored — but this guard is now
          // scoped to curBest, never to text from an earlier segment.
          if (!value.startsWith(curBest)) {
            // Citation-aware retry: the divergence may be nothing more than the
            // server resolving 【1-xxxx】 into the \ue200cite\ue202…\ue201 form
            // at a position already inside curBest. With citation spans
            // neutralized, a genuine append still shows the current segment as a
            // prefix; if it does, this is a citation resolution (real forward
            // progress, no tokens retracted) and we adopt the newer, resolved
            // value. Only a normalized non-prefix is a true retraction/rewrite
            // that append-only cannot represent, so only that is dropped.
            if (!stripCites(value).startsWith(stripCites(curBest))) {
              console.warn("[VE-m365] ignored non-prefix answer snapshot", {
                source,
                messageId: curId,
                currentSegmentLength: curBest.length,
                incomingLength: value.length,
              });
              return;
            }
          }
          curBest = value;
          sawSnapshot = true;
          post({
            type: "M365_DELTA",
            id,
            text: totalText(),
            deltaSource: source,
          });
        };
        const publishCursorDelta = (delta) => {
          const value = String(delta || "");
          if (!value || !activeAnswerMessageId) return;
          leaveCodeExecuting("cursor-delta");
          // writeAtCursor is only used as a low-latency fallback for servers
          // that stream *without* periodic cumulative snapshots. The moment a
          // single authoritative snapshot has been seen for this segment, the
          // snapshot stream owns it and we must stop appending provisional
          // deltas — otherwise unresolved fragments (citation placeholders,
          // pre-redaction text) get permanently baked in and the corrective
          // snapshots that follow are rejected by the prefix guard.
          ensureSegment(activeAnswerMessageId);
          if (sawSnapshot) return;
          curBest += value;
          post({
            type: "M365_DELTA",
            id,
            text: totalText(),
            deltaSource: "writeAtCursor",
            messageId: activeAnswerMessageId,
          });
        };
        const finalTextFromType2 = (item) => {
          const result = (item && item.result) || {};
          if (typeof result.message === "string" && result.message)
            return result.message;
          if (Array.isArray(item && item.messages)) {
            const exact = item.messages.find(
              (message) =>
                message &&
                message.messageId === activeAnswerMessageId &&
                String(message.messageType || "").toLowerCase() !==
                  "progress" &&
                typeof message.text === "string",
            );
            if (exact) return exact.text;
          }
          return "";
        };
        const done = (payload) => {
          if (terminal) return;
          terminal = true;
          leaveCodeExecuting((payload && payload.type) || "terminal");
          try {
            dbg(
              "done id=%s type=%s convId=%s signal=%s authoritative=%s turnState=%s textLen=%d%s",
              id,
              payload && payload.type,
              (payload && payload.conversationId) || convId,
              (payload && payload.completionSignal) || "",
              !!(payload && payload.authoritative),
              (payload && payload.turnState) || "",
              payload && typeof payload.text === "string"
                ? payload.text.length
                : 0,
              payload && payload.error ? " error=" + payload.error : "",
            );
          } catch (_) {}
          // Fire-and-forget: harvest any artifact download URLs seen this turn.
          // Runs over HTTP, independent of the ws we are about to close; posts
          // M365_ARTIFACT messages that background uploads to SharePoint. Never
          // blocks or alters the terminal payload below.
          try {
            harvestArtifacts();
          } catch (_) {}
          // Tell background how many artifacts this turn produced BEFORE the
          // terminal is forwarded. Each M365_ARTIFACT / M365_ARTIFACT_ERROR
          // for this id arrives asynchronously (harvestOne polls page-tee'd
          // bytes for up to ~6s, sometimes AFTER this DONE), so background
          // cannot otherwise know whether to wait for SharePoint links before
          // finalizing. 0 (no artifacts) keeps the streaming path a strict
          // no-op. Only meaningful on the authoritative M365_DONE payload.
          try {
            if (payload && payload.type === "M365_DONE") {
              payload.artifactCount = artifactUrls.size;
              // Ride the per-turn artScan counters ON the DONE payload so they
              // reach the PYTHON [relay-artifact] log (via background
              // m365ForwardToPy) — no page console needed. Diagnostic only;
              // background does not act on it and it never touches answer text.
              payload.artScanSummary = {
                frames: artScan.frames,
                strictUrls: artifactUrls.size,
                looseAsyncgw: artScan.looseAsyncgw,
                looseObjects: artScan.looseObjects,
                missed: artScan.missed,
                missedSnippets: artScan.missedSnippets,
              };
            }
          } catch (_) {}
          // OBSERVATION-ONLY per-turn artifact-detection summary. Read it on the
          // authoritative DONE alongside the Python [relay-artifact] line:
          //   strict_urls=0 & loose_asyncgw=0  -> upstream emitted NO artifact
          //     URL this turn (not a plugin bug; nothing to harvest).
          //   strict_urls=0 & loose_asyncgw>0  -> the URL WAS present but strict
          //     detection/cleanAmsUrl missed it (a plugin gap); the per-frame
          //     "MISSED" dbg lines show the exact shape to fix with.
          //   strict_urls>0                    -> detection worked; a later
          //     missing link is a harvest/upload/hold issue, not detection.
          try {
            if (payload && payload.type === "M365_DONE") {
              dbg(
                "[artifact-scan] summary",
                "id=" + id,
                "frames=" + artScan.frames,
                "strict_urls=" + artifactUrls.size,
                "loose_asyncgw=" + artScan.looseAsyncgw,
                "loose_objects=" + artScan.looseObjects,
                "missed=" + artScan.missed,
              );
            }
          } catch (_) {}
          if (timeoutTimer !== null) clearTimeout(timeoutTimer);
          if (completionFallbackTimer !== null)
            clearTimeout(completionFallbackTimer);
          try {
            ws.close();
          } catch (_) {}
          _veActiveAsks.delete(id);
          post(payload);
        };
        // Socket-level teardown for a bridge M365_STOP: mark this turn terminal,
        // stop timers, close the Chathub socket so the model stops generating.
        // Setting terminal=true FIRST makes ws.onclose skip its "closed before
        // terminal" error and makes done() a no-op, so aborting posts NO terminal
        // answer frame to Python (whose relay queue for this id is already gone).
        // It DOES post a diagnostic-only M365_STOPPED receipt (background dlogs
        // it; never forwarded to Python, never enters the answer stream).
        _veAbortLocal = () => {
          if (terminal) return;
          terminal = true;
          if (timeoutTimer !== null) clearTimeout(timeoutTimer);
          if (completionFallbackTimer !== null)
            clearTimeout(completionFallbackTimer);
          try {
            if (ws) ws.close();
          } catch (_) {}
          _veActiveAsks.delete(id);
          post({ type: "M365_STOPPED", id });
          dbg("doAsk aborted by M365_STOP id=%s", id);
        };
        // A stop observed between socket creation and handler wiring is honored
        // now, before the read loop is armed.
        if (_veAborted) {
          _veAbortLocal();
          return;
        }
        // Idle watchdog, not a wall-clock cap. The previous implementation armed
        // a single 180s timeout at socket-open and never reset it, so a healthy
        // but long/slow answer (e.g. Claude Opus reasoning with web/work search
        // that legitimately streams for more than 180s of wall-clock) was killed
        // mid-stream with "M365 websocket timeout", truncating the reply — the
        // relay/consumer both showed the answer was byte-perfect up to the cut,
        // then an extension-side M365_ERROR. Reset on EVERY received frame
        // (including server type:6 pings) so this only fires when the socket is
        // physically silent for the full window. Whether answer *content* is
        // still arriving is a separate concern already enforced by the Python
        // relay's content-based idle timeout, so the two layers do not overlap.
        // Python computes one size/count-aware budget from the original upload
        // metadata and sends it with M365_ASK. Clamp here so a malformed relay
        // cannot disable the watchdog; older senders retain the relaxed 600s
        // behavior. Keep the 30-minute ceiling aligned with Python's
        // default M365_IDLE_MAX_SECONDS while still bounding malformed input.
        const IDLE_TIMEOUT_MS = Math.min(
          1800000,
          Math.max(600000, Number(requestedIdleTimeoutMs) || 600000),
        );
        const armIdleTimeout = () => {
          if (terminal) return;
          if (timeoutTimer !== null) clearTimeout(timeoutTimer);
          timeoutTimer = setTimeout(
            () =>
              done({
                type: "M365_ERROR",
                id,
                error:
                  "M365 websocket idle timeout (no frames for " +
                  Math.round(IDLE_TIMEOUT_MS / 1000) +
                  "s)",
              }),
            IDLE_TIMEOUT_MS,
          );
        };
        const finishFromType2Fallback = () => {
          completionFallbackTimer = null;
          done({
            type: "M365_DONE",
            id,
            text: totalText(),
            conversationId: type2ConversationId || convId,
            completionSignal: "type2-completed",
            authoritative: true,
            turnState: type2TurnState,
          });
        };

        ws.onopen = () => {
          ws.send(JSON.stringify({ protocol: "json", version: 1 }) + RS);
        };
        ws.onerror = () => done({ type: "M365_ERROR", id, error: "ws error" });
        ws.onclose = () => {
          if (!terminal)
            done({
              type: "M365_ERROR",
              id,
              error: "socket closed before terminal frame",
            });
        };
        ws.onmessage = (ev) => {
          // Any inbound frame proves the socket is alive; reset the idle
          // watchdog before processing it so a long/slow-but-healthy stream is
          // never cut off mid-answer.
          armIdleTimeout();
          const run = (s) => {
            for (const f of parseFrames(s)) {
              // 仅对【本轮实时流帧 type=1】做整帧深扫：type=1 只携带当轮流式消息，安全。
              // type=2 终帧含整段会话历史，若在此整帧深扫会把历史旧链接一并收入，故改到下面
              // 的 type===2 分支里按【本轮消息 + 权威 result】限定扫描。type=2 分支在同一
              // for 循环内先于随后的 type=3 done() 执行，读取 artifactUrls.size 的时序仍正确。
              if (f && f.type === 1) collectArtifactsDeep(f);
              if (!handshook) {
                // 首个 {} 是握手 OK
                handshook = true;
                const now = new Date().toISOString();
                const chatFrame = {
                  arguments: [args],
                  invocationId: invId,
                  target: "chat",
                  type: 4,
                };
                const metrics = {
                  arguments: [
                    {
                      Timestamps: {
                        ConnectionStart: now,
                        UserInputStart: now,
                        ConnectionEstablished: now,
                        UserInputSubmit: now,
                      },
                    },
                  ],
                  target: "Metrics",
                  type: 1,
                };
                ws.send(
                  JSON.stringify(chatFrame) + RS + JSON.stringify(metrics) + RS,
                );
                continue;
              }
              if (f.type === 6) {
                ws.send(JSON.stringify({ type: 6 }) + RS);
                continue;
              } // ping→pong
              if (f.type === 1 && f.target === "update") {
                const a = (f.arguments && f.arguments[0]) || {};
                const selectedMessageId = cursorMessageId(a.cursor);
                if (selectedMessageId) {
                  activeAnswerMessageId = selectedMessageId;
                  currentTurnMsgIds.add(String(selectedMessageId));
                }
                if (Array.isArray(a.messages))
                  for (const message of a.messages) {
                    // 记录本轮实时流经的每个 messageId(答案/进度/CoT 都属于当轮),供 type=2
                    // 终帧按本轮范围过滤 item.messages,隔离历史旧消息携带的旧下载链接。
                    if (message && message.messageId)
                      currentTurnMsgIds.add(String(message.messageId));
                    const messageType = String(
                      message.messageType || "",
                    ).toLowerCase();
                    if (messageType === "disengaged") {
                      done({ type: "M365_ERROR", id, error: "Disengaged" });
                      return;
                    }
                    // Scan every message for CodeInterpreter artifact download
                    // URLs (sourceAttributions.seeMoreUrl / references.targetLink);
                    // harmless on progress messages, which carry none.
                    collectArtifacts(message);
                    if (messageType === "progress") {
                      if (
                        String(message.contentType || "").toLowerCase() ===
                          "code" &&
                        typeof message.hiddenText === "string" &&
                        message.hiddenText.length > 0
                      )
                        enterCodeExecuting(message);
                      // Preserve protocol visibility without contaminating the
                      // append-only answer stream. The text is deliberately not
                      // interpreted, so localization/new wording is irrelevant.
                      post({
                        type: "M365_PROGRESS",
                        id,
                        text: String(message.text || ""),
                        messageId: String(message.messageId || ""),
                        contentType: String(message.contentType || ""),
                        // Forward the chain-of-thought flag so the Python relay
                        // can distinguish the model's actual reasoning stream
                        // (addToChainOfThought === true, cumulative per
                        // messageId) from mere status placeholders such as
                        // "请稍候…"/"正在思考..." (EarlyProgress, flag false).
                        // Only the former should be surfaced as reasoning.
                        addToChainOfThought:
                          message.addToChainOfThought === true,
                      });
                      // Additionally surface the model's REAL reasoning as an
                      // append-only M365_REASONING stream. Only frames flagged
                      // addToChainOfThought:true carry genuine chain-of-thought;
                      // EarlyProgress placeholders (flag false) stay liveness-
                      // only above. The M365_PROGRESS post is left untouched, so
                      // the existing keepalive/liveness path is unchanged — this
                      // is purely additive and never touches the answer channel.
                      if (
                        message.addToChainOfThought === true &&
                        typeof message.text === "string" &&
                        message.text
                      ) {
                        publishReasoning(
                          message.text,
                          String(message.messageId || ""),
                        );
                      }
                      continue;
                    }
                    if (messageType === "generatedcode")
                      leaveCodeExecuting("generated-code-result");
                    // Adopt the first genuine answer message even if the
                    // cursor frame has not arrived yet. Real answer messages
                    // carry an empty messageType and empty contentType; this
                    // deliberately excludes Progress / EscapeHatch ("Hide")
                    // and typed cards. Without this, any answer token that
                    // precedes the cursor frame is silently dropped — the
                    // classic "first character of the reply is missing" bug.
                    if (
                      !activeAnswerMessageId &&
                      !messageType &&
                      !message.contentType &&
                      message.messageId &&
                      typeof message.text === "string"
                    )
                      activeAnswerMessageId = message.messageId;
                    if (
                      activeAnswerMessageId &&
                      message.messageId === activeAnswerMessageId &&
                      typeof message.text === "string"
                    )
                      publishSnapshot(
                        message.text,
                        "messages-snapshot",
                        message.messageId,
                      );
                    // Remember that this answer segment issued a tool call. The
                    // `invocation` arrives on the segment's own frame (after its
                    // text), so curId already points at this segment; ensureSegment
                    // uses the flag to delimit it once a new segment supersedes it.
                    if ("invocation" in message) curSegHadInvocation = true;
                  }
                if (typeof a.writeAtCursor === "string")
                  publishCursorDelta(a.writeAtCursor);
              } else if (f.type === 2) {
                leaveCodeExecuting("authoritative-result");
                const item = f.item || {};
                const result = (item && item.result) || {};
                // Artifact download URLs (asyncgw/v1/objects/…) frequently live
                // ONLY on the authoritative type=2 result — its resolved
                // sourceAttributions/references — with no preceding type=1
                // snapshot carrying them (type=3 can follow type=2 in the same
                // frame). So we MUST scan the terminal here, but STRICTLY within
                // the current turn:
                //   - `result` 是本轮权威答案(其 sourceAttributions/references 已解析),
                //     浅扫两字段 + 深扫整棵 result 子树即可拿到当轮对象 URL(不论落在哪个字段)。
                //   - `item.messages` 携带【整段会话历史】——只扫 messageId ∈ currentTurnMsgIds
                //     的本轮消息;历史旧消息(及其旧 asyncgw 对象 URL)一律跳过,根除跨轮污染与
                //     "旧链接挤掉新链接"。此前无条件 collectArtifacts(item) + 遍历全部
                //     item.messages 正是旧链接来源,已移除。
                // artifactUrls 为 Set,与上面 type=1 的命中天然幂等去重。
                collectArtifacts(result);
                collectArtifactsDeep(result);
                if (Array.isArray(item.messages))
                  for (const m of item.messages) {
                    if (!m || !currentTurnMsgIds.has(String(m.messageId || "")))
                      continue;
                    collectArtifacts(m);
                    collectArtifactsDeep(m);
                  }
                if (typeof result.message === "string" && result.message) {
                  // AUTHORITATIVE text with citations resolved. CAPTURE-VERIFIED
                  // : result.message carries ONLY the LAST answer segment,
                  // NOT the whole turn — the earlier segment is absent from it.
                  //
                  // The type=1 accumulator has ALREADY streamed the full turn as
                  // committed(prior segments) + curBest(current segment); for a
                  // multi-segment / tool-preamble turn that is A + B. So the old
                  // `committed = result.message` OVERWROTE the whole accumulated
                  // A + B with B-only, dropping every earlier segment — the actual
                  // text-loss root cause. (It only looked fine when the Python relay
                  // happened to reject the shorter B-only post as non-prefix.)
                  //
                  // Reconcile instead of overwrite: result.message is the resolved
                  // form of the CURRENT segment. Keep the already-committed prior
                  // segments and adopt result.message as curBest. stripCites lets a
                  // pure citation resolution (【1-xxxx】 → \ue200cite…\ue201) of the
                  // current segment still be recognized, and the whole-turn case
                  // (result.message already spans committed) still adopts wholesale.
                  if (
                    committed &&
                    !stripCites(result.message).startsWith(
                      stripCites(committed),
                    )
                  ) {
                    // Multi-segment: result.message == authoritative CURRENT segment
                    // only. Preserve prior segments (committed), refresh curBest.
                    curBest = result.message;
                  } else {
                    // Single-segment turn, or result.message already spans the whole
                    // answer: adopt wholesale (original citation-resolution fix).
                    committed = result.message;
                    curBest = "";
                  }
                  sawSnapshot = true;
                  post({
                    type: "M365_DELTA",
                    id,
                    text: totalText(),
                    deltaSource: "type2-final",
                  });
                } else {
                  // No authoritative result.message this turn; fall back to the
                  // per-messageId text via the (now citation-aware) snapshot path.
                  const finalText = finalTextFromType2(item);
                  if (finalText)
                    publishSnapshot(
                      finalText,
                      "type2-final",
                      activeAnswerMessageId,
                    );
                }
                type2ConversationId = item.conversationId || convId;
                type2TurnState = String(item.turnState || "");
                // type=2 is the streamed item result, not the Hub invocation
                // completion. Keep parsing: captures show type=3 can follow in
                // the same WebSocket message after the RS separator.
                if (completionFallbackTimer !== null)
                  clearTimeout(completionFallbackTimer);
                completionFallbackTimer = setTimeout(
                  finishFromType2Fallback,
                  3000,
                );
                continue;
              } else if (f.type === 3) {
                if (f.error)
                  done({
                    type: "M365_ERROR",
                    id,
                    error: "COMPLETION: " + String(f.error),
                  });
                else
                  done({
                    type: "M365_DONE",
                    id,
                    text: totalText(),
                    conversationId: type2ConversationId || convId,
                    completionSignal: "signalr-type-3",
                    authoritative: true,
                    turnState: type2TurnState,
                  });
                return;
              }
            }
          };
          const d = ev.data;
          if (typeof d === "string") run(d);
          else if (d instanceof Blob)
            d.text()
              .then(run)
              .catch(() => {});
          else if (d instanceof ArrayBuffer) run(new TextDecoder().decode(d));
        };

        // 兜底超时(空闲看门狗:每收到一帧都会在 ws.onmessage 中重置)
        armIdleTimeout();
      } catch (e) {
        try {
          if (ws) ws.close();
        } catch (_) {}
        _veActiveAsks.delete(id);
        post({ type: "M365_ERROR", id, error: String((e && e.message) || e) });
      }
    }

    hookState.capabilities = {
      messageListener: true,
      tokenBridge: true,
      chatBuilder: typeof buildChatArgs === "function",
      doAsk: typeof doAsk === "function",
      websocket: typeof WebSocket === "function",
      attachments: true,
    };
    hookState.ready =
      hookState.capabilities.messageListener &&
      hookState.capabilities.chatBuilder &&
      hookState.capabilities.doAsk &&
      hookState.capabilities.websocket;
    if (!hookState.ready) {
      delete window.__veM365Hook;
      post({
        type: "M365_HOOK_ERROR",
        error: "M365 page hook initialization incomplete",
      });
      return;
    }
    if (frameMayHostSocket()) {
      post({
        type: "M365_FRAME_READY",
        capabilities: hookState.capabilities,
        frameOrigin: location.origin,
        frameUrl: location.href,
      });
    }
  }

  // Register relays before injection so PAGE_HOOK's immediate READY cannot race past us.
  window.addEventListener("message", (ev) => {
    const d = ev.data;
    if (!d || d.__veM365 !== true || d.dir !== "fromPage") return;
    try {
      browser.runtime.sendMessage(d);
    } catch (_) {}
  });
  browser.runtime.onMessage.addListener((msg) => {
    if (!msg || !msg.__veM365ToPage) return;
    const payload = msg.payload || {};
    window.postMessage(
      Object.assign({ __veM365: true, dir: "toPage" }, payload),
      "*",
    );
    return Promise.resolve({
      ok: true,
      relayed: true,
      frameUrl: location.href,
      attachmentCount: Array.isArray(payload.attachments)
        ? payload.attachments.length
        : -1,
    });
  });

  // Inject page world only after the relays exist.
  try {
    const s = document.createElement("script");
    s.textContent = "(" + PAGE_HOOK.toString() + ")();";
    (document.head || document.documentElement).appendChild(s);
    s.remove();
  } catch (e) {
    console.error("[VE-m365] inject failed", e);
  }

  // Explicit probe covers both fresh injection and extension reload on an existing page hook.
  window.postMessage(
    { __veM365: true, dir: "toPage", type: "M365_PROBE" },
    "*",
  );

  if (
    /^https:\/\/outlook\.cloud\.microsoft\/host\/[0-9a-f-]{36}\/entity1-[0-9a-f-]{36}/i.test(
      location.href,
    )
  ) {
    browser.runtime
      .sendMessage({
        __veM365: true,
        type: "M365_ENTRY_DISCOVERED",
        entryUrl: location.href.split(/[?#]/, 1)[0],
      })
      .catch(() => {});
  }
  console.log("[VE-m365] content script (self-built WSS) loaded");
})();
