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
 *       → 在正确的 outlook.office.com frame 建立 page-world Chathub WebSocket
 *       → 握手 {"protocol":"json","version":1}\x1e → 收 {} → 发 chat+Metrics 帧
 *       → 按 cursor messageId 读取正文快照/writeAtCursor 增量→ type2/type3 收尾
 * ========================================================================= */
(() => {
  "use strict";

  function PAGE_HOOK() {
    "use strict";
    // Origin self-gate. The manifest injects this script into every frame of
    // three match origins (outlook.cloud.microsoft, m365.cloud.microsoft,
    // outlook.office.com) with all_frames:true, so every host-shell subframe
    // also advertises itself as a Chathub-socket candidate. Only the
    // outlook.office.com frame ever builds the working socket (observed
    // frameOrigin in every VE-FRAME-READY). We therefore BLOCK the two proven
    // host-shell origins from advertising as socket candidates, and FAIL OPEN
    // for outlook.office.com and any unforeseen origin — gating can only ever
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
    function requestAmsToken() {
      return new Promise((resolve, reject) => {
        if (_amsTokenCache.token && _amsTokenCache.exp - Date.now() > 90000) {
          return resolve(_amsTokenCache.token);
        }
        _amsTokenWaiters.push({ resolve, reject });
        post({ type: "M365_NEED_AMS_TOKEN" });
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
        );
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

    async function doAsk(id, text, tone, conversationId, attachments = []) {
      let ws = null;
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
            curBest = "";
            curId = mid;
            sawSnapshot = false; // snapshot-vs-writeAtCursor is per segment
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
        // fetchable with a bare, credential-less GET from this outlook.office.com
        // page origin (probe verified: default GET -> 200 + bytes;
        // credentials:include -> 401). URLs are collected while parsing frames
        // and fetched once at terminal, off the answer path.
        const artifactUrls = new Set();
        let artifactsHarvested = false;
        const AMS_OBJECT_RE =
          /https:\/\/[^"'\s]*asyncgw[^"'\s]*\/v1\/objects\/[^"'\s]*/i;
        const collectArtifacts = (message) => {
          if (!message || typeof message !== "object") return;
          const scan = (url) => {
            const s = String(url || "");
            if (AMS_OBJECT_RE.test(s)) artifactUrls.add(s);
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
          /https:\/\/[^"'\s\\]*asyncgw[^"'\s\\]*\/v1\/objects\/[^"'\s\\]*/gi;
        const collectArtifactsDeep = (frame) => {
          if (frame == null) return;
          let s;
          try {
            s = typeof frame === "string" ? frame : JSON.stringify(frame);
          } catch (_) {
            return;
          }
          if (!s) return;
          const hits = s.match(AMS_OBJECT_RE_G);
          if (hits) for (const u of hits) artifactUrls.add(u);
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
          const headers = { Accept: "*/*", "MS-IC3-Product": "Copilot" };
          if (auth) headers["Authorization"] = auth;
          try {
            const res = await fetch(fetchUrl, {
              method: "GET",
              credentials: "omit",
              headers,
            });
            if (!res.ok)
              return post({
                type: "M365_ARTIFACT_ERROR",
                id,
                url,
                fetchUrl,
                error: "fetch " + res.status,
                usedPageAuth: auth === _amsPageAuth && !!_amsPageAuth,
              });
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
              // AMS 下载端点常返回 octet-stream，直接采信会漏判图片。改为按 magic bytes +
              // 扩展名推断，仅在两者都无结论时才回退到响应头 / octet-stream。
              mimeType: resolveMime(head, nm, res.headers.get("content-type")),
              data: bufToB64(buf),
              source: auth ? "authed-fetch" : "bare-fetch",
            });
          } catch (e) {
            post({
              type: "M365_ARTIFACT_ERROR",
              id,
              url,
              fetchUrl,
              error: String((e && e.message) || e),
            });
          }
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
            }
          } catch (_) {}
          if (timeoutTimer !== null) clearTimeout(timeoutTimer);
          if (completionFallbackTimer !== null)
            clearTimeout(completionFallbackTimer);
          try {
            ws.close();
          } catch (_) {}
          post(payload);
        };
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
        const IDLE_TIMEOUT_MS = 180000;
        const armIdleTimeout = () => {
          if (terminal) return;
          if (timeoutTimer !== null) clearTimeout(timeoutTimer);
          timeoutTimer = setTimeout(
            () =>
              done({
                type: "M365_ERROR",
                id,
                error: "M365 websocket idle timeout (no frames for 180s)",
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
              // 每帧先做无死角 URL 扫描，保证 type=2/type=3 的 done() 读取
              // artifactUrls.size 时，本帧携带的对象 URL 已全部入集（type=2 与
              // type=3 常在同一 ws 消息内先后到达，循环里 type=2 先被扫，再轮到
              // type=3 的 done()，时序正确）。
              collectArtifactsDeep(f);
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
                if (selectedMessageId)
                  activeAnswerMessageId = selectedMessageId;
                if (Array.isArray(a.messages))
                  for (const message of a.messages) {
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
                  }
                if (typeof a.writeAtCursor === "string")
                  publishCursorDelta(a.writeAtCursor);
              } else if (f.type === 2) {
                const item = f.item || {};
                const result = (item && item.result) || {};
                // Artifact download URLs (asyncgw/v1/objects/…) frequently live
                // ONLY on the authoritative type=2 result — its resolved
                // sourceAttributions/references — with no preceding type=1
                // snapshot carrying them (type=3 can follow type=2 in the same
                // frame). collectArtifacts was previously called only in the
                // type=1 update loop, so those turns finished with
                // artifactUrls.size===0 -> artifactCount=0 -> background never
                // held the terminal -> the SharePoint link block was silently
                // dropped (links only, answer text unaffected). Scan the
                // authoritative result/item/messages here too. artifactUrls is
                // a Set, so this is idempotent with any type=1 hits — no dupes.
                collectArtifacts(result);
                collectArtifacts(item);
                if (Array.isArray(item.messages))
                  for (const m of item.messages) collectArtifacts(m);
                if (typeof result.message === "string" && result.message) {
                  // AUTHORITATIVE full-turn text: this is the server's final
                  // message with citations already resolved — the single source
                  // of truth for the whole answer. Adopt it WHOLESALE and bypass
                  // the append-only prefix guard (that guard exists for
                  // provisional streaming snapshots, not for the final message).
                  // This is the core fix: previously result.message was fed back
                  // through publishSnapshot and rejected as "non-prefix" when it
                  // resolved 【1-xxxx】 into the \ue200cite…\ue201 form, freezing
                  // the answer at the citation and dropping the tail. totalText()
                  // (used by the type=3 DONE below) now equals this full text.
                  committed = result.message;
                  curBest = "";
                  sawSnapshot = true;
                  post({
                    type: "M365_DELTA",
                    id,
                    text: committed,
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
