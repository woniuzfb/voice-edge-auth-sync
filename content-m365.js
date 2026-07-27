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
 *       → 读 update 流(writeAtCursor / messages 快照取最长)→ type2/type3 收尾
 * ========================================================================= */
(() => {
  "use strict";

  function PAGE_HOOK() {
    "use strict";
    const priorHook = window.__veM365Hook;
    if (priorHook && priorHook.ready === true) {
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

    // ---- 捕获页面 token fetch 的成功响应，保存服务端轮换后的 refresh_token ----
    (function hookTokenFetch() {
      const originalFetch = window.fetch;
      if (originalFetch.__veM365TokenHook) return;

      async function wrappedFetch(input, init) {
        const response = await originalFetch.apply(this, arguments);
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
      XMLHttpRequest.prototype.open = function (method, url) {
        try {
          this.__veUrl = String(url || "");
        } catch (_) {}
        return O.apply(this, arguments);
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
    // background 回传 token
    window.addEventListener("message", (ev) => {
      const d = ev.data;
      if (!d || d.__veM365 !== true || d.dir !== "toPage") return;
      if (d.type === "M365_PROBE") {
        post({
          type: "M365_FRAME_READY",
          frameOrigin: location.origin,
          frameUrl: location.href,
        });
      } else if (d.type === "M365_TOKEN") {
        _tokenCache = { token: String(d.token || ""), exp: Number(d.exp || 0) };
        while (_tokenWaiters.length) {
          const w = _tokenWaiters.shift();
          _tokenCache.token
            ? w.resolve(_tokenCache.token)
            : w.reject(new Error(d.error || "no token"));
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
      const messageAnnotations = attachments
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
      const hasAttachments = messageAnnotations.length > 0;
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
                ...(messageAnnotations.length ? ["cwc_fileupload_odb"] : []),
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

        const sid = uuid(),
          reqSess = uuid();
        const hasAttachments =
          Array.isArray(attachments) && attachments.length > 0;
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
        let best = "";
        let cursorText = "";
        let handshook = false;
        let terminal = false;
        let timeoutTimer = null;
        const publishLongest = (candidate) => {
          const value = String(candidate || "");
          if (value.length > best.length) {
            best = value;
            post({ type: "M365_DELTA", id, text: best });
          }
        };
        const done = (payload) => {
          if (terminal) return;
          terminal = true;
          if (timeoutTimer !== null) clearTimeout(timeoutTimer);
          try {
            ws.close();
          } catch (_) {}
          post(payload);
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
          const run = (s) => {
            for (const f of parseFrames(s)) {
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
                if (Array.isArray(a.messages))
                  for (const message of a.messages) {
                    if (message.messageType === "Disengaged") {
                      done({ type: "M365_ERROR", id, error: "Disengaged" });
                      return;
                    }
                    if (typeof message.text === "string")
                      publishLongest(message.text);
                  }
                if (typeof a.writeAtCursor === "string") {
                  cursorText += a.writeAtCursor;
                  publishLongest(cursorText);
                }
              } else if (f.type === 2) {
                const res = (f.item && f.item.result) || {};
                if (
                  typeof res.message === "string" &&
                  res.message.length > best.length
                )
                  best = res.message;
                const cid = (f.item && f.item.conversationId) || convId;
                done({
                  type: "M365_DONE",
                  id,
                  text: best,
                  conversationId: cid,
                });
                return;
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
                    text: best,
                    conversationId: convId,
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

        // 兜底超时
        timeoutTimer = setTimeout(
          () =>
            done({ type: "M365_ERROR", id, error: "M365 websocket timeout" }),
          180000,
        );
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
    post({
      type: "M365_FRAME_READY",
      capabilities: hookState.capabilities,
      frameOrigin: location.origin,
      frameUrl: location.href,
    });
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
