/*
 * Voice Edge · Doubao completion capture (Chromium replacement for Firefox's
 * webRequest.filterResponseData).
 *
 * Chromium's webRequest can never read a response body, so we capture the
 * /chat/completion SSE stream in the PAGE world by tee-ing window.fetch (and,
 * defensively, XMLHttpRequest). The raw SSE text is relayed:
 *
 *   page world (fetch hook) --window.postMessage--> this content script
 *     --browser.runtime.sendMessage--> background (parseSseBody + auto-sync)
 *
 * Background parses the body exactly as before; the only thing that changed is
 * where the bytes come from. The page's own reader is untouched (we tee), so
 * Doubao keeps streaming normally.
 */
(function () {
  "use strict";

  var COMPLETION_RE = /^https:\/\/www\.doubao\.com\/chat\/completion/;
  var MAX_CAPTURE_BYTES = 4 * 1024 * 1024;

  // ---- Page-world hook (serialized via toString and injected) --------------
  function PAGE_HOOK() {
    var RE = /^https:\/\/www\.doubao\.com\/chat\/completion/;
    var MAX = 4 * 1024 * 1024;

    function post(url, startedAt, body) {
      try {
        window.postMessage(
          {
            __veDoubao: true,
            dir: "fromPage",
            type: "DOUBAO_CAPTURE",
            url: String(url || ""),
            startedAt: startedAt,
            completedAt: Date.now(),
            body: body,
          },
          "*"
        );
      } catch (_) {}
    }

    // --- fetch() -------------------------------------------------------------
    try {
      var origFetch = window.fetch;
      if (typeof origFetch === "function" && !origFetch.__veHooked) {
        var hookedFetch = function (input, init) {
          var url =
            typeof input === "string"
              ? input
              : (input && input.url) || "";
          var method = String(
            (init && init.method) || (input && input.method) || "GET"
          ).toUpperCase();
          var startedAt = Date.now();
          var p = origFetch.apply(this, arguments);
          if (!RE.test(url) || method !== "POST") return p;
          return p.then(function (resp) {
            try {
              if (!resp || !resp.body) return resp;
              var branches = resp.body.tee();
              var passthrough = branches[0];
              var capture = branches[1];
              (function () {
                var reader = capture.getReader();
                var dec = new TextDecoder();
                var body = "";
                var bytes = 0;
                var overflow = false;
                function pump() {
                  reader.read().then(
                    function (r) {
                      if (r.done) {
                        if (!overflow) {
                          body += dec.decode();
                          post(url, startedAt, body);
                        }
                        return;
                      }
                      if (!overflow) {
                        bytes += r.value.byteLength;
                        if (bytes > MAX) {
                          overflow = true;
                          body = "";
                        } else {
                          body += dec.decode(r.value, { stream: true });
                        }
                      }
                      pump();
                    },
                    function () {}
                  );
                }
                pump();
              })();
              return new Response(passthrough, {
                status: resp.status,
                statusText: resp.statusText,
                headers: resp.headers,
              });
            } catch (_) {
              return resp;
            }
          });
        };
        hookedFetch.__veHooked = true;
        window.fetch = hookedFetch;
      }
    } catch (_) {}

    // --- XMLHttpRequest (defensive fallback) --------------------------------
    try {
      var XHR = window.XMLHttpRequest;
      if (XHR && XHR.prototype && !XHR.prototype.__veHooked) {
        var origOpen = XHR.prototype.open;
        var origSend = XHR.prototype.send;
        XHR.prototype.open = function (method, url) {
          try {
            this.__veUrl = String(url || "");
            this.__veMethod = String(method || "GET").toUpperCase();
          } catch (_) {}
          return origOpen.apply(this, arguments);
        };
        XHR.prototype.send = function () {
          try {
            if (RE.test(this.__veUrl || "") && this.__veMethod === "POST") {
              var started = Date.now();
              var xhr = this;
              this.addEventListener("load", function () {
                try {
                  var text =
                    typeof xhr.responseText === "string"
                      ? xhr.responseText
                      : "";
                  if (text && text.length <= MAX) {
                    post(xhr.__veUrl, started, text);
                  }
                } catch (_) {}
              });
            }
          } catch (_) {}
          return origSend.apply(this, arguments);
        };
        try {
          Object.defineProperty(XHR.prototype, "__veHooked", { value: true });
        } catch (_) {}
      }
    } catch (_) {}
  }

  // ---- Isolated-world relay ------------------------------------------------
  window.addEventListener("message", function (ev) {
    var d = ev && ev.data;
    if (!d || d.__veDoubao !== true || d.dir !== "fromPage") return;
    if (d.type !== "DOUBAO_CAPTURE") return;
    if (!COMPLETION_RE.test(String(d.url || ""))) return;
    var body = String(d.body || "");
    if (body.length > MAX_CAPTURE_BYTES) return;
    try {
      browser.runtime
        .sendMessage({
          type: "DOUBAO_CAPTURE",
          url: d.url,
          startedAt: Number(d.startedAt || Date.now()),
          completedAt: Number(d.completedAt || Date.now()),
          body: body,
        })
        .catch(function () {});
    } catch (_) {}
  });

  // Inject the page-world hook as early as possible (document_start).
  try {
    var s = document.createElement("script");
    s.textContent = "(" + PAGE_HOOK.toString() + ")();";
    (document.head || document.documentElement).appendChild(s);
    s.remove();
  } catch (e) {
    // If the page CSP blocks inline injection, capture silently degrades and
    // the user can still fall back to a manual popup sync.
    try {
      console.warn("[VE-doubao] page hook injection failed", e);
    } catch (_) {}
  }
})();
