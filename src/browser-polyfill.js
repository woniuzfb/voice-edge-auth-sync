/*
 * Minimal WebExtension shim — safe on BOTH Firefox and Chromium.
 *
 * Voice Edge is written entirely against the promise-based `browser.*`
 * namespace (0 `chrome.*` call sites).
 *
 *   - Firefox already ships a native promise-based `browser.*`, so this file
 *     DETECTS that and becomes a no-op (it must NOT overwrite the native one
 *     with Chromium's callback-based `chrome`).
 *   - Chromium MV3's `chrome.*` returns Promises when called without a
 *     callback, so aliasing `browser = chrome` covers ~99% of the surface.
 *     The one gap is `runtime.onMessage`: Firefox lets a listener RETURN a
 *     Promise/value to reply, whereas Chromium needs `sendResponse()` +
 *     `return true`. We wrap onMessage to bridge that.
 *
 * Load order matters: this file MUST run before background.js / content
 * scripts / popup.js.
 */
(function (glob) {
  "use strict";

  var chromeApi = glob.chrome;

  // Firefox ships a native promise-based `browser` that is a DISTINCT object
  // from `chrome`. Detect that (works in both background and content contexts)
  // and do nothing, so we never clobber the native namespace with the
  // callback-based `chrome`. UA check is a belt-and-suspenders fallback.
  var isFirefox =
    (typeof glob.browser !== "undefined" &&
      glob.browser &&
      glob.browser !== chromeApi &&
      glob.browser.runtime) ||
    (typeof navigator !== "undefined" &&
      /firefox/i.test(navigator.userAgent || ""));
  if (isFirefox) return;

  if (!chromeApi || !chromeApi.runtime) return; // not a WebExtension context
  if (glob.browser && glob.browser.__vePolyfill) return; // idempotent

  // --- Bridge runtime.onMessage Firefox semantics onto Chromium -------------
  try {
    var evt = chromeApi.runtime.onMessage;
    if (evt && !evt.__vePolyfillWrapped) {
      var rawAdd = evt.addListener.bind(evt);
      var rawRemove = evt.removeListener.bind(evt);
      var rawHas = evt.hasListener ? evt.hasListener.bind(evt) : null;
      var wrappedMap = new WeakMap();

      evt.addListener = function (listener) {
        if (typeof listener !== "function") return rawAdd.apply(null, arguments);
        var wrapped = wrappedMap.get(listener);
        if (!wrapped) {
          wrapped = function (message, sender, sendResponse) {
            var result;
            try {
              result = listener(message, sender, sendResponse);
            } catch (e) {
              return false; // let other listeners handle it
            }
            // Firefox: a Promise return means "respond asynchronously".
            if (result && typeof result.then === "function") {
              result.then(
                function (value) {
                  try { sendResponse(value); } catch (_) {}
                },
                function (err) {
                  try {
                    sendResponse({
                      __error: String((err && err.message) || err),
                    });
                  } catch (_) {}
                }
              );
              return true; // keep the message channel open
            }
            if (result === true) return true; // legacy sendResponse style
            if (typeof result !== "undefined") {
              try { sendResponse(result); } catch (_) {}
              return false;
            }
            return false; // not handled by this listener
          };
          wrappedMap.set(listener, wrapped);
        }
        return rawAdd(wrapped);
      };

      evt.removeListener = function (listener) {
        return rawRemove(wrappedMap.get(listener) || listener);
      };
      if (rawHas) {
        evt.hasListener = function (listener) {
          return rawHas(wrappedMap.get(listener) || listener);
        };
      }
      try {
        Object.defineProperty(evt, "__vePolyfillWrapped", { value: true });
      } catch (_) {}
    }
  } catch (_) {}

  // --- Expose the `browser` namespace ---------------------------------------
  try {
    Object.defineProperty(chromeApi, "__vePolyfill", { value: true });
  } catch (_) {}
  glob.browser = chromeApi;
})(typeof self !== "undefined" ? self : this);
