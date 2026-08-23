"use strict";

const MODEL_TYPES = new Set(["default", "expert", "vision"]);
const MODEL_LABELS = new Map([
  ["快速模式", "default"],
  ["专家模式", "expert"],
  ["识图模式", "vision"],
]);
const CONVERSATION_PATH_RE = /^\/a\/chat\/s\/([^/?#]+)/;
const SNAPSHOT_DELAY_MS = 120;
const ROUTE_POLL_MS = 500;

let lastSnapshotKey = "";
let lastPathname = location.pathname;
let snapshotTimer = null;
let modeObserver = null;
let observedModeRoot = null;

function currentConversationId() {
  const match = location.pathname.match(CONVERSATION_PATH_RE);
  return match ? decodeURIComponent(match[1]) : "";
}

function textModelType(root = document) {
  for (const element of root.querySelectorAll("span")) {
    const label = String(element.textContent || "").trim();
    const modelType = MODEL_LABELS.get(label);
    if (modelType) return modelType;
  }
  return "";
}

function hasToggleLabel(label) {
  return [...document.querySelectorAll("[aria-pressed]")].some((element) =>
    [...element.querySelectorAll("span")].some(
      (span) => String(span.textContent || "").trim() === label,
    ),
  );
}

function capabilityModelType() {
  // The composer capability row is a stable fallback for an existing locked
  // conversation: default has search + upload, vision has upload only, and
  // expert has neither. Deep Thinking is intentionally ignored because all
  // three modes expose it.
  const hasSearch = hasToggleLabel("智能搜索");
  const hasUpload = Boolean(
    document.querySelector('input[type="file"][multiple]'),
  );
  if (hasSearch && hasUpload) return "default";
  if (!hasSearch && hasUpload) return "vision";
  if (!hasSearch && !hasUpload) return "expert";
  return "";
}

function detectModelType() {
  // New/unsubmitted conversations expose the interactive radio group.
  const selected = document.querySelector(
    '[role="radio"][data-model-type][aria-checked="true"]',
  );
  const selectedType = selected && selected.getAttribute("data-model-type");
  if (MODEL_TYPES.has(selectedType)) return selectedType;

  // Existing conversations normally render an immutable compact mode badge.
  const badgeType = textModelType(document);
  if (MODEL_TYPES.has(badgeType)) return badgeType;

  // Fall back to the composer capability row if the badge is delayed, hidden,
  // or changed by a future DeepSeek layout revision.
  return capabilityModelType();
}

function currentDeepSeekConversationSnapshot() {
  const conversationId = currentConversationId();
  if (!conversationId) return null;
  const modelType = detectModelType();
  if (!MODEL_TYPES.has(modelType)) return null;
  return {
    type: "DEEPSEEK_CONVERSATION_SNAPSHOT",
    conversation_id: conversationId,
    model_type: modelType,
    page_url: location.href,
    captured_at: Date.now(),
  };
}

function publishConversationSnapshot() {
  const snapshot = currentDeepSeekConversationSnapshot();
  if (!snapshot) return false;
  const key = `${snapshot.conversation_id}:${snapshot.model_type}`;
  if (key === lastSnapshotKey) return true;
  lastSnapshotKey = key;
  browser.runtime.sendMessage(snapshot).catch(() => {});
  return true;
}

function scheduleSnapshot() {
  if (snapshotTimer) clearTimeout(snapshotTimer);
  snapshotTimer = setTimeout(() => {
    snapshotTimer = null;
    publishConversationSnapshot();
    attachModeObserver();
  }, SNAPSHOT_DELAY_MS);
}

function findModeRoot() {
  const selected = document.querySelector(
    '[role="radio"][data-model-type][aria-checked="true"]',
  );
  if (selected)
    return selected.closest('[role="radiogroup"]') || selected.parentElement;

  for (const span of document.querySelectorAll("span")) {
    if (!MODEL_LABELS.has(String(span.textContent || "").trim())) continue;
    // The compact existing-conversation badge is small; observe only its
    // nearest stable content wrapper instead of the full streaming transcript.
    return span.closest("div") || span.parentElement;
  }
  return null;
}

function attachModeObserver() {
  const root = findModeRoot();
  if (!root || root === observedModeRoot) return;
  if (modeObserver) modeObserver.disconnect();
  observedModeRoot = root;
  modeObserver = new MutationObserver(scheduleSnapshot);
  modeObserver.observe(root, {
    subtree: true,
    childList: true,
    characterData: true,
    attributes: true,
    attributeFilter: ["aria-checked", "data-model-type"],
  });
}

document.addEventListener(
  "click",
  (event) => {
    if (event.target.closest('[role="radio"][data-model-type]'))
      scheduleSnapshot();
  },
  true,
);

// DeepSeek uses history-based SPA navigation. Poll only the pathname, then
// re-discover the small mode subtree after a conversation change.
setInterval(() => {
  if (location.pathname === lastPathname) return;
  lastPathname = location.pathname;
  lastSnapshotKey = "";
  observedModeRoot = null;
  if (modeObserver) modeObserver.disconnect();
  modeObserver = null;
  scheduleSnapshot();
}, ROUTE_POLL_MS);

window.addEventListener("popstate", scheduleSnapshot);
window.addEventListener("hashchange", scheduleSnapshot);
scheduleSnapshot();
