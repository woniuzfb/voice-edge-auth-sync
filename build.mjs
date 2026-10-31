#!/usr/bin/env node
/*
 * build.mjs — build the Firefox (MV2) or Chrome (MV3) extension from a single
 * source tree.
 *
 *   node build.mjs                 # build both (firefox + chrome)
 *   node build.mjs firefox         # build one target
 *   node build.mjs chrome --zip    # build + zip
 *   node build.mjs all --zip       # build both + zip each
 *   node build.mjs --no-clean      # keep dist/ contents, only overwrite
 *
 * Layout:
 *   src/                     shared source (background.js is platform-neutral)
 *   manifest.base.json       shared manifest fields
 *   manifest.firefox.json    MV2 overrides
 *   manifest.chrome.json     MV3 overrides
 *   dist/<target>/           build output (this script writes here)
 *
 * The merge is a shallow top-level override: every platform-specific block
 * (permissions, background, content_scripts, action/browser_action, …) lives
 * whole in exactly one override file, so there is no deep-merge ambiguity.
 * Only the files a target's manifest actually references get copied, so the
 * Firefox build never ships content-doubao.js and the Chrome build never
 * ships a stray MV2 leftover.
 */
import {
  readFileSync,
  writeFileSync,
  mkdirSync,
  rmSync,
  cpSync,
  existsSync,
  readdirSync,
} from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { execSync } from "node:child_process";

const ROOT = dirname(fileURLToPath(import.meta.url));
const SRC = join(ROOT, "src");
const DIST = join(ROOT, "dist");
const PKG_NAME = "voice-edge-auth-sync";

const TARGETS = {
  firefox: "manifest.firefox.json",
  chrome: "manifest.chrome.json",
};

function readJSON(p) {
  return JSON.parse(readFileSync(p, "utf8"));
}

function mergeManifest(target) {
  const base = readJSON(join(ROOT, "manifest.base.json"));
  const override = readJSON(join(ROOT, TARGETS[target]));
  return { ...base, ...override };
}

// Collect every source file a manifest references, so we copy exactly what the
// target needs (and nothing it doesn't).
function referencedFiles(manifest) {
  const files = new Set();

  const bg = manifest.background || {};
  if (Array.isArray(bg.scripts)) bg.scripts.forEach((f) => files.add(f));
  if (bg.service_worker) files.add(bg.service_worker);

  for (const cs of manifest.content_scripts || []) {
    for (const f of cs.js || []) files.add(f);
    for (const f of cs.css || []) files.add(f);
  }

  const action = manifest.action || manifest.browser_action || {};
  if (action.default_popup) {
    const popup = action.default_popup;
    files.add(popup);
    // Pull in assets the popup HTML references (kept simple + explicit).
    const html = readFileSync(join(SRC, popup), "utf8");
    for (const m of html.matchAll(/(?:src|href)="([^"]+)"/g)) {
      const ref = m[1];
      if (/^https?:|^data:|^\/\//.test(ref)) continue;
      files.add(ref);
    }
  }

  return files;
}

function copyIcons(outDir) {
  const iconsDir = join(SRC, "icons");
  if (!existsSync(iconsDir)) return;
  mkdirSync(join(outDir, "icons"), { recursive: true });
  for (const f of readdirSync(iconsDir)) {
    if (f.startsWith("._")) continue; // skip AppleDouble sidecars
    cpSync(join(iconsDir, f), join(outDir, "icons", f));
  }
}

function build(target) {
  if (!TARGETS[target]) {
    throw new Error(`unknown target "${target}" (use firefox|chrome|all)`);
  }
  const manifest = mergeManifest(target);
  const outDir = join(DIST, target);
  // `--no-clean` keeps the output directory and only overwrites files: sandboxes
  // and permission prompts often block recursive deletes, and a dev build does
  // not need a pristine tree.
  if (!process.argv.includes("--no-clean")) {
    rmSync(outDir, { recursive: true, force: true });
  }
  mkdirSync(outDir, { recursive: true });

  const files = referencedFiles(manifest);
  for (const rel of files) {
    if (rel.startsWith("icons/")) continue; // handled in bulk below
    const from = join(SRC, rel);
    if (!existsSync(from)) throw new Error(`missing source file: ${rel}`);
    const to = join(outDir, rel);
    mkdirSync(dirname(to), { recursive: true });
    cpSync(from, to);
  }
  copyIcons(outDir);

  writeFileSync(
    join(outDir, "manifest.json"),
    JSON.stringify(manifest, null, 2) + "\n",
  );

  return { target, outDir, files: [...files].sort() };
}

function zip(target, outDir) {
  const zipPath = join(DIST, `${PKG_NAME}-${target}.zip`);
  rmSync(zipPath, { force: true });
  // -X strips extra attributes; run from inside outDir for clean root paths.
  execSync(`cd "${outDir}" && zip -qrX "${zipPath}" . -x '*.DS_Store'`, {
    stdio: "inherit",
  });
  return zipPath;
}

function main() {
  const args = process.argv.slice(2);
  const doZip = args.includes("--zip");
  const positional = args.filter((a) => !a.startsWith("--"));
  let targets = positional.length ? positional : ["all"];
  if (targets.includes("all")) targets = ["firefox", "chrome"];

  mkdirSync(DIST, { recursive: true });
  for (const t of targets) {
    const res = build(t);
    console.log(`✓ built ${res.target} → ${res.outDir}`);
    console.log(`  files: ${res.files.join(", ")}`);
    if (doZip) {
      const z = zip(t, res.outDir);
      console.log(`  zipped → ${z}`);
    }
  }
}

main();
