// Renders the Stream Deck key icons (144x144 PNG) from Tabler outline icons.
// Run from a folder where `npm i @tabler/icons@3.31.0 @resvg/resvg-js@2.6.2` is installed:
//   node C:\GSADUs\Tools\StreamDeck\icons\make-icons.mjs <that folder>
import { readFileSync, writeFileSync } from "node:fs";
import { createRequire } from "node:module";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";

const deps = process.argv[2] ?? ".";
const { Resvg } = createRequire(join(deps, "package.json"))("@resvg/resvg-js");
const tablerDir = join(deps, "node_modules", "@tabler", "icons", "icons", "outline");
const outDir = dirname(fileURLToPath(import.meta.url));

// Ink by what the key does: opens a page or folder, runs a command, stops something.
const ink = { open: "#AFA9EC", run: "#5DCAA5", stop: "#F0997B" };

// [file, kind, tabler icon, label]
const keys = [
  ["gmail", "open", "mail", "Gmail"],
  ["chat", "open", "messages", "Chat"],
  ["pm", "open", "calendar-event", "PM"],
  ["webapp", "open", "world", "WebApp"],
  ["site", "open", "home", "gsadus.com"],
  ["pipedrive", "open", "filter", "Pipedrive"],
  ["proposals", "open", "file-dollar", "Proposals"],
  ["projects", "open", "folder", "Projects"],
  ["helpdesk", "open", "lifebuoy", "Helpdesk"],
  ["board", "open", "layout-kanban", "Board"],
  ["orchestrator", "open", "sparkles", "Orchestrator"],
  ["gws-login", "run", "brand-google", "gws-login"],
  ["gcloud-login", "run", "cloud-lock", "gcloud login"],
  ["remote", "run", "device-desktop", "remote"],
  ["stop-servers", "stop", "player-stop", "Stop servers"],
];

for (const [file, kind, icon, label] of keys) {
  const src = readFileSync(join(tablerDir, `${icon}.svg`), "utf8");
  const body = src.replace(/^[\s\S]*?<svg[^>]*>/, "").replace(/<\/svg>\s*$/, "");
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="144" height="144" viewBox="0 0 144 144">
  <rect width="144" height="144" fill="#1A1A19"/>
  <g transform="translate(42 22) scale(2.5)" fill="none" stroke="${ink[kind]}" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round">${body}</g>
  <text x="72" y="122" text-anchor="middle" font-family="Segoe UI" font-weight="600" font-size="19" fill="#F1EFE8">${label}</text>
</svg>`;
  const png = new Resvg(svg, { font: { loadSystemFonts: true, defaultFontFamily: "Segoe UI" } }).render().asPng();
  writeFileSync(join(outDir, `${file}.png`), png);
}
console.log(`${keys.length} icons written to ${outDir}`);
