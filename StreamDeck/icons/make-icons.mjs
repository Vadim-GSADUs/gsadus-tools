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

const ink = { dev: "#5DCAA5", revit: "#F0997B", office: "#AFA9EC", nav: "#B4B2A9" };

// [file, page, tabler icon, label]
const keys = [
  ["dev-webapp", "dev", "world", "WebApp 3009"],
  ["dev-site", "dev", "home", "Site 3019"],
  ["dev-pm", "dev", "calendar-event", "PM 3200"],
  ["dev-claude", "dev", "sparkles", "Claude"],
  ["dev-docker", "dev", "brand-docker", "Docker"],
  ["dev-unwip", "dev", "cloud-download", "unwip-all"],
  ["dev-wip", "dev", "cloud-upload", "wip-all"],
  ["dev-vercel", "dev", "triangle", "Vercel"],
  ["dev-supabase", "dev", "database", "Supabase"],
  ["revit-open", "revit", "building", "Revit 2026"],
  ["revit-vg", "revit", "eye", "VG"],
  ["revit-thin", "revit", "line-dashed", "Thin lines"],
  ["revit-hide", "revit", "eye-off", "Hide EH"],
  ["revit-isolate", "revit", "focus-2", "Isolate HI"],
  ["revit-reset", "revit", "refresh", "Reset HR"],
  ["revit-3d", "revit", "box", "Default 3D"],
  ["revit-tile", "revit", "layout-grid", "Tile WT"],
  ["revit-pyrevit", "revit", "reload", "pyRevit RL"],
  ["office-mute", "office", "microphone-off", "Meet mute"],
  ["office-camera", "office", "video-off", "Meet camera"],
  ["office-meet", "office", "video", "Open Meet"],
  ["office-pipedrive", "office", "filter", "Pipedrive"],
  ["office-pm", "office", "calendar-event", "PM live"],
  ["office-projects", "office", "folder", "Projects"],
  ["office-business", "office", "folders", "Business"],
  ["nav-dev", "nav", "code", "Dev"],
  ["nav-revit", "nav", "building", "Revit"],
  ["nav-office", "nav", "briefcase", "Office"],
];

for (const [file, page, icon, label] of keys) {
  const src = readFileSync(join(tablerDir, `${icon}.svg`), "utf8");
  const body = src.replace(/^[\s\S]*?<svg[^>]*>/, "").replace(/<\/svg>\s*$/, "");
  const bg = page === "nav" ? "#2C2C2A" : "#1A1A19";
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="144" height="144" viewBox="0 0 144 144">
  <rect width="144" height="144" fill="${bg}"/>
  <g transform="translate(42 22) scale(2.5)" fill="none" stroke="${ink[page]}" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round">${body}</g>
  <text x="72" y="122" text-anchor="middle" font-family="Segoe UI" font-weight="600" font-size="19" fill="#F1EFE8">${label}${page === "nav" ? " ›" : ""}</text>
</svg>`;
  const png = new Resvg(svg, { font: { loadSystemFonts: true, defaultFontFamily: "Segoe UI" } }).render().asPng();
  writeFileSync(join(outDir, `${file}.png`), png);
}
console.log(`${keys.length} icons written to ${outDir}`);
