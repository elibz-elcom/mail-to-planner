// Build: bundles the task pane, copies static files and writes manifest.xml into dist/.
// Values come from config.json; env vars CLIENT_ID / BASE_URL override (used by GitHub Actions).
import { build } from "esbuild";
import { readFileSync, writeFileSync, mkdirSync, cpSync, rmSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const root = dirname(fileURLToPath(import.meta.url));
const cfg = JSON.parse(readFileSync(join(root, "config.json"), "utf8"));

const clientId = process.env.CLIENT_ID || cfg.clientId;
const baseUrl = (
  process.env.BASE_URL ||
  cfg.customBaseUrl ||
  `https://${cfg.githubUser.toLowerCase()}.github.io/${cfg.repoName}`
).replace(/\/+$/, "");
const origin = new URL(baseUrl).origin;

const problems = [];
if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(clientId))
  problems.push("clientId in config.json is not a valid GUID (Application (client) ID from Entra)");
if (/PUT-YOUR/.test(baseUrl)) problems.push("githubUser in config.json is not set");
if (problems.length && !process.env.ALLOW_PLACEHOLDERS) {
  console.error("\n✖ Config problems:\n  - " + problems.join("\n  - ") + "\n");
  process.exit(1);
}

const dist = join(root, "dist");
rmSync(dist, { recursive: true, force: true });
mkdirSync(dist, { recursive: true });

await build({
  entryPoints: { taskpane: join(root, "taskpane.js"), redirect: join(root, "redirect.js") },
  outdir: dist,
  bundle: true,
  minify: true,
  format: "iife",
  target: ["es2020"],
  define: { __CONFIG__: JSON.stringify({ clientId, baseUrl }) },
  logLevel: "warning",
});

for (const f of [
  "taskpane.html", "taskpane.css", "commands.html", "redirect.html",
  "elcom-logo.png", "elcom-logo-dark.png",
])
  cpSync(join(root, f), join(dist, f));
cpSync(join(root, "assets"), join(dist, "assets"), { recursive: true });
writeFileSync(join(dist, ".nojekyll"), "");
writeFileSync(
  join(dist, "index.html"),
  `<!doctype html><meta charset="utf-8"><title>Send to Planner</title><p>Elcom IT – Send to Planner Outlook add-in. <a href="manifest.xml">manifest.xml</a></p>`
);

const manifest = readFileSync(join(root, "manifest.template.xml"), "utf8")
  .replaceAll("{{ADDIN_ID}}", cfg.addinId)
  .replaceAll("{{VERSION}}", cfg.version)
  .replaceAll("{{PROVIDER}}", cfg.providerName)
  .replaceAll("{{SUPPORT}}", cfg.supportUrl)
  .replaceAll("{{BASE}}", baseUrl)
  .replaceAll("{{ORIGIN}}", origin);
writeFileSync(join(dist, "manifest.xml"), manifest);

console.log(`✔ Built for ${baseUrl}`);
console.log(`  NAA redirect URI for Entra (SPA): brk-multihub://${new URL(baseUrl).host}`);
console.log(`  Fallback redirect URI for Entra (SPA): ${baseUrl}/redirect.html`);
