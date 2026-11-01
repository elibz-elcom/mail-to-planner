// Landing page for two flows:
//  1. Admin consent (…/adminconsent?…&redirect_uri=…/redirect.html) — Entra returns
//     ?admin_consent=True&tenant=… or ?error=… with no MSAL state. Show a clear result.
//  2. Fallback MSAL popup sign-in (Outlook clients without Nested App Auth) — hand the
//     response back to the task pane.
import { broadcastResponseToMainFrame } from "@azure/msal-browser/redirect-bridge";

const $ = (id) => document.getElementById(id);

function render(kind, title, msg, tenant) {
  $("card").className = "card " + kind;
  $("icon").textContent = kind === "ok" ? "✔" : "✖";
  $("title").textContent = title;
  $("msg").textContent = msg || "";
  if (tenant) {
    $("tenant").textContent = "Tenant: " + tenant;
    $("tenant").classList.remove("hidden");
  }
  document.title = title;
}

const params = new URLSearchParams(window.location.search);
const hash = new URLSearchParams(window.location.hash.replace(/^#/, ""));
const get = (k) => params.get(k) || hash.get(k);
const isMsal = !!get("state"); // MSAL always sends state; admin consent links don't.

if (!isMsal && get("admin_consent")) {
  const ok = String(get("admin_consent")).toLowerCase() === "true";
  if (ok) {
    render(
      "ok",
      "האישור נשמר בהצלחה",
      "התוסף \"שלח ל-Planner\" אושר בארגון. אפשר לסגור את החלון ולהמשיך להפצת התוסף (Integrated Apps).",
      get("tenant")
    );
  } else {
    render("err", "האישור לא הושלם", "נסה שוב עם חשבון Global Admin או Cloud Application Administrator.", get("tenant"));
  }
} else if (!isMsal && get("error")) {
  const desc = (get("error_description") || get("error") || "").replace(/\+/g, " ");
  const cancelled = get("error") === "access_denied";
  render(
    "err",
    cancelled ? "האישור בוטל" : "האישור נכשל",
    cancelled ? "לא ניתן אישור לתוסף. אפשר לפתוח את הקישור שוב ולאשר." : desc.split(/\r?\n/)[0],
    get("tenant")
  );
} else {
  broadcastResponseToMainFrame().catch((e) => {
    render("err", "שגיאת התחברות", e && e.message ? e.message : String(e));
  });
}
