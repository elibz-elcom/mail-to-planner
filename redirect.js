// Landing page for three flows:
//  1. Admin consent (…/adminconsent?…&redirect_uri=…/redirect.html) — Entra returns
//     ?admin_consent=True&tenant=… or ?error=… with no MSAL state. Show a clear result.
//  2. Office dialog sign-in (Outlook builds without Nested App Auth, e.g. Office 2021,
//     where MSAL popups are blocked). The task pane opens redirect.html?m2p=dialog with
//     Office.context.ui.displayDialogAsync; this page runs an MSAL redirect sign-in inside
//     the dialog and returns the token with Office.context.ui.messageParent.
//  3. MSAL popup sign-in — hand the response back to the task pane.
import { broadcastResponseToMainFrame } from "@azure/msal-browser/redirect-bridge";
import { createStandardPublicClientApplication } from "@azure/msal-browser";

const CONFIG = __CONFIG__;
const SCOPES = ["User.Read", "User.ReadBasic.All", "Tasks.ReadWrite", "GroupMember.Read.All", "Mail.ReadWrite"];
const DIALOG_FLAG = "m2p-dialog";

const $ = (id) => document.getElementById(id);

function render(kind, title, msg, tenant) {
  $("card").className = "card " + kind;
  $("icon").textContent = kind === "ok" ? "✔" : kind === "err" ? "✖" : "⏳";
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

function ssGet(k) {
  try { return sessionStorage.getItem(k); } catch { return null; }
}
function ssSet(k, v) {
  try { v == null ? sessionStorage.removeItem(k) : sessionStorage.setItem(k, v); } catch { /* ignore */ }
}

function loadOffice() {
  return new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = "https://appsforoffice.microsoft.com/lib/1/hosted/office.js";
    s.onload = () => Office.onReady(() => resolve());
    s.onerror = () => reject(new Error("office.js failed to load"));
    document.head.appendChild(s);
  });
}

async function runDialog() {
  render("wait", "מתחבר…", "חלון ההתחברות של Microsoft ייפתח כאן.");
  const officeReady = loadOffice();
  const pca = await createStandardPublicClientApplication({
    auth: {
      clientId: CONFIG.clientId,
      authority: "https://login.microsoftonline.com/organizations",
      redirectUri: CONFIG.baseUrl + "/redirect.html",
    },
    cache: { cacheLocation: "localStorage" },
  });

  const send = async (payload) => {
    ssSet(DIALOG_FLAG, null);
    await officeReady;
    Office.context.ui.messageParent(JSON.stringify(payload));
  };
  const ok = (r) =>
    send({
      accessToken: r.accessToken,
      expiresOn: r.expiresOn ? new Date(r.expiresOn).getTime() : Date.now() + 50 * 60 * 1000,
      tenantId: r.tenantId || (r.account && r.account.tenantId) || "",
      username: (r.account && r.account.username) || "",
    });

  try {
    const back = await pca.handleRedirectPromise();
    if (back && back.accessToken) {
      render("ok", "מחובר", "חוזר לתוסף…");
      return ok(back);
    }
    if (isMsal) {
      // Came back from Microsoft without a token (e.g. cancelled).
      const err = get("error_description") || get("error") || "ההתחברות לא הושלמה";
      return send({ error: String(err).replace(/\+/g, " ").split(/\r?\n/)[0] });
    }
    // Starting: try the cached account first, otherwise go to the Microsoft sign-in page.
    const account = pca.getAllAccounts()[0];
    if (account) {
      try {
        return ok(await pca.acquireTokenSilent({ scopes: SCOPES, account }));
      } catch {
        /* fall through to interactive */
      }
    }
    ssSet(DIALOG_FLAG, "1");
    await pca.acquireTokenRedirect({ scopes: SCOPES, prompt: account ? undefined : "select_account" });
  } catch (e) {
    render("err", "שגיאת התחברות", e && e.message ? e.message : String(e));
    try {
      await send({ error: e && e.message ? e.message : String(e) });
    } catch {
      /* not inside an Office dialog */
    }
  }
}

if (params.get("m2p") === "dialog" || (isMsal && ssGet(DIALOG_FLAG))) {
  runDialog();
} else if (!isMsal && get("admin_consent")) {
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
