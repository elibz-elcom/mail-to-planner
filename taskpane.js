/* Elcom – Send to Planner (Outlook add-in task pane)
 * Turns the open email into a Planner task in a plan/bucket the user picks.
 * Works in Outlook desktop (new + classic), web, iOS and Android.
 */
import {
  createNestablePublicClientApplication,
  createStandardPublicClientApplication,
} from "@azure/msal-browser";

/* global Office, __CONFIG__ */
const CONFIG = __CONFIG__;
const GRAPH = "https://graph.microsoft.com/v1.0";
const SCOPES = ["User.Read", "Tasks.ReadWrite", "GroupMember.Read.All", "Mail.Read"];
const MAX_DESC = 25000;
const PLAN_CACHE_KEY = "m2p.plans.v1";

const $ = (id) => document.getElementById(id);
const state = {
  pca: null,
  account: null,
  tenantId: null,
  me: null,
  plans: [],          // [{id, title, groupId, groupName}]
  favorites: [],      // [{id, title, groupId}]
  lastBucket: {},     // planId -> bucketId
  mail: null,         // {subject, fromName, fromAddr, received, body, restId}
};

/* ---------------- UI helpers ---------------- */
function show(id) {
  ["status", "signin", "taskForm", "done", "error"].forEach((s) =>
    $(s).classList.toggle("hidden", s !== id)
  );
}
function setStatus(text) {
  $("statusText").textContent = text;
  show("status");
}
function showError(err) {
  console.error(err);
  let msg = typeof err === "string" ? err : err?.message || "שגיאה לא צפויה";
  if (/AADSTS65001|consent/i.test(msg)) {
    msg = "האפליקציה עדיין לא אושרה בארגון. מנהל המערכת צריך לתת Admin consent (ראה מדריך ההתקנה).";
  } else if (/403/.test(msg)) {
    msg = "אין הרשאה לפעולה הזו. ייתכן שאינך חבר בקבוצה של התוכנית.";
  }
  $("errorText").textContent = msg;
  show("error");
}
function option(value, text) {
  const o = document.createElement("option");
  o.value = value;
  o.textContent = text;
  return o;
}

/* ---------------- Roaming settings (sync PC <-> phone) ---------------- */
function rsGet(key, fallback) {
  try {
    const v = Office.context.roamingSettings.get(key);
    return v === undefined || v === null ? fallback : v;
  } catch {
    return fallback;
  }
}
function rsSet(values) {
  try {
    Object.entries(values).forEach(([k, v]) => Office.context.roamingSettings.set(k, v));
    Office.context.roamingSettings.saveAsync(() => {});
  } catch {
    /* not critical */
  }
}

/* ---------------- Auth ---------------- */
async function initAuth() {
  const naa =
    Office.context.requirements &&
    Office.context.requirements.isSetSupported("NestedAppAuth", "1.1");
  const auth = {
    clientId: CONFIG.clientId,
    authority: "https://login.microsoftonline.com/organizations",
    redirectUri: CONFIG.baseUrl + "/redirect.html",
  };
  state.pca = naa
    ? await createNestablePublicClientApplication({ auth })
    : await createStandardPublicClientApplication({
        auth,
        cache: { cacheLocation: "localStorage" },
      });
  state.naa = !!naa;
}

async function getToken(interactive = false) {
  const accounts = state.pca.getAllAccounts();
  const account = state.account || accounts[0];
  const req = { scopes: SCOPES, account };
  if (!interactive) {
    try {
      const r = await state.pca.acquireTokenSilent(req);
      return remember(r);
    } catch (e) {
      // NAA can usually complete interactively without a user gesture.
      if (!state.naa) throw Object.assign(new Error("interaction_required"), { needsSignIn: true });
    }
  }
  const r = await state.pca.acquireTokenPopup({ scopes: SCOPES });
  return remember(r);
}
function remember(r) {
  state.account = r.account;
  state.tenantId = r.tenantId || r.account?.tenantId || state.tenantId;
  return r.accessToken;
}

/* ---------------- Graph ---------------- */
async function graph(path, { method = "GET", body, headers = {}, raw = false } = {}, attempt = 0) {
  const token = await getToken();
  const url = path.startsWith("http") ? path : GRAPH + path;
  const res = await fetch(url, {
    method,
    headers: {
      Authorization: "Bearer " + token,
      ...(body ? { "Content-Type": "application/json" } : {}),
      ...headers,
    },
    body: body ? JSON.stringify(body) : undefined,
  });
  if ((res.status === 429 || res.status === 503) && attempt < 3) {
    const wait = Number(res.headers.get("Retry-After") || 2) * 1000;
    await new Promise((r) => setTimeout(r, wait));
    return graph(path, { method, body, headers, raw }, attempt + 1);
  }
  if (raw) return res;
  if (!res.ok) {
    let detail = "";
    try {
      detail = (await res.json())?.error?.message || "";
    } catch {}
    throw new Error(`Graph ${res.status}: ${detail}`);
  }
  if (res.status === 204) return null;
  return res.json();
}

async function graphAll(path) {
  let out = [];
  let next = path;
  while (next) {
    const page = await graph(next);
    out = out.concat(page.value || []);
    next = page["@odata.nextLink"] || null;
  }
  return out;
}

/* ---------------- Plans ---------------- */
async function fetchPlans() {
  const groups = (
    await graphAll("/me/memberOf/microsoft.graph.group?$select=id,displayName,groupTypes&$top=999")
  ).filter((g) => (g.groupTypes || []).includes("Unified"));

  const plans = [];
  for (let i = 0; i < groups.length; i += 20) {
    const chunk = groups.slice(i, i + 20);
    const batch = await graph("/$batch", {
      method: "POST",
      body: {
        requests: chunk.map((g, idx) => ({
          id: String(idx),
          method: "GET",
          url: `/groups/${g.id}/planner/plans?$select=id,title,owner`,
        })),
      },
    });
    for (const r of batch.responses || []) {
      if (r.status !== 200) continue;
      const g = chunk[Number(r.id)];
      for (const p of r.body?.value || []) {
        plans.push({ id: p.id, title: p.title, groupId: g.id, groupName: g.displayName });
      }
    }
  }
  plans.sort((a, b) => a.title.localeCompare(b.title, "he"));
  return plans;
}

function readPlanCache() {
  try {
    return JSON.parse(localStorage.getItem(PLAN_CACHE_KEY) || "null");
  } catch {
    return null;
  }
}
function writePlanCache(plans) {
  try {
    localStorage.setItem(PLAN_CACHE_KEY, JSON.stringify(plans));
  } catch {}
}

function renderPlans(selectedId) {
  const sel = $("planSelect");
  sel.innerHTML = "";
  if (!state.plans.length) {
    sel.appendChild(option("", "לא נמצאו תוכניות"));
    return;
  }
  for (const p of state.plans) {
    const label = p.groupName && p.groupName !== p.title ? `${p.title} · ${p.groupName}` : p.title;
    sel.appendChild(option(p.id, label));
  }
  const pick =
    (selectedId && state.plans.find((p) => p.id === selectedId)) ||
    state.plans.find((p) => p.id === rsGet("lastPlanId")) ||
    state.plans[0];
  sel.value = pick.id;
}

function renderFavorites() {
  const wrap = $("favChips");
  wrap.innerHTML = "";
  const favs = state.favorites.filter((f) => state.plans.length === 0 || state.plans.some((p) => p.id === f.id));
  $("favRow").classList.toggle("hidden", favs.length === 0);
  for (const f of favs) {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "chip" + ($("planSelect").value === f.id ? " active" : "");
    b.textContent = f.title;
    b.addEventListener("click", () => {
      $("planSelect").value = f.id;
      onPlanChange();
    });
    wrap.appendChild(b);
  }
  const isFav = state.favorites.some((f) => f.id === $("planSelect").value);
  $("favBtn").textContent = isFav ? "★" : "☆";
  $("favBtn").classList.toggle("on", isFav);
}

function toggleFavorite() {
  const plan = currentPlan();
  if (!plan) return;
  const exists = state.favorites.some((f) => f.id === plan.id);
  state.favorites = exists
    ? state.favorites.filter((f) => f.id !== plan.id)
    : [...state.favorites, { id: plan.id, title: plan.title, groupId: plan.groupId }];
  rsSet({ favorites: state.favorites });
  renderFavorites();
}

function currentPlan() {
  return state.plans.find((p) => p.id === $("planSelect").value);
}

let planChangeSeq = 0;
async function onPlanChange() {
  const plan = currentPlan();
  renderFavorites();
  if (!plan) return;
  const seq = ++planChangeSeq;
  const bSel = $("bucketSelect");
  const aSel = $("assigneeSelect");
  bSel.innerHTML = "";
  bSel.appendChild(option("", "טוען…"));
  bSel.disabled = true;

  try {
    const [buckets, members] = await Promise.all([
      graphAll(`/planner/plans/${plan.id}/buckets?$select=id,name,orderHint`),
      graphAll(`/groups/${plan.groupId}/members/microsoft.graph.user?$select=id,displayName&$top=200`).catch(() => []),
    ]);
    if (seq !== planChangeSeq) return;

    buckets.sort((a, b) => (a.orderHint < b.orderHint ? 1 : -1));
    bSel.innerHTML = "";
    if (!buckets.length) bSel.appendChild(option("", "ללא Bucket"));
    buckets.forEach((b) => bSel.appendChild(option(b.id, b.name)));
    const lastB = state.lastBucket[plan.id];
    if (lastB && buckets.some((b) => b.id === lastB)) bSel.value = lastB;
    bSel.disabled = false;

    const prevAssignee = aSel.value || state.me.id;
    aSel.innerHTML = "";
    aSel.appendChild(option(state.me.id, "אני"));
    aSel.appendChild(option("", "ללא שיוך"));
    members
      .filter((m) => m.id !== state.me.id)
      .sort((a, b) => (a.displayName || "").localeCompare(b.displayName || "", "he"))
      .forEach((m) => aSel.appendChild(option(m.id, m.displayName)));
    aSel.value = [...aSel.options].some((o) => o.value === prevAssignee) ? prevAssignee : state.me.id;
  } catch (e) {
    if (seq === planChangeSeq) showError(e);
  }
}

/* ---------------- Mail ---------------- */
function loadMail() {
  const item = Office.context.mailbox.item;
  const from = item.from || item.sender || {};
  const mail = {
    subject: item.subject || "(ללא נושא)",
    fromName: from.displayName || "",
    fromAddr: from.emailAddress || "",
    received: item.dateTimeCreated ? new Date(item.dateTimeCreated) : null,
    body: "",
    restId: null,
  };
  try {
    const host = Office.context.mailbox.diagnostics.hostName || "";
    if (/iOS|Android/i.test(host) || !Office.context.mailbox.convertToRestId) {
      mail.restId = item.itemId;
    } else {
      mail.restId = Office.context.mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0);
    }
  } catch {
    mail.restId = item.itemId;
  }
  return new Promise((resolve) => {
    item.body.getAsync(Office.CoercionType.Text, (r) => {
      if (r.status === Office.AsyncResultStatus.Succeeded) mail.body = (r.value || "").trim();
      resolve(mail);
    });
  });
}

async function getWebLink(restId) {
  if (!restId) return null;
  try {
    const m = await graph(`/me/messages/${encodeURIComponent(restId)}?$select=webLink`);
    return m.webLink || null;
  } catch {
    return null;
  }
}

/* ---------------- Create task ---------------- */
function encodeRefKey(url) {
  return url.replace(/%/g, "%25").replace(/\./g, "%2E").replace(/:/g, "%3A").replace(/@/g, "%40").replace(/#/g, "%23");
}

function buildDescription(mail, includeBody) {
  const lines = [];
  lines.push(`מאת: ${mail.fromName}${mail.fromAddr ? ` <${mail.fromAddr}>` : ""}`);
  if (mail.received) {
    lines.push(
      "התקבל: " +
        mail.received.toLocaleString("he-IL", { dateStyle: "short", timeStyle: "short" })
    );
  }
  lines.push(`נושא: ${mail.subject}`);
  if (includeBody && mail.body) {
    lines.push("", "────────────", "");
    let body = mail.body.replace(/\r\n/g, "\n").replace(/\n{3,}/g, "\n\n");
    const room = MAX_DESC - lines.join("\n").length - 40;
    if (body.length > room) body = body.slice(0, room) + "\n…(קוצר)";
    lines.push(body);
  }
  return lines.join("\n");
}

async function getDetailsEtag(taskId) {
  for (let i = 0; i < 6; i++) {
    const res = await graph(`/planner/tasks/${taskId}/details`, { raw: true });
    if (res.ok) {
      const d = await res.json();
      return d["@odata.etag"];
    }
    await new Promise((r) => setTimeout(r, 700 * (i + 1)));
  }
  throw new Error("לא הצלחתי לעדכן את פרטי המשימה (המשימה עצמה נוצרה).");
}

async function createTask(ev) {
  ev.preventDefault();
  const plan = currentPlan();
  const title = $("titleInput").value.trim();
  if (!plan) return showError("בחר תוכנית.");
  if (!title) {
    $("titleInput").focus();
    return;
  }
  const bucketId = $("bucketSelect").value || undefined;
  const assignee = $("assigneeSelect").value;
  const due = $("dueInput").value;

  const btn = $("createBtn");
  btn.disabled = true;
  btn.textContent = "יוצר משימה…";
  try {
    const payload = { planId: plan.id, title: title.slice(0, 255) };
    if (bucketId) payload.bucketId = bucketId;
    if (due) payload.dueDateTime = new Date(due + "T17:00:00").toISOString();
    if (assignee) {
      payload.assignments = {
        [assignee]: { "@odata.type": "#microsoft.graph.plannerAssignment", orderHint: " !" },
      };
    }
    const task = await graph("/planner/tasks", { method: "POST", body: payload });

    const [etag, webLink] = await Promise.all([
      getDetailsEtag(task.id),
      getWebLink(state.mail.restId),
    ]);
    const details = {
      description: buildDescription(state.mail, $("includeBody").checked),
      previewType: "description",
    };
    if (webLink) {
      details.references = {
        [encodeRefKey(webLink)]: {
          "@odata.type": "microsoft.graph.plannerExternalReference",
          alias: "📧 פתח את המייל המקורי",
          type: "Other",
          previewPriority: " !",
        },
      };
    }
    await graph(`/planner/tasks/${task.id}/details`, {
      method: "PATCH",
      body: details,
      headers: { "If-Match": etag, Prefer: "return=minimal" },
    });

    state.lastBucket[plan.id] = bucketId || "";
    rsSet({ lastPlanId: plan.id, lastBucket: state.lastBucket });

    $("doneDetails").textContent = `${plan.title}${
      $("bucketSelect").selectedOptions[0] && bucketId ? " · " + $("bucketSelect").selectedOptions[0].textContent : ""
    }`;
    $("openTaskLink").href = `https://tasks.office.com/${state.tenantId}/Home/Task/${task.id}`;
    show("done");
  } catch (e) {
    showError(e);
  } finally {
    btn.disabled = false;
    btn.textContent = "צור משימה";
  }
}

/* ---------------- Boot ---------------- */
async function loadEverything() {
  setStatus("מתחבר…");
  await getToken();
  state.me = state.me || (await graph("/me?$select=id,displayName"));

  setStatus("קורא את המייל…");
  state.mail = await loadMail();
  $("titleInput").value = state.mail.subject;

  const cached = readPlanCache();
  if (cached && cached.length) {
    state.plans = cached;
    renderPlans();
    show("taskForm");
    onPlanChange();
    // Refresh plan list quietly in the background.
    fetchPlans()
      .then((fresh) => {
        const keep = $("planSelect").value;
        state.plans = fresh;
        writePlanCache(fresh);
        renderPlans(keep);
        renderFavorites();
      })
      .catch(() => {});
  } else {
    setStatus("טוען את התוכניות שלך…");
    state.plans = await fetchPlans();
    writePlanCache(state.plans);
    renderPlans();
    show("taskForm");
    onPlanChange();
  }
}

async function start() {
  try {
    await loadEverything();
  } catch (e) {
    if (e && e.needsSignIn) show("signin");
    else showError(e);
  }
}

Office.onReady(async () => {
  state.favorites = rsGet("favorites", []);
  state.lastBucket = rsGet("lastBucket", {});

  $("planSelect").addEventListener("change", onPlanChange);
  $("favBtn").addEventListener("click", toggleFavorite);
  $("taskForm").addEventListener("submit", createTask);
  $("retryBtn").addEventListener("click", start);
  $("againBtn").addEventListener("click", () => show("taskForm"));
  $("signinBtn").addEventListener("click", async () => {
    try {
      await getToken(true);
      await start();
    } catch (e) {
      showError(e);
    }
  });

  // Pinned task pane: refresh when the user selects another email.
  try {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, async () => {
      if (!Office.context.mailbox.item) return;
      state.mail = await loadMail();
      $("titleInput").value = state.mail.subject;
      show("taskForm");
    });
  } catch {}

  try {
    await initAuth();
  } catch (e) {
    return showError(e);
  }
  start();
});
