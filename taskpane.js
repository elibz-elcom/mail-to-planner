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
const SCOPES = ["User.Read", "User.ReadBasic.All", "Tasks.ReadWrite", "GroupMember.Read.All", "Mail.ReadWrite"];
const MAX_DESC = 25000;
const PLAN_CACHE_KEY = "m2p.plans.v2";
const MAIL_CATEGORY = "משימה ב-Planner";
const NEW_BUCKET = "__new__";
const LABEL_COLORS = [
  "", "#e8478b", "#d13438", "#f2c80f", "#13a10e", "#0f6cbd", "#8764b8", "#a87b45", "#8cbd18",
  "#00b7c3", "#69797e", "#a0aeb2", "#8e562e", "#c50f1f", "#f7630c", "#ff8c00", "#eaa300",
  "#7fba00", "#0b6a0b", "#038387", "#4f9fe0", "#003966", "#b4a0ff", "#5c2e91", "#c8c6c4", "#3b3a39",
];

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
  members: new Set(), // user ids of the current plan's group
  catDesc: {},        // plan label names: category1..25 -> name
  selectedLabels: new Set(),
  mailInfo: null,     // {webLink, categories}
  mailInfoPromise: null,
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

async function getToken(interactive = false, scopes = SCOPES) {
  const account = state.account || state.pca.getAllAccounts()[0];
  if (!interactive) {
    try {
      return remember(await state.pca.acquireTokenSilent({ scopes, account }));
    } catch (e) {
      // NAA can usually complete interactively without a user gesture.
      if (!state.naa) throw Object.assign(new Error("interaction_required"), { needsSignIn: true });
    }
  }
  return remember(await state.pca.acquireTokenPopup({ scopes }));
}
function remember(r) {
  state.account = r.account;
  state.tenantId = r.tenantId || r.account?.tenantId || state.tenantId;
  return r.accessToken;
}

/* ---------------- Graph ---------------- */
async function graph(path, opts = {}, attempt = 0) {
  const { method = "GET", body, headers = {}, raw = false, scopes = SCOPES } = opts;
  const token = await getToken(false, scopes);
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
    return graph(path, opts, attempt + 1);
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
  const text = await res.text();
  return text ? JSON.parse(text) : null;
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
function planGroupId(p) {
  if (p.container && p.container.type === "group") return p.container.containerId;
  return p.owner || null;
}

async function fetchPlans() {
  const byId = new Map();
  const add = (p, groupName) => {
    if (!p || byId.has(p.id)) return;
    byId.set(p.id, { id: p.id, title: p.title, groupId: planGroupId(p), groupName: groupName || "" });
  };

  // 1. Groups I'm a member of + groups I own (owners aren't always members).
  const [memberOf, owned] = await Promise.all([
    graphAll("/me/memberOf/microsoft.graph.group?$select=id,displayName,groupTypes&$top=999").catch(() => []),
    graphAll("/me/ownedObjects/microsoft.graph.group?$select=id,displayName,groupTypes&$top=999").catch(() => []),
  ]);
  const groups = new Map();
  [...memberOf, ...owned]
    .filter((g) => (g.groupTypes || []).includes("Unified"))
    .forEach((g) => groups.set(g.id, g));
  const list = [...groups.values()];

  for (let i = 0; i < list.length; i += 20) {
    const chunk = list.slice(i, i + 20);
    const batch = await graph("/$batch", {
      method: "POST",
      body: {
        requests: chunk.map((g, idx) => ({
          id: String(idx),
          method: "GET",
          url: `/groups/${g.id}/planner/plans?$select=id,title,owner,container`,
        })),
      },
    });
    for (const r of batch.responses || []) {
      if (r.status !== 200) continue;
      const g = chunk[Number(r.id)];
      for (const p of r.body?.value || []) add(p, g.displayName);
    }
  }

  // 2. Plans shared with me directly (incl. ones outside my group list).
  const shared = await graphAll("/me/planner/plans?$select=id,title,owner,container").catch(() => []);
  for (const p of shared) {
    const gid = planGroupId(p);
    add(p, gid && groups.get(gid) ? groups.get(gid).displayName : "");
  }

  const plans = [...byId.values()];
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

function planLabel(p) {
  return p.groupName && p.groupName !== p.title ? `${p.title} · ${p.groupName}` : p.title;
}

function renderPlans(selectedId) {
  const sel = $("planSelect");
  const keep = selectedId || sel.value;
  const q = ($("planFilter").value || "").trim().toLowerCase();
  sel.innerHTML = "";
  $("planCount").textContent = state.plans.length ? `(${state.plans.length})` : "";
  if (!state.plans.length) {
    sel.appendChild(option("", "לא נמצאו תוכניות"));
    return;
  }
  const shown = q ? state.plans.filter((p) => planLabel(p).toLowerCase().includes(q)) : state.plans;
  if (!shown.length) {
    sel.appendChild(option("", "אין תוכנית שמתאימה לחיפוש"));
    return;
  }
  for (const p of shown) sel.appendChild(option(p.id, planLabel(p)));
  const pick =
    (keep && shown.find((p) => p.id === keep)) ||
    shown.find((p) => p.id === rsGet("lastPlanId")) ||
    shown[0];
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
      $("planFilter").value = "";
      renderPlans(f.id);
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
  $("labelChips").innerHTML = '<span class="muted small">טוען…</span>';

  try {
    const [buckets, members, details] = await Promise.all([
      graphAll(`/planner/plans/${plan.id}/buckets?$select=id,name,orderHint`),
      plan.groupId
        ? graphAll(`/groups/${plan.groupId}/members/microsoft.graph.user?$select=id,displayName&$top=200`).catch(() => [])
        : Promise.resolve([]),
      graph(`/planner/plans/${plan.id}/details?$select=categoryDescriptions`).catch(() => null),
    ]);
    if (seq !== planChangeSeq) return;

    buckets.sort((a, b) => (a.orderHint < b.orderHint ? 1 : -1));
    bSel.innerHTML = "";
    if (!buckets.length) bSel.appendChild(option("", "ללא מיכל"));
    buckets.forEach((b) => bSel.appendChild(option(b.id, b.name)));
    bSel.appendChild(option(NEW_BUCKET, "➕ מיכל חדש…"));
    const lastB = state.lastBucket[plan.id];
    if (lastB && buckets.some((b) => b.id === lastB)) bSel.value = lastB;
    bSel.disabled = false;
    onBucketChange();

    state.catDesc = (details && details.categoryDescriptions) || {};
    state.selectedLabels = new Set();
    $("labelErr").classList.add("hidden");
    renderLabels();

    state.members = new Set(members.map((m) => m.id));
    const prevAssignee = aSel.value || state.me.id;
    const prevLabel = aSel.selectedOptions[0] ? aSel.selectedOptions[0].textContent : "";
    aSel.innerHTML = "";
    aSel.appendChild(option(state.me.id, "אני"));
    aSel.appendChild(option("", "ללא שיוך"));
    members
      .filter((m) => m.id !== state.me.id)
      .sort((a, b) => (a.displayName || "").localeCompare(b.displayName || "", "he"))
      .forEach((m) => aSel.appendChild(option(m.id, m.displayName)));
    if (prevAssignee && ![...aSel.options].some((o) => o.value === prevAssignee) && prevLabel) {
      aSel.appendChild(option(prevAssignee, prevLabel));
    }
    aSel.value = prevAssignee;
    updateAssigneeWarning();
  } catch (e) {
    if (seq === planChangeSeq) showError(e);
  }
}

/* ---------------- Buckets ---------------- */
function onBucketChange() {
  const isNew = $("bucketSelect").value === NEW_BUCKET;
  $("newBucketInput").classList.toggle("hidden", !isNew);
  if (isNew) $("newBucketInput").focus();
}

/* ---------------- Labels (Planner categories) ---------------- */
function labelEntries() {
  return Object.entries(state.catDesc)
    .filter(([k, v]) => /^category\d+$/.test(k) && v)
    .map(([k, v]) => ({ key: k, name: v, n: Number(k.slice(8)) }))
    .sort((a, b) => a.n - b.n);
}

function renderLabels() {
  const wrap = $("labelChips");
  wrap.innerHTML = "";
  const list = labelEntries();
  if (!list.length) {
    wrap.innerHTML = '<span class="muted small">אין עדיין תוויות בתוכנית הזו</span>';
    return;
  }
  for (const l of list) {
    const on = state.selectedLabels.has(l.key);
    const b = document.createElement("button");
    b.type = "button";
    b.className = "chip label-chip" + (on ? " on" : "");
    b.setAttribute("aria-pressed", on ? "true" : "false");
    const dot = document.createElement("span");
    dot.className = "dot";
    dot.style.background = LABEL_COLORS[l.n] || "#888";
    b.append(dot, document.createTextNode(l.name));
    b.addEventListener("click", () => {
      if (state.selectedLabels.has(l.key)) state.selectedLabels.delete(l.key);
      else state.selectedLabels.add(l.key);
      renderLabels();
    });
    wrap.appendChild(b);
  }
}

async function addLabel() {
  const input = $("newLabelInput");
  const name = input.value.trim();
  const plan = currentPlan();
  const err = $("labelErr");
  err.classList.add("hidden");
  if (!name || !plan) return;
  const existing = labelEntries().find((l) => l.name === name);
  if (existing) {
    state.selectedLabels.add(existing.key);
    input.value = "";
    renderLabels();
    return;
  }
  const btn = $("addLabelBtn");
  btn.disabled = true;
  try {
    const res = await graph(`/planner/plans/${plan.id}/details`, { raw: true });
    if (!res.ok) throw new Error(`Graph ${res.status}`);
    const d = await res.json();
    const desc = d.categoryDescriptions || {};
    let slot = null;
    for (let i = 1; i <= 25; i++) {
      if (!desc["category" + i]) { slot = "category" + i; break; }
    }
    if (!slot) throw new Error("כל 25 התוויות בתוכנית כבר בשימוש.");
    await graph(`/planner/plans/${plan.id}/details`, {
      method: "PATCH",
      body: { categoryDescriptions: { [slot]: name } },
      headers: { "If-Match": d["@odata.etag"], Prefer: "return=minimal" },
    });
    state.catDesc = { ...desc, [slot]: name };
    state.selectedLabels.add(slot);
    input.value = "";
    renderLabels();
  } catch (e) {
    err.textContent = "לא הצלחתי ליצור את התווית: " + (e.message || e);
    err.classList.remove("hidden");
  } finally {
    btn.disabled = false;
  }
}

/* ---------------- People search (whole organization) ---------------- */
function notMemberMsg(plan) {
  const sel = $("assigneeSelect").selectedOptions[0];
  const name = sel ? sel.textContent : "המשתמש";
  return `${name} לא משויך לתוכנית "${plan.title}". יש להוסיף אותו כחבר בקבוצה של התוכנית ואז לנסות שוב, או לבחור משתמש אחר.`;
}

function assigneeNotMember() {
  const id = $("assigneeSelect").value;
  const plan = currentPlan();
  return !!(id && plan && plan.groupId && state.members.size && !state.members.has(id) && id !== state.me.id);
}

function updateAssigneeWarning() {
  const bad = assigneeNotMember();
  if (bad) $("assigneeWarn").textContent = notMemberMsg(currentPlan());
  $("assigneeWarn").classList.toggle("hidden", !bad);
}

let peopleTimer = null;
let peopleSeq = 0;
function onPeopleInput() {
  clearTimeout(peopleTimer);
  const q = $("peopleSearch").value.trim();
  if (q.length < 2) {
    $("peopleResults").classList.add("hidden");
    return;
  }
  peopleTimer = setTimeout(() => searchPeople(q), 300);
}

async function searchPeople(q) {
  const seq = ++peopleSeq;
  const box = $("peopleResults");
  const safe = q.replace(/"/g, "");
  try {
    const res = await graph(
      `/users?$search=${encodeURIComponent(`"displayName:${safe}" OR "mail:${safe}"`)}` +
        `&$select=id,displayName,mail,userPrincipalName&$top=15&$count=true`,
      { headers: { ConsistencyLevel: "eventual" } }
    );
    if (seq !== peopleSeq) return;
    box.innerHTML = "";
    const people = (res.value || []).filter((u) => u.mail);
    if (!people.length) {
      box.innerHTML = '<div class="muted small" style="padding:8px 10px">לא נמצאו אנשים</div>';
    }
    for (const u of people) {
      const b = document.createElement("button");
      b.type = "button";
      b.textContent = u.displayName;
      const sub = document.createElement("span");
      sub.className = "sub";
      sub.textContent = u.mail || u.userPrincipalName || "";
      b.appendChild(sub);
      b.addEventListener("click", () => pickPerson(u));
      box.appendChild(b);
    }
    box.classList.remove("hidden");
  } catch (e) {
    if (seq !== peopleSeq) return;
    box.innerHTML = '<div class="muted small" style="padding:8px 10px">החיפוש נכשל: ' +
      String(e.message || e).replace(/</g, "&lt;") + "</div>";
    box.classList.remove("hidden");
  }
}

function pickPerson(u) {
  const aSel = $("assigneeSelect");
  if (![...aSel.options].some((o) => o.value === u.id)) aSel.appendChild(option(u.id, u.displayName));
  aSel.value = u.id;
  $("peopleSearch").value = "";
  $("peopleResults").classList.add("hidden");
  updateAssigneeWarning();
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
    attachments: (item.attachments || [])
      .filter((a) => !a.isInline && a.attachmentType !== "cloud")
      .map((a) => a.name)
      .filter(Boolean),
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

let mailSeq = 0;
async function loadMailInfo() {
  const seq = ++mailSeq;
  state.mailInfo = null;
  $("doneBefore").classList.add("hidden");
  const empty = { webLink: null, categories: [] };
  if (!state.mail || !state.mail.restId) {
    state.mailInfo = empty;
    return;
  }
  const id = encodeURIComponent(state.mail.restId);
  try {
    const m = await graph(`/me/messages/${id}?$select=webLink,categories`);
    if (seq !== mailSeq) return;
    state.mailInfo = { webLink: m.webLink || null, categories: m.categories || [] };
  } catch {
    if (seq !== mailSeq) return;
    state.mailInfo = empty;
  }
  const info = state.mailInfo;
  $("doneBefore").classList.toggle("hidden", !info.categories.includes(MAIL_CATEGORY));
}

// The green "משימה ב-Planner" category is created once per mailbox (manually or by macro),
// so the add-in only needs Mail.ReadWrite to tag the message.
async function markMail() {
  const id = encodeURIComponent(state.mail.restId);
  const current = (state.mailInfo && state.mailInfo.categories) || [];
  const cats = Array.from(new Set([...current, MAIL_CATEGORY]));
  await graph(`/me/messages/${id}`, { method: "PATCH", body: { categories: cats } });
  if (state.mailInfo) state.mailInfo.categories = cats;
}

function friendly(e, prefix) {
  const m = String((e && e.message) || e);
  if (/40[13]|consent|AADSTS65001|interaction_required|insufficient|Access ?Denied/i.test(m)) {
    return `${prefix}: חסרה הרשאה (Admin consent ב-Entra).`;
  }
  return `${prefix}: ${m}`;
}

/* ---------------- Create task ---------------- */
function encodeRefKey(url) {
  return url.replace(/%/g, "%25").replace(/\./g, "%2E").replace(/:/g, "%3A").replace(/@/g, "%40").replace(/#/g, "%23");
}

function buildDescription(mail, body) {
  const lines = [];
  lines.push(`מאת: ${mail.fromName}${mail.fromAddr ? ` <${mail.fromAddr}>` : ""}`);
  if (mail.received) {
    lines.push(
      "התקבל: " +
        mail.received.toLocaleString("he-IL", { dateStyle: "short", timeStyle: "short" })
    );
  }
  lines.push(`נושא: ${mail.subject}`);
  if (mail.attachments && mail.attachments.length) {
    lines.push(`קבצים מצורפים (במייל המקורי): ${mail.attachments.join(", ")}`);
  }
  if (body) {
    lines.push("", "────────────", "");
    body = body.replace(/\r\n/g, "\n").replace(/\n{3,}/g, "\n\n");
    const room = MAX_DESC - lines.join("\n").length - 40;
    if (body.length > room) body = body.slice(0, room) + "\n…(קוצר)";
    lines.push(body);
  }
  return lines.join("\n");
}

function fillMailFields() {
  $("titleInput").value = state.mail.subject;
  $("bodyInput").value = state.mail.body;
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

function checkDates() {
  const start = $("startInput").value;
  const due = $("dueInput").value;
  const bad = !!(start && due && due < start);
  $("dateWarn").classList.toggle("hidden", !bad);
  return !bad;
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
  if (!checkDates()) {
    $("dueInput").focus();
    return;
  }
  if (assigneeNotMember()) {
    updateAssigneeWarning();
    $("assigneeSelect").focus();
    return;
  }
  const bSel = $("bucketSelect");
  let bucketId = bSel.value || undefined;
  const newBucketName = $("newBucketInput").value.trim();
  if (bucketId === NEW_BUCKET && !newBucketName) {
    $("newBucketInput").focus();
    return;
  }
  const assignee = $("assigneeSelect").value;
  const start = $("startInput").value;
  const due = $("dueInput").value;
  const warnings = [];

  const btn = $("createBtn");
  const setBtn = (t) => (btn.textContent = t);
  btn.disabled = true;
  setBtn("יוצר משימה…");
  try {
    if (state.mailInfoPromise) await state.mailInfoPromise;

    let bucketName = bSel.selectedOptions[0] ? bSel.selectedOptions[0].textContent : "";
    if (bucketId === NEW_BUCKET) {
      const b = await graph("/planner/buckets", {
        method: "POST",
        body: { name: newBucketName, planId: plan.id, orderHint: " !" },
      });
      bucketId = b.id;
      bucketName = b.name;
      bSel.insertBefore(option(b.id, b.name), bSel.querySelector(`option[value="${NEW_BUCKET}"]`));
      bSel.value = b.id;
      $("newBucketInput").value = "";
      onBucketChange();
    }

    const payload = {
      planId: plan.id,
      title: title.slice(0, 255),
      priority: Number($("prioritySelect").value) || 5,
    };
    if (bucketId) payload.bucketId = bucketId;
    if (start) payload.startDateTime = new Date(start + "T08:00:00").toISOString();
    if (due) payload.dueDateTime = new Date(due + "T17:00:00").toISOString();
    if (assignee) {
      payload.assignments = {
        [assignee]: { "@odata.type": "#microsoft.graph.plannerAssignment", orderHint: " !" },
      };
    }
    if (state.selectedLabels.size) {
      payload.appliedCategories = Object.fromEntries([...state.selectedLabels].map((k) => [k, true]));
    }
    let task;
    try {
      task = await graph("/planner/tasks", { method: "POST", body: payload });
    } catch (e) {
      if (assignee && assignee !== state.me.id && /Graph (400|403)/.test(e.message || "")) {
        throw new Error(notMemberMsg(plan));
      }
      throw e;
    }

    const refs = {};
    const info = state.mailInfo || {};

    setBtn("משלים פרטים…");
    const etag = await getDetailsEtag(task.id);
    const details = {
      description: buildDescription(
        state.mail,
        $("includeBody").checked ? $("bodyInput").value.trim() : ""
      ),
      previewType: "description",
    };
    if (info.webLink) {
      refs[encodeRefKey(info.webLink)] = {
        "@odata.type": "microsoft.graph.plannerExternalReference",
        alias: "📧 פתח את המייל המקורי",
        type: "Other",
        previewPriority: " !",
      };
    }
    if (Object.keys(refs).length) details.references = refs;
    await graph(`/planner/tasks/${task.id}/details`, {
      method: "PATCH",
      body: details,
      headers: { "If-Match": etag, Prefer: "return=minimal" },
    });

    if ($("markMail").checked) {
      try {
        await markMail();
        $("doneBefore").classList.remove("hidden");
      } catch (e) {
        warnings.push(friendly(e, "המייל לא סומן בקטגוריה"));
      }
    }

    state.lastBucket[plan.id] = bucketId || "";
    rsSet({ lastPlanId: plan.id, lastBucket: state.lastBucket });

    $("doneDetails").textContent = plan.title + (bucketId && bucketName ? " · " + bucketName : "");
    $("doneWarn").textContent = warnings.join(" ");
    $("doneWarn").classList.toggle("hidden", !warnings.length);
    $("openTaskLink").href = `https://tasks.office.com/${state.tenantId}/Home/Task/${task.id}`;
    show("done");
  } catch (e) {
    showError(e);
  } finally {
    btn.disabled = false;
    setBtn("צור משימה");
  }
}

/* ---------------- Boot ---------------- */
async function loadEverything() {
  setStatus("מתחבר…");
  await getToken();
  state.me = state.me || (await graph("/me?$select=id,displayName"));

  setStatus("קורא את המייל…");
  state.mail = await loadMail();
  fillMailFields();
  state.mailInfoPromise = loadMailInfo();

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
  if (!Office.context || !Office.context.mailbox) {
    showError("יש לפתוח את התוסף מתוך Outlook (מתוך מייל פתוח).");
    return;
  }
  state.favorites = rsGet("favorites", []);
  state.lastBucket = rsGet("lastBucket", {});

  $("planSelect").addEventListener("change", onPlanChange);
  $("planFilter").addEventListener("input", () => {
    const before = $("planSelect").value;
    renderPlans();
    if ($("planSelect").value !== before) onPlanChange();
    else renderFavorites();
  });
  $("assigneeSelect").addEventListener("change", updateAssigneeWarning);
  $("bucketSelect").addEventListener("change", onBucketChange);
  $("addLabelBtn").addEventListener("click", addLabel);
  $("newLabelInput").addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      addLabel();
    }
  });
  $("startInput").addEventListener("change", checkDates);
  $("dueInput").addEventListener("change", checkDates);
  $("peopleSearch").addEventListener("input", onPeopleInput);
  $("includeBody").addEventListener("change", () =>
    $("bodyInput").classList.toggle("hidden", !$("includeBody").checked)
  );
  $("favBtn").addEventListener("click", toggleFavorite);
  $("taskForm").addEventListener("submit", createTask);
  $("retryBtn").addEventListener("click", start);
  $("againBtn").addEventListener("click", () => show("taskForm"));
  const canClose = !!(Office.context.ui && Office.context.ui.closeContainer);
  $("closeBtn").classList.toggle("hidden", !canClose);
  $("closeBtn").addEventListener("click", () => {
    try {
      Office.context.ui.closeContainer();
    } catch {
      $("closeBtn").classList.add("hidden");
    }
  });
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
      fillMailFields();
      state.mailInfoPromise = loadMailInfo();
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
