const C = window.CONFIG, $ = s => document.querySelector(s);
const SC = ["User.Read", "Mail.Read", "Files.Read.All", "Calendars.Read"];
const pca = new msal.PublicClientApplication({
  auth: { clientId: C.clientId, authority: "https://login.microsoftonline.com/" + C.tenant, redirectUri: location.origin + location.pathname },
  cache: { cacheLocation: "sessionStorage" }
});
let acct = pca.getAllAccounts()[0] || null, busy = false;

function ui() { $("#who").textContent = acct ? acct.username : "Signed out"; $("#login").hidden = !!acct; $("#logout").hidden = !acct; $("#in").disabled = $("#send").disabled = !acct; }
function say(cls, text, src) {
  const d = document.createElement("div"); d.className = cls; d.textContent = text;
  (src || []).forEach((s, i) => { const l = document.createElement("span"); l.className = "src"; l.append("[" + (i + 1) + "] " + s.type + " · ");
    if (/^https:/.test(s.link)) { const a = document.createElement("a"); a.href = s.link; a.target = "_blank"; a.rel = "noopener"; a.textContent = s.title; l.append(a); } else l.append(s.title);
    l.append(" · " + (s.date || "").slice(0, 10)); d.append(l); });
  $("#col").append(d); $("#feed").scrollTop = 1e9; return d;
}
async function login() { try { acct = (await pca.loginPopup({ scopes: SC })).account; ui(); } catch (e) { say("err", "Sign-in failed: " + e.message); } }
function logout() { pca.logoutPopup({ account: acct }).finally(() => { acct = null; ui(); }); }
async function token() {
  try { return (await pca.acquireTokenSilent({ scopes: SC, account: acct })).accessToken; }
  catch (e) { return (await pca.acquireTokenPopup({ scopes: SC })).accessToken; }
}
async function graphSearch(tok, q) {
  const body = { requests: ["message", "driveItem", "event"].map(t => ({ entityTypes: [t], query: { queryString: q }, from: 0, size: 5 })) };
  const r = await fetch("https://graph.microsoft.com/v1.0/search/query", { method: "POST", headers: { authorization: "Bearer " + tok, "content-type": "application/json" }, body: JSON.stringify(body) });
  if (!r.ok) throw new Error("Microsoft Graph search failed (HTTP " + r.status + ")");
  const d = await r.json(), out = [];
  for (const v of d.value || []) for (const h of v.hitsContainers || []) for (const x of h.hits || []) {
    const o = x.resource || {}, t = o["@odata.type"] || "";
    out.push({ type: t.includes("message") ? "Email" : t.includes("event") ? "Event" : "File", title: o.subject || o.name || "(untitled)", link: o.webLink || "",
      date: o.receivedDateTime || o.lastModifiedDateTime || (o.start && o.start.dateTime) || "", text: (x.summary || o.bodyPreview || "").replace(/<[^>]+>/g, "") });
  }
  return out;
}
async function call(mode, body, tok) {
  const r = await fetch(C.workerUrl, { method: "POST", headers: { authorization: "Bearer " + tok, "content-type": "application/json" }, body: JSON.stringify({ mode, ...body }) });
  const d = await r.json().catch(() => ({})); if (!r.ok) throw new Error(d.error || "Worker HTTP " + r.status); return d;
}
async function ask(q) {
  const tok = await token();
  const kw = (await call("keywords", { question: q }, tok)).text.replace(/["']/g, "").trim();
  const src = (await graphSearch(tok, kw)).slice(0, 15);
  if (!src.length) return say("ai", "I searched your Microsoft 365 for “" + kw + "” and found nothing, so I can’t answer this from your data.");
  const a = await call("ask", { question: q, sources: src }, tok); say("ai", a.text, src);
}
$("#login").onclick = login; $("#logout").onclick = logout;
$("#f").onsubmit = async e => {
  e.preventDefault(); const q = $("#in").value.trim(); if (!q || busy || !acct) return;
  busy = true; $("#in").value = ""; say("me", q); const w = say("ai", "Searching your Microsoft 365…");
  try { await ask(q); } catch (err) { say("err", err.message); } finally { w.remove(); busy = false; }
};
ui();
