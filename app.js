const C = window.CONFIG || {}, $ = s => document.querySelector(s);
const SC = ["User.Read", "Mail.Read", "Files.Read.All", "Calendars.Read"];

function isConfigured() {
  return !!C.clientId && C.clientId !== "YOUR-ENTRA-APP-CLIENT-ID" && !!C.workerUrl && C.workerUrl !== "https://YOUR-WORKER.YOUR-SUBDOMAIN.workers.dev";
}

const pca = window.msal ? new msal.PublicClientApplication({
  auth: { clientId: C.clientId, authority: "https://login.microsoftonline.com/" + (C.tenant || "common"), redirectUri: location.origin + location.pathname },
  cache: { cacheLocation: "sessionStorage" }
}) : null;

let acct = pca ? pca.getAllAccounts()[0] || null : null, busy = false;

function setTheme(theme) {
  const next = theme === "light" ? "light" : "dark";
  document.body.setAttribute("data-theme", next);
  localStorage.setItem("theme", next);
  const toggle = $("#themeToggle");
  if (toggle) toggle.textContent = next === "dark" ? "Toggle light mode" : "Toggle dark mode";
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function renderMarkdown(text) {
  const safe = escapeHtml(String(text || ""));
  const blocks = safe
    .replace(/```([\s\S]*?)```/g, (_, code) => `<pre><code>${code.trim()}</code></pre>`)
    .replace(/\*\*([^*]+)\*\*/g, "<strong>$1</strong>")
    .replace(/\*([^*]+)\*/g, "<em>$1</em>")
    .replace(/`([^`]+)`/g, "<code>$1</code>")
    .replace(/\[(.+?)\]\((https?:\/\/[^\s)]+)\)/g, '<a href="$2" target="_blank" rel="noopener">$1</a>')
    .replace(/\n\n+/g, "</p><p>")
    .replace(/\n/g, "<br>");
  return `<p>${blocks}</p>`;
}

function ui() {
  const configured = isConfigured();
  $("#who").textContent = acct ? acct.username : (configured ? "Signed out" : "Setup required");
  $("#login").hidden = !!acct || !configured;
  $("#logout").hidden = !acct || !configured;
  $("#login").disabled = !configured;
  $("#in").disabled = $("#send").disabled = !acct || !configured;
  $("#userInitial").textContent = acct ? (acct.username || "A").charAt(0).toUpperCase() : "A";
  const welcome = $("#welcomeCard");
  if (welcome) welcome.style.display = $("#col").children.length > 1 ? "none" : "flex";
}

function say(cls, text, src) {
  const d = document.createElement("div");
  d.className = cls;
  d.innerHTML = renderMarkdown(text);

  if ((src || []).length) {
    const sbox = document.createElement("div");
    sbox.className = "src";
    (src || []).forEach((s, i) => {
      const item = document.createElement("span");
      item.className = "src-item";
      const label = document.createElement("span");
      label.textContent = "[" + (i + 1) + "] " + (s.type || "Result");
      item.appendChild(label);
      if (/^https:/.test(s.link)) {
        const a = document.createElement("a");
        a.href = s.link;
        a.target = "_blank";
        a.rel = "noopener";
        a.textContent = s.title || "Source";
        item.appendChild(a);
      } else {
        const n = document.createElement("span");
        n.textContent = s.title || "Source";
        item.appendChild(n);
      }
      sbox.appendChild(item);
    });
    d.appendChild(sbox);
  }

  $("#col").append(d);
  $("#feed").scrollTop = 1e9;
  ui();
  return d;
}

function loadingMessage(label) {
  const wrap = document.createElement("div");
  wrap.className = "loading";
  wrap.innerHTML = '<span>' + label + '</span><span class="dot-loader"><span></span><span></span><span></span></span>';
  $("#col").appendChild(wrap);
  $("#feed").scrollTop = 1e9;
  return wrap;
}

async function login() {
  if (!isConfigured()) return say("err", "Set your Microsoft Entra app client ID and Cloudflare worker URL in config.js before signing in.");
  if (!pca) return say("err", "The Microsoft authentication library did not load. Check the page script tags.");
  try {
    const result = await pca.loginPopup({ scopes: SC, prompt: "select_account" });
    acct = result.account;
    ui();
  } catch (e) {
    say("err", "Sign-in failed: " + (e && e.message ? e.message : String(e)));
  }
}

function logout() {
  if (!pca || !acct) { acct = null; ui(); return; }
  pca.logoutPopup({ account: acct }).catch(() => {}).finally(() => { acct = null; ui(); });
}

async function token() {
  if (!acct || !pca) throw new Error("Sign in with your Microsoft account before asking the AI.");
  try { return (await pca.acquireTokenSilent({ scopes: SC, account: acct })).accessToken; }
  catch (e) { return (await pca.acquireTokenPopup({ scopes: SC, account: acct })).accessToken; }
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

async function ask(q, w) {
  const tok = await token();
  if (w) { w.innerHTML = '<span>Searching</span><span class="dot-loader"><span></span><span></span><span></span></span>'; }
  const kw = (await call("keywords", { question: q }, tok)).text.replace(/["']/g, "").trim();
  if (w) { w.innerHTML = '<span>Analyzing</span><span class="dot-loader"><span></span><span></span><span></span></span>'; }
  const src = (await graphSearch(tok, kw)).slice(0, 15);
  if (!src.length) return say("ai", "I searched your Microsoft 365 for “" + kw + "” and found nothing, so I can’t answer this from your data.");
  if (w) { w.innerHTML = '<span>Generating answer</span><span class="dot-loader"><span></span><span></span><span></span></span>'; }
  const a = await call("ask", { question: q, sources: src }, tok); say("ai", a.text, src);
}

$("#login").onclick = login;
$("#logout").onclick = logout;
$("#themeToggle").onclick = () => setTheme(document.body.dataset.theme === "dark" ? "light" : "dark");
document.querySelectorAll(".suggestion").forEach((btn) => {
  btn.onclick = () => {
    const q = btn.dataset.question || btn.textContent.trim();
    $("#in").value = q;
    $("#f").requestSubmit ? $("#f").requestSubmit() : $("#f").dispatchEvent(new Event("submit", { cancelable: true, bubbles: true }));
  };
});

$("#f").onsubmit = async e => {
  e.preventDefault();
  const q = $("#in").value.trim();
  if (!q || busy || !acct) return;
  busy = true;
  $("#in").value = "";
  const welcome = $("#welcomeCard");
  if (welcome) welcome.style.display = "none";
  say("me", q);
  const w = loadingMessage("Searching");
  try {
    await ask(q, w);
  } catch (err) {
    say("err", err.message);
  } finally {
    w.remove();
    busy = false;
  }
};

const savedTheme = localStorage.getItem("theme") || "dark";
setTheme(savedTheme);
ui();

