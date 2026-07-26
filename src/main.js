import { InteractionRequiredAuthError, PublicClientApplication } from "@azure/msal-browser";
import "./style.css";

const CONFIG = Object.freeze({
  clientId: "e6230b6d-43ff-4f47-bdca-d6212960099f",
  tenantId: "c6ec425d-54a9-492e-a1f6-6640678bb37b",
  siteId:
    "bama.sharepoint.com,8007680f-1a54-409c-b3ca-69abe3d2612d,43c2d80a-4f7a-4db4-85ce-7b9df2b7d3ae",
  filePath: "Produksjonsoversikt/produksjon2.json",
  refreshMinutes: 5,
});

const GROUP_META = Object.freeze({
  Dagligvare: { color: "#52d273", background: "#102d1b" },
  Storkjøkken: { color: "#65a7ff", background: "#10233d" },
  Beger: { color: "#f47dbc", background: "#351329" },
  Industri: { color: "#f6b94b", background: "#35230c" },
  Vakuum: { color: "#c999ff", background: "#26133c" },
  "HF Trim": { color: "#54c9f3", background: "#0d2934" },
  "HF Kuttetorg": { color: "#aa91f5", background: "#241a3c" },
});

const FILTERS = [
  ["ALL", "Alle"],
  ["DONE", "Ferdig"],
  ["ACTIVE", "Pågår"],
  ["SCHEDULED", "Planlagt"],
];

const els = {
  loginView: document.querySelector("#loginView"),
  appView: document.querySelector("#appView"),
  loginButton: document.querySelector("#loginButton"),
  loginMessage: document.querySelector("#loginMessage"),
  logoutButton: document.querySelector("#logoutButton"),
  accountButton: document.querySelector("#accountButton"),
  accountMenu: document.querySelector("#accountMenu"),
  accountInitials: document.querySelector("#accountInitials"),
  accountName: document.querySelector("#accountName"),
  accountEmail: document.querySelector("#accountEmail"),
  refreshButton: document.querySelector("#refreshButton"),
  clock: document.querySelector("#clock"),
  statusBanner: document.querySelector("#statusBanner"),
  statusText: document.querySelector("#statusText"),
  lastUpdated: document.querySelector("#lastUpdated"),
  summaryGrid: document.querySelector("#summaryGrid"),
  etaCard: document.querySelector("#etaCard"),
  groupGrid: document.querySelector("#groupGrid"),
  filterBar: document.querySelector("#filterBar"),
  searchInput: document.querySelector("#searchInput"),
  ordersBody: document.querySelector("#ordersBody"),
  emptyMessage: document.querySelector("#emptyMessage"),
};

const redirectUri = `${window.location.origin}${window.location.pathname}`;
const msal = new PublicClientApplication({
  auth: {
    clientId: CONFIG.clientId,
    authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
    redirectUri,
    postLogoutRedirectUri: redirectUri,
    navigateToLoginRequestUrl: false,
  },
  cache: {
    cacheLocation: "localStorage",
  },
});

let activeAccount = null;
let allData = [];
let currentFilter = "ALL";
let currentSearch = "";
let refreshTimer = null;

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function number(value) {
  const parsed = Number.parseFloat(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function formatNumber(value) {
  return Math.round(value).toLocaleString("nb-NO");
}

function percentage(done, planned) {
  return planned > 0 ? Math.round((done / planned) * 100) : 0;
}

function formatDate(value) {
  if (!value || String(value).startsWith("1900")) return "—";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "—";
  return date.toLocaleDateString("nb-NO", { day: "2-digit", month: "2-digit" });
}

function statusLabel(status) {
  return (
    {
      Completed: "Ferdig",
      ReportedFinished: "Ferdigmeldt",
      StartedUp: "Pågår",
      Scheduled: "Planlagt",
      CostEstimated: "Kostnadsestimert",
    }[status] || status || "Ukjent"
  );
}

function statusClass(status) {
  if (status === "Completed" || status === "ReportedFinished") return "status-done";
  if (status === "StartedUp") return "status-active";
  if (status === "Scheduled" || status === "CostEstimated") return "status-scheduled";
  return "status-neutral";
}

function processData(raw) {
  if (!Array.isArray(raw)) throw new Error("JSON-filen har ikke forventet format.");
  return raw
    .filter((row) => !row.Site || row.Site === "BI-Tranby")
    .map((row) => {
      const planned = number(row.QTY_scheduled);
      const done = number(row.QTY_good);
      const remaining = number(row.QTY_remaining);
      return {
        ...row,
        planned,
        done,
        remaining,
        progress: percentage(done, planned),
        isDone:
          row.Status === "Completed" ||
          row.Status === "ReportedFinished" ||
          (remaining === 0 && done > 0),
        isScheduled: row.Status === "Scheduled" || row.Status === "CostEstimated",
      };
    });
}

function summarize(data) {
  const sum = (rows, key) => rows.reduce((total, row) => total + row[key], 0);
  const fvGroups = new Set(["Dagligvare", "Storkjøkken", "Beger", "Industri", "Vakuum"]);
  const hfGroups = new Set(["HF Trim", "HF Kuttetorg"]);
  const fv = data.filter((row) => fvGroups.has(row.Prod_group));
  const hf = data.filter((row) => hfGroups.has(row.Prod_group));
  const groups = {};

  for (const row of data) {
    const group = row.Prod_group || "Annet";
    groups[group] ||= { planned: 0, done: 0, remaining: 0 };
    groups[group].planned += row.planned;
    groups[group].done += row.done;
    groups[group].remaining += row.remaining;
  }

  return {
    totalOrders: data.length,
    doneOrders: data.filter((row) => row.isDone).length,
    activeOrders: data.filter((row) => row.Status === "StartedUp").length,
    totalPlanned: sum(data, "planned"),
    totalDone: sum(data, "done"),
    totalRemaining: sum(data, "remaining"),
    fvPlanned: sum(fv, "planned"),
    fvDone: sum(fv, "done"),
    fvRemaining: sum(fv, "remaining"),
    hfPlanned: sum(hf, "planned"),
    hfDone: sum(hf, "done"),
    hfRemaining: sum(hf, "remaining"),
    groups,
  };
}

function progressColor(value) {
  if (value < 30) return "#ef6a6a";
  if (value < 70) return "#f6b94b";
  return "#52d273";
}

function summaryCard(label, value, detail, progress, tone = "", valueClass = "") {
  const color = progressColor(progress);
  return `
    <article class="summary-card ${tone}">
      <p>${escapeHtml(label)}</p>
      <strong class="${valueClass}">${escapeHtml(value)}</strong>
      <span>${escapeHtml(detail)}</span>
      <div class="progress-track"><i style="width:${Math.min(progress, 100)}%;background:${color}"></i></div>
    </article>`;
}

function renderSummary(summary) {
  const fvProgress = Math.min(100, percentage(summary.fvDone, summary.fvPlanned));
  const hfProgress = Math.min(100, percentage(summary.hfDone, summary.hfPlanned));
  const orderProgress = percentage(summary.doneOrders, summary.totalOrders);
  const totalProgress = Math.min(100, percentage(summary.totalDone, summary.totalPlanned));
  els.summaryGrid.innerHTML = [
    summaryCard(
      "Plan FV",
      formatNumber(summary.fvPlanned),
      `${fvProgress}% ferdig · stk`,
      fvProgress,
      "tone-fv",
    ),
    summaryCard(
      "Ferdigmeldt FV",
      formatNumber(summary.fvDone),
      `Gjenstår ${formatNumber(summary.fvRemaining)} stk`,
      fvProgress,
      "tone-fv",
      "value-good",
    ),
    summaryCard(
      "Plan HF",
      formatNumber(summary.hfPlanned),
      `${hfProgress}% ferdig · kg`,
      hfProgress,
      "tone-hf",
    ),
    summaryCard(
      "Ferdigmeldt HF",
      formatNumber(summary.hfDone),
      `Gjenstår ${formatNumber(summary.hfRemaining)} kg`,
      hfProgress,
      "tone-hf",
      "value-good",
    ),
    summaryCard(
      "Ordrestatus",
      `${summary.doneOrders} / ${summary.totalOrders}`,
      `${summary.activeOrders} pågår`,
      orderProgress,
    ),
    summaryCard(
      "Speedometer total",
      `${totalProgress}%`,
      "BI Tranby totalt",
      totalProgress,
      "",
      "value-good",
    ),
  ].join("");
}

function calculateEta(planned, done) {
  const now = new Date();
  const nowDecimal = now.getHours() + now.getMinutes() / 60;
  const start = 6;
  const neededRate = Math.round(planned / 16);
  const remaining = Math.max(0, planned - done);
  if (done >= planned && planned > 0) return { eta: "Ferdig", status: "Ferdig", rate: 0, neededRate, remaining };
  if (done <= 0 || nowDecimal <= start) {
    return { eta: "Ikke startet", status: "Ikke startet", rate: 0, neededRate, remaining };
  }
  const rate = Math.round(done / (nowDecimal - start));
  if (rate <= 0) return { eta: "Beregner …", status: "Ikke startet", rate, neededRate, remaining };
  const etaDecimal = nowDecimal + remaining / rate;
  const hours = Math.floor(etaDecimal) % 24;
  const minutes = Math.round((etaDecimal % 1) * 60);
  const difference = rate - neededRate;
  const status = Math.abs(difference) <= 300 ? "På plan" : difference < 0 ? "Bak plan" : "Foran plan";
  return {
    eta: `Est. ferdig ${String(hours).padStart(2, "0")}:${String(minutes).padStart(2, "0")} ${
      etaDecimal < 24 ? "i dag" : "i morgen"
    }`,
    status,
    rate,
    neededRate,
    remaining,
  };
}

function makeGauge(progress, color) {
  const clamped = Math.max(0, Math.min(100, progress));
  const radius = 58;
  const centerX = 76;
  const centerY = 70;
  const start = -Math.PI;
  const end = start + Math.PI * (clamped / 100);
  const x = centerX + radius * Math.cos(end);
  const y = centerY + radius * Math.sin(end);
  const largeArc = clamped > 50 ? 1 : 0;
  return `
    <svg class="gauge" viewBox="0 0 152 84" role="img" aria-label="${progress} prosent ferdig">
      <path d="M18 70 A58 58 0 0 1 134 70" fill="none" stroke="#252a34" stroke-width="13" stroke-linecap="round"/>
      ${
        clamped > 0
          ? `<path d="M18 70 A58 58 0 ${largeArc} 1 ${x.toFixed(2)} ${y.toFixed(
              2,
            )}" fill="none" stroke="${color}" stroke-width="13" stroke-linecap="round"/>`
          : ""
      }
      <line x1="${centerX}" y1="${centerY}" x2="${x.toFixed(2)}" y2="${y.toFixed(
        2,
      )}" stroke="#ecebe7" stroke-width="3" stroke-linecap="round"/>
      <circle cx="${centerX}" cy="${centerY}" r="5" fill="#ecebe7"/>
    </svg>`;
}

function renderProgress(summary) {
  const daily = summary.groups.Dagligvare || { planned: 0, done: 0 };
  const kitchen = summary.groups.Storkjøkken || { planned: 0, done: 0 };
  const finishedPlanned = daily.planned + kitchen.planned;
  const finishedDone = daily.done + kitchen.done;
  const eta = calculateEta(finishedPlanned, finishedDone);
  els.etaCard.innerHTML = `
    <p class="eta-title">Ferdigvare (DV + SK)</p>
    <div class="eta-totals">
      <span><small>Plan totalt</small><strong>${formatNumber(finishedPlanned)} <i>stk</i></strong></span>
      <span><small>Gjenstår</small><strong class="text-danger">${formatNumber(eta.remaining)} <i>stk</i></strong></span>
    </div>
    <h2>${escapeHtml(eta.eta)}</h2>
    <div class="eta-stats">
      <span><small>Status</small><strong class="${eta.status === "Bak plan" ? "text-danger" : "eta-status"}">${escapeHtml(
        eta.status,
      )}</strong></span>
      <span><small>Takt/time nå</small><strong>${formatNumber(eta.rate)}</strong></span>
      <span><small>Behov takt/time</small><strong>${formatNumber(eta.neededRate)}</strong></span>
    </div>`;

  const gaugeGroups = ["Dagligvare", "Storkjøkken", "Beger", "HF Trim", "HF Kuttetorg"];
  els.groupGrid.innerHTML = gaugeGroups
    .map((name) => {
      const meta = GROUP_META[name];
      const group = summary.groups[name] || { planned: 0, done: 0, remaining: 0 };
      const progress = percentage(group.done, group.planned);
      const unit = name.startsWith("HF") ? "kg" : "stk";
      return `
        <article class="group-card" style="--group-color:${meta.color};--group-background:${meta.background}">
          <h3>${escapeHtml(name)}</h3>
          ${makeGauge(progress, meta.color)}
          <strong class="gauge-value">${progress}%</strong>
          <p>${formatNumber(group.done)} / ${formatNumber(group.planned)} ${unit}</p>
          <div class="gauge-numbers">
            <span><b class="text-good">${formatNumber(group.done)}</b>Ferdig</span>
            <span><b class="text-danger">${formatNumber(group.remaining)}</b>Gjenstår</span>
          </div>
        </article>`;
    })
    .join("");
}

function renderFilters() {
  const availableGroups = new Set(allData.map((row) => row.Prod_group).filter(Boolean));
  const buttons = [
    ...FILTERS,
    ...Object.keys(GROUP_META)
      .filter((group) => availableGroups.has(group))
      .map((group) => [group, group]),
  ];
  els.filterBar.innerHTML = buttons
    .map(
      ([key, label]) =>
        `<button type="button" data-filter="${escapeHtml(key)}" class="${
          currentFilter === key ? "active" : ""
        }">${escapeHtml(label)}</button>`,
    )
    .join("");
}

function matchesFilter(row) {
  if (currentFilter === "DONE") return row.isDone;
  if (currentFilter === "ACTIVE") return row.Status === "StartedUp";
  if (currentFilter === "SCHEDULED") return row.isScheduled;
  if (currentFilter === "ALL") return true;
  return row.Prod_group === currentFilter;
}

function renderOrders() {
  const query = currentSearch.trim().toLocaleLowerCase("nb-NO");
  const rows = allData
    .filter(matchesFilter)
    .filter((row) => {
      if (!query) return true;
      return [row.Order_name, row.Item_nr, row.Order_nr]
        .filter(Boolean)
        .some((value) => String(value).toLocaleLowerCase("nb-NO").includes(query));
    })
    .sort((a, b) => String(a.Order_name || "").localeCompare(String(b.Order_name || ""), "nb-NO"));

  els.ordersBody.innerHTML = rows
    .map((row) => {
      const meta = GROUP_META[row.Prod_group] || { color: "#a8a8a8", background: "#232630" };
      return `
        <tr class="${row.isDone ? "completed-row" : ""}">
          <td data-label="Produkt">${escapeHtml(row.Order_name || "—")}<small>Ordre ${escapeHtml(row.Order_nr || "—")}</small></td>
          <td data-label="Varenr.">${escapeHtml(row.Item_nr || "—")}</td>
          <td data-label="Gruppe"><span class="group-tag" style="color:${meta.color};background:${meta.background}">${escapeHtml(
            row.Prod_group || "—",
          )}</span></td>
          <td data-label="Status"><span class="order-status ${statusClass(row.Status)}">${escapeHtml(
            statusLabel(row.Status),
          )}</span></td>
          <td data-label="Planlagt">${formatNumber(row.planned)}</td>
          <td data-label="Ferdigmeldt" class="text-good">${row.done ? formatNumber(row.done) : "—"}</td>
          <td data-label="Gjenstår">${row.remaining ? formatNumber(row.remaining) : "—"}</td>
          <td data-label="Prosent"><strong>${row.progress ? `${row.progress}%` : "—"}</strong></td>
          <td data-label="Levering">${formatDate(row.Date_delivery)}</td>
        </tr>`;
    })
    .join("");
  els.emptyMessage.hidden = rows.length !== 0;
}

function renderDashboard() {
  const summary = summarize(allData);
  renderSummary(summary);
  renderProgress(summary);
  renderFilters();
  renderOrders();
}

function showLogin(message = "") {
  els.loginMessage.textContent = message;
  els.loginView.hidden = false;
  els.appView.hidden = true;
}

function showApp() {
  els.loginView.hidden = true;
  els.appView.hidden = false;
  const name = activeAccount?.name || "Innlogget bruker";
  const email = activeAccount?.username || "";
  els.accountName.textContent = name;
  els.accountEmail.textContent = email;
  els.accountInitials.textContent = name
    .split(/\s+/)
    .slice(0, 2)
    .map((part) => part[0])
    .join("")
    .toUpperCase();
}

function setStatus(message, type = "loading") {
  els.statusText.textContent = message;
  els.statusBanner.dataset.type = type;
}

async function getAccessToken() {
  const request = { scopes: ["Files.Read"], account: activeAccount };
  try {
    return (await msal.acquireTokenSilent(request)).accessToken;
  } catch (error) {
    if (error instanceof InteractionRequiredAuthError) {
      await msal.acquireTokenRedirect(request);
      return null;
    }
    throw error;
  }
}

async function fetchSharePointData() {
  const token = await getAccessToken();
  if (!token) return null;
  const encodedPath = CONFIG.filePath
    .split("/")
    .map((part) => encodeURIComponent(part))
    .join("/");
  const metadataUrl =
    `https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(CONFIG.siteId)}` +
    `/drive/root:/${encodedPath}?$select=id,name,lastModifiedDateTime,@microsoft.graph.downloadUrl`;
  const metadataResponse = await fetch(metadataUrl, {
    headers: { Authorization: `Bearer ${token}` },
    cache: "no-store",
  });
  if (!metadataResponse.ok) {
    const details = await metadataResponse.json().catch(() => null);
    throw new Error(details?.error?.message || `Microsoft Graph svarte ${metadataResponse.status}.`);
  }
  const metadata = await metadataResponse.json();
  const downloadUrl = metadata["@microsoft.graph.downloadUrl"];
  if (!downloadUrl) throw new Error("SharePoint returnerte ingen nedlastingsadresse.");
  const dataResponse = await fetch(downloadUrl, { cache: "no-store" });
  if (!dataResponse.ok) throw new Error(`Kunne ikke laste produksjonsfilen (${dataResponse.status}).`);
  return { raw: await dataResponse.json(), modified: metadata.lastModifiedDateTime };
}

async function loadData() {
  if (!activeAccount) return;
  els.refreshButton.disabled = true;
  setStatus("Henter ferske produksjonsdata …");
  try {
    const result = await fetchSharePointData();
    if (!result) return;
    allData = processData(result.raw);
    renderDashboard();
    const modified = result.modified
      ? new Date(result.modified).toLocaleString("nb-NO", { dateStyle: "short", timeStyle: "short" })
      : new Date().toLocaleTimeString("nb-NO");
    els.lastUpdated.textContent = `Sist oppdatert ${modified}`;
    setStatus("Produksjonsdata er oppdatert", "success");
  } catch (error) {
    console.error(error);
    setStatus(`Kunne ikke hente data: ${error.message}`, "error");
  } finally {
    els.refreshButton.disabled = false;
  }
}

async function signIn() {
  els.loginButton.disabled = true;
  els.loginMessage.textContent = "Åpner sikker Microsoft-innlogging …";
  try {
    await msal.loginRedirect({ scopes: ["User.Read", "Files.Read"], redirectStartPage: redirectUri });
  } catch (error) {
    console.error(error);
    els.loginButton.disabled = false;
    els.loginMessage.textContent = "Innloggingen kunne ikke startes. Prøv igjen.";
  }
}

async function initialize() {
  await msal.initialize();
  const redirectResult = await msal.handleRedirectPromise();
  activeAccount = redirectResult?.account || msal.getAllAccounts()[0] || null;
  if (!activeAccount) {
    showLogin();
    return;
  }
  msal.setActiveAccount(activeAccount);
  showApp();
  await loadData();
  refreshTimer = window.setInterval(loadData, CONFIG.refreshMinutes * 60 * 1000);
}

els.loginButton.addEventListener("click", signIn);
els.logoutButton.addEventListener("click", () => {
  window.clearInterval(refreshTimer);
  msal.logoutRedirect({ account: activeAccount, postLogoutRedirectUri: redirectUri });
});
els.refreshButton.addEventListener("click", loadData);
els.accountButton.addEventListener("click", () => {
  const shouldOpen = els.accountMenu.hidden;
  els.accountMenu.hidden = !shouldOpen;
  els.accountButton.setAttribute("aria-expanded", String(shouldOpen));
});
els.filterBar.addEventListener("click", (event) => {
  const button = event.target.closest("button[data-filter]");
  if (!button) return;
  currentFilter = button.dataset.filter;
  renderFilters();
  renderOrders();
});
els.searchInput.addEventListener("input", (event) => {
  currentSearch = event.target.value;
  renderOrders();
});

function updateClock() {
  els.clock.textContent = new Date().toLocaleTimeString("nb-NO");
}
updateClock();
window.setInterval(updateClock, 1000);

if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("./sw.js").catch((error) => console.warn("Service worker:", error));
  });
}

initialize().catch((error) => {
  console.error(error);
  showLogin("Appen kunne ikke startes. Last siden på nytt.");
});
