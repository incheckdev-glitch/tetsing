(() => {
  const { WEBAPP_EXEC_URL, DEFAULT_URGENT_TOP_N } = window.APP_CONFIG;
  const U = window.Utils;

  const els = {
    year: document.getElementById("year"),
    buildInfo: document.getElementById("buildInfo"),
    statusDot: document.getElementById("statusDot"),
    lastUpdatedText: document.getElementById("lastUpdatedText"),
    thresholdHint: document.getElementById("thresholdHint"),

    kpiOver: document.getElementById("kpiOver"),
    kpiNear: document.getElementById("kpiNear"),
    kpiOk: document.getElementById("kpiOk"),
    kpiNoCommit: document.getElementById("kpiNoCommit"),

    btnRefresh: document.getElementById("btnRefresh"),
    btnExportCsv: document.getElementById("btnExportCsv"),
    btnResetFilters: document.getElementById("btnResetFilters"),

    searchInput: document.getElementById("searchInput"),
    statusSelect: document.getElementById("statusSelect"),
    clientSelect: document.getElementById("clientSelect"),
    sortSelect: document.getElementById("sortSelect"),
    pageSizeSelect: document.getElementById("pageSizeSelect"),

    resultCount: document.getElementById("resultCount"),
    tableBody: document.getElementById("tableBody"),

    btnPrev: document.getElementById("btnPrev"),
    btnNext: document.getElementById("btnNext"),
    pageMeta: document.getElementById("pageMeta"),
  };

  const state = {
    raw: [],
    filtered: [],
    clients: [],
    threshold: 120,
    page: 1,
    pageSize: 25,
    sort: "remainingAsc",
    search: "",
    status: "ALL",
    client: "ALL",
    charts: { urgent: null, status: null },
    tableSortOverride: null, // for clicking column headers
  };

  const API_SUMMARY = () => `${WEBAPP_EXEC_URL}?action=summary`;

  function setStatus(mode) {
    // modes: idle, loading, ok, error
    const dot = els.statusDot;
    dot.classList.remove("dot--idle","dot--loading","dot--ok","dot--error");
    dot.classList.add(`dot--${mode}`);
  }

  async function fetchSummary() {
    setStatus("loading");
    els.lastUpdatedText.textContent = "Loading…";

    const res = await fetch(API_SUMMARY(), { method: "GET" });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json();
    return data;
  }

  function normalizeRows(rows) {
    return (rows || []).map(r => ({
      client: (r.client ?? "").toString(),
      account: (r.account ?? "").toString(),
      committed: U.toInt(r.committed, 0),
      used: U.toInt(r.used, 0),
      remaining: U.toInt(r.remaining, 0),
      status: (r.status ?? "").toString()
    }));
  }

  function fillClientFilter(clients) {
    // preserve selection if possible
    const prev = els.clientSelect.value || "ALL";
    els.clientSelect.innerHTML = `<option value="ALL">All clients</option>`;
    for (const c of clients) {
      const opt = document.createElement("option");
      opt.value = c;
      opt.textContent = c;
      els.clientSelect.appendChild(opt);
    }
    els.clientSelect.value = clients.includes(prev) ? prev : "ALL";
  }

  function computeClients(rows) {
    const set = new Set(rows.map(r => r.client).filter(Boolean));
    return [...set].sort((a,b) => a.localeCompare(b));
  }

  function applyFilters() {
    const q = state.search.trim().toLowerCase();
    const status = state.status;
    const client = state.client;

    let rows = [...state.raw];

    if (status !== "ALL") rows = rows.filter(r => r.status === status);
    if (client !== "ALL") rows = rows.filter(r => r.client === client);

    if (q) {
      rows = rows.filter(r =>
        r.client.toLowerCase().includes(q) ||
        r.account.toLowerCase().includes(q) ||
        r.status.toLowerCase().includes(q)
      );
    }

    // Sorting
    const sortKey = state.tableSortOverride || state.sort;
    rows.sort(getSorter(sortKey));

    state.filtered = rows;

    // Paging clamp
    const totalPages = Math.max(1, Math.ceil(rows.length / state.pageSize));
    state.page = U.clamp(state.page, 1, totalPages);

    renderTable();
    renderMeta();
    renderCharts();
  }

  function getSorter(key) {
    switch (key) {
      case "remainingAsc": return (a,b) => a.remaining - b.remaining || a.client.localeCompare(b.client);
      case "remainingDesc": return (a,b) => b.remaining - a.remaining || a.client.localeCompare(b.client);
      case "usedDesc": return (a,b) => b.used - a.used || a.client.localeCompare(b.client);
      case "clientAsc": return (a,b) => a.client.localeCompare(b.client) || a.account.localeCompare(b.account);

      // Table header click sorts:
      case "client": return (a,b) => a.client.localeCompare(b.client) || a.account.localeCompare(b.account);
      case "account": return (a,b) => a.account.localeCompare(b.account) || a.client.localeCompare(b.client);
      case "committed": return (a,b) => b.committed - a.committed || a.client.localeCompare(b.client);
      case "used": return (a,b) => b.used - a.used || a.client.localeCompare(b.client);
      case "remaining": return (a,b) => a.remaining - b.remaining || a.client.localeCompare(b.client);
      case "status": return (a,b) => a.status.localeCompare(b.status) || a.client.localeCompare(b.client);
      default: return (a,b) => a.remaining - b.remaining;
    }
  }

  function slicePage(rows) {
    const start = (state.page - 1) * state.pageSize;
    return rows.slice(start, start + state.pageSize);
  }

  function badge(status) {
    const cls =
      status === "OVER LIMIT" ? "badge badge--danger" :
      status === "NEAR LIMIT" ? "badge badge--warning" :
      status === "OK" ? "badge badge--ok" :
      "badge badge--neutral";
    return `<span class="${cls}">${U.escapeHtml(status)}</span>`;
  }

  function renderTable() {
    const rows = slicePage(state.filtered);
    const html = rows.map(r => `
      <tr>
        <td>${U.escapeHtml(r.client)}</td>
        <td>${U.escapeHtml(r.account)}</td>
        <td class="num">${U.fmtNumber(r.committed)}</td>
        <td class="num">${U.fmtNumber(r.used)}</td>
        <td class="num ${r.remaining < 0 ? "neg" : ""}">${U.fmtNumber(r.remaining)}</td>
        <td>${badge(r.status)}</td>
      </tr>
    `).join("");

    els.tableBody.innerHTML = html || `<tr><td colspan="6" class="muted">No results</td></tr>`;
  }

  function renderMeta() {
    const total = state.filtered.length;
    const totalPages = Math.max(1, Math.ceil(total / state.pageSize));
    els.resultCount.textContent = `${total.toLocaleString()} result${total === 1 ? "" : "s"}`;

    els.pageMeta.textContent = `Page ${state.page} of ${totalPages}`;
    els.btnPrev.disabled = state.page <= 1;
    els.btnNext.disabled = state.page >= totalPages;
  }

  function renderKpis(counts) {
    els.kpiOver.textContent = (counts?.over ?? 0).toLocaleString();
    els.kpiNear.textContent = (counts?.near ?? 0).toLocaleString();
    els.kpiOk.textContent = (counts?.ok ?? 0).toLocaleString();
    els.kpiNoCommit.textContent = (counts?.noCommit ?? 0).toLocaleString();
  }

  function renderCharts() {
    renderUrgentChart();
    renderStatusChart();
  }

  function renderUrgentChart() {
    const topN = DEFAULT_URGENT_TOP_N || 12;
    const urgent = [...state.filtered]
      .filter(r => r.status !== "NO COMMIT SET")
      .sort((a,b) => a.remaining - b.remaining)
      .slice(0, topN);

    const labels = urgent.map(r => `${r.client} / ${r.account}`);
    const values = urgent.map(r => r.remaining);

    const ctx = document.getElementById("chartUrgent");
    if (state.charts.urgent) state.charts.urgent.destroy();

    state.charts.urgent = new Chart(ctx, {
      type: "bar",
      data: {
        labels,
        datasets: [{ label: "Remaining minutes", data: values }]
      },
      options: {
        responsive: true,
        plugins: {
          legend: { display: false },
          tooltip: { mode: "index", intersect: false }
        },
        scales: {
          x: { ticks: { maxRotation: 50, minRotation: 0 } },
          y: { beginAtZero: true }
        }
      }
    });
  }

  function renderStatusChart() {
    const counts = {
      "OVER LIMIT": 0,
      "NEAR LIMIT": 0,
      "OK": 0,
      "NO COMMIT SET": 0
    };
    for (const r of state.raw) {
      if (counts[r.status] !== undefined) counts[r.status]++;
    }

    const labels = Object.keys(counts);
    const values = labels.map(k => counts[k]);

    const ctx = document.getElementById("chartStatus");
    if (state.charts.status) state.charts.status.destroy();

    state.charts.status = new Chart(ctx, {
      type: "doughnut",
      data: {
        labels,
        datasets: [{ label: "Accounts", data: values }]
      },
      options: {
        responsive: true,
        plugins: {
          legend: { position: "bottom" }
        }
      }
    });
  }

  function wireEvents() {
    els.btnRefresh.addEventListener("click", () => load());

    els.btnExportCsv.addEventListener("click", () => {
      const csv = U.toCsv(state.filtered);
      const ts = new Date().toISOString().slice(0,19).replaceAll(":","-");
      U.downloadText(`incheck-lite-dashboard-${ts}.csv`, csv);
    });

    els.btnResetFilters.addEventListener("click", () => {
      state.search = "";
      state.status = "ALL";
      state.client = "ALL";
      state.sort = "remainingAsc";
      state.page = 1;
      state.tableSortOverride = null;

      els.searchInput.value = "";
      els.statusSelect.value = "ALL";
      els.clientSelect.value = "ALL";
      els.sortSelect.value = "remainingAsc";
      applyFilters();
    });

    els.searchInput.addEventListener("input", (e) => {
      state.search = e.target.value || "";
      state.page = 1;
      applyFilters();
    });

    els.statusSelect.addEventListener("change", (e) => {
      state.status = e.target.value;
      state.page = 1;
      applyFilters();
    });

    els.clientSelect.addEventListener("change", (e) => {
      state.client = e.target.value;
      state.page = 1;
      applyFilters();
    });

    els.sortSelect.addEventListener("change", (e) => {
      state.sort = e.target.value;
      state.tableSortOverride = null;
      state.page = 1;
      applyFilters();
    });

    els.pageSizeSelect.addEventListener("change", (e) => {
      state.pageSize = Number(e.target.value) || 25;
      state.page = 1;
      applyFilters();
    });

    els.btnPrev.addEventListener("click", () => {
      state.page = Math.max(1, state.page - 1);
      renderTable();
      renderMeta();
      renderCharts();
    });

    els.btnNext.addEventListener("click", () => {
      const totalPages = Math.max(1, Math.ceil(state.filtered.length / state.pageSize));
      state.page = Math.min(totalPages, state.page + 1);
      renderTable();
      renderMeta();
      renderCharts();
    });

    // Click to sort columns
    const ths = document.querySelectorAll("#accountsTable thead th");
    ths.forEach(th => {
      th.addEventListener("click", () => {
        const key = th.getAttribute("data-sort");
        if (!key) return;
        state.tableSortOverride = key;
        state.page = 1;
        applyFilters();
      });
    });
  }

  async function load() {
    try {
      const data = await fetchSummary();

      state.threshold = U.toInt(data.thresholdNearLimitMinutes, 120);
      els.thresholdHint.textContent = `Near limit threshold: ${state.threshold} minutes (≤ threshold).`;

      renderKpis(data.counts);

      state.raw = normalizeRows(data.rows);
      state.clients = computeClients(state.raw);
      fillClientFilter(state.clients);

      state.page = 1;
      applyFilters();

      setStatus("ok");
      els.lastUpdatedText.textContent = `Updated: ${U.fmtDateTime(data.generatedAt)}`;
    } catch (err) {
      console.error(err);
      setStatus("error");
      els.lastUpdatedText.textContent = "Failed to load";
      els.tableBody.innerHTML = `<tr><td colspan="6" class="muted">Could not load data. Check WEBAPP_EXEC_URL and deployment access.</td></tr>`;
    }
  }

  function init() {
    els.year.textContent = new Date().getFullYear();
    els.buildInfo.textContent = "v1.0";
    wireEvents();

    // default page size
    state.pageSize = Number(els.pageSizeSelect.value) || 25;

    // quick config validation
    if (!WEBAPP_EXEC_URL || WEBAPP_EXEC_URL.includes("PASTE_YOUR_WEBAPP_EXEC_URL_HERE")) {
      setStatus("error");
      els.lastUpdatedText.textContent = "Config missing";
      els.tableBody.innerHTML = `<tr><td colspan="6" class="muted">Set WEBAPP_EXEC_URL in assets/js/config.js</td></tr>`;
      return;
    }

    load();
  }

  init();
})();
