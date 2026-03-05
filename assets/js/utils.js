window.Utils = (() => {
  function escapeHtml(v) {
    return (v ?? "").toString()
      .replaceAll("&","&amp;")
      .replaceAll("<","&lt;")
      .replaceAll(">","&gt;")
      .replaceAll('"',"&quot;")
      .replaceAll("'","&#039;");
  }

  function clamp(n, min, max) {
    return Math.max(min, Math.min(max, n));
  }

  function toInt(v, fallback = 0) {
    const n = Number(v);
    return Number.isFinite(n) ? Math.trunc(n) : fallback;
  }

  function fmtNumber(n) {
    const x = Number(n);
    if (!Number.isFinite(x)) return "0";
    return x.toLocaleString();
  }

  function fmtDateTime(iso) {
    try {
      const d = new Date(iso);
      return d.toLocaleString(undefined, { year: "numeric", month: "short", day: "2-digit", hour: "2-digit", minute: "2-digit" });
    } catch {
      return "—";
    }
  }

  function downloadText(filename, text) {
    const blob = new Blob([text], { type: "text/plain;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  function toCsv(rows) {
    const headers = ["client","account","committed","used","remaining","status"];
    const lines = [headers.join(",")];
    for (const r of rows) {
      const line = headers.map(h => {
        const val = (r[h] ?? "").toString().replaceAll('"','""');
        return `"${val}"`;
      }).join(",");
      lines.push(line);
    }
    return lines.join("\n");
  }

  return { escapeHtml, clamp, toInt, fmtNumber, fmtDateTime, downloadText, toCsv };
})();
