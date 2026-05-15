// src/workers/aaqWorker.js
import * as XLSX from "xlsx";

// Compute actual used range (all populated cells) and override ws['!ref'].
// This still reads *all* data cells, but skips empty formatted tails.
function normalizeRef(ws) {
  const cellRegex = /^[A-Z]+[0-9]+$/;
  let minR = Infinity, minC = Infinity, maxR = -1, maxC = -1;

  for (const addr of Object.keys(ws)) {
    if (!cellRegex.test(addr)) continue;
    const cell = ws[addr];
    // treat cells with a value or formula as "real"
    if (!cell || (cell.v == null && cell.f == null)) continue;

    const { r, c } = XLSX.utils.decode_cell(addr);
    if (r < minR) minR = r;
    if (c < minC) minC = c;
    if (r > maxR) maxR = r;
    if (c > maxC) maxC = c;
  }

  if (maxR >= 0 && maxC >= 0) {
    ws["!ref"] = XLSX.utils.encode_range({ s: { r: minR, c: minC }, e: { r: maxR, c: maxC } });
  }
}

// Your existing parseAAQ logic should be moved into the worker,
// OR imported if you split it to a separate module.
// For clarity: we’ll inline a minimal hook — you can paste your existing parseAAQ here.
function classifyUsage(count) {
  const c = parseInt(count, 10);
  if (Number.isNaN(c)) return "UNKNOWN";
  if (c >= 50000) return "HIGH";
  if (c >= 1000) return "MEDIUM";
  if (c >= 100) return "LOW";
  return "VERY LOW";
}

function parseAAQ(workbook) {
  const notes = [];
  const sheetNames = workbook.SheetNames;

  notes.push(`Sheets found: ${sheetNames.join(", ")}`);

  const findSheet = (...patterns) =>
    sheetNames.find(s => patterns.some(p => s.toLowerCase().includes(p.toLowerCase())));

  // --- SERVER ---
  const servers = [];
  const serverSheet = findSheet("server");
  if (!serverSheet) notes.push("Server sheet not found");
  else {
    const ws = workbook.Sheets[serverSheet];
    normalizeRef(ws);
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, blankrows: false });

    // find header row
    let headerIdx = -1;
    for (let i = 0; i < raw.length; i++) {
      const rowText = raw[i].join(" ").toLowerCase().replace(/\s/g, "");
      if (rowText.includes("workloadname") || rowText.includes("cpucount")) { headerIdx = i; break; }
    }
    if (headerIdx === -1) notes.push("Server header row not found");
    else {
      const headers = raw[headerIdx].map(h => String(h ?? "").trim());
      const norm = headers.map(h => h.replace(/[\s_]/g, "").toLowerCase());
      const findCol = (...keys) => {
        for (const k of keys) {
          const kk = k.replace(/[\s_]/g, "").toLowerCase();
          const idx = norm.findIndex(c => c.includes(kk));
          if (idx !== -1) return headers[idx];
        }
        return null;
      };
      const workloadCol = findCol("workloadname");
      const functionCol = findCol("function");
      const cpuCol = findCol("cpucount") || findCol("totalcpucores", "cpu");
      const ramCol = findCol("ram", "memory");

      notes.push(`Server CPU column: ${cpuCol}`);
      notes.push(`Server RAM column: ${ramCol}`);

      const safeNum = v => {
        if (v == null || v === "") return null;
        const s = String(v).replace(/[^0-9.]/g, "");
        const n = parseFloat(s);
        return Number.isNaN(n) ? null : n;
      };

      for (let r = headerIdx + 1; r < raw.length; r++) {
        const row = raw[r];
        if (!row || row.length === 0) continue;

        const rec = {};
        headers.forEach((h, i) => { rec[h] = row[i]; });

        const name = rec[workloadCol];
        if (!name || String(name).trim() === "") continue;
        if (String(name).toLowerCase().includes("dependency")) break;

        const cpu = safeNum(rec[cpuCol]);
        const ramRaw = safeNum(rec[ramCol]);
        const mem = ramRaw == null ? null : (ramRaw > 1e9 ? ramRaw / (1024 ** 3) : ramRaw);

        servers.push({
          server: String(name).trim(),
          application: String(rec[functionCol] ?? "").trim(),
          vcpu: cpu != null ? Math.round(cpu) : null,
          memory_gib: mem != null ? Math.round(mem * 10) / 10 : null,
        });
      }
      notes.push(`Extracted ${servers.length} application servers`);
    }
  }

  // --- FIREWALL ---
  const firewall = [];
  const fwSheet = findSheet("firewall");
  if (!fwSheet) notes.push("Firewall sheet not found");
  else {
    const ws = workbook.Sheets[fwSheet];
    normalizeRef(ws);
    const rows = XLSX.utils.sheet_to_json(ws, { defval: null, blankrows: false });
    const normKey = obj => {
      const out = {};
      for (const [k, v] of Object.entries(obj)) out[k.toLowerCase().replace(/\s+/g, "_")] = v;
      return out;
    };

    for (const rawRow of rows) {
      const r = normKey(rawRow);

      const srcIp = r["src_ip"] ? String(r["src_ip"]).trim() : null;
      const dstIp = r["dest_ip"] ? String(r["dest_ip"]).trim() : null;
      if (!srcIp || !dstIp) continue;

      const count = parseInt(String(r["netstat_count"] ?? "0").replace(/,/g, ""), 10) || 0;

      firewall.push({
        source_ip: srcIp,
        destination_ip: dstIp,
        source_hostname: r["source_hostname"] ? String(r["source_hostname"]).trim() : null,
        destination_hostname: r["dest_hostname"] ? String(r["dest_hostname"]).trim() : null,
        source_application: r["source_stackname"] ? String(r["source_stackname"]).trim() : null,
        destination_application: r["dest_stackname"] ? String(r["dest_stackname"]).trim() : null,
        source_port: r["src_port"] ? String(r["src_port"]).trim() : null,
        destination_port: r["dest_port"] ? String(r["dest_port"]).trim() : null,
        protocol: r["protocol"] ? String(r["protocol"]).trim() : null,
        netstat_count: count,
        usage: classifyUsage(count),
      });
    }
    notes.push(`Extracted ${firewall.length} firewall rules`);
  }

  // --- DATABASE ---
  let databases = [];
  const dbSheet = findSheet("database");
  if (!dbSheet) notes.push("Database sheet not found");
  else {
    const ws = workbook.Sheets[dbSheet];
    normalizeRef(ws);
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, blankrows: false });

    let headerIdx = -1;
    for (let i = 0; i < raw.length; i++) {
      const txt = raw[i].join(" ").toLowerCase();
      if (txt.includes("database name") && txt.includes("db server name")) { headerIdx = i; break; }
    }
    if (headerIdx === -1) notes.push("Database header row not found");
    else {
      const headers = raw[headerIdx].map(h => String(h ?? "").trim().toLowerCase());
      const sizeCol = headers.find(h => h.includes("database size"));
      const sizeUnit = sizeCol?.includes("mb") ? "MB" : "GB";
      notes.push(`DB size column: ${sizeCol} (${sizeUnit})`);

      const safeNum = v => {
        if (v == null || v === "") return null;
        const s = String(v).replace(/[^0-9.]/g, "");
        const n = parseFloat(s);
        return Number.isNaN(n) ? null : n;
      };

      for (let r = headerIdx + 1; r < raw.length; r++) {
        const row = raw[r];
        if (!row || row.length === 0) continue;

        const rec = {};
        headers.forEach((h, i) => { rec[h] = row[i]; });

        const sizeRaw = safeNum(rec[sizeCol]);
        const sizeGb = sizeRaw == null ? null : (sizeUnit === "MB" ? sizeRaw / 1024 : sizeRaw);

        databases.push({
          db_server: rec["db server name"] ? String(rec["db server name"]).trim() : null,
          database_name: rec["database name"] ? String(rec["database name"]).trim() : null,
          database_instance: rec["database instance"] ? String(rec["database instance"]).trim() : null,
          database_size_gb: sizeGb != null ? Math.round(sizeGb * 100) / 100 : null,
          database_type: rec["database type"] ? String(rec["database type"]).trim() : null,
        });
      }
      databases = databases.filter(d => d.db_server || d.database_name);
      notes.push(`Extracted ${databases.length} databases`);
    }
  }

  return { servers, firewall, databases, notes };
}

self.onmessage = (e) => {
  const { buffer } = e.data;
  try {
    // XLSX must read entire file in memory (XLSX is a ZIP container) [3](https://github.com/SheetJS/sheetjs/issues/1136)
    const wb = XLSX.read(new Uint8Array(buffer), { type: "array" });
    const parsed = parseAAQ(wb);
    self.postMessage({ ok: true, result: parsed });
  } catch (err) {
    self.postMessage({ ok: false, error: String(err?.message || err) });
  }
};