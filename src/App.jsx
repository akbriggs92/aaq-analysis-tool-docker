import { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import logo from "./assets/kyndryl_logo.webp";
import Worker from "./workers/aaqWorker.js?worker";

// ─── Palette & Theme ───────────────────────────────────────────────────────
const CSS = `
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');


  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

  :root {
    --bg: #0a0c10;
    --surface: #111318;
    --surface2: #181c24;
    --border: #1f2533;
    --accent: #00e5ff;
    --accent2: #ff3d71;
    --accent3: #a259ff;
    --text: #e8eaf0;
    --muted: #5a6275;
    --high: #ff3d71;
    --med: #ffaa00;
    --low: #00e5ff;
    --vlow: #3a4155;
    --font-head: 'Inter', sans-serif;
    --font-mono: 'Inter', sans-serif;
    --radius: 8px;
    --transition: 180ms cubic-bezier(.4,0,.2,1);
  }

  body { background: var(--bg); color: var(--text); font-family: var(--font-mono); }

  .app {
    min-height: 100vh;
    background:
      radial-gradient(ellipse 80% 40% at 10% 0%, rgba(0,229,255,.05) 0%, transparent 60%),
      radial-gradient(ellipse 60% 30% at 90% 100%, rgba(162,89,255,.06) 0%, transparent 60%),
      var(--bg);
  }

  /* ── Header ── */
  .header {
    display: flex; align-items: center; gap: 14px;
    padding: 20px 36px;
    border-bottom: 1px solid var(--border);
    background: rgba(17,19,24,.85);
    backdrop-filter: blur(12px);
    position: sticky; top: 0; z-index: 100;
    box-shadow: 0 6px 24px rgba(0,0,0,.45);
  }

  .logo-img {
    height: 34px;
    width: auto;
    object-fit: contain;
    filter: drop-shadow(0 0 6px rgba(0,229,255,0.35));
    flex-shrink: 0;
  }

  .header-text h1 {
  font-family: var(--font-head);
  font-size: 18px;
  font-weight: 600;
  letter-spacing: -0.2px;
  color: var(--text);
}


  .header-text p { font-size: 11px; color: var(--muted); letter-spacing: .5px; margin-top: 1px; }

  .privacy-badge {
    margin-left: auto;
    display: flex; align-items: center; gap: 7px;
    background: rgba(0,229,255,.08);
    border: 1px solid rgba(0,229,255,.2);
    border-radius: 20px;
    padding: 6px 14px;
    font-size: 11px; color: var(--accent); letter-spacing: .4px;
  }

  .privacy-dot {
    width: 7px; height: 7px; border-radius: 50%;
    background: var(--accent);
    box-shadow: 0 0 8px var(--accent);
    animation: pulse 2s infinite;
  }

  @keyframes pulse { 0%,100%{opacity:1} 50%{opacity:.4} }

  /* ── Upload ── */
  .upload-zone {
    margin: 40px auto; max-width: 660px;
    border: 2px dashed var(--border);
    border-radius: 16px;
    padding: 56px 40px;
    text-align: center;
    cursor: pointer;
    transition: var(--transition);
    background: var(--surface);
    position: relative; overflow: hidden;
    box-shadow: 0 0 60px rgba(0,229,255,0.04);
  }
    
{loading && <div style={{color:"var(--muted)", marginTop: 12}}>Processing… large files may take a moment.</div>}
{error && <div style={{color:"var(--accent2)", marginTop: 12}}>Error: {error}</div>}

  .upload-zone::before {
    content:''; position:absolute; inset:0;
    background: linear-gradient(135deg, rgba(0,229,255,.04), rgba(162,89,255,.04));
    opacity:0; transition: var(--transition);
  }

  .upload-zone:hover, .upload-zone.drag { border-color: var(--accent); }
  .upload-zone:hover::before, .upload-zone.drag::before { opacity:1; }

  .upload-icon { font-size: 44px; margin-bottom: 14px; }
  .upload-title { font-family: var(--font-head); font-size: 20px; font-weight: 700; color: var(--text); }
  .upload-sub { font-size: 12px; color: var(--muted); margin-top: 8px; }

  .upload-btn {
    display: inline-block; margin-top: 22px;
    background: linear-gradient(135deg, var(--accent), var(--accent3));
    color: #000; font-family: var(--font-head); font-weight: 700;
    font-size: 13px; letter-spacing: .5px;
    padding: 10px 26px; border-radius: 8px; border: none; cursor: pointer;
    transition: var(--transition);
  }

  .upload-btn:hover { transform: translateY(-1px); box-shadow: 0 6px 24px rgba(0,229,255,.3); }

  .file-name {
    margin-top: 14px; font-size: 12px; color: var(--accent);
    display: flex; align-items: center; justify-content: center; gap: 8px;
  }

  /* ── Nav ── */
  .nav-bar {
    display: flex; gap: 4px;
    padding: 0 36px;
    border-bottom: 1px solid var(--border);
    background: var(--surface);
  }

  .nav-btn {
    padding: 14px 22px;
    font-family: var(--font-mono); font-size: 12px; font-weight: 500;
    color: var(--muted); background: transparent; border: none; cursor: pointer;
    border-bottom: 2px solid transparent;
    transition: var(--transition);
    letter-spacing: .3px;
  }

  .nav-btn:hover { color: var(--text); }
  .nav-btn.active { color: var(--accent); border-bottom-color: var(--accent); box-shadow: 0 3px 15px rgba(0,229,255,0.35); }

  /* ── Content ── */
  .content { padding: 60px 80px; max-width: 1500px; margin: 0 auto; }

  /* ── Cards / Metrics ── */
  .metrics { display: grid; grid-template-columns: repeat(3, 1fr); gap: 16px; margin-bottom: 32px; }

  .metric-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 24px 28px;
    position: relative;
    overflow: hidden;
    box-shadow: 0 10px 40px rgba(0,0,0,0.4);
  }

  .metric-card::after {
    content:''; position:absolute; top:0; left:0; right:0; height:2px;
  }

  .metric-card:nth-child(1)::after { background: var(--accent); }
  .metric-card:nth-child(2)::after { background: var(--accent3); }
  .metric-card:nth-child(3)::after { background: var(--accent2); }

  
.metric-label {
  font-size: 10px;
  color: var(--muted);
  letter-spacing: 0.5px;
  text-transform: uppercase;
}

  
.metric-value {
  font-family: var(--font-head);
  font-size: 32px;
  font-weight: 600;
  letter-spacing: -0.5px;
  color: var(--text);
  margin-top: 6px;
}


  /* ── Nav buttons on home ── */
  .nav-cards { display: grid; grid-template-columns: repeat(3,1fr); gap: 16px; margin-bottom: 32px; }

  .nav-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 20px 24px;
    cursor: pointer;
    transition: var(--transition);
    text-align: left;
    font-family: var(--font-mono);
  }

  .nav-card:hover {
    border-color: var(--accent);
    background: var(--surface2);
    transform: translateY(-4px);
    box-shadow: 0 10px 30px rgba(0,229,255,0.2);
  }

  .nav-card-icon { font-size: 26px; margin-bottom: 10px; }
  .nav-card-title { font-family: var(--font-head); font-size: 15px; font-weight: 700; color: var(--text); }
  .nav-card-sub { font-size: 11px; color: var(--muted); margin-top: 4px; }

  /* ── Section titles ── */
  
.section-title {
  font-family: var(--font-head);
  font-size: 22px;
  font-weight: 600;
  letter-spacing: -0.3px;
  margin-bottom: 16px;
}


  .section-title span {
    background: linear-gradient(90deg, var(--accent), var(--accent3));
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
  }

  /* ── Table ── */
  .table-wrap {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    overflow: auto;
    max-height: 520px;
    box-shadow: 0 20px 60px rgba(0,0,0,0.5);
  }

  table { border-collapse: collapse; width: 100%; font-size: 12px; }
  thead { position: sticky; top: 0; z-index: 2; }

  th {
    background: var(--surface2);
    padding: 11px 14px;
    text-align: left; font-weight: 500; color: var(--muted);
    text-transform: uppercase; letter-spacing: .6px; font-size: 10px;
    border-bottom: 1px solid var(--border);
    white-space: nowrap;
  }

  td {
    padding: 10px 14px;
    border-bottom: 1px solid rgba(31,37,51,.6);
    color: var(--text); font-size: 12px;
    white-space: nowrap; max-width: 220px;
    overflow: hidden; text-overflow: ellipsis;
  }

  tr:last-child td { border-bottom: none; }
  tr:hover td { background: rgba(255,255,255,.02); }

  /* ── Usage badges ── */
  .badge {
    display: inline-block; padding: 3px 9px;
    border-radius: 4px; font-size: 10px; font-weight: 600;
    letter-spacing: .4px;
  }

  .badge-HIGH { background: rgba(255,61,113,.15); color: var(--accent2); border: 1px solid rgba(255,61,113,.3); }
  .badge-MEDIUM { background: rgba(255,170,0,.12); color: #ffaa00; border: 1px solid rgba(255,170,0,.3); }
  .badge-LOW { background: rgba(0,229,255,.1); color: var(--accent); border: 1px solid rgba(0,229,255,.3); }
  .badge-VERYLOW { background: rgba(58,65,85,.3); color: var(--muted); border: 1px solid var(--border); }
  .badge-UNKNOWN { background: rgba(58,65,85,.2); color: var(--muted); border: 1px solid var(--border); }

  /* ── Sizing cards ── */
  .server-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 24px 28px;
    margin-bottom: 20px;
  }

  .server-name { font-family: var(--font-head); font-size: 17px; font-weight: 700; color: var(--text); }
  .server-meta { font-size: 11px; color: var(--muted); margin-top: 4px; margin-bottom: 16px; }
  .instance-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(240px,1fr)); gap: 12px; }

  .instance-card {
    background: var(--surface2);
    border: 1px solid var(--border);
    border-radius: 6px;
    padding: 14px 16px;
    transition: var(--transition);
  }

  .instance-card:hover { border-color: var(--accent3); }
  .instance-type { font-family: var(--font-head); font-size: 15px; font-weight: 700; color: var(--accent); }
  .instance-specs { font-size: 11px; color: var(--muted); margin-top: 4px; }
  .instance-price { font-size: 13px; color: var(--text); margin-top: 6px; font-weight: 500; }

  /* ── Export button ── */
  .export-btn {
    display: inline-flex; align-items: center; gap: 8px;
    background: linear-gradient(135deg, var(--accent), var(--accent3));
    color: #000; font-family: var(--font-head); font-weight: 700;
    font-size: 13px; letter-spacing: .3px;
    padding: 11px 24px; border-radius: 8px; border: none; cursor: pointer;
    transition: var(--transition); margin-top: 20px;
  }

  .export-btn:hover { transform: translateY(-1px); box-shadow: 0 6px 24px rgba(0,229,255,.3); }

  /* ── Parser notes ── */
  .notes-box {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 20px 24px;
    margin-top: 32px;
  }

  .note-item { font-size: 12px; color: var(--muted); padding: 4px 0; display: flex; gap: 10px; }
  .note-item::before { content: '›'; color: var(--accent3); }

  /* ── Search / filter ── */
  .filter-row { display: flex; gap: 12px; margin-bottom: 16px; flex-wrap: wrap; }

  .filter-input {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 6px;
    padding: 9px 14px;
    color: var(--text);
    font-family: var(--font-mono);
    font-size: 12px;
    outline: none;
    transition: var(--transition);
    min-width: 200px;
  }

  .filter-input:focus { border-color: var(--accent); }
  .filter-input::placeholder { color: var(--muted); }

  .filter-select {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 6px;
    padding: 9px 14px;
    color: var(--text);
    font-family: var(--font-mono);
    font-size: 12px;
    outline: none;
    cursor: pointer;
  }

  .filter-select option { background: var(--surface2); }

  .divider { border: none; border-top: 1px solid var(--border); margin: 32px 0; }
  .empty { text-align: center; padding: 48px; color: var(--muted); font-size: 13px; }

  /* ── Scrollbar ── */
  ::-webkit-scrollbar { width: 6px; height: 6px; }
  ::-webkit-scrollbar-track { background: var(--surface); }
  ::-webkit-scrollbar-thumb { background: var(--border); border-radius: 3px; }
  ::-webkit-scrollbar-thumb:hover { background: var(--muted); }
`;

// ─── Helpers ───────────────────────────────────────────────────────────────
function classifyUsage(count) {
  const c = parseInt(count, 10);
  if (isNaN(c)) return "UNKNOWN";
  if (c >= 50000) return "HIGH";
  if (c >= 1000) return "MEDIUM";
  if (c >= 100) return "LOW";
  return "VERY LOW";
}

function safeNum(v) {
  if (v == null || v === "") return null;
  const s = String(v).replace(/[^0-9.]/g, "");
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}

function findCol(cols, ...keys) {
  const norm = cols.map(c => c.replace(/[\s_]/g, "").toLowerCase());
  for (const key of keys) {
    const idx = norm.findIndex(c => c.includes(key.toLowerCase().replace(/[\s_]/g, "")));
    if (idx !== -1) return cols[idx];
  }
  return null;
}

// ─── Parser ────────────────────────────────────────────────────────────────
function parseAAQ(workbook) {
  const notes = [];
  const sheetNames = workbook.SheetNames;

  const findSheet = (...patterns) =>
    sheetNames.find(s => patterns.some(p => s.toLowerCase().includes(p.toLowerCase())));

  // ── Servers ──
  let servers = [];
  const serverSheet = findSheet("Server");
  if (!serverSheet) {
    notes.push("Server sheet not found");
  } else {
    const raw = XLSX.utils.sheet_to_json(workbook.Sheets[serverSheet], { header: 1 });
    let headerIdx = -1;
    for (let i = 0; i < raw.length; i++) {
      const rowText = raw[i].join(" ").toLowerCase().replace(/\s/g, "");
      if (rowText.includes("workloadname") || rowText.includes("cpucount")) {
        headerIdx = i; break;
      }
    }
    if (headerIdx === -1) {
      notes.push("Server header row not found");
    } else {
      const headers = raw[headerIdx].map(h => String(h ?? "").trim());
      const workloadCol = findCol(headers, "workloadname");
      const functionCol = findCol(headers, "function");
      const cpuCol = findCol(headers, "cpucount") || findCol(headers, "totalcpucores", "cpu");
      const ramCol = findCol(headers, "ram", "memory");

      notes.push(`Server CPU column: ${cpuCol}`);
      notes.push(`Server RAM column: ${ramCol}`);

      for (let r = headerIdx + 1; r < raw.length; r++) {
        const row = {};
        headers.forEach((h, i) => { row[h] = raw[r][i]; });
        const name = row[workloadCol];
        if (!name || String(name).trim() === "") continue;
        if (String(name).toLowerCase().includes("dependency")) break;

        const cpu = safeNum(row[cpuCol]);
        const ramRaw = safeNum(row[ramCol]);
        let mem = null;
        if (ramRaw != null) mem = ramRaw > 1e9 ? ramRaw / (1024 ** 3) : ramRaw;

        servers.push({
          server: String(name).trim(),
          application: String(row[functionCol] ?? "").trim(),
          vcpu: cpu != null ? Math.round(cpu) : null,
          memory_gib: mem != null ? Math.round(mem * 10) / 10 : null,
        });
      }
      notes.push(`Extracted ${servers.length} application servers`);
    }
  }

  // ── Firewall ──
  let firewall = [];
  const fwSheet = findSheet("firewall", "Firewall");
  if (!fwSheet) {
    notes.push("Firewall sheet not found");
  } else {
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[fwSheet], { defval: null });
    const normKey = obj => {
      const out = {};
      for (const [k, v] of Object.entries(obj)) out[k.toLowerCase().replace(/\s/g, "_")] = v;
      return out;
    };

    for (const rawRow of rows) {
      const r = normKey(rawRow);
      const srcIp = r["src_ip"] ? String(r["src_ip"]).trim() : null;
      const dstIp = r["dest_ip"] ? String(r["dest_ip"]).trim() : null;
      if (!srcIp || !dstIp) continue;

      const count = parseInt(String(r["netstat_count"] ?? "0").replace(",", ""), 10) || 0;
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

  // ── Databases ──
  let databases = [];
  const dbSheet = findSheet("Database");
  if (!dbSheet) {
    notes.push("Database sheet not found");
  } else {
    const raw = XLSX.utils.sheet_to_json(workbook.Sheets[dbSheet], { header: 1 });
    let headerIdx = -1;
    for (let i = 0; i < raw.length; i++) {
      const txt = raw[i].join(" ").toLowerCase();
      if (txt.includes("database name") && txt.includes("db server name")) {
        headerIdx = i; break;
      }
    }

    if (headerIdx === -1) {
      notes.push("Database header row not found");
    } else {
      const headers = raw[headerIdx].map(h => String(h ?? "").trim().toLowerCase());
      const sizeCol = headers.find(h => h.includes("database size"));
      const sizeUnit = sizeCol?.includes("mb") ? "MB" : "GB";
      notes.push(`DB size column: ${sizeCol} (${sizeUnit})`);

      for (let r = headerIdx + 1; r < raw.length; r++) {
        const row = {};
        headers.forEach((h, i) => { row[h] = raw[r][i]; });
        const sizeRaw = safeNum(row[sizeCol]);
        const sizeGb = sizeRaw == null ? null : (sizeUnit === "MB" ? sizeRaw / 1024 : sizeRaw);
        databases.push({
          db_server: row["db server name"] ? String(row["db server name"]).trim() : null,
          database_name: row["database name"] ? String(row["database name"]).trim() : null,
          database_instance: row["database instance"] ? String(row["database instance"]).trim() : null,
          database_size_gb: sizeGb != null ? Math.round(sizeGb * 100) / 100 : null,
          database_type: row["database type"] ? String(row["database type"]).trim() : null,
        });
      }

      databases = databases.filter(d => d.db_server || d.database_name);
      notes.push(`Extracted ${databases.length} databases`);
    }
  }

  return { servers, firewall, databases, notes };
}

// ─── AWS Sizing (static fallback data, no server call) ─────────────────────
const INSTANCE_DATA = [
  { instance_type: "t3.small", vCPU: 2, memory: 2, price_usd: 0.0208 },
  { instance_type: "t3.medium", vCPU: 2, memory: 4, price_usd: 0.0416 },
  { instance_type: "t3.large", vCPU: 2, memory: 8, price_usd: 0.0832 },
  { instance_type: "t3.xlarge", vCPU: 4, memory: 16, price_usd: 0.1664 },
  { instance_type: "t3.2xlarge", vCPU: 8, memory: 32, price_usd: 0.3328 },
  { instance_type: "m5.large", vCPU: 2, memory: 8, price_usd: 0.096 },
  { instance_type: "m5.xlarge", vCPU: 4, memory: 16, price_usd: 0.192 },
  { instance_type: "m5.2xlarge", vCPU: 8, memory: 32, price_usd: 0.384 },
  { instance_type: "m5.4xlarge", vCPU: 16, memory: 64, price_usd: 0.768 },
  { instance_type: "m5.8xlarge", vCPU: 32, memory: 128, price_usd: 1.536 },
  { instance_type: "m5.12xlarge", vCPU: 48, memory: 192, price_usd: 2.304 },
  { instance_type: "c5.large", vCPU: 2, memory: 4, price_usd: 0.085 },
  { instance_type: "c5.xlarge", vCPU: 4, memory: 8, price_usd: 0.17 },
  { instance_type: "c5.2xlarge", vCPU: 8, memory: 16, price_usd: 0.34 },
  { instance_type: "c5.4xlarge", vCPU: 16, memory: 32, price_usd: 0.68 },
  { instance_type: "r5.large", vCPU: 2, memory: 16, price_usd: 0.126 },
  { instance_type: "r5.xlarge", vCPU: 4, memory: 32, price_usd: 0.252 },
  { instance_type: "r5.2xlarge", vCPU: 8, memory: 64, price_usd: 0.504 },
  { instance_type: "r5.4xlarge", vCPU: 16, memory: 128, price_usd: 1.008 },
  { instance_type: "r5.8xlarge", vCPU: 32, memory: 256, price_usd: 2.016 },
];

function recommendInstances(vcpu, memGib) {
  if (!vcpu || !memGib) return [];
  return INSTANCE_DATA
    .filter(i => i.vCPU >= vcpu && i.memory >= memGib)
    .sort((a, b) => a.vCPU - b.vCPU || a.memory - b.memory)
    .slice(0, 3);
}

// ─── Export ───────────────────────────────────────────────────────────────
function exportToExcel(servers, databases, firewall) {
  const wb = XLSX.utils.book_new();

  if (servers.length) {
    const ws = XLSX.utils.json_to_sheet(servers);
    XLSX.utils.book_append_sheet(wb, ws, "Servers");
  }

  if (databases.length) {
    const ws = XLSX.utils.json_to_sheet(databases);
    XLSX.utils.book_append_sheet(wb, ws, "Databases");
  }

  if (firewall.length) {
    const agg = aggregateFirewall(firewall);
    const ws1 = XLSX.utils.json_to_sheet(agg);
    XLSX.utils.book_append_sheet(wb, ws1, "Firewall_Aggregated");

    const perSrv = perServerFirewall(firewall);
    const ws2 = XLSX.utils.json_to_sheet(perSrv);
    XLSX.utils.book_append_sheet(wb, ws2, "Firewall_Per_Server");
  }

  XLSX.writeFile(wb, "aaq_full_analysis.xlsx");
}

function aggregateFirewall(rows) {
  const key = r => [r.source_hostname, r.destination_hostname, r.source_application,
    r.destination_application, r.source_port, r.protocol].join("\n");

  const map = {};
  for (const r of rows) {
    const k = key(r);
    if (!map[k]) map[k] = { ...r, destination_port: new Set(), netstat_count: 0 };
    map[k].netstat_count += r.netstat_count;
    if (r.destination_port) map[k].destination_port.add(r.destination_port);
  }

  return Object.values(map).map(r => ({
    ...r,
    destination_port: [...r.destination_port].sort().join(", "),
    usage: classifyUsage(r.netstat_count),
  }));
}

function perServerFirewall(rows) {
  const key = r => [r.source_hostname, r.destination_application, r.source_port].join("\n");
  const map = {};

  for (const r of rows) {
    const k = key(r);
    if (!map[k]) map[k] = {
      server: r.source_hostname,
      destination_application: r.destination_application,
      source_port: r.source_port,
      netstat_count: 0,
      destination_ports: new Set(),
    };

    map[k].netstat_count += r.netstat_count;
    if (r.destination_port) map[k].destination_ports.add(r.destination_port);
  }

  return Object.values(map).map(r => ({
    server: r.server,
    destination_application: r.destination_application,
    source_port: r.source_port,
    destination_port: [...r.destination_ports].sort().join(", "),
    netstat_count: r.netstat_count,
    usage: classifyUsage(r.netstat_count),
  }));
}

// ─── Components ────────────────────────────────────────────────────────────
function UsageBadge({ usage }) {
  const cls = usage === "VERY LOW" ? "badge badge-VERYLOW"
    : usage === "HIGH" ? "badge badge-HIGH"
      : usage === "MEDIUM" ? "badge badge-MEDIUM"
        : usage === "LOW" ? "badge badge-LOW"
          : "badge badge-UNKNOWN";
  return <span className={cls}>{usage}</span>;
}

function DataTable({ columns, rows, emptyMsg = "No data" }) {
  if (!rows.length) return <div className="empty">{emptyMsg}</div>;
  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>{columns.map(c => <th key={c.key}>{c.label}</th>)}</tr>
        </thead>
        <tbody>
          {rows.map((row, i) => (
            <tr key={i}>
              {columns.map(c => (
                <td key={c.key} title={String(row[c.key] ?? "")}>
                  {c.render ? c.render(row[c.key], row) : (row[c.key] ?? "—")}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

// ─── Sections ──────────────────────────────────────────────────────────────
function HomeSection({ servers, databases, firewall, setNav }) {
  return (
    <>
      <div className="metrics">
        <div className="metric-card">
          <div className="metric-label">Servers</div>
          <div className="metric-value">{servers.length}</div>
        </div>
        <div className="metric-card">
          <div className="metric-label">Databases</div>
          <div className="metric-value">{databases.length}</div>
        </div>
        <div className="metric-card">
          <div className="metric-label">Firewall Records</div>
          <div className="metric-value">{firewall.length}</div>
        </div>
      </div>

      <div className="nav-cards">
        <div className="nav-card" onClick={() => setNav("sizing")}
        >
          <div className="nav-card-icon">💻</div>
          <div className="nav-card-title">Sizing</div>
          <div className="nav-card-sub">AWS EC2 recommendations per server</div>
        </div>
        <div className="nav-card" onClick={() => setNav("databases")}
        >
          <div className="nav-card-icon">🗄️</div>
          <div className="nav-card-title">Databases</div>
          <div className="nav-card-sub">Database inventory from AAQ</div>
        </div>
        <div className="nav-card" onClick={() => setNav("firewall")}
        >
          <div className="nav-card-icon">🔥</div>
          <div className="nav-card-title">Firewall</div>
          <div className="nav-card-sub">Aggregated network flow analysis</div>
        </div>
      </div>

      <hr className="divider" />
      <div className="section-title">Export <span>Full Analysis</span></div>
      <button className="export-btn" onClick={() => exportToExcel(servers, databases, firewall)}>
        📥 Download Full Analysis (.xlsx)
      </button>
    </>
  );
}

function SizingSection({ servers }) {
  if (!servers.length) return <div className="empty">No server data found in the uploaded file.</div>;
  return (
    <>
      <div className="section-title">AWS <span>Recommendations</span></div>
      <p style={{ fontSize: 12, color: "var(--muted)", marginBottom: 24 }}>
        Instance data: eu-west-2 on-demand pricing (static snapshot). Cheapest fit shown first.
      </p>
      {servers.map((s, i) => {
        const recs = recommendInstances(s.vcpu, s.memory_gib);
        return (
          <div className="server-card" key={i}>
            <div className="server-name">{s.server}</div>
            <div className="server-meta">
              {s.application} · {s.vcpu ?? "?"} vCPU / {s.memory_gib ?? "?"} GiB RAM
            </div>
            {recs.length ? (
              <div className="instance-grid">
                {recs.map((r, j) => (
                  <div className="instance-card" key={j}>
                    <div className="instance-type">{r.instance_type}</div>
                    <div className="instance-specs">{r.vCPU} vCPU · {r.memory} GiB RAM</div>
                    <div className="instance-price">${r.price_usd.toFixed(4)}/hr on-demand</div>
                  </div>
                ))}
              </div>
            ) : (
              <div style={{ fontSize: 12, color: "var(--muted)" }}>
                {s.vcpu && s.memory_gib ? "No matching instance found." : "Insufficient CPU/RAM data to size."}
              </div>
            )}
          </div>
        );
      })}
    </>
  );
}

function DatabasesSection({ databases }) {
  const [search, setSearch] = useState("");
  const filtered = databases.filter(d =>
    !search || Object.values(d).some(v => String(v ?? "").toLowerCase().includes(search.toLowerCase()))
  );

  const exportDb = () => {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(databases);
    XLSX.utils.book_append_sheet(wb, ws, "Databases");
    XLSX.writeFile(wb, "database_inventory.xlsx");
  };

  return (
    <>
      <div className="section-title">Database <span>Inventory</span></div>
      <div className="filter-row">
        <input
          className="filter-input"
          placeholder="Search databases…"
          value={search}
          onChange={e => setSearch(e.target.value)}
        />
      </div>
      <DataTable
        columns={[
          { key: "db_server", label: "DB Server" },
          { key: "database_name", label: "Database Name" },
          { key: "database_instance", label: "Instance" },
          { key: "database_size_gb", label: "Size (GB)", render: v => v != null ? `${v} GB` : "—" },
          { key: "database_type", label: "Type" },
        ]}
        rows={filtered}
        emptyMsg="No database records found."
      />
      {databases.length > 0 && (
        <button className="export-btn" onClick={exportDb}>📥 Download Database Inventory</button>
      )}
    </>
  );
}

function FirewallSection({ firewall }) {
  const [search, setSearch] = useState("");
  const [usageFilter, setUsageFilter] = useState("ALL");
  const [view, setView] = useState("aggregated");

  const agg = aggregateFirewall(firewall);
  const perSrv = perServerFirewall(firewall);

  const filterRows = rows => rows.filter(r => {
    const matchSearch = !search || Object.values(r).some(v => String(v ?? "").toLowerCase().includes(search.toLowerCase()));
    const matchUsage = usageFilter === "ALL" || r.usage === usageFilter;
    return matchSearch && matchUsage;
  });

  const filteredAgg = filterRows(agg);
  const filteredPerSrv = filterRows(perSrv);

  const exportFw = () => {
    const wb = XLSX.utils.book_new();
    const ws1 = XLSX.utils.json_to_sheet(agg);
    XLSX.utils.book_append_sheet(wb, ws1, "Firewall_Aggregated");
    const ws2 = XLSX.utils.json_to_sheet(perSrv);
    XLSX.utils.book_append_sheet(wb, ws2, "Firewall_Per_Server");
    XLSX.writeFile(wb, "firewall_analysis.xlsx");
  };

  return (
    <>
      <div className="section-title">Firewall <span>Analysis</span></div>

      <div className="filter-row">
        <input
          className="filter-input"
          placeholder="Search…"
          value={search}
          onChange={e => setSearch(e.target.value)}
        />
        <select className="filter-select" value={usageFilter} onChange={e => setUsageFilter(e.target.value)}>
          <option value="ALL">All Usage Levels</option>
          <option value="HIGH">HIGH</option>
          <option value="MEDIUM">MEDIUM</option>
          <option value="LOW">LOW</option>
          <option value="VERY LOW">VERY LOW</option>
        </select>
        <select className="filter-select" value={view} onChange={e => setView(e.target.value)}>
          <option value="aggregated">Aggregated View</option>
          <option value="perserver">Per Server View</option>
        </select>
      </div>

      {view === "aggregated" ? (
        <DataTable
          columns={[
            { key: "source_hostname", label: "Source Host" },
            { key: "destination_hostname", label: "Dest Host" },
            { key: "source_application", label: "Source App" },
            { key: "destination_application", label: "Dest App" },
            { key: "source_port", label: "Src Port" },
            { key: "destination_port", label: "Dst Port" },
            { key: "protocol", label: "Protocol" },
            { key: "netstat_count", label: "Count", render: v => v?.toLocaleString() },
            { key: "usage", label: "Usage", render: v => <UsageBadge usage={v} /> },
          ]}
          rows={filteredAgg.slice(0, 500)}
          emptyMsg="No firewall records found."
        />
      ) : (
        <DataTable
          columns={[
            { key: "server", label: "Server" },
            { key: "destination_application", label: "Dest App" },
            { key: "source_port", label: "Src Port" },
            { key: "destination_port", label: "Dst Port" },
            { key: "netstat_count", label: "Count", render: v => v?.toLocaleString() },
            { key: "usage", label: "Usage", render: v => <UsageBadge usage={v} /> },
          ]}
          rows={filteredPerSrv.slice(0, 500)}
          emptyMsg="No firewall records found."
        />
      )}

      {firewall.length > 0 && (
        <button className="export-btn" onClick={exportFw}>📥 Download Firewall Analysis</button>
      )}
    </>
  );
}

// ─── Main App ──────────────────────────────────────────────────────────────
export default function App() {
  const [result, setResult] = useState(null);
  const [nav, setNav] = useState("home");
  const [fileName, setFileName] = useState(null);
  const [drag, setDrag] = useState(false);
  const fileRef = useRef();

  
const [loading, setLoading] = useState(false);
const [error, setError] = useState(null);

const handleFile = useCallback((file) => {
  if (!file) return;
  setFileName(file.name);
  setLoading(true);
  setError(null);

  const reader = new FileReader();
  reader.onload = () => {
    const buffer = reader.result;

    const worker = new Worker();
    worker.onmessage = (evt) => {
      const { ok, result, error } = evt.data;
      if (ok) {
        setResult(result);
        setNav("home");
      } else {
        setError(error);
      }
      setLoading(false);
      worker.terminate();
    };

    // Transfer the ArrayBuffer (fast, avoids copying)
    worker.postMessage({ buffer }, [buffer]);
  };

  reader.onerror = () => {
    setError("Failed to read file.");
    setLoading(false);
  };

  reader.readAsArrayBuffer(file);
}, []);


  const onDrop = useCallback(e => {
    e.preventDefault();
    setDrag(false);
    const file = e.dataTransfer.files[0];
    if (file?.name.endsWith(".xlsx")) handleFile(file);
  }, [handleFile]);

  const NAVS = [
    { id: "home", label: "🏠 Home" },
    { id: "sizing", label: "💻 Sizing" },
    { id: "databases", label: "🗄️ Databases" },
    { id: "firewall", label: "🔥 Firewall" },
  ];

  return (
    <>
      <style>{CSS}</style>
      <div className="app">
        {/* Header */}
        <div className="header">
          <img src={logo} alt="Kyndryl" className="logo-img" />
          <div className="header-text">
            <h1>AAQ Analysis Tool</h1>
            <p>Cloud Migration Analysis Tool</p>
          </div>
          <div className="privacy-badge">
            <div className="privacy-dot" />
            100% Client-Side · No Data Uploaded
          </div>
        </div>

        {/* Upload */}
        {!result && (
          <div className="content">
            <div
              className={`upload-zone${drag ? " drag" : ""}`}
              onDragOver={e => { e.preventDefault(); setDrag(true); }}
              onDragLeave={() => setDrag(false)}
              onDrop={onDrop}
              onClick={() => fileRef.current.click()}
            >
              <div className="upload-icon">📊</div>
              <div className="upload-title">Upload AAQ Excel File</div>
              <div className="upload-sub">Drag & drop or click to browse · .xlsx files only</div>
              <button className="upload-btn" onClick={e => { e.stopPropagation(); fileRef.current.click(); }}>
                Select File
              </button>
              {fileName && <div className="file-name">✓ {fileName}</div>}
            </div>
            <input
              ref={fileRef}
              type="file"
              accept=".xlsx"
              style={{ display: "none" }}
              onChange={e => handleFile(e.target.files[0])}
            />
          </div>
        )}

        {/* Main content */}
        {result && (
          <>
            <div className="nav-bar">
              {NAVS.map(n => (
                <button
                  key={n.id}
                  className={`nav-btn${nav === n.id ? " active" : ""}`}
                  onClick={() => setNav(n.id)}
                >
                  {n.label}
                </button>
              ))}
              <button
                className="nav-btn"
                style={{ marginLeft: "auto", color: "var(--accent2)" }}
                onClick={() => { setResult(null); setFileName(null); }}
              >
                ✕ New File
              </button>
            </div>

            <div className="content">
              {nav === "home" && <HomeSection servers={result.servers} databases={result.databases} firewall={result.firewall} setNav={setNav} />}
              {nav === "sizing" && <SizingSection servers={result.servers} />}
              {nav === "databases" && <DatabasesSection databases={result.databases} />}
              {nav === "firewall" && <FirewallSection firewall={result.firewall} />}

              {/* Parser notes */}
              <div className="notes-box">
                <div className="section-title" style={{ fontSize: 14, marginBottom: 12 }}>
                  Parser <span>Notes</span>
                </div>
                {result.notes.map((n, i) => <div className="note-item" key={i}>{n}</div>)}
              </div>
            </div>
          </>
        )}
      </div>
    </>
  );
}
