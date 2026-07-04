import { useState, useCallback, useRef, useMemo } from "react";
import * as XLSX from "xlsx";
import logo from "./assets/kyndryl_logo.webp";
import Worker from "./workers/aaqWorker.js?worker";
import Login from "./Login.jsx";

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

  .status-line { margin-top: 14px; font-size: 12px; color: var(--muted); }
  .status-error { margin-top: 14px; font-size: 12px; color: var(--accent2); }

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
  .metrics { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 16px; margin-bottom: 32px; }

  .metric-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 24px 28px;
    position: relative;
    overflow: hidden;
    box-shadow: 0 10px 40px rgba(0,0,0,0.4);
  }

  .metric-card::after { content:''; position:absolute; top:0; left:0; right:0; height:2px; }
  .metric-card:nth-child(1)::after { background: var(--accent); }
  .metric-card:nth-child(2)::after { background: var(--accent3); }
  .metric-card:nth-child(3)::after { background: var(--accent2); }
  .metric-card:nth-child(4)::after { background: rgba(255,255,255,.2); }
  .metric-card:nth-child(5)::after { background: rgba(255,255,255,.15); }

  .metric-label { font-size: 10px; color: var(--muted); letter-spacing: 0.5px; text-transform: uppercase; }

  .metric-value {
    font-family: var(--font-head);
    font-size: 32px;
    font-weight: 600;
    letter-spacing: -0.5px;
    color: var(--text);
    margin-top: 6px;
  }

  /* ── Nav buttons on home ── */
  .nav-cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(240px,1fr)); gap: 16px; margin-bottom: 32px; }

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
    white-space: nowrap; max-width: 260px;
    overflow: hidden; text-overflow: ellipsis;
  }

  tr:last-child td { border-bottom: none; }
  tr:hover td { background: rgba(255,255,255,.02); }

  /* ── Export button ── */
  .export-btn {
    display: inline-flex; align-items: center; gap: 8px;
    background: linear-gradient(135deg, var(--accent), var(--accent3));
    color: #000; font-family: var(--font-head); font-weight: 700;
    font-size: 13px; letter-spacing: .3px;
    padding: 11px 24px; border-radius: 8px; border: none; cursor: pointer;
    transition: var(--transition); margin-top: 0;
  }

  .export-btn:hover { transform: translateY(-1px); box-shadow: 0 6px 24px rgba(0,229,255,.3); }

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
  .filter-row { display: flex; gap: 12px; margin-bottom: 16px; flex-wrap: wrap; align-items: center; }

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
  if (Number.isNaN(c)) return "UNKNOWN";
  if (c >= 50000) return "HIGH";
  if (c >= 1000) return "MEDIUM";
  if (c >= 100) return "LOW";
  return "VERY LOW";
}

function UsageBadge({ usage }) {
  const cls = usage === "VERY LOW" ? "badge badge-VERYLOW"
    : usage === "HIGH" ? "badge badge-HIGH"
      : usage === "MEDIUM" ? "badge badge-MEDIUM"
        : usage === "LOW" ? "badge badge-LOW"
          : "badge badge-UNKNOWN";
  return <span className={cls}>{usage}</span>;
}

function buildSizingExport(servers) {
  const rows = [];
  servers.forEach(server => {
    const recs = recommendInstances(server.vcpu, server.memory_gib);
    recs.forEach(rec => {
      rows.push({
        server: server.server,
        application: server.application,
        required_vcpu: server.vcpu,
        required_memory_gib: server.memory_gib,
        instance_type: rec.instance_type,
        instance_vcpu: rec.vCPU,
        instance_memory_gib: rec.memory,
        price_usd_per_hr: rec.price_usd,
      });
    });
  });
  return rows;
}

// Deduplicate raw firewall rows by a stable key and sum counts.
function dedupeFirewallRows(rows) {
  const map = new Map();
  const norm = (v) => String(v ?? "").trim();

  for (const r of rows || []) {
    const key = [
      norm(r.source_hostname),
      norm(r.destination_hostname),
      norm(r.source_application),
      norm(r.destination_application),
      norm(r.source_port),
      norm(r.destination_port),
      norm(r.protocol),
    ].join("\n");

    const count = Number(r.netstat_count ?? 0) || 0;

    if (!map.has(key)) {
      map.set(key, {
        ...r,
        netstat_count: count,
      });
    } else {
      const existing = map.get(key);
      existing.netstat_count = (Number(existing.netstat_count ?? 0) || 0) + count;
      // prefer non-empty fields
      for (const k of ["source_ip","destination_ip","source_hostname","destination_hostname","source_application","destination_application","source_port","destination_port","protocol"]) {
        if (!existing[k] && r[k]) existing[k] = r[k];
      }
    }
  }

  // Recompute usage classification on the deduped totals
  return Array.from(map.values()).map(r => ({
    ...r,
    usage: classifyUsage(r.netstat_count),
  }));
}

// Aggregate firewall by endpoint/app/port/protocol and merge destination ports.
function aggregateFirewall(rows) {
  const map = new Map();
  const norm = (v) => String(v ?? "").trim();

  for (const r of rows || []) {
    const key = [
      norm(r.source_hostname),
      norm(r.destination_hostname),
      norm(r.source_application),
      norm(r.destination_application),
      norm(r.source_port),
      norm(r.protocol),
    ].join("\n");

    if (!map.has(key)) {
      map.set(key, {
        source_hostname: r.source_hostname ?? null,
        destination_hostname: r.destination_hostname ?? null,
        source_application: r.source_application ?? null,
        destination_application: r.destination_application ?? null,
        source_port: r.source_port ?? null,
        destination_port: new Set(),
        protocol: r.protocol ?? null,
        netstat_count: 0,
      });
    }

    const entry = map.get(key);
    entry.netstat_count += Number(r.netstat_count ?? 0) || 0;
    if (r.destination_port) entry.destination_port.add(String(r.destination_port).trim());
  }

  return Array.from(map.values()).map(r => ({
    ...r,
    destination_port: Array.from(r.destination_port).sort().join(", "),
    usage: classifyUsage(r.netstat_count),
  }));
}

// Per-server view: group by source host + destination app + source port
function perServerFirewall(rows) {
  const map = new Map();
  const norm = (v) => String(v ?? "").trim();

  for (const r of rows || []) {
    const key = [
      norm(r.source_hostname),
      norm(r.destination_application),
      norm(r.source_port),
      norm(r.protocol),
    ].join("\n");

    if (!map.has(key)) {
      map.set(key, {
        server: r.source_hostname ?? null,
        destination_application: r.destination_application ?? null,
        source_port: r.source_port ?? null,
        protocol: r.protocol ?? null,
        destination_port: new Set(),
        netstat_count: 0,
      });
    }

    const entry = map.get(key);
    entry.netstat_count += Number(r.netstat_count ?? 0) || 0;
    if (r.destination_port) entry.destination_port.add(String(r.destination_port).trim());
  }

  return Array.from(map.values()).map(r => ({
    server: r.server,
    destination_application: r.destination_application,
    source_port: r.source_port,
    protocol: r.protocol,
    destination_port: Array.from(r.destination_port).sort().join(", "),
    netstat_count: r.netstat_count,
    usage: classifyUsage(r.netstat_count),
  }));
}

// ─── AWS Sizing (static fallback data) ─────────────────────────────────────
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

// ─── Export helpers ───────────────────────────────────────────────────────
function exportSheet(data, sheetName, fileName) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(data || []), sheetName);
  XLSX.writeFile(wb, fileName);
}

function exportFullAnalysis(servers, databases, firewallRaw, stakeholders, appDrBackup) {
  const wb = XLSX.utils.book_new();

  if (servers?.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(servers), "Servers");
  if (databases?.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(databases), "Databases");

  // Firewall: include deduped raw + aggregated + per-server
  const fwDeduped = dedupeFirewallRows(firewallRaw || []);
  if (fwDeduped.length) {
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(fwDeduped), "Firewall_Raw_Deduped");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(aggregateFirewall(fwDeduped)), "Firewall_Aggregated");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(perServerFirewall(fwDeduped)), "Firewall_Per_Server");
  }

  if (servers?.length) {
    const sizing = buildSizingExport(servers);
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sizing), "Sizing");
  }

  if (stakeholders?.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(stakeholders), "Stakeholders");
  if (appDrBackup?.length) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(appDrBackup), "Applications_DR_Backup");

  XLSX.writeFile(wb, "aaq_full_analysis.xlsx");
}

// ─── Components ────────────────────────────────────────────────────────────
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
function HomeSection({ servers, databases, firewall, stakeholders, appDrBackup, setNav }) {
  const fwDeduped = useMemo(() => dedupeFirewallRows(firewall), [firewall]);

  return (
    <>
      <div className="metrics">
        <div className="metric-card"><div className="metric-label">Servers</div><div className="metric-value">{servers.length}</div></div>
        <div className="metric-card"><div className="metric-label">Databases</div><div className="metric-value">{databases.length}</div></div>
        <div className="metric-card"><div className="metric-label">Firewall (deduped)</div><div className="metric-value">{fwDeduped.length}</div></div>
        <div className="metric-card"><div className="metric-label">Stakeholders</div><div className="metric-value">{stakeholders.length}</div></div>
        <div className="metric-card"><div className="metric-label">Apps (DR/Backup)</div><div className="metric-value">{appDrBackup.length}</div></div>
      </div>

      <div className="nav-cards">
        <div className="nav-card" onClick={() => setNav("sizing")}>
          <div className="nav-card-icon">💻</div>
          <div className="nav-card-title">Sizing</div>
          <div className="nav-card-sub">AWS EC2 recommendations per server</div>
        </div>
        <div className="nav-card" onClick={() => setNav("databases")}>
          <div className="nav-card-icon">🗄️</div>
          <div className="nav-card-title">Databases</div>
          <div className="nav-card-sub">Database inventory from AAQ</div>
        </div>
        <div className="nav-card" onClick={() => setNav("firewall")}>
          <div className="nav-card-icon">🔥</div>
          <div className="nav-card-title">Firewall</div>
          <div className="nav-card-sub">Aggregated + per-server views (deduped)</div>
        </div>
        <div className="nav-card" onClick={() => setNav("stakeholders")}>
          <div className="nav-card-icon">👥</div>
          <div className="nav-card-title">Stakeholders</div>
          <div className="nav-card-sub">Contacts from Application Stakeholder sheet</div>
        </div>
        <div className="nav-card" onClick={() => setNav("apps")}>
          <div className="nav-card-icon">📦</div>
          <div className="nav-card-title">Apps DR/Backup</div>
          <div className="nav-card-sub">RPO/RTO, DR & backup fields</div>
        </div>
      </div>

      <hr className="divider" />
      <div className="filter-row" style={{ justifyContent: "space-between" }}>
        <div className="section-title" style={{ marginBottom: 0 }}>Export <span>Full Analysis</span></div>
        <button className="export-btn" onClick={() => exportFullAnalysis(servers, databases, firewall, stakeholders, appDrBackup)}>
          📥 Download Full Analysis (.xlsx)
        </button>
      </div>
    </>
  );
}

function SizingSection({ servers }) {
  const rows = useMemo(() => buildSizingExport(servers), [servers]);

  return (
    <>
      <div className="filter-row" style={{ justifyContent: "space-between" }}>
        <div className="section-title" style={{ marginBottom: 0 }}>AWS <span>Recommendations</span></div>
        <button className="export-btn" onClick={() => exportSheet(rows, "Sizing", "sizing_recommendations.xlsx")}>
          📥 Export Sizing
        </button>
      </div>

      <DataTable
        columns={[
          { key: "server", label: "Server" },
          { key: "application", label: "Application" },
          { key: "required_vcpu", label: "Required vCPU" },
          { key: "required_memory_gib", label: "Required RAM (GiB)" },
          { key: "instance_type", label: "Instance" },
          { key: "instance_vcpu", label: "vCPU" },
          { key: "instance_memory_gib", label: "Memory" },
          { key: "price_usd_per_hr", label: "Price/hr", render: v => v != null ? `$${Number(v).toFixed(4)}` : "—" },
        ]}
        rows={rows}
        emptyMsg="No server data found in the uploaded file."
      />
    </>
  );
}

function DatabasesSection({ databases }) {
  const [search, setSearch] = useState("");
  const filtered = databases.filter(d =>
    !search || Object.values(d).some(v => String(v ?? "").toLowerCase().includes(search.toLowerCase()))
  );

  return (
    <>
      <div className="filter-row" style={{ justifyContent: "space-between" }}>
        <div className="section-title" style={{ marginBottom: 0 }}>Database <span>Inventory</span></div>
        <button className="export-btn" onClick={() => exportSheet(databases, "Databases", "database_inventory.xlsx")}>
          📥 Export
        </button>
      </div>

      <div className="filter-row">
        <input className="filter-input" placeholder="Search databases…" value={search} onChange={e => setSearch(e.target.value)} />
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
    </>
  );
}

function FirewallSection({ firewall }) {
  const [search, setSearch] = useState("");
  const [usageFilter, setUsageFilter] = useState("ALL");
  const [view, setView] = useState("aggregated");

  const deduped = useMemo(() => dedupeFirewallRows(firewall), [firewall]);
  const agg = useMemo(() => aggregateFirewall(deduped), [deduped]);
  const perSrv = useMemo(() => perServerFirewall(deduped), [deduped]);

  const filterRows = (rows) => rows.filter(r => {
    const matchSearch = !search || Object.values(r).some(v => String(v ?? "").toLowerCase().includes(search.toLowerCase()));
    const matchUsage = usageFilter === "ALL" || r.usage === usageFilter;
    return matchSearch && matchUsage;
  });

  const filteredAgg = useMemo(() => filterRows(agg), [agg, search, usageFilter]);
  const filteredPerSrv = useMemo(() => filterRows(perSrv), [perSrv, search, usageFilter]);

  const exportFw = () => {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(aggregateFirewall(deduped)), "Firewall_Aggregated");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(perServerFirewall(deduped)), "Firewall_Per_Server");
    XLSX.writeFile(wb, "firewall_analysis.xlsx");
  };

  return (
    <>
      <div className="filter-row" style={{ justifyContent: "space-between" }}>
        <div className="section-title" style={{ marginBottom: 0 }}>Firewall <span>Analysis</span></div>
        <button className="export-btn" onClick={exportFw}>📥 Export Firewall</button>
      </div>

      <div className="filter-row">
        <input className="filter-input" placeholder="Search…" value={search} onChange={e => setSearch(e.target.value)} />
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
            { key: "destination_port", label: "Dst Port(s)" },
            { key: "protocol", label: "Protocol" },
            { key: "netstat_count", label: "Count", render: v => (v?.toLocaleString ? v.toLocaleString() : v) },
            { key: "usage", label: "Usage", render: v => <UsageBadge usage={v} /> },
          ]}
          rows={filteredAgg.slice(0, 500)}
          emptyMsg="No aggregated firewall records found."
        />
      ) : (
        <DataTable
          columns={[
            { key: "server", label: "Server" },
            { key: "destination_application", label: "Dest App" },
            { key: "source_port", label: "Src Port" },
            { key: "destination_port", label: "Dst Port(s)" },
            { key: "protocol", label: "Protocol" },
            { key: "netstat_count", label: "Count", render: v => (v?.toLocaleString ? v.toLocaleString() : v) },
            { key: "usage", label: "Usage", render: v => <UsageBadge usage={v} /> },
          ]}
          rows={filteredPerSrv.slice(0, 500)}
          emptyMsg="No per-server firewall records found."
        />
      )}

      <div style={{ marginTop: 10, color: "var(--muted)", fontSize: 11 }}>
        Showing deduplicated firewall data (duplicates merged; counts summed).
      </div>
    </>
  );
}

function StakeholdersSection({ stakeholders }) {
  const [search, setSearch] = useState("");
  const filtered = stakeholders.filter(s =>
    !search || Object.values(s).some(v => String(v ?? "").toLowerCase().includes(search.toLowerCase()))
  );

  return (
    <>
      <div className="filter-row" style={{ justifyContent: "space-between" }}>
        <div className="section-title" style={{ marginBottom: 0 }}>Application <span>Stakeholders</span></div>
        <button className="export-btn" onClick={() => exportSheet(stakeholders, "Stakeholders", "stakeholders.xlsx")}>📥 Export</button>
      </div>

      <div className="filter-row">
        <input className="filter-input" placeholder="Search stakeholders…" value={search} onChange={e => setSearch(e.target.value)} />
      </div>

      <DataTable
        columns={[
          { key: "application", label: "Application" },
          { key: "detail", label: "Role / Detail" },
          { key: "contact_name", label: "Contact Name" },
          { key: "email", label: "Email" },
          { key: "contact_number", label: "Contact Number" },
        ]}
        rows={filtered}
        emptyMsg="No stakeholder records found."
      />
    </>
  );
}

function AppsDRBackupSection({ appDrBackup }) {
  const [search, setSearch] = useState("");
  const filtered = appDrBackup.filter(a =>
    !search || Object.values(a).some(v => String(v ?? "").toLowerCase().includes(search.toLowerCase()))
  );

  return (
    <>
      <div className="filter-row" style={{ justifyContent: "space-between" }}>
        <div className="section-title" style={{ marginBottom: 0 }}>Applications <span>DR & Backup</span></div>
        <button className="export-btn" onClick={() => exportSheet(appDrBackup, "Applications_DR_Backup", "applications_dr_backup.xlsx")}>📥 Export</button>
      </div>

      <div className="filter-row">
        <input className="filter-input" placeholder="Search…" value={search} onChange={e => setSearch(e.target.value)} />
      </div>

      <DataTable
        columns={[
          { key: "sheet", label: "Sheet" },
          { key: "app_id", label: "App Id" },
          { key: "application_name", label: "Application Name" },
          { key: "rpo", label: "RPO" },
          { key: "rto", label: "RTO" },
          { key: "backup_size", label: "Backup Size" },
          { key: "backup_retention", label: "Backup Retention" },
          { key: "dr_in_place", label: "DR In Place" },
          { key: "dr_description", label: "DR Description" },
          { key: "dr_plans_exist", label: "DR Plans Exist" },
          { key: "last_dr_test", label: "Last DR Test" },
        ]}
        rows={filtered}
        emptyMsg="No application DR/Backup records found."
      />
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
        const { ok, result: parsed, error: workerError } = evt.data;
        if (ok) {
          setResult({
            servers: parsed?.servers ?? [],
            databases: parsed?.databases ?? [],
            firewall: parsed?.firewall ?? [],
            stakeholders: parsed?.stakeholders ?? [],
            app_dr_backup: parsed?.app_dr_backup ?? [],
            notes: parsed?.notes ?? [],
          });
          setNav("home");
        } else {
          setError(workerError || "Failed to parse file.");
        }
        setLoading(false);
        worker.terminate();
      };

      worker.postMessage({ buffer }, [buffer]);
    };

    reader.onerror = () => {
      setError("Failed to read file.");
      setLoading(false);
    };

    reader.readAsArrayBuffer(file);
  }, []);

  const onDrop = useCallback((e) => {
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
    { id: "stakeholders", label: "👥 Stakeholders" },
    { id: "apps", label: "📦 Apps DR/Backup" },
  ];

return (
  <>
    <style>{CSS}</style>

    <Login />

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
              onDragOver={(e) => { e.preventDefault(); setDrag(true); }}
              onDragLeave={() => setDrag(false)}
              onDrop={onDrop}
              onClick={() => fileRef.current.click()}
            >
              <div className="upload-icon">📊</div>
              <div className="upload-title">Upload AAQ Excel File</div>
              <div className="upload-sub">Drag & drop or click to browse · .xlsx files only</div>
              <button className="upload-btn" onClick={(e) => { e.stopPropagation(); fileRef.current.click(); }}>
                Select File
              </button>
              {fileName && <div className="file-name">✓ {fileName}</div>}
              {loading && <div className="status-line">Processing… large files may take a moment.</div>}
              {error && <div className="status-error">Error: {error}</div>}
            </div>
            <input
              ref={fileRef}
              type="file"
              accept=".xlsx"
              style={{ display: "none" }}
              onChange={(e) => handleFile(e.target.files[0])}
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
                onClick={() => { setResult(null); setFileName(null); setError(null); setLoading(false); }}
              >
                ✕ New File
              </button>
            </div>

            <div className="content">
              {nav === "home" && (
                <HomeSection
                  servers={result.servers}
                  databases={result.databases}
                  firewall={result.firewall}
                  stakeholders={result.stakeholders}
                  appDrBackup={result.app_dr_backup}
                  setNav={setNav}
                />
              )}
              {nav === "sizing" && <SizingSection servers={result.servers} />}
              {nav === "databases" && <DatabasesSection databases={result.databases} />}
              {nav === "firewall" && <FirewallSection firewall={result.firewall} />}
              {nav === "stakeholders" && <StakeholdersSection stakeholders={result.stakeholders} />}
              {nav === "apps" && <AppsDRBackupSection appDrBackup={result.app_dr_backup} />}

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
