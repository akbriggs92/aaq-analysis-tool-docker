import * as XLSX from "xlsx";

// Compute actual used range (all populated cells) and override ws['!ref'].
function normalizeRef(ws) {
  const cellRegex = /^[A-Z]+[0-9]+$/;
  let minR = Infinity, minC = Infinity, maxR = -1, maxC = -1;

  for (const addr of Object.keys(ws)) {
    if (!cellRegex.test(addr)) continue;
    const cell = ws[addr];
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

function classifyUsage(count) {
  const c = parseInt(count, 10);
  if (Number.isNaN(c)) return "UNKNOWN";
  if (c >= 50000) return "HIGH";
  if (c >= 1000) return "MEDIUM";
  if (c >= 100) return "LOW";
  return "VERY LOW";
}

function safeNum(v) {
  if (v == null || v === "") return null;
  const s = String(v).replace(/[^0-9.]/g, "");
  const n = parseFloat(s);
  return Number.isNaN(n) ? null : n;
}

function findCol(cols, ...keys) {
  const norm = cols.map(c => c.replace(/[\s_]/g, "").toLowerCase());
  for (const key of keys) {
    const idx = norm.findIndex(c => c.includes(key.toLowerCase().replace(/[\s_]/g, "")));
    if (idx !== -1) return cols[idx];
  }
  return null;
}

function parseStakeholdersFromSheet(ws) {
  const out = [];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

  let currentApp = null;
  let headerMap = null;

  const norm = (s) => String(s || "").trim().toLowerCase();

  for (let i = 0; i < raw.length; i++) {
    const row = raw[i] || [];
    const rowText = row.map(v => String(v || "").trim()).filter(Boolean).join(" ");

    if (/^STAR:/i.test(rowText)) {
      currentApp = rowText.trim();
      headerMap = null;
      continue;
    }

    const cols = row.map(v => String(v || "").trim());
    const colsNorm = cols.map(norm);

    const detailsIdx = colsNorm.findIndex(c => c === "details");
    const nameIdx = colsNorm.findIndex(c => c.includes("contact") && c.includes("name"));
    const emailIdx = colsNorm.findIndex(c => c.includes("e-mail") || c.includes("email"));
    const phoneIdx = colsNorm.findIndex(c => c.includes("contact") && c.includes("number"));

    if (detailsIdx !== -1 && nameIdx !== -1) {
      headerMap = { detailsIdx, nameIdx, emailIdx, phoneIdx };
      continue;
    }

    if (currentApp && headerMap) {
      const detail = row[headerMap.detailsIdx] ?? "";
      const contact = row[headerMap.nameIdx] ?? "";
      if (!String(detail).trim() && !String(contact).trim()) continue;
      if (!String(detail).trim()) continue;

      out.push({
        application: currentApp,
        detail: String(detail).trim(),
        contact_name: String(contact).trim(),
        email: headerMap.emailIdx !== -1 ? String(row[headerMap.emailIdx] || "").trim() : "",
        contact_number: headerMap.phoneIdx !== -1 ? String(row[headerMap.phoneIdx] || "").trim() : "",
      });
    }
  }

  return out;
}

// --- Application DR/Backup header detection ---
// Many AAQ sheets have a "question" row and then a "label" row (e.g. yellow header)
// We score candidates and choose the best match.
function detectAppHeaderRow(raw, maxScan = 200) {
  const norm = (s) => String(s || "").toLowerCase().replace(/\s+/g, " ").trim();

  const preferred = [
    { token: "application name", weight: 5 },
    { token: "app id", weight: 2 },
    { token: "rpo", weight: 4 },
    { token: "rto", weight: 4 },
    { token: "backup retention", weight: 4 },
    { token: "full backup size", weight: 4 },
    { token: "backup size", weight: 3 },
    { token: "is dr solution currently in place", weight: 5 },
    { token: "dr solution currently in place", weight: 4 },
    { token: "dr description", weight: 4 },
    { token: "do dr plans exist", weight: 4 },
    { token: "dr plans exist", weight: 3 },
    { token: "last dr test", weight: 4 },
  ];

  const weak = [
    { token: "recovery point objective", weight: 1 },
    { token: "recovery time objective", weight: 1 },
    { token: "size of backup", weight: 1 },
    { token: "backup retention duration", weight: 1 },
  ];

  let bestIdx = -1;
  let bestScore = -1;

  for (let i = 0; i < Math.min(raw.length, maxScan); i++) {
    const row = raw[i] || [];
    const headers = row.map(norm);

    // Must have app name somewhere in row
    if (!headers.some(h => h.includes("application name"))) continue;

    let score = 0;
    for (const p of preferred) if (headers.some(h => h.includes(p.token))) score += p.weight;
    for (const w of weak) if (headers.some(h => h.includes(w.token))) score += w.weight;

    // Prefer rows that have literal RPO/RTO (the label header row)
    const hasLiteralRpoRto = headers.some(h => h === "rpo" || h.includes(" rpo")) && headers.some(h => h === "rto" || h.includes(" rto"));
    if (hasLiteralRpoRto) score += 3;

    if (score > bestScore) {
      bestScore = score;
      bestIdx = i;
    }
  }

  // Require at least a meaningful score beyond just application name
  if (bestScore < 7) return -1;
  return bestIdx;
}

function parseAppDRBackupSheet(ws, sheetName) {
  const out = [];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (!raw.length) return out;

  const headerRowIdx = detectAppHeaderRow(raw);
  if (headerRowIdx === -1) return out;

  const headers = raw[headerRowIdx].map(h => String(h || "").trim().toLowerCase().replace(/\s+/g, " "));

  const findIdx = (...names) => {
    const tokens = names.map(n => n.toLowerCase());
    return headers.findIndex(h => tokens.some(t => h.includes(t)));
  };

  const idxAppId = findIdx("app id", "application id");
  const idxAppName = findIdx("application name");

  const idxRPO = findIdx("rpo", "recovery point objective");
  const idxRTO = findIdx("rto", "recovery time objective");

  const idxBackupSize = findIdx("full backup size", "backup size", "size of backup");
  const idxBackupRetention = findIdx("backup retention", "backup retention duration");

  const idxDRInPlace = findIdx(
    "is dr solution currently in place",
    "dr solution currently in place",
    "dr solution",
    "dr in place",
    "disaster recovery plan in place"
  );

  const idxDRDesc = findIdx("dr description", "description of disaster recovery");
  const idxDRPlans = findIdx("do dr plans exist", "dr plans exist", "dr plans");
  const idxLastDRTest = findIdx("last dr test", "recent test date");

  for (let r = headerRowIdx + 1; r < raw.length; r++) {
    const row = raw[r] || [];

    const appName = idxAppName !== -1 ? String(row[idxAppName] || "").trim() : "";
    const appId = idxAppId !== -1 ? String(row[idxAppId] || "").trim() : "";

    // Skip rows with no app identifier
    if (!appName && !appId) continue;

    const rec = {
      sheet: sheetName,
      app_id: appId || null,
      application_name: appName || sheetName,
      rpo: idxRPO !== -1 ? String(row[idxRPO] || "").trim() : "",
      rto: idxRTO !== -1 ? String(row[idxRTO] || "").trim() : "",
      backup_size: idxBackupSize !== -1 ? String(row[idxBackupSize] || "").trim() : "",
      backup_retention: idxBackupRetention !== -1 ? String(row[idxBackupRetention] || "").trim() : "",
      dr_in_place: idxDRInPlace !== -1 ? String(row[idxDRInPlace] || "").trim() : "",
      dr_description: idxDRDesc !== -1 ? String(row[idxDRDesc] || "").trim() : "",
      dr_plans_exist: idxDRPlans !== -1 ? String(row[idxDRPlans] || "").trim() : "",
      last_dr_test: idxLastDRTest !== -1 ? String(row[idxLastDRTest] || "").trim() : "",
    };

    // Only keep if at least one of the fields is populated (avoid random matches)
    const hasAny = [rec.rpo, rec.rto, rec.backup_size, rec.backup_retention, rec.dr_in_place, rec.dr_description, rec.dr_plans_exist, rec.last_dr_test]
      .some(v => String(v || "").trim() !== "");

    if (hasAny) out.push(rec);
  }

  return out;
}

function isLikelyAppSheet(ws) {
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  return detectAppHeaderRow(raw) !== -1;
}

function parseAAQ(workbook) {
  const notes = [];
  const sheetNames = workbook.SheetNames;

  const findSheet = (...patterns) =>
    sheetNames.find(s => patterns.some(p => s.toLowerCase().includes(p.toLowerCase())));

  notes.push(`Sheets found: ${sheetNames.join(", ")}`);

  // ── Servers ──
  let servers = [];
  const serverSheet = findSheet("server");
  if (!serverSheet) {
    notes.push("Server sheet not found");
  } else {
    const ws = workbook.Sheets[serverSheet];
    normalizeRef(ws);
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, blankrows: false });

    let headerIdx = -1;
    for (let i = 0; i < raw.length; i++) {
      const rowText = raw[i].join(" ").toLowerCase().replace(/\s/g, "");
      if (rowText.includes("workloadname") || rowText.includes("cpucount")) { headerIdx = i; break; }
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
        headers.forEach((h, i) => { row[h] = raw[r]?.[i]; });
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
  const fwSheet = findSheet("firewall");
  if (!fwSheet) {
    notes.push("Firewall sheet not found");
  } else {
    const ws = workbook.Sheets[fwSheet];
    normalizeRef(ws);
    const rows = XLSX.utils.sheet_to_json(ws, { defval: null, blankrows: false });

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

  // ── Databases ──
  let databases = [];
  const dbSheet = findSheet("database");
  if (!dbSheet) {
    notes.push("Database sheet not found");
  } else {
    const ws = workbook.Sheets[dbSheet];
    normalizeRef(ws);
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, blankrows: false });

    let headerIdx = -1;
    for (let i = 0; i < raw.length; i++) {
      const txt = raw[i].join(" ").toLowerCase();
      if (txt.includes("database name") && txt.includes("db server name")) { headerIdx = i; break; }
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
        headers.forEach((h, i) => { row[h] = raw[r]?.[i]; });

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

  // ── Stakeholders ──
  let stakeholders = [];
  const stakeholderSheet = sheetNames.find(s => s.toLowerCase().includes("application stakeholder"));
  if (!stakeholderSheet) {
    notes.push("Application Stakeholder sheet not found");
  } else {
    const ws = workbook.Sheets[stakeholderSheet];
    normalizeRef(ws);
    stakeholders = parseStakeholdersFromSheet(ws);
    notes.push(`Extracted ${stakeholders.length} stakeholder rows`);
  }

  // ── App DR/Backup across all application sheets ──
  const app_dr_backup = [];
  for (const name of sheetNames) {
    const lname = name.toLowerCase();
    if (["server", "database", "firewall", "application stakeholder", "logs"].some(t => lname.includes(t))) continue;

    const ws = workbook.Sheets[name];
    try {
      normalizeRef(ws);
      if (!isLikelyAppSheet(ws)) continue;

      const rows = parseAppDRBackupSheet(ws, name);
      if (rows.length) {
        app_dr_backup.push(...rows);
        notes.push(`Extracted ${rows.length} app DR/Backup rows from ${name}`);
      }
    } catch (e) {
      notes.push(`Skipped app parse for ${name}`);
    }
  }

  return { servers, firewall, databases, stakeholders, app_dr_backup, notes };
}

self.onmessage = (e) => {
  const { buffer } = e.data;
  try {
    const wb = XLSX.read(new Uint8Array(buffer), { type: "array" });
    const parsed = parseAAQ(wb);
    self.postMessage({ ok: true, result: parsed });
  } catch (err) {
    self.postMessage({ ok: false, error: String(err?.message || err) });
  }
};
