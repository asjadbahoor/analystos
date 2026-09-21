import React, { useState, useMemo, useCallback, useRef, useEffect, createContext, useContext } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  ScatterChart, Scatter,
} from "recharts";
import {
  UploadCloud, Database, Sparkles, MessageSquare, BarChart3, Trash2,
  Download, AlertTriangle, CheckCircle2, Filter, Activity,
  Loader2, Send, FileWarning, ChevronRight, ChevronUp, Table2,
  Sun, Moon, Plus, X, History as HistoryIcon, RotateCcw,
} from "lucide-react";

// ---------------------------------------------------------- theme palettes
const dark = {
  bg: "#0A0D12", surface: "#12161F", surface2: "#1A2029", border: "#262E3A",
  borderLight: "#333D4B", text: "#E7ECF3", muted: "#8592A3", faint: "#57626F",
  signal: "#4FE3C1", signalDim: "#2A6E60", onSignal: "#04211C",
  chatUserBg: "#2A6E60", chatUserText: "#EAFFFA",
  amber: "#F0A94E", danger: "#F0684F", violet: "#8B93F8", name: "dark",
};
const light = {
  bg: "#F6F7F9", surface: "#FFFFFF", surface2: "#F0F2F5", border: "#E3E7ED",
  borderLight: "#CBD2DC", text: "#171B21", muted: "#5B6472", faint: "#98A1AC",
  signal: "#0E9C82", signalDim: "#D6F3EC", onSignal: "#FFFFFF",
  chatUserBg: "#0E9C82", chatUserText: "#FFFFFF",
  amber: "#B8720A", danger: "#C43F2E", violet: "#5B62D6", name: "light",
};
const ThemeCtx = createContext(dark);
function useT() { return useContext(ThemeCtx); }
function rgba(hex, a) {
  const h = hex.replace("#", "");
  const r = parseInt(h.substring(0, 2), 16), g = parseInt(h.substring(2, 4), 16), b = parseInt(h.substring(4, 6), 16);
  return `rgba(${r},${g},${b},${a})`;
}

// ---------------- helpers: type inference & stats ----------------
function toNum(v) {
  if (v === null || v === undefined || v === "") return NaN;
  const n = parseFloat(v);
  return isFinite(n) ? n : NaN;
}
function inferType(values) {
  const nonEmpty = values.filter((v) => v !== null && v !== undefined && String(v).trim() !== "");
  if (nonEmpty.length === 0) return "empty";
  const numCount = nonEmpty.filter((v) => !isNaN(toNum(v))).length;
  if (numCount / nonEmpty.length > 0.9) return "numeric";
  const dateCount = nonEmpty.filter((v) => !isNaN(Date.parse(v)) && isNaN(toNum(v))).length;
  if (dateCount / nonEmpty.length > 0.9) return "date";
  const uniq = new Set(nonEmpty.map(String)).size;
  if (uniq <= 30 || uniq / nonEmpty.length < 0.5) return "categorical";
  return "text";
}
function mean(arr) { return arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : NaN; }
function median(arr) {
  if (!arr.length) return NaN;
  const s = [...arr].sort((a, b) => a - b);
  const mid = Math.floor(s.length / 2);
  return s.length % 2 ? s[mid] : (s[mid - 1] + s[mid]) / 2;
}
function std(arr) {
  if (arr.length < 2) return NaN;
  const m = mean(arr);
  return Math.sqrt(arr.reduce((a, b) => a + (b - m) ** 2, 0) / (arr.length - 1));
}
function mode(arr) {
  const counts = new Map();
  arr.forEach((v) => counts.set(v, (counts.get(v) || 0) + 1));
  let best = null, bestC = -1;
  counts.forEach((c, v) => { if (c > bestC) { best = v; bestC = c; } });
  return best;
}
function pearson(xs, ys) {
  const n = xs.length;
  if (n < 2) return NaN;
  const mx = mean(xs), my = mean(ys);
  let num = 0, dx = 0, dy = 0;
  for (let i = 0; i < n; i++) {
    num += (xs[i] - mx) * (ys[i] - my);
    dx += (xs[i] - mx) ** 2;
    dy += (ys[i] - my) ** 2;
  }
  const denom = Math.sqrt(dx * dy);
  return denom === 0 ? 0 : num / denom;
}
function iqrBounds(arr) {
  const s = [...arr].sort((a, b) => a - b);
  const q = (p) => {
    const idx = (s.length - 1) * p;
    const lo = Math.floor(idx), hi = Math.ceil(idx);
    return s[lo] + (s[hi] - s[lo]) * (idx - lo);
  };
  const q1 = q(0.25), q3 = q(0.75), iqr = q3 - q1;
  return { low: q1 - 1.5 * iqr, high: q3 + 1.5 * iqr, q1, q3 };
}
function fmt(n) {
  if (n === null || n === undefined || Number.isNaN(n)) return "—";
  if (Math.abs(n) >= 1000) return n.toLocaleString(undefined, { maximumFractionDigits: 1 });
  return Math.round(n * 1000) / 1000;
}
function fmtDate(ts) {
  const d = new Date(ts);
  return d.toLocaleDateString(undefined, { month: "short", day: "numeric" }) + " · " +
    d.toLocaleTimeString(undefined, { hour: "2-digit", minute: "2-digit" });
}
function download(filename, content, mime) {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}
function computeProfile(rows) {
  if (!rows || !rows.length) return null;
  const columns = Object.keys(rows[0] || {});
  const rowCount = rows.length;
  const seen = new Set();
  let duplicates = 0;
  rows.forEach((r) => { const key = JSON.stringify(r); if (seen.has(key)) duplicates++; else seen.add(key); });
  let totalMissing = 0;
  const cols = columns.map((col) => {
    const raw = rows.map((r) => r[col]);
    const missing = raw.filter((v) => v === null || v === undefined || String(v).trim() === "").length;
    totalMissing += missing;
    const type = inferType(raw);
    const nonEmpty = raw.filter((v) => v !== null && v !== undefined && String(v).trim() !== "");
    const base = { name: col, type, missing, missingPct: (missing / rowCount) * 100, unique: new Set(nonEmpty.map(String)).size };
    if (type === "numeric") {
      const nums = nonEmpty.map(toNum).filter((n) => !isNaN(n));
      const { low, high } = nums.length > 3 ? iqrBounds(nums) : { low: -Infinity, high: Infinity };
      const outliers = nums.filter((n) => n < low || n > high).length;
      return { ...base, min: Math.min(...nums), max: Math.max(...nums), mean: mean(nums), median: median(nums), std: std(nums), outliers, bounds: { low, high } };
    }
    if (type === "categorical") {
      const counts = new Map();
      nonEmpty.forEach((v) => counts.set(String(v), (counts.get(String(v)) || 0) + 1));
      const top = [...counts.entries()].sort((a, b) => b[1] - a[1]).slice(0, 8);
      return { ...base, top };
    }
    return base;
  });
  const numericCols = cols.filter((c) => c.type === "numeric");
  let corr = [];
  if (numericCols.length >= 2) {
    corr = numericCols.map((c1) => numericCols.map((c2) => {
      const pairsX = [], pairsY = [];
      rows.forEach((r) => {
        const x = toNum(r[c1.name]), y = toNum(r[c2.name]);
        if (!isNaN(x) && !isNaN(y)) { pairsX.push(x); pairsY.push(y); }
      });
      return pearson(pairsX, pairsY);
    }));
  }
  return { columns: cols, rowCount, colCount: columns.length, duplicates, totalMissing, numericCols, corr };
}
function buildSummary(profile, fileName) {
  if (!profile) return "";
  const colLines = profile.columns.map((c) => {
    if (c.type === "numeric") return `- ${c.name} (numeric): min=${fmt(c.min)}, max=${fmt(c.max)}, mean=${fmt(c.mean)}, median=${fmt(c.median)}, std=${fmt(c.std)}, missing=${c.missing}, outliers=${c.outliers}`;
    if (c.type === "categorical") return `- ${c.name} (categorical): ${c.unique} unique, top values: ${c.top.map(([v, n]) => `${v}(${n})`).join(", ")}, missing=${c.missing}`;
    return `- ${c.name} (${c.type}): ${c.unique} unique, missing=${c.missing}`;
  }).join("\n");
  let corrLines = "";
  if (profile.numericCols.length >= 2) {
    const pairs = [];
    profile.numericCols.forEach((c1, i) => profile.numericCols.forEach((c2, j) => {
      if (j > i) pairs.push(`${c1.name} vs ${c2.name}: r=${fmt(profile.corr[i][j])}`);
    }));
    corrLines = "\nCorrelations:\n" + pairs.join("\n");
  }
  return `Dataset "${fileName}": ${profile.rowCount} rows, ${profile.colCount} columns, ${profile.duplicates} duplicate rows, ${profile.totalMissing} missing values total.\n\nColumns:\n${colLines}${corrLines}`;
}
// This build is wired for self-hosting: BACKEND_URL = "" means every request
// goes to a same-origin "/api/chat" route. server.py serves this app's built
// files AND that route, proxying it to the real Anthropic API using your own
// ANTHROPIC_API_KEY (see server.py / README.md). That's what makes the AI
// features work outside the Claude.ai artifact preview.
const BACKEND_URL = "";

async function askClaude(prompt) {
  if (BACKEND_URL !== null) {
    const res = await fetch(`${BACKEND_URL}/api/chat`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ prompt }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data?.error || "Request failed");
    return (data.text || "").trim();
  }
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ model: "claude-sonnet-4-6", max_tokens: 1000, messages: [{ role: "user", content: prompt }] }),
  });
  const data = await res.json();
  if (!res.ok) throw new Error(data?.error?.message || "Request failed — if you're self-hosting this app, set BACKEND_URL near the top of this file to point at server.py.");
  return (data.content || []).map((b) => (b.type === "text" ? b.text : "")).join("\n").trim();
}

// ---------------------------------------------------------- storage helpers
// `window.storage` only exists inside a Claude.ai artifact preview. When
// self-hosting, shim it with plain localStorage so history/theme still
// persist across reloads.
if (typeof window !== "undefined" && !window.storage) {
  const STORAGE_PREFIX = "ai-data-analyst:";
  window.storage = {
    async get(key) {
      const raw = localStorage.getItem(STORAGE_PREFIX + key);
      if (raw === null) throw new Error("key not found");
      return { key, value: raw, shared: false };
    },
    async set(key, value) {
      localStorage.setItem(STORAGE_PREFIX + key, value);
      return { key, value, shared: false };
    },
    async delete(key) {
      localStorage.removeItem(STORAGE_PREFIX + key);
      return { key, deleted: true, shared: false };
    },
    async list(prefix = "") {
      const keys = Object.keys(localStorage)
        .filter((k) => k.startsWith(STORAGE_PREFIX + prefix))
        .map((k) => k.slice(STORAGE_PREFIX.length));
      return { keys, prefix, shared: false };
    },
  };
}
const HISTORY_KEY = "analysis-history";
async function loadPersistedHistory() {
  try {
    const res = await window.storage.get(HISTORY_KEY, false);
    return res ? JSON.parse(res.value) : [];
  } catch { return []; }
}
async function savePersistedHistory(list) {
  try { await window.storage.set(HISTORY_KEY, JSON.stringify(list.slice(0, 15)), false); } catch (e) { console.error(e); }
}
async function loadThemePref() {
  try { const res = await window.storage.get("theme-pref", false); return res ? res.value : null; } catch { return null; }
}
async function saveThemePref(v) {
  try { await window.storage.set("theme-pref", v, false); } catch (e) { console.error(e); }
}
function summarize(analysis, profile) {
  return {
    id: analysis.id, fileName: analysis.fileName, timestamp: analysis.timestamp,
    rowCount: profile?.rowCount ?? 0, colCount: profile?.colCount ?? 0,
    missing: profile?.totalMissing ?? 0, duplicates: profile?.duplicates ?? 0,
    insights: analysis.insights || "", log: analysis.log || [],
  };
}

let uid = 0;
const newId = () => `a${Date.now()}_${uid++}`;

// ==================================================================
function AppInner() {
  const T = useT();
  const [analyses, setAnalyses] = useState([]);       // in-session, full data
  const [activeId, setActiveId] = useState(null);
  const [persistedHistory, setPersistedHistory] = useState([]);
  const [historyLoaded, setHistoryLoaded] = useState(false);
  const [tab, setTab] = useState("overview");
  const [dragOver, setDragOver] = useState(false);
  const [parseError, setParseError] = useState("");
  const fileInput = useRef(null);
  const contentRef = useRef(null);
  const [showScrollTop, setShowScrollTop] = useState(false);

  const [insightsLoading, setInsightsLoading] = useState(false);
  const [insightsErr, setInsightsErr] = useState("");
  const [question, setQuestion] = useState("");
  const [askLoading, setAskLoading] = useState(false);
  const [histCol, setHistCol] = useState(null);
  const [catCol, setCatCol] = useState(null);
  const [scatterX, setScatterX] = useState(null);
  const [scatterY, setScatterY] = useState(null);

  const active = analyses.find((a) => a.id === activeId) || null;
  const rows = active?.rows || null;
  const profile = useMemo(() => (rows ? computeProfile(rows) : null), [rows]);

  useEffect(() => {
    (async () => { setPersistedHistory(await loadPersistedHistory()); setHistoryLoaded(true); })();
  }, []);

  useEffect(() => { if (contentRef.current) contentRef.current.scrollTop = 0; }, [tab, activeId]);

  const persist = (analysisObj, prof) => {
    const entry = summarize(analysisObj, prof);
    setPersistedHistory((prev) => {
      const updated = [entry, ...prev.filter((e) => e.id !== entry.id)].slice(0, 15);
      savePersistedHistory(updated);
      return updated;
    });
  };

  // -------------------- upload handling --------------------
  const handleFile = useCallback((file) => {
    setParseError("");
    const name = file.name.toLowerCase();
    const finish = (data) => {
      if (!Array.isArray(data) || data.length === 0) { setParseError("Couldn't find any rows in that file."); return; }
      const id = newId();
      const newAnalysis = {
        id, fileName: file.name, timestamp: Date.now(),
        origRows: data, rows: data, log: [], insights: "", chat: [],
      };
      setAnalyses((prev) => [...prev, newAnalysis]);
      setActiveId(id);
      setTab("overview");
      setHistCol(null); setCatCol(null); setScatterX(null); setScatterY(null);
      setInsightsErr("");
      persist(newAnalysis, computeProfile(data));
    };
    if (name.endsWith(".csv") || name.endsWith(".tsv")) {
      Papa.parse(file, { header: true, skipEmptyLines: true, complete: (r) => finish(r.data), error: (e) => setParseError(e.message) });
    } else if (name.endsWith(".json")) {
      const reader = new FileReader();
      reader.onload = () => {
        try {
          let parsed = JSON.parse(reader.result);
          if (!Array.isArray(parsed)) parsed = parsed.data || parsed.rows || parsed.records || Object.values(parsed)[0];
          finish(parsed);
        } catch (e) { setParseError("That JSON file couldn't be parsed."); }
      };
      reader.readAsText(file);
    } else if (name.endsWith(".xlsx") || name.endsWith(".xls")) {
      const reader = new FileReader();
      reader.onload = () => {
        try {
          const wb = XLSX.read(reader.result, { type: "array" });
          finish(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" }));
        } catch (e) { setParseError("That Excel file couldn't be read."); }
      };
      reader.readAsArrayBuffer(file);
    } else {
      setParseError("Supported formats: CSV, TSV, JSON, XLS, XLSX. For a SQL export, download it as CSV first.");
    }
  }, []);

  const updateActive = (patch, newRowsForProfile) => {
    setAnalyses((prev) => prev.map((a) => (a.id === activeId ? { ...a, ...patch } : a)));
    const merged = { ...active, ...patch };
    persist(merged, computeProfile(newRowsForProfile || merged.rows));
  };
  const addLog = (msg, newRows) => updateActive({ log: [msg, ...active.log].slice(0, 30), ...(newRows ? { rows: newRows } : {}) }, newRows);

  // -------------------- cleaning ops --------------------
  const fillMissing = (colName, method, customVal) => {
    const col = profile.columns.find((c) => c.name === colName);
    let fillVal;
    if (method === "mean") fillVal = fmt(col.mean);
    else if (method === "median") fillVal = fmt(col.median);
    else if (method === "mode") fillVal = mode(active.rows.map((r) => r[colName]).filter((v) => v !== "" && v != null));
    else fillVal = customVal ?? "";
    const newRows = active.rows.map((r) => ((r[colName] === "" || r[colName] === null || r[colName] === undefined) ? { ...r, [colName]: fillVal } : r));
    addLog(`Filled missing values in "${colName}" using ${method}`, newRows);
  };
  const dropMissing = (colName) => {
    const newRows = active.rows.filter((r) => !(r[colName] === "" || r[colName] === null || r[colName] === undefined));
    addLog(`Dropped rows with missing "${colName}"`, newRows);
  };
  const removeDuplicates = () => {
    const seen = new Set(); const out = [];
    active.rows.forEach((r) => { const k = JSON.stringify(r); if (!seen.has(k)) { seen.add(k); out.push(r); } });
    addLog("Removed duplicate rows", out);
  };
  const removeOutliers = (colName) => {
    const col = profile.columns.find((c) => c.name === colName);
    const newRows = active.rows.filter((r) => { const n = toNum(r[colName]); return isNaN(n) || (n >= col.bounds.low && n <= col.bounds.high); });
    addLog(`Removed outliers in "${colName}" (IQR method)`, newRows);
  };
  const convertType = (colName, type) => {
    const newRows = active.rows.map((r) => { let v = r[colName]; if (type === "number") v = toNum(v); else if (type === "string") v = String(v); return { ...r, [colName]: v }; });
    addLog(`Converted "${colName}" to ${type}`, newRows);
  };
  const resetData = () => updateActive({ rows: active.origRows, log: [] }, active.origRows);

  // -------------------- AI insights & ask --------------------
  const generateInsights = async () => {
    setInsightsLoading(true); setInsightsErr("");
    try {
      const summary = buildSummary(profile, active.fileName);
      const text = await askClaude(
        `You are a data analyst. Here is a profile of a dataset:\n\n${summary}\n\nWrite a concise executive summary (6-9 bullet points, each starting with "- ") covering the most important trends, notable correlations, anomalies/outliers, and data quality issues. Be specific and reference actual column names and numbers. No preamble or closing remarks, just the bullets.`
      );
      updateActive({ insights: text });
    } catch (e) { setInsightsErr(e.message || "Something went wrong generating insights."); }
    setInsightsLoading(false);
  };

  const askQuestion = async () => {
    if (!question.trim()) return;
    const q = question.trim();
    const withUser = [...active.chat, { role: "user", text: q }];
    setAnalyses((prev) => prev.map((a) => (a.id === activeId ? { ...a, chat: withUser } : a)));
    setQuestion(""); setAskLoading(true);
    try {
      const summary = buildSummary(profile, active.fileName);
      const sample = active.rows.slice(0, 25);
      const history = active.chat.slice(-6).map((m) => `${m.role === "user" ? "Q" : "A"}: ${m.text}`).join("\n");
      const text = await askClaude(
        `You are answering questions about a dataset for a non-technical user.\n\nDataset profile:\n${summary}\n\nSample rows (first ${sample.length} of ${active.rows.length}):\n${JSON.stringify(sample)}\n\n${history ? "Recent conversation:\n" + history + "\n\n" : ""}Question: ${q}\n\nAnswer directly and concisely using the profile stats where possible. If the question needs the full dataset and you only have a sample, say so and give your best estimate from the stats provided.`
      );
      setAnalyses((prev) => prev.map((a) => (a.id === activeId ? { ...a, chat: [...withUser, { role: "assistant", text }] } : a)));
    } catch (e) {
      setAnalyses((prev) => prev.map((a) => (a.id === activeId ? { ...a, chat: [...withUser, { role: "assistant", text: "I couldn't reach the AI service: " + (e.message || "unknown error") }] } : a)));
    }
    setAskLoading(false);
  };

  // -------------------- downloads --------------------
  const downloadCleaned = () => download((active.fileName.replace(/\.[^.]+$/, "") || "dataset") + "_cleaned.csv", Papa.unparse(active.rows), "text/csv");
  const downloadReport = () => {
    const md = `# Data Analysis Report\n\n**Dataset:** ${active.fileName}\n**Rows:** ${profile.rowCount}  **Columns:** ${profile.colCount}  **Duplicates:** ${profile.duplicates}  **Missing values:** ${profile.totalMissing}\n\n## Column Profile\n\n${profile.columns.map((c) => `### ${c.name} (${c.type})\n${c.type === "numeric" ? `- min: ${fmt(c.min)}, max: ${fmt(c.max)}, mean: ${fmt(c.mean)}, median: ${fmt(c.median)}, std: ${fmt(c.std)}\n- outliers (IQR): ${c.outliers}` : c.type === "categorical" ? `- unique values: ${c.unique}\n- top values: ${c.top.map(([v, n]) => `${v} (${n})`).join(", ")}` : `- unique values: ${c.unique}`}\n- missing: ${c.missing} (${fmt(c.missingPct)}%)\n`).join("\n")}\n\n## Cleaning Actions Applied\n${active.log.length ? active.log.map((l) => `- ${l}`).join("\n") : "- none"}\n\n## AI-Generated Insights\n${active.insights || "_Not generated yet — visit the Insights tab._"}\n`;
    download((active.fileName.replace(/\.[^.]+$/, "") || "dataset") + "_report.md", md, "text/markdown");
  };

  // -------------------- derived chart data --------------------
  const histData = useMemo(() => {
    if (!profile || !histCol || !rows) return [];
    const nums = rows.map((r) => toNum(r[histCol])).filter((n) => !isNaN(n));
    if (!nums.length) return [];
    const bins = 12;
    const min = Math.min(...nums), max = Math.max(...nums);
    const width = (max - min) / bins || 1;
    const counts = Array(bins).fill(0);
    nums.forEach((n) => { let idx = Math.floor((n - min) / width); if (idx >= bins) idx = bins - 1; if (idx < 0) idx = 0; counts[idx]++; });
    return counts.map((c, i) => ({ bucket: `${fmt(min + i * width)}`, count: c }));
  }, [profile, histCol, rows]);
  const catData = useMemo(() => {
    if (!profile || !catCol) return [];
    const col = profile.columns.find((c) => c.name === catCol);
    return (col?.top || []).map(([v, n]) => ({ name: v, count: n }));
  }, [profile, catCol]);
  const scatterData = useMemo(() => {
    if (!scatterX || !scatterY || !rows) return [];
    return rows.map((r) => ({ x: toNum(r[scatterX]), y: toNum(r[scatterY]) })).filter((p) => !isNaN(p.x) && !isNaN(p.y));
  }, [scatterX, scatterY, rows]);

  const deletePersisted = (id) => setPersistedHistory((prev) => { const u = prev.filter((e) => e.id !== id); savePersistedHistory(u); return u; });
  const clearPersisted = () => { setPersistedHistory([]); savePersistedHistory([]); };
  const removeFromSession = (id) => {
    setAnalyses((prev) => prev.filter((a) => a.id !== id));
    if (activeId === id) { const rest = analyses.filter((a) => a.id !== id); setActiveId(rest.length ? rest[rest.length - 1].id : null); }
  };

  const navItems = [
    { id: "overview", label: "Overview", icon: Database },
    { id: "preview", label: "Data Preview", icon: Table2 },
    { id: "clean", label: "Cleaning", icon: Filter },
    { id: "charts", label: "Visual Analytics", icon: BarChart3 },
    { id: "insights", label: "AI Insights", icon: Sparkles },
    { id: "ask", label: "Ask Your Data", icon: MessageSquare },
  ];

  return (
    <div style={{ background: T.bg, height: "100vh", overflow: "hidden", color: T.text, fontFamily: "'Inter', sans-serif", display: "flex" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500;600&display=swap');
        * { box-sizing: border-box; }
        .mono { font-family: 'JetBrains Mono', monospace; }
        .display { font-family: 'Space Grotesk', sans-serif; }
        ::-webkit-scrollbar { width: 8px; height: 8px; }
        ::-webkit-scrollbar-thumb { background: ${T.borderLight}; border-radius: 4px; }
        button { font-family: inherit; cursor: pointer; }
        input, select { font-family: inherit; }
        .navbtn:hover { background: ${T.surface2} !important; }
        .navbtn:focus-visible, button:focus-visible, input:focus-visible { outline: 2px solid ${T.signal}; outline-offset: 2px; }
        .content-scroll { scroll-behavior: smooth; }
        table { border-collapse: collapse; width: 100%; }
        th, td { border-bottom: 1px solid ${T.border}; padding: 8px 12px; text-align: left; white-space: nowrap; }
        th { color: ${T.muted}; font-weight: 500; font-size: 11px; text-transform: uppercase; letter-spacing: 0.05em; position: sticky; top: 0; background: ${T.surface}; }
        .pulse { animation: pulse 2s ease-in-out infinite; }
        @keyframes pulse { 0%,100%{opacity:1} 50%{opacity:0.35} }
        @media (prefers-reduced-motion: reduce) { .pulse { animation: none; } }
      `}</style>

      {/* ---------------- sidebar ---------------- */}
      <div style={{ width: 230, borderRight: `1px solid ${T.border}`, padding: "20px 12px", display: "flex", flexDirection: "column", flexShrink: 0, overflowY: "auto", height: "100%" }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "0 8px 20px" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <Activity size={20} color={T.signal} />
            <span className="display" style={{ fontSize: 15, fontWeight: 700, letterSpacing: "0.02em" }}>ANALYST.OS</span>
          </div>
          <ThemeToggle />
        </div>

        {navItems.map((t) => (
          <button key={t.id} className="navbtn" onClick={() => setTab(t.id)} disabled={!active}
            style={{
              display: "flex", alignItems: "center", gap: 10, padding: "10px 12px", marginBottom: 2,
              background: tab === t.id ? T.surface2 : "transparent", border: "none", borderRadius: 8,
              color: !active ? T.faint : tab === t.id ? T.text : T.muted, fontSize: 13.5, fontWeight: 500,
              textAlign: "left", opacity: !active ? 0.5 : 1,
            }}>
            <t.icon size={16} /> {t.label}
          </button>
        ))}
        <button className="navbtn" onClick={() => setTab("history")}
          style={{
            display: "flex", alignItems: "center", gap: 10, padding: "10px 12px", marginBottom: 2,
            background: tab === "history" ? T.surface2 : "transparent", border: "none", borderRadius: 8,
            color: tab === "history" ? T.text : T.muted, fontSize: 13.5, fontWeight: 500, textAlign: "left",
          }}>
          <HistoryIcon size={16} /> History
          {persistedHistory.length > 0 && <span className="mono" style={{ marginLeft: "auto", fontSize: 10.5, color: T.faint }}>{persistedHistory.length}</span>}
        </button>

        {/* session dataset switcher */}
        <div style={{ marginTop: 18, marginBottom: 8, display: "flex", justifyContent: "space-between", alignItems: "center", padding: "0 4px" }}>
          <span className="mono" style={{ fontSize: 10, color: T.faint, letterSpacing: "0.05em" }}>DATASETS</span>
          <button onClick={() => fileInput.current?.click()} title="Upload another dataset"
            style={{ background: "transparent", border: `1px solid ${T.borderLight}`, borderRadius: 5, color: T.muted, padding: 2, display: "flex" }}>
            <Plus size={12} />
          </button>
        </div>
        <div style={{ display: "flex", flexDirection: "column", gap: 3, overflowY: "auto", maxHeight: 180 }}>
          {analyses.map((a) => (
            <div key={a.id} onClick={() => setActiveId(a.id)} style={{
              display: "flex", alignItems: "center", gap: 6, padding: "7px 8px", borderRadius: 7, cursor: "pointer",
              background: a.id === activeId ? T.surface2 : "transparent",
            }}>
              <span className="pulse" style={{ width: 6, height: 6, borderRadius: "50%", background: a.id === activeId ? T.signal : T.faint, flexShrink: 0 }} />
              <span style={{ fontSize: 12, color: a.id === activeId ? T.text : T.muted, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", flex: 1 }}>{a.fileName}</span>
              <button onClick={(e) => { e.stopPropagation(); removeFromSession(a.id); }} style={{ background: "transparent", border: "none", color: T.faint, display: "flex", flexShrink: 0 }}><X size={11} /></button>
            </div>
          ))}
          {analyses.length === 0 && <div style={{ fontSize: 11.5, color: T.faint, padding: "4px 8px" }}>No datasets yet</div>}
        </div>
        <div style={{ flex: 1 }} />
      </div>

      {/* ---------------- main ---------------- */}
      <div style={{ flex: 1, display: "flex", flexDirection: "column", minWidth: 0, minHeight: 0 }}>
        <div style={{ borderBottom: `1px solid ${T.border}`, padding: "14px 28px", display: "flex", alignItems: "center", gap: 24, flexWrap: "wrap" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <span className="pulse" style={{ width: 8, height: 8, borderRadius: "50%", background: active ? T.signal : T.faint, display: "inline-block" }} />
            <span className="mono" style={{ fontSize: 11, color: T.muted }}>{active ? "DATA LOADED" : "AWAITING INPUT"}</span>
          </div>
          {profile && (
            <div className="mono" style={{ display: "flex", gap: 20, fontSize: 12, color: T.text }}>
              <span>ROWS <b style={{ color: T.signal }}>{profile.rowCount.toLocaleString()}</b></span>
              <span>COLS <b style={{ color: T.signal }}>{profile.colCount}</b></span>
              <span>MISSING <b style={{ color: profile.totalMissing ? T.amber : T.signal }}>{profile.totalMissing}</b></span>
              <span>DUPES <b style={{ color: profile.duplicates ? T.danger : T.signal }}>{profile.duplicates}</b></span>
            </div>
          )}
          <div style={{ flex: 1 }} />
          {active && (
            <div style={{ display: "flex", gap: 8 }}>
              <button onClick={downloadCleaned} style={{ display: "flex", alignItems: "center", gap: 6, padding: "7px 12px", background: T.surface2, border: `1px solid ${T.border}`, borderRadius: 7, color: T.text, fontSize: 12.5 }}><Download size={13} /> Cleaned CSV</button>
              <button onClick={downloadReport} style={{ display: "flex", alignItems: "center", gap: 6, padding: "7px 12px", background: T.signal, border: "none", borderRadius: 7, color: T.onSignal, fontSize: 12.5, fontWeight: 600 }}><Download size={13} /> Report</button>
            </div>
          )}
        </div>

        <div ref={contentRef} className="content-scroll" onScroll={(e) => setShowScrollTop(e.target.scrollTop > 300)} style={{ flex: 1, minHeight: 0, overflowY: "auto", overflowX: "hidden", padding: 28, position: "relative", WebkitOverflowScrolling: "touch" }}>
          {tab === "history" ? (
            <HistoryTab persistedHistory={persistedHistory} analyses={analyses} setActiveId={setActiveId} setTab={setTab} deletePersisted={deletePersisted} clearPersisted={clearPersisted} historyLoaded={historyLoaded} />
          ) : !active ? (
            <UploadScreen dragOver={dragOver} setDragOver={setDragOver} handleFile={handleFile} fileInput={fileInput} parseError={parseError} />
          ) : tab === "overview" ? (
            <Overview profile={profile} />
          ) : tab === "preview" ? (
            <Preview rows={rows} />
          ) : tab === "clean" ? (
            <Cleaning profile={profile} fillMissing={fillMissing} dropMissing={dropMissing} removeDuplicates={removeDuplicates} removeOutliers={removeOutliers} convertType={convertType} resetData={resetData} log={active.log} />
          ) : tab === "charts" ? (
            <Charts profile={profile} histCol={histCol} setHistCol={setHistCol} catCol={catCol} setCatCol={setCatCol} scatterX={scatterX} setScatterX={setScatterX} scatterY={scatterY} setScatterY={setScatterY} histData={histData} catData={catData} scatterData={scatterData} />
          ) : tab === "insights" ? (
            <Insights insights={active.insights} loading={insightsLoading} err={insightsErr} generate={generateInsights} />
          ) : tab === "ask" ? (
            <AskData chat={active.chat} question={question} setQuestion={setQuestion} askQuestion={askQuestion} loading={askLoading} />
          ) : null}

          {showScrollTop && (
            <button onClick={() => contentRef.current.scrollTo({ top: 0, behavior: "smooth" })}
              style={{ position: "sticky", bottom: 20, left: "100%", transform: "translateX(-56px)", width: 36, height: 36, borderRadius: "50%", background: T.surface2, border: `1px solid ${T.borderLight}`, color: T.text, display: "flex", alignItems: "center", justifyContent: "center", boxShadow: "0 4px 14px rgba(0,0,0,0.25)" }}>
              <ChevronUp size={16} />
            </button>
          )}
        </div>
      </div>

      <input ref={fileInput} type="file" accept=".csv,.tsv,.json,.xlsx,.xls" style={{ display: "none" }}
        onChange={(e) => e.target.files[0] && handleFile(e.target.files[0])} />
    </div>
  );
}

export default function App() {
  const [theme, setTheme] = useState("dark");
  useEffect(() => { (async () => { const p = await loadThemePref(); if (p === "light" || p === "dark") setTheme(p); })(); }, []);
  const palette = theme === "dark" ? dark : light;
  const toggle = () => { const next = theme === "dark" ? "light" : "dark"; setTheme(next); saveThemePref(next); };
  return (
    <ThemeCtx.Provider value={{ ...palette, toggleTheme: toggle, themeName: theme }}>
      <AppInner />
    </ThemeCtx.Provider>
  );
}

function ThemeToggle() {
  const T = useT();
  return (
    <button onClick={T.toggleTheme} title="Toggle theme"
      style={{ width: 28, height: 28, borderRadius: 7, background: T.surface2, border: `1px solid ${T.border}`, color: T.muted, display: "flex", alignItems: "center", justifyContent: "center" }}>
      {T.themeName === "dark" ? <Sun size={14} /> : <Moon size={14} />}
    </button>
  );
}

// ==================== sub-components ====================

function UploadScreen({ dragOver, setDragOver, handleFile, fileInput, parseError }) {
  const T = useT();
  return (
    <div style={{ maxWidth: 640, margin: "40px auto", textAlign: "center" }}>
      <div className="display" style={{ fontSize: 28, fontWeight: 700, marginBottom: 6 }}>Bring in your data</div>
      <div style={{ color: T.muted, fontSize: 14, marginBottom: 28 }}>CSV, TSV, JSON, or Excel. Everything runs in your browser — nothing leaves this tab except when you ask for AI insights. Each upload is kept as its own analysis you can switch back to.</div>
      <div
        onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
        onDragLeave={() => setDragOver(false)}
        onDrop={(e) => { e.preventDefault(); setDragOver(false); if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]); }}
        onClick={() => fileInput.current?.click()}
        style={{
          border: `1.5px dashed ${dragOver ? T.signal : T.borderLight}`, borderRadius: 16, padding: "56px 24px",
          background: dragOver ? rgba(T.signal, 0.08) : T.surface, cursor: "pointer", transition: "all 0.15s",
        }}>
        <UploadCloud size={34} color={dragOver ? T.signal : T.muted} style={{ marginBottom: 14 }} />
        <div style={{ fontSize: 15, fontWeight: 600, marginBottom: 4 }}>Drop a file here, or click to browse</div>
        <div className="mono" style={{ fontSize: 11, color: T.faint }}>.CSV · .TSV · .JSON · .XLSX · .XLS</div>
      </div>
      {parseError && (
        <div style={{ marginTop: 16, display: "flex", gap: 8, alignItems: "flex-start", background: rgba(T.danger, 0.08), border: `1px solid ${T.danger}`, borderRadius: 10, padding: "12px 14px", textAlign: "left" }}>
          <FileWarning size={16} color={T.danger} style={{ flexShrink: 0, marginTop: 1 }} />
          <span style={{ fontSize: 13, color: T.text }}>{parseError}</span>
        </div>
      )}
    </div>
  );
}

function Card({ children, style }) {
  const T = useT();
  return <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 12, padding: 18, ...style }}>{children}</div>;
}

function Overview({ profile }) {
  const T = useT();
  return (
    <div>
      <div className="display" style={{ fontSize: 20, fontWeight: 700, marginBottom: 16 }}>Dataset Overview</div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px,1fr))", gap: 12, marginBottom: 24 }}>
        {[
          ["Rows", profile.rowCount.toLocaleString(), T.signal],
          ["Columns", profile.colCount, T.signal],
          ["Missing values", profile.totalMissing, profile.totalMissing ? T.amber : T.signal],
          ["Duplicate rows", profile.duplicates, profile.duplicates ? T.danger : T.signal],
        ].map(([label, val, color]) => (
          <Card key={label}>
            <div style={{ fontSize: 11, color: T.muted, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 8 }}>{label}</div>
            <div className="mono" style={{ fontSize: 26, fontWeight: 600, color }}>{val}</div>
          </Card>
        ))}
      </div>
      <div className="display" style={{ fontSize: 15, fontWeight: 600, marginBottom: 10 }}>Column Profile</div>
      <Card style={{ padding: 0, overflow: "auto" }}>
        <table>
          <thead><tr><th>Column</th><th>Type</th><th>Missing</th><th>Unique</th><th>Detail</th></tr></thead>
          <tbody>
            {profile.columns.map((c) => (
              <tr key={c.name}>
                <td style={{ fontWeight: 500 }}>{c.name}</td>
                <td><TypeBadge type={c.type} /></td>
                <td style={{ color: c.missing ? T.amber : T.muted }} className="mono">{c.missing} ({fmt(c.missingPct)}%)</td>
                <td className="mono">{c.unique}</td>
                <td className="mono" style={{ color: T.muted, fontSize: 12 }}>
                  {c.type === "numeric" ? `μ=${fmt(c.mean)} σ=${fmt(c.std)} range=[${fmt(c.min)}, ${fmt(c.max)}]${c.outliers ? ` · ${c.outliers} outliers` : ""}` :
                   c.type === "categorical" ? `top: ${c.top.slice(0, 3).map(([v, n]) => `${v}(${n})`).join(", ")}` : "—"}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </Card>
    </div>
  );
}

function TypeBadge({ type }) {
  const T = useT();
  const colors = { numeric: T.signal, categorical: T.violet, date: T.amber, text: T.muted, empty: T.danger };
  return <span className="mono" style={{ fontSize: 10.5, padding: "2px 8px", borderRadius: 20, border: `1px solid ${colors[type] || T.border}`, color: colors[type] || T.muted, textTransform: "uppercase" }}>{type}</span>;
}

function Preview({ rows }) {
  const T = useT();
  const cols = Object.keys(rows[0] || {});
  return (
    <div>
      <div className="display" style={{ fontSize: 20, fontWeight: 700, marginBottom: 4 }}>Data Preview</div>
      <div style={{ color: T.muted, fontSize: 13, marginBottom: 16 }}>Showing first 100 of {rows.length.toLocaleString()} rows.</div>
      <Card style={{ padding: 0, overflow: "auto", maxHeight: "70vh" }}>
        <table>
          <thead><tr>{cols.map((c) => <th key={c}>{c}</th>)}</tr></thead>
          <tbody>
            {rows.slice(0, 100).map((r, i) => (
              <tr key={i}>{cols.map((c) => <td key={c} className="mono" style={{ fontSize: 12.5 }}>{String(r[c] ?? "")}</td>)}</tr>
            ))}
          </tbody>
        </table>
      </Card>
    </div>
  );
}

function Cleaning({ profile, fillMissing, dropMissing, removeDuplicates, removeOutliers, convertType, resetData, log }) {
  const T = useT();
  const [selCol, setSelCol] = useState(profile.columns[0]?.name);
  const col = profile.columns.find((c) => c.name === selCol) || profile.columns[0];
  const [customVal, setCustomVal] = useState("");
  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
        <div className="display" style={{ fontSize: 20, fontWeight: 700 }}>Data Cleaning</div>
        <button onClick={resetData} style={{ display: "flex", gap: 6, alignItems: "center", padding: "6px 12px", background: "transparent", border: `1px solid ${T.borderLight}`, borderRadius: 7, color: T.muted, fontSize: 12.5 }}><RotateCcw size={13} /> Reset to original</button>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "260px 1fr", gap: 18 }}>
        <Card>
          <div style={{ fontSize: 11, color: T.muted, textTransform: "uppercase", marginBottom: 10 }}>Select column</div>
          {profile.columns.map((c) => (
            <button key={c.name} onClick={() => setSelCol(c.name)} style={{
              display: "flex", justifyContent: "space-between", width: "100%", padding: "8px 10px", marginBottom: 2, borderRadius: 7,
              background: selCol === c.name ? T.surface2 : "transparent", border: "none", color: T.text, fontSize: 13,
            }}>
              <span>{c.name}</span>
              {c.missing > 0 && <span className="mono" style={{ color: T.amber, fontSize: 11 }}>{c.missing}</span>}
            </button>
          ))}
        </Card>
        <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
          {col && (
            <Card>
              <div style={{ fontWeight: 600, marginBottom: 12 }}>{col.name} <TypeBadge type={col.type} /></div>
              {col.missing > 0 && (
                <div style={{ marginBottom: 16 }}>
                  <div style={{ fontSize: 12.5, color: T.muted, marginBottom: 8 }}>{col.missing} missing values ({fmt(col.missingPct)}%)</div>
                  <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                    {col.type === "numeric" && <>
                      <ActionBtn onClick={() => fillMissing(col.name, "mean")}>Fill with mean</ActionBtn>
                      <ActionBtn onClick={() => fillMissing(col.name, "median")}>Fill with median</ActionBtn>
                    </>}
                    {col.type !== "numeric" && <ActionBtn onClick={() => fillMissing(col.name, "mode")}>Fill with most common</ActionBtn>}
                    <input value={customVal} onChange={(e) => setCustomVal(e.target.value)} placeholder="custom value" style={{ background: T.surface2, border: `1px solid ${T.border}`, borderRadius: 6, padding: "6px 10px", color: T.text, fontSize: 12.5, width: 110 }} />
                    <ActionBtn onClick={() => fillMissing(col.name, "custom", customVal)}>Fill with value</ActionBtn>
                    <ActionBtn danger onClick={() => dropMissing(col.name)}>Drop rows</ActionBtn>
                  </div>
                </div>
              )}
              {col.type === "numeric" && col.outliers > 0 && (
                <div style={{ marginBottom: 16 }}>
                  <div style={{ fontSize: 12.5, color: T.muted, marginBottom: 8 }}>{col.outliers} outliers detected (IQR method, outside [{fmt(col.bounds.low)}, {fmt(col.bounds.high)}])</div>
                  <ActionBtn danger onClick={() => removeOutliers(col.name)}>Remove outlier rows</ActionBtn>
                </div>
              )}
              <div>
                <div style={{ fontSize: 12.5, color: T.muted, marginBottom: 8 }}>Convert type</div>
                <div style={{ display: "flex", gap: 8 }}>
                  <ActionBtn onClick={() => convertType(col.name, "number")}>To number</ActionBtn>
                  <ActionBtn onClick={() => convertType(col.name, "string")}>To text</ActionBtn>
                </div>
              </div>
            </Card>
          )}
          <Card>
            <div style={{ fontWeight: 600, marginBottom: 10 }}>Whole dataset</div>
            <ActionBtn onClick={removeDuplicates}>Remove {profile.duplicates} duplicate row{profile.duplicates === 1 ? "" : "s"}</ActionBtn>
          </Card>
          {log.length > 0 && (
            <Card>
              <div style={{ fontWeight: 600, marginBottom: 10 }}>Action log</div>
              {log.map((l, i) => (
                <div key={i} style={{ display: "flex", gap: 8, fontSize: 12.5, color: T.muted, marginBottom: 6 }}>
                  <CheckCircle2 size={13} color={T.signal} style={{ flexShrink: 0, marginTop: 1 }} /> {l}
                </div>
              ))}
            </Card>
          )}
        </div>
      </div>
    </div>
  );
}

function ActionBtn({ children, onClick, danger }) {
  const T = useT();
  return (
    <button onClick={onClick} style={{
      padding: "7px 12px", borderRadius: 7, fontSize: 12.5, fontWeight: 500,
      background: danger ? rgba(T.danger, 0.1) : T.surface2, border: `1px solid ${danger ? T.danger : T.borderLight}`,
      color: danger ? T.danger : T.text,
    }}>{children}</button>
  );
}

function Charts({ profile, histCol, setHistCol, catCol, setCatCol, scatterX, setScatterX, scatterY, setScatterY, histData, catData, scatterData }) {
  const T = useT();
  const numeric = profile.columns.filter((c) => c.type === "numeric");
  const categorical = profile.columns.filter((c) => c.type === "categorical");
  const Sel = ({ value, onChange, options, placeholder }) => (
    <select value={value || ""} onChange={(e) => onChange(e.target.value)} style={{ background: T.surface2, border: `1px solid ${T.border}`, borderRadius: 7, color: T.text, fontSize: 12.5, padding: "6px 10px" }}>
      <option value="">{placeholder}</option>
      {options.map((o) => <option key={o} value={o}>{o}</option>)}
    </select>
  );
  return (
    <div>
      <div className="display" style={{ fontSize: 20, fontWeight: 700, marginBottom: 16 }}>Visual Analytics</div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 16 }}>
        <Card>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
            <div style={{ fontWeight: 600, fontSize: 13.5 }}>Distribution</div>
            <Sel value={histCol} onChange={setHistCol} options={numeric.map((c) => c.name)} placeholder="Choose numeric column" />
          </div>
          {histData.length ? (
            <ResponsiveContainer width="100%" height={220}>
              <BarChart data={histData}>
                <CartesianGrid stroke={T.border} vertical={false} />
                <XAxis dataKey="bucket" tick={{ fill: T.faint, fontSize: 10 }} />
                <YAxis tick={{ fill: T.faint, fontSize: 10 }} />
                <Tooltip contentStyle={{ background: T.surface2, border: `1px solid ${T.border}`, fontSize: 12 }} />
                <Bar dataKey="count" fill={T.signal} radius={[3, 3, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          ) : <EmptyChart />}
        </Card>
        <Card>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
            <div style={{ fontWeight: 600, fontSize: 13.5 }}>Category counts</div>
            <Sel value={catCol} onChange={setCatCol} options={categorical.map((c) => c.name)} placeholder="Choose category column" />
          </div>
          {catData.length ? (
            <ResponsiveContainer width="100%" height={220}>
              <BarChart data={catData} layout="vertical">
                <CartesianGrid stroke={T.border} horizontal={false} />
                <XAxis type="number" tick={{ fill: T.faint, fontSize: 10 }} />
                <YAxis dataKey="name" type="category" width={90} tick={{ fill: T.faint, fontSize: 10 }} />
                <Tooltip contentStyle={{ background: T.surface2, border: `1px solid ${T.border}`, fontSize: 12 }} />
                <Bar dataKey="count" fill={T.violet} radius={[0, 3, 3, 0]} />
              </BarChart>
            </ResponsiveContainer>
          ) : <EmptyChart />}
        </Card>
      </div>
      <Card style={{ marginBottom: 16 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
          <div style={{ fontWeight: 600, fontSize: 13.5 }}>Scatter — relationship between two numeric columns</div>
          <div style={{ display: "flex", gap: 8 }}>
            <Sel value={scatterX} onChange={setScatterX} options={numeric.map((c) => c.name)} placeholder="X axis" />
            <Sel value={scatterY} onChange={setScatterY} options={numeric.map((c) => c.name)} placeholder="Y axis" />
          </div>
        </div>
        {scatterData.length ? (
          <ResponsiveContainer width="100%" height={260}>
            <ScatterChart>
              <CartesianGrid stroke={T.border} />
              <XAxis dataKey="x" type="number" name={scatterX} tick={{ fill: T.faint, fontSize: 10 }} />
              <YAxis dataKey="y" type="number" name={scatterY} tick={{ fill: T.faint, fontSize: 10 }} />
              <Tooltip cursor={{ strokeDasharray: "3 3" }} contentStyle={{ background: T.surface2, border: `1px solid ${T.border}`, fontSize: 12 }} />
              <Scatter data={scatterData} fill={T.amber} />
            </ScatterChart>
          </ResponsiveContainer>
        ) : <EmptyChart />}
      </Card>
      {numeric.length >= 2 && (
        <Card>
          <div style={{ fontWeight: 600, fontSize: 13.5, marginBottom: 14 }}>Correlation matrix</div>
          <div style={{ overflow: "auto" }}>
            <table>
              <thead><tr><th></th>{numeric.map((c) => <th key={c.name}>{c.name}</th>)}</tr></thead>
              <tbody>
                {numeric.map((c1, i) => (
                  <tr key={c1.name}>
                    <th style={{ position: "static" }}>{c1.name}</th>
                    {numeric.map((c2, j) => {
                      const v = profile.corr[i][j];
                      const intensity = Math.min(Math.abs(v), 1);
                      const bg = v >= 0 ? rgba(T.signal, intensity * 0.55) : rgba(T.danger, intensity * 0.55);
                      return <td key={c2.name} className="mono" style={{ background: bg, textAlign: "center", fontSize: 12 }}>{fmt(v)}</td>;
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Card>
      )}
    </div>
  );
}
function EmptyChart() { const T = useT(); return <div style={{ height: 220, display: "flex", alignItems: "center", justifyContent: "center", color: T.faint, fontSize: 12.5 }}>Choose a column above to plot it</div>; }

function Insights({ insights, loading, err, generate }) {
  const T = useT();
  const bullets = insights.split("\n").filter((l) => l.trim().startsWith("-"));
  return (
    <div style={{ maxWidth: 760 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
        <div className="display" style={{ fontSize: 20, fontWeight: 700 }}>AI Insights</div>
        <button onClick={generate} disabled={loading} style={{ display: "flex", gap: 8, alignItems: "center", padding: "9px 16px", background: T.signal, border: "none", borderRadius: 8, color: T.onSignal, fontWeight: 600, fontSize: 13, opacity: loading ? 0.7 : 1 }}>
          {loading ? <Loader2 size={15} className="pulse" /> : <Sparkles size={15} />}
          {loading ? "Analyzing…" : insights ? "Regenerate" : "Generate insights"}
        </button>
      </div>
      {err && <Card style={{ borderColor: T.danger, marginBottom: 16 }}><div style={{ color: T.danger, fontSize: 13, display: "flex", gap: 8 }}><AlertTriangle size={16} />{err}</div></Card>}
      {!insights && !loading && !err && (
        <Card><div style={{ color: T.muted, fontSize: 13.5 }}>Generate a plain-English summary of trends, correlations, anomalies, and data quality issues found in your dataset.</div></Card>
      )}
      {bullets.length > 0 && (
        <Card>
          {bullets.map((b, i) => (
            <div key={i} style={{ display: "flex", gap: 10, marginBottom: 12, fontSize: 14, lineHeight: 1.5 }}>
              <ChevronRight size={15} color={T.signal} style={{ flexShrink: 0, marginTop: 2 }} />
              <span>{b.replace(/^-\s*/, "")}</span>
            </div>
          ))}
        </Card>
      )}
    </div>
  );
}

function AskData({ chat, question, setQuestion, askQuestion, loading }) {
  const T = useT();
  const scrollRef = useRef(null);
  useEffect(() => { if (scrollRef.current) scrollRef.current.scrollTop = scrollRef.current.scrollHeight; }, [chat.length, loading]);
  return (
    <div style={{ maxWidth: 760, display: "flex", flexDirection: "column", height: "calc(100vh - 180px)" }}>
      <div className="display" style={{ fontSize: 20, fontWeight: 700, marginBottom: 16 }}>Ask Your Data</div>
      <div ref={scrollRef} style={{ flex: 1, overflow: "auto", marginBottom: 14 }}>
        {chat.length === 0 && (
          <Card><div style={{ color: T.muted, fontSize: 13.5 }}>Ask things like "what are the most important trends?", "which column correlates most with X?", or "are there any anomalies?"</div></Card>
        )}
        {chat.map((m, i) => (
          <div key={i} style={{ display: "flex", justifyContent: m.role === "user" ? "flex-end" : "flex-start", marginBottom: 10 }}>
            <div style={{
              maxWidth: "80%", padding: "10px 14px", borderRadius: 12, fontSize: 13.5, lineHeight: 1.5,
              background: m.role === "user" ? T.chatUserBg : T.surface, border: m.role === "user" ? "none" : `1px solid ${T.border}`,
              color: m.role === "user" ? T.chatUserText : T.text,
            }}>{m.text}</div>
          </div>
        ))}
        {loading && <div style={{ display: "flex", gap: 8, alignItems: "center", color: T.muted, fontSize: 12.5 }}><Loader2 size={14} className="pulse" /> Thinking…</div>}
      </div>
      <div style={{ display: "flex", gap: 10 }}>
        <input value={question} onChange={(e) => setQuestion(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && !loading && askQuestion()}
          placeholder="Ask a question about your data…"
          style={{ flex: 1, background: T.surface, border: `1px solid ${T.border}`, borderRadius: 9, padding: "11px 14px", color: T.text, fontSize: 13.5 }} />
        <button onClick={askQuestion} disabled={loading} style={{ padding: "0 18px", background: T.signal, border: "none", borderRadius: 9, color: T.onSignal }}><Send size={16} /></button>
      </div>
    </div>
  );
}

function HistoryTab({ persistedHistory, analyses, setActiveId, setTab, deletePersisted, clearPersisted, historyLoaded }) {
  const T = useT();
  const liveIds = new Set(analyses.map((a) => a.id));
  return (
    <div style={{ maxWidth: 820 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
        <div>
          <div className="display" style={{ fontSize: 20, fontWeight: 700 }}>Analysis History</div>
          <div style={{ color: T.muted, fontSize: 13, marginTop: 2 }}>Saved on this device, across sessions. Datasets still open in this tab can be reopened directly; others need re-uploading to continue working on them.</div>
        </div>
        {persistedHistory.length > 0 && (
          <button onClick={clearPersisted} style={{ display: "flex", gap: 6, alignItems: "center", padding: "7px 12px", background: "transparent", border: `1px solid ${T.borderLight}`, borderRadius: 7, color: T.muted, fontSize: 12.5, flexShrink: 0 }}>
            <Trash2 size={13} /> Clear all
          </button>
        )}
      </div>
      {!historyLoaded ? (
        <Card><div style={{ color: T.muted, fontSize: 13.5, display: "flex", gap: 8, alignItems: "center" }}><Loader2 size={14} className="pulse" /> Loading history…</div></Card>
      ) : persistedHistory.length === 0 ? (
        <Card><div style={{ color: T.muted, fontSize: 13.5 }}>No past analyses yet — upload a dataset to get started.</div></Card>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {persistedHistory.map((e) => {
            const isLive = liveIds.has(e.id);
            const bullets = (e.insights || "").split("\n").filter((l) => l.trim().startsWith("-")).slice(0, 3);
            return (
              <Card key={e.id}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 8 }}>
                  <div>
                    <div style={{ fontWeight: 600, fontSize: 14 }}>{e.fileName}</div>
                    <div className="mono" style={{ fontSize: 11, color: T.faint, marginTop: 2 }}>{fmtDate(e.timestamp)}</div>
                  </div>
                  <div style={{ display: "flex", gap: 8, flexShrink: 0 }}>
                    {isLive && (
                      <button onClick={() => { setActiveId(e.id); setTab("overview"); }} style={{ padding: "6px 12px", background: T.signal, border: "none", borderRadius: 7, color: T.onSignal, fontSize: 12, fontWeight: 600 }}>Reopen</button>
                    )}
                    <button onClick={() => deletePersisted(e.id)} style={{ padding: "6px 8px", background: "transparent", border: `1px solid ${T.borderLight}`, borderRadius: 7, color: T.faint }}><Trash2 size={13} /></button>
                  </div>
                </div>
                <div className="mono" style={{ display: "flex", gap: 16, fontSize: 11.5, color: T.muted, marginBottom: bullets.length ? 10 : 0 }}>
                  <span>{e.rowCount.toLocaleString()} rows</span>
                  <span>{e.colCount} cols</span>
                  <span style={{ color: e.missing ? T.amber : T.muted }}>{e.missing} missing</span>
                  <span style={{ color: e.duplicates ? T.danger : T.muted }}>{e.duplicates} dupes</span>
                  {e.log?.length > 0 && <span>{e.log.length} cleaning action{e.log.length === 1 ? "" : "s"}</span>}
                </div>
                {bullets.map((b, i) => (
                  <div key={i} style={{ display: "flex", gap: 8, fontSize: 12.5, color: T.text, marginBottom: 4 }}>
                    <ChevronRight size={12} color={T.signal} style={{ flexShrink: 0, marginTop: 3 }} />
                    <span>{b.replace(/^-\s*/, "")}</span>
                  </div>
                ))}
                {!isLive && (
                  <div style={{ fontSize: 11.5, color: T.faint, marginTop: 6, fontStyle: "italic" }}>Re-upload this file to continue analyzing it.</div>
                )}
              </Card>
            );
          })}
        </div>
      )}
    </div>
  );
}
