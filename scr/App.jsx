import React, { useState, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip,
  ResponsiveContainer, LabelList, Legend
} from "recharts";
import { UploadCloud, FileSpreadsheet, AlertTriangle, TrendingUp, TrendingDown, X } from "lucide-react";

// ---------- constants ----------
const INK = "#12213A";
const SLATE = "#5B6B82";
const LINE = "#E1E5EC";
const PAPER = "#F6F7FA";
const GREY_BAR = "#A7AFBC";
const RED_BAR = "#B23A2E";
const NAVY_LINE = "#12213A";

const fmtEUR = (v) =>
  new Intl.NumberFormat("sk-SK", { maximumFractionDigits: 0 }).format(Math.round(v || 0)) + " €";
const fmtNum = (v) => new Intl.NumberFormat("sk-SK").format(Math.round(v || 0));
const fmtPct = (v) => (v * 100).toFixed(1).replace(".", ",") + " %";

// ---------- parsing ----------
function parsePodklad(ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
  if (!rows.length) throw new Error("Záložka PODKLAD je prázdna.");
  const header = rows[0].map((h) => (h == null ? "" : String(h).trim()));
  const idxs = [];
  header.forEach((h, i) => { if (h === "Označenia riadkov") idxs.push(i); });
  if (idxs.length < 2) {
    throw new Error(
      "V záložke PODKLAD sa nenašli dve pivot tabuľky (stĺpce 'Označenia riadkov'). Skontrolujte štruktúru súboru."
    );
  }
  const [ia, ib] = idxs;
  const pivotA = {}, pivotB = {};
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;
    const ka = row[ia];
    if (ka != null && ka !== "" && ka !== "Celkový súčet") {
      pivotA[ka] = { suma: Number(row[ia + 1]) || 0, pocet: Number(row[ia + 2]) || 0 };
    }
    const kb = row[ib];
    if (kb != null && kb !== "" && kb !== "Celkový súčet") {
      pivotB[kb] = { suma: Number(row[ib + 1]) || 0, pocet: Number(row[ib + 2]) || 0 };
    }
  }
  const sumPocet = (p) => Object.values(p).reduce((s, v) => s + v.pocet, 0);
  const [pivotAll, pivotOver45] = sumPocet(pivotA) >= sumPocet(pivotB) ? [pivotA, pivotB] : [pivotB, pivotA];

  const keys = Object.keys(pivotAll).filter((k) => pivotOver45[k]);
  if (!keys.length) throw new Error("Prevádzky sa v oboch pivot tabuľkách nezhodujú.");

  return keys.map((k) => ({
    prevadzka: k,
    sumaAll: pivotAll[k].suma,
    pocetAll: pivotAll[k].pocet,
    suma45: pivotOver45[k].suma,
    pocet45: pivotOver45[k].pocet,
    podiel: pivotAll[k].pocet ? pivotOver45[k].pocet / pivotAll[k].pocet : 0,
  }));
}

function parseTrend(ws) {
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, cellDates: true });
  if (rows.length < 4) return null;
  const dateRow = rows[1] || [];
  const labelRow = rows[2] || [];
  let spoluRow = null;
  for (let r = 3; r < rows.length; r++) {
    if (rows[r] && rows[r][0] === "Spolu") { spoluRow = rows[r]; break; }
  }
  if (!spoluRow) return null;

  const months = [];
  for (let c = 1; c < labelRow.length - 1; c++) {
    if (labelRow[c] === "Suma" && labelRow[c + 1] && String(labelRow[c + 1]).startsWith("Počet")) {
      const dateVal = dateRow[c];
      if (!dateVal) { c++; continue; }
      const suma = Number(spoluRow[c]) || 0;
      const pocet = Number(spoluRow[c + 1]) || 0;
      if (suma === 0 && pocet === 0) { c++; continue; }
      const d = dateVal instanceof Date ? dateVal : new Date(dateVal);
      const label = isNaN(d) ? String(dateVal) : d.toLocaleDateString("sk-SK", { month: "short", year: "numeric" });
      months.push({ label, suma, pocet });
      c++;
    }
  }
  return months.length ? months : null;
}

// ---------- small UI bits ----------
function AngledTick({ x, y, payload }) {
  return (
    <text x={x} y={y + 6} textAnchor="end" fill={SLATE} fontSize={11}
      transform={`rotate(-40 ${x} ${y})`}>
      {payload.value}
    </text>
  );
}

function ValueLabel({ euro }) {
  return (props) => {
    const { x, y, width, value } = props;
    if (value == null) return null;
    return (
      <text x={x + width / 2} y={y - 6} textAnchor="middle" fontSize={10} fill={INK}
        transform={`rotate(-45 ${x + width / 2} ${y - 6})`}>
        {euro ? fmtEUR(value) : fmtNum(value)}
      </text>
    );
  };
}

function StatCard({ label, value, sub, tone = "ink" }) {
  return (
    <div style={{
      flex: 1, minWidth: 180, background: "#fff", border: `1px solid ${LINE}`,
      borderRadius: 4, padding: "16px 18px",
    }}>
      <div style={{ fontSize: 11, color: SLATE, letterSpacing: 0.3, marginBottom: 6 }}>{label}</div>
      <div style={{ fontSize: 24, fontWeight: 700, color: tone === "red" ? RED_BAR : INK }}>{value}</div>
      {sub && <div style={{ fontSize: 12, color: SLATE, marginTop: 4 }}>{sub}</div>}
    </div>
  );
}

// ---------- main component ----------
export default function App() {
  const [fileName, setFileName] = useState(null);
  const [rows, setRows] = useState(null);
  const [trend, setTrend] = useState(null);
  const [sourceSheet, setSourceSheet] = useState("");
  const [error, setError] = useState(null);
  const [loading, setLoading] = useState(false);
  const [dragOver, setDragOver] = useState(false);
  const [sortBy, setSortBy] = useState("pocet"); // 'pocet' | 'suma'

  const handleFile = useCallback((file) => {
    if (!file) return;
    setLoading(true);
    setError(null);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: "array", cellDates: true });
        const names = wb.SheetNames;
        const podkladName = names.find((n) => /podklad/i.test(n));
        if (!podkladName) {
          throw new Error("V súbore sa nenašla záložka obsahujúca 'PODKLAD'.");
        }
        const parsedRows = parsePodklad(wb.Sheets[podkladName]);
        setRows(parsedRows);
        setSourceSheet(podkladName);

        const trendName = names.find((n) => /nad 45.*od|ha-?nad 45/i.test(n));
        setTrend(trendName ? parseTrend(wb.Sheets[trendName]) : null);

        setFileName(file.name);
      } catch (err) {
        setError(err.message || "Súbor sa nepodarilo spracovať.");
        setRows(null);
        setTrend(null);
      } finally {
        setLoading(false);
      }
    };
    reader.onerror = () => { setError("Súbor sa nepodarilo načítať."); setLoading(false); };
    reader.readAsArrayBuffer(file);
  }, []);

  const sorted = useMemo(() => {
    if (!rows) return [];
    const key = sortBy === "pocet" ? "pocet45" : "suma45";
    return [...rows].sort((a, b) => b[key] - a[key]);
  }, [rows, sortBy]);

  const totals = useMemo(() => {
    if (!rows) return null;
    return rows.reduce(
      (acc, r) => ({
        sumaAll: acc.sumaAll + r.sumaAll,
        pocetAll: acc.pocetAll + r.pocetAll,
        suma45: acc.suma45 + r.suma45,
        pocet45: acc.pocet45 + r.pocet45,
      }),
      { sumaAll: 0, pocetAll: 0, suma45: 0, pocet45: 0 }
    );
  }, [rows]);

  const reset = () => { setRows(null); setTrend(null); setFileName(null); setError(null); };

  return (
    <div style={{
      minHeight: "100%", background: PAPER, fontFamily:
        "'Inter', -apple-system, 'Segoe UI', sans-serif", color: INK,
    }}>
      {/* header */}
      <div style={{ background: INK, color: "#fff", padding: "22px 28px" }}>
        <div style={{ fontSize: 18, fontWeight: 700, letterSpacing: 0.2 }}>
          Otvorené objednávky — manažérsky prehľad
        </div>
        <div style={{ fontSize: 13, color: "#AEB8CB", marginTop: 4 }}>
          Nahrajte podkladový xlsx súbor a dashboard automaticky pripraví grafy podľa prevádzok a vývoj v čase.
        </div>
      </div>

      <div style={{ maxWidth: 1080, margin: "0 auto", padding: "28px 24px 60px" }}>
        {/* upload zone */}
        {!rows && (
          <div
            onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
            onDragLeave={() => setDragOver(false)}
            onDrop={(e) => { e.preventDefault(); setDragOver(false); handleFile(e.dataTransfer.files[0]); }}
            style={{
              border: `1.5px dashed ${dragOver ? INK : "#C3CADA"}`,
              borderRadius: 6, background: dragOver ? "#EDEFF4" : "#fff",
              padding: "48px 24px", textAlign: "center", transition: "all .15s",
            }}
          >
            <UploadCloud size={30} color={SLATE} style={{ marginBottom: 10 }} />
            <div style={{ fontSize: 15, fontWeight: 600, marginBottom: 4 }}>
              Presuňte sem xlsx súbor alebo ho vyberte
            </div>
            <div style={{ fontSize: 12.5, color: SLATE, marginBottom: 16 }}>
              Očakávaná štruktúra: záložka s "PODKLAD" (otvorené objednávky) a voliteľne záložka "…nad 45 dní od…" pre vývoj v čase.
            </div>
            <label style={{
              display: "inline-block", background: INK, color: "#fff", padding: "9px 18px",
              borderRadius: 4, fontSize: 13, fontWeight: 600, cursor: "pointer",
            }}>
              Vybrať súbor
              <input type="file" accept=".xlsx" style={{ display: "none" }}
                onChange={(e) => handleFile(e.target.files[0])} />
            </label>
            {loading && <div style={{ marginTop: 14, fontSize: 12.5, color: SLATE }}>Spracúvam súbor…</div>}
            {error && (
              <div style={{
                marginTop: 16, display: "inline-flex", gap: 8, alignItems: "flex-start", textAlign: "left",
                background: "#FBEAE8", border: "1px solid #EFC3BE", borderRadius: 4, padding: "10px 14px",
                color: "#7A2A20", fontSize: 12.5, maxWidth: 480,
              }}>
                <AlertTriangle size={16} style={{ flexShrink: 0, marginTop: 1 }} />
                <span>{error}</span>
              </div>
            )}
          </div>
        )}

        {rows && (
          <>
            {/* file bar */}
            <div style={{
              display: "flex", alignItems: "center", justifyContent: "space-between",
              background: "#fff", border: `1px solid ${LINE}`, borderRadius: 4,
              padding: "10px 16px", marginBottom: 20, fontSize: 12.5,
            }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, color: SLATE }}>
                <FileSpreadsheet size={15} />
                <span>{fileName}</span>
                <span style={{ color: "#C3CADA" }}>·</span>
                <span>záložka „{sourceSheet}"</span>
              </div>
              <button onClick={reset} style={{
                display: "flex", alignItems: "center", gap: 4, background: "none", border: "none",
                color: SLATE, fontSize: 12.5, cursor: "pointer",
              }}>
                <X size={14} /> Nahrať iný súbor
              </button>
            </div>

            {/* stat cards */}
            <div style={{ display: "flex", gap: 14, flexWrap: "wrap", marginBottom: 28 }}>
              <StatCard label="VŠETKY OTVORENÉ — POČET" value={fmtNum(totals.pocetAll)} />
              <StatCard label="VŠETKY OTVORENÉ — SUMA" value={fmtEUR(totals.sumaAll)} />
              <StatCard label="NAD 45 DNÍ — POČET" value={fmtNum(totals.pocet45)} tone="red"
                sub={`${fmtPct(totals.pocet45 / totals.pocetAll)} zo všetkých`} />
              <StatCard label="NAD 45 DNÍ — SUMA" value={fmtEUR(totals.suma45)} tone="red"
                sub={`${fmtPct(totals.suma45 / totals.sumaAll)} zo všetkých`} />
            </div>

            {/* sort toggle */}
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
              <span style={{ fontSize: 12.5, color: SLATE }}>Zoradiť podľa:</span>
              {[["pocet", "počtu"], ["suma", "sumy"]].map(([val, lbl]) => (
                <button key={val} onClick={() => setSortBy(val)} style={{
                  fontSize: 12.5, padding: "5px 12px", borderRadius: 20, cursor: "pointer",
                  border: `1px solid ${sortBy === val ? INK : LINE}`,
                  background: sortBy === val ? INK : "#fff",
                  color: sortBy === val ? "#fff" : SLATE, fontWeight: 600,
                }}>
                  {lbl}
                </button>
              ))}
              <span style={{ fontSize: 12, color: SLATE }}>(zostupne, od najväčšej po najmenšiu)</span>
            </div>

            {/* bar chart: count or sum */}
            <ChartCard
              title={sortBy === "pocet"
                ? "Počet otvorených objednávok vs. nad 45 dní podľa prevádzok"
                : "Suma otvorených objednávok vs. nad 45 dní podľa prevádzok"}
            >
              <ResponsiveContainer width="100%" height={420}>
                <BarChart data={sorted} margin={{ top: 30, right: 20, left: 10, bottom: 70 }} barGap={4}>
                  <CartesianGrid vertical={false} stroke={LINE} />
                  <XAxis dataKey="prevadzka" interval={0} height={70} tick={<AngledTick />} />
                  <YAxis tick={{ fontSize: 11, fill: SLATE }}
                    tickFormatter={sortBy === "pocet" ? fmtNum : (v) => fmtEUR(v)} width={70} />
                  <Tooltip formatter={(v) => (sortBy === "pocet" ? fmtNum(v) : fmtEUR(v))} />
                  <Legend wrapperStyle={{ fontSize: 12.5 }} />
                  <Bar dataKey={sortBy === "pocet" ? "pocetAll" : "sumaAll"} name="Všetky otvorené" fill={GREY_BAR} radius={[2, 2, 0, 0]}>
                    <LabelList content={ValueLabel({ euro: sortBy !== "pocet" })} />
                  </Bar>
                  <Bar dataKey={sortBy === "pocet" ? "pocet45" : "suma45"} name="Nad 45 dní" fill={RED_BAR} radius={[2, 2, 0, 0]}>
                    <LabelList content={ValueLabel({ euro: sortBy !== "pocet" })} />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </ChartCard>

            {/* table */}
            <ChartCard title="Podklad podľa prevádzok">
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
                  <thead>
                    <tr style={{ background: INK, color: "#fff" }}>
                      {["Prevádzka", "Suma všetkých (€)", "Počet všetkých", "Suma nad 45 dní (€)", "Počet nad 45 dní", "Podiel nad 45 dní"].map((h) => (
                        <th key={h} style={{ padding: "8px 10px", textAlign: "left", fontWeight: 600 }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {sorted.map((r, i) => (
                      <tr key={r.prevadzka} style={{ background: i % 2 ? "#fff" : "#FAFBFC", borderBottom: `1px solid ${LINE}` }}>
                        <td style={{ padding: "7px 10px", fontWeight: 600 }}>{r.prevadzka}</td>
                        <td style={{ padding: "7px 10px" }}>{fmtEUR(r.sumaAll)}</td>
                        <td style={{ padding: "7px 10px" }}>{fmtNum(r.pocetAll)}</td>
                        <td style={{ padding: "7px 10px", color: RED_BAR, fontWeight: 600 }}>{fmtEUR(r.suma45)}</td>
                        <td style={{ padding: "7px 10px", color: RED_BAR, fontWeight: 600 }}>{fmtNum(r.pocet45)}</td>
                        <td style={{ padding: "7px 10px" }}>{fmtPct(r.podiel)}</td>
                      </tr>
                    ))}
                  </tbody>
                  <tfoot>
                    <tr style={{ fontWeight: 700, borderTop: `2px solid ${INK}` }}>
                      <td style={{ padding: "8px 10px" }}>Spolu</td>
                      <td style={{ padding: "8px 10px" }}>{fmtEUR(totals.sumaAll)}</td>
                      <td style={{ padding: "8px 10px" }}>{fmtNum(totals.pocetAll)}</td>
                      <td style={{ padding: "8px 10px", color: RED_BAR }}>{fmtEUR(totals.suma45)}</td>
                      <td style={{ padding: "8px 10px", color: RED_BAR }}>{fmtNum(totals.pocet45)}</td>
                      <td style={{ padding: "8px 10px" }}>{fmtPct(totals.pocet45 / totals.pocetAll)}</td>
                    </tr>
                  </tfoot>
                </table>
              </div>
            </ChartCard>

            {/* trend section */}
            {trend && trend.length > 0 && (
              <>
                <div style={{ fontSize: 15, fontWeight: 700, margin: "36px 0 14px" }}>
                  Vývoj za celú skupinu od 1.1.2026
                </div>
                <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
                  <ChartCard title="Vývoj počtu objednávok nad 45 dní" flex>
                    <ResponsiveContainer width="100%" height={260}>
                      <LineChart data={trend} margin={{ top: 24, right: 20, left: 0, bottom: 8 }}>
                        <CartesianGrid vertical={false} stroke={LINE} />
                        <XAxis dataKey="label" tick={{ fontSize: 11.5, fill: SLATE }} />
                        <YAxis tick={{ fontSize: 11, fill: SLATE }} tickFormatter={fmtNum} width={50} />
                        <Tooltip formatter={(v) => fmtNum(v)} />
                        <Line type="monotone" dataKey="pocet" name="Počet" stroke={RED_BAR} strokeWidth={2.5} dot={{ r: 4 }}>
                          <LabelList dataKey="pocet" position="top" fontSize={11} formatter={fmtNum} />
                        </Line>
                      </LineChart>
                    </ResponsiveContainer>
                  </ChartCard>
                  <ChartCard title="Vývoj sumy objednávok nad 45 dní" flex>
                    <ResponsiveContainer width="100%" height={260}>
                      <LineChart data={trend} margin={{ top: 24, right: 20, left: 0, bottom: 8 }}>
                        <CartesianGrid vertical={false} stroke={LINE} />
                        <XAxis dataKey="label" tick={{ fontSize: 11.5, fill: SLATE }} />
                        <YAxis tick={{ fontSize: 11, fill: SLATE }} tickFormatter={(v) => fmtEUR(v)} width={70} />
                        <Tooltip formatter={(v) => fmtEUR(v)} />
                        <Line type="monotone" dataKey="suma" name="Suma" stroke={NAVY_LINE} strokeWidth={2.5} dot={{ r: 4 }}>
                          <LabelList dataKey="suma" position="top" fontSize={11} formatter={fmtEUR} />
                        </Line>
                      </LineChart>
                    </ResponsiveContainer>
                  </ChartCard>
                </div>
              </>
            )}
            {!trend && (
              <div style={{ marginTop: 30, fontSize: 12.5, color: SLATE, display: "flex", gap: 6, alignItems: "center" }}>
                <TrendingUp size={14} /> Záložka s mesačným vývojom (napr. „…nad 45 dní od…") sa v súbore nenašla — sekcia vývoja bola vynechaná.
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
}

function ChartCard({ title, children, flex }) {
  return (
    <div style={{
      background: "#fff", border: `1px solid ${LINE}`, borderRadius: 4,
      padding: "18px 20px 10px", marginBottom: 20, flex: flex ? "1 1 420px" : undefined,
    }}>
      <div style={{ fontSize: 13.5, fontWeight: 700, marginBottom: 6 }}>{title}</div>
      {children}
    </div>
  );
}
