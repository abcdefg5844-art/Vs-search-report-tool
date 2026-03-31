import { useState, useCallback, useRef, useMemo } from "react";
import { AreaChart, Area, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Legend, BarChart, Bar } from "recharts";

const HISTORY = [
  { month: "2023-11월", members: 327, saves: 398 },
  { month: "2023-12월", members: 375, saves: 490 },
  { month: "2024-1월", members: 463, saves: 633 },
  { month: "2024-2월", members: 520, saves: 686 },
  { month: "2024-3월", members: 588, saves: 795 },
  { month: "2024-4월", members: 490, saves: 699 },
  { month: "2024-5월", members: 465, saves: 669 },
  { month: "2024-6월", members: 397, saves: 524 },
  { month: "2024-7월", members: 444, saves: 648 },
  { month: "2024-8월", members: 493, saves: 635 },
  { month: "2024-9월", members: 421, saves: 539 },
  { month: "2024-10월", members: 424, saves: 553 },
  { month: "2024-11월", members: 464, saves: 605 },
  { month: "2024-12월", members: 428, saves: 556 },
  { month: "2025-1월", members: 487, saves: 643 },
  { month: "2025-2월", members: 537, saves: 685 },
  { month: "2025-3월", members: 491, saves: 659 },
  { month: "2025-4월", members: 429, saves: 585 },
  { month: "2025-5월", members: 551, saves: 742 },
  { month: "2025-6월", members: 469, saves: 608 },
  { month: "2025-7월", members: 538, saves: 723 },
  { month: "2025-8월", members: 463, saves: 587 },
  { month: "2025-9월", members: 466, saves: 632 },
  { month: "2025-10월", members: 405, saves: 532 },
  { month: "2025-11월", members: 411, saves: 550 },
  { month: "2025-12월", members: 378, saves: 477 },
  { month: "2026-1월", members: 454, saves: 605 },
  { month: "2026-2월", members: 385, saves: 514 },
];

function fmt(n) { if (n == null || isNaN(n)) return "-"; return Number(n).toLocaleString(); }
function signFmt(n) { if (n == null || isNaN(n)) return "-"; return (n > 0 ? "+" : "") + Number(n).toLocaleString(); }
function pctStr(cur, prev) {
  if (!prev || prev === 0) return "-";
  const p = ((cur - prev) / prev * 100).toFixed(0);
  return (p > 0 ? "+" : "") + p + "%";
}

function parseXlsx(arrayBuffer) {
  return new Promise((resolve) => {
    const doWork = () => {
      const wb = window.XLSX.read(arrayBuffer, { type: "array" });
      const sheetName = wb.SheetNames.find(s => s !== "Sheet1") || wb.SheetNames[0];
      const data = window.XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1 });
      resolve(data);
    };
    if (window.XLSX) { doWork(); return; }
    const s = document.createElement("script");
    s.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
    s.onload = doWork;
    document.head.appendChild(s);
  });
}

function processPvData(rows) {
  const dataRows = rows.filter(r => r[0] && r[1] && typeof r[2] === "number");
  let pcTotal = 0, moTotal = 0;
  const days = new Set();
  const daily = {};
  dataRows.forEach(r => {
    const d = r[0] instanceof Date ? r[0].toISOString().slice(0, 10) : String(r[0]).slice(0, 10);
    const plat = String(r[1]).toLowerCase();
    const cnt = Number(r[2]);
    days.add(d);
    if (!daily[d]) daily[d] = { date: d, pc: 0, mo: 0 };
    if (plat === "pc") { pcTotal += cnt; daily[d].pc = cnt; }
    else { moTotal += cnt; daily[d].mo = cnt; }
  });
  const n = days.size;
  const sortedDates = [...days].sort();
  const dailyArr = sortedDates.map(d => ({ ...daily[d], total: daily[d].pc + daily[d].mo }));
  return {
    pcPvSum: pcTotal, moPvSum: moTotal, pvTotal: pcTotal + moTotal,
    pcPvAvg: n ? Math.round(pcTotal / n) : 0,
    moPvAvg: n ? Math.round(moTotal / n) : 0,
    numDays: n, dailyArr,
    monthLabel: sortedDates[0]?.slice(0, 7) || "",
  };
}

function parsePastedData(text) {
  if (!text.trim()) return null;
  const lines = text.trim().split("\n").map(l => l.trim()).filter(Boolean);
  const dailyData = [];
  let total = null, cumulative = null;
  for (const line of lines) {
    const parts = line.split(/\t+|\s{2,}/);
    if (parts.length < 2) continue;
    const label = parts[0].trim();
    const rawVal = parts[parts.length - 1].replace(/,/g, "").trim();
    const value = parseInt(rawVal);
    if (isNaN(value)) continue;
    if (/^\d{4}-\d{2}-\d{2}$/.test(label)) dailyData.push({ date: label, count: value });
    else if (/합계/i.test(label)) total = value;
    else if (/전체|누적/i.test(label)) cumulative = value;
  }
  if (dailyData.length === 0 && total == null) return null;
  return { dailyData, total: total ?? dailyData.reduce((s, d) => s + d.count, 0), cumulative, count: dailyData.length };
}

function toMonthStr(label) {
  if (!label) return "";
  return `${label.slice(0, 4)}년 ${parseInt(label.slice(5, 7))}월`;
}

const tdS = { padding: "8px 10px", textAlign: "right", borderBottom: "1px solid var(--border)", whiteSpace: "nowrap", fontSize: 12 };
const thS = { ...tdS, color: "var(--sub)", fontWeight: 500, background: "var(--bg)" };

const CustomTooltip = ({ active, payload, label }) => {
  if (!active || !payload?.length) return null;
  return (
    <div style={{ background: "#1a1a1e", border: "1px solid #333", borderRadius: 8, padding: "10px 14px", fontSize: 12 }}>
      <div style={{ fontWeight: 600, marginBottom: 6, color: "#e4e4e7" }}>{label}</div>
      {payload.map((p, i) => (
        <div key={i} style={{ color: p.color, marginBottom: 2 }}>
          {p.name}: {fmt(p.value)}
        </div>
      ))}
    </div>
  );
};

export default function App() {
  const [prevPv, setPrevPv] = useState(null);
  const [curPv, setCurPv] = useState(null);
  const [prevFileName, setPrevFileName] = useState("");
  const [curFileName, setCurFileName] = useState("");
  const [dragOverPrev, setDragOverPrev] = useState(false);
  const [dragOverCur, setDragOverCur] = useState(false);
  const [membersText, setMembersText] = useState("");
  const [savesText, setSavesText] = useState("");
  const [copied, setCopied] = useState(false);
  const prevFileRef = useRef();
  const curFileRef = useRef();

  const loadFile = useCallback(async (file, setter, nameSetter) => {
    if (!file) return;
    nameSetter(file.name);
    const buf = await file.arrayBuffer();
    const rows = await parseXlsx(buf);
    setter(processPvData(rows));
  }, []);

  const membersData = parsePastedData(membersText);
  const savesData = parsePastedData(savesText);

  const prevMonthStr = prevPv ? toMonthStr(prevPv.monthLabel) : "";
  const curMonthStr = curPv ? toMonthStr(curPv.monthLabel) : "";

  const pcDiff = prevPv && curPv ? curPv.pcPvSum - prevPv.pcPvSum : null;
  const moDiff = prevPv && curPv ? curPv.moPvSum - prevPv.moPvSum : null;
  const pvDiff = prevPv && curPv ? curPv.pvTotal - prevPv.pvTotal : null;

  // Find prev month members/saves from HISTORY based on prev file month
  const prevMembersSaves = useMemo(() => {
    if (!prevPv) return HISTORY[HISTORY.length - 1];
    const m = parseInt(prevPv.monthLabel?.slice(5, 7));
    const y = prevPv.monthLabel?.slice(0, 4);
    const key = `${y}-${m}월`;
    return HISTORY.find(h => h.month === key) || HISTORY[HISTORY.length - 1];
  }, [prevPv]);

  const membersTotal = membersData?.total;
  const savesTotal = savesData?.total;
  const membersDiff = membersTotal != null ? membersTotal - prevMembersSaves.members : null;
  const savesDiff = savesTotal != null ? savesTotal - prevMembersSaves.saves : null;

  // Chart data: combine prev + cur daily
  const chartData = useMemo(() => {
    const all = [];
    if (prevPv) {
      prevPv.dailyArr.forEach(d => all.push({
        date: d.date.slice(5),
        fullDate: d.date,
        prevPc: d.pc, prevMo: d.mo,
      }));
    }
    if (curPv) {
      curPv.dailyArr.forEach(d => {
        const existing = all.find(a => a.date === d.date.slice(5));
        if (existing) {
          existing.curPc = d.pc; existing.curMo = d.mo;
        } else {
          all.push({ date: d.date.slice(5), fullDate: d.date, curPc: d.pc, curMo: d.mo });
        }
      });
    }
    return all.sort((a, b) => a.date.localeCompare(b.date));
  }, [prevPv, curPv]);

  const generateReport = () => {
    if (!curPv) return "";
    const pL = prevMonthStr || "전월";
    const cL = curMonthStr;

    let r = `### 1. VS검색페이지 PV (PC/MO 구분, uxlog 기준)\n\n`;
    r += `| 날짜 | PV 합계 | PC PV 합계 | PC PV 합계 전월비교 | PC PV 일평균 | MO PV 합계 | MO PV 합계 전월비교 | MO PV 일평균 |\n`;
    r += `|------|---------|-----------|-------------------|------------|-----------|-------------------|------------|\n`;
    if (prevPv) {
      r += `| ${pL} | ${fmt(prevPv.pvTotal)} | ${fmt(prevPv.pcPvSum)} | | ${fmt(prevPv.pcPvAvg)} | ${fmt(prevPv.moPvSum)} | | ${fmt(prevPv.moPvAvg)} |\n`;
    }
    r += `| ${cL} | ${fmt(curPv.pvTotal)} | ${fmt(curPv.pcPvSum)} | ${pcDiff != null ? signFmt(pcDiff) : ""} | ${fmt(curPv.pcPvAvg)} | ${fmt(curPv.moPvSum)} | ${moDiff != null ? signFmt(moDiff) : ""} | ${fmt(curPv.moPvAvg)} |\n\n`;

    r += `### 2. VS검색 저장 회원수 및 저장 건수 (DB 기준)\n\n`;
    r += `| 날짜 | 저장 회원수 | 저장 회원수 전월비교 | 저장건수 | 저장건수 전월비교 |\n`;
    r += `|------|-----------|-------------------|---------|----------------|\n`;
    r += `| ${prevMembersSaves.month.replace("-", "년 ")} | ${fmt(prevMembersSaves.members)} | | ${fmt(prevMembersSaves.saves)} | |\n`;
    r += `| ${cL} | ${membersTotal != null ? fmt(membersTotal) : ""} | ${membersDiff != null ? signFmt(membersDiff) : ""} | ${savesTotal != null ? fmt(savesTotal) : ""} | ${savesDiff != null ? signFmt(savesDiff) : ""} |\n`;
    return r;
  };

  const copyReport = () => {
    navigator.clipboard.writeText(generateReport());
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  return (
    <div style={{
      minHeight: "100vh", fontFamily: "'Pretendard', -apple-system, sans-serif",
      background: "var(--bg)", color: "var(--text)", padding: "32px 20px",
    }}>
      <style>{`
        @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
        :root {
          --bg: #0a0a0b; --card: #161618; --border: #2a2a2d;
          --accent: #4f8cff; --accent2: #36d399; --red: #ff6b6b;
          --text: #e4e4e7; --sub: #8b8b95;
        }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        textarea:focus { outline: 2px solid var(--accent); outline-offset: -2px; }
        textarea { resize: vertical; }
        .recharts-cartesian-grid-horizontal line, .recharts-cartesian-grid-vertical line { stroke: #222 !important; }
      `}</style>

      <div style={{ maxWidth: 920, margin: "0 auto" }}>
        <div style={{ marginBottom: 32 }}>
          <h1 style={{ fontSize: 22, fontWeight: 700, letterSpacing: -0.5 }}>VS검색 월별 지표 자동화</h1>
          <p style={{ color: "var(--sub)", fontSize: 13, marginTop: 6 }}>
            전월/당월 PV 파일 업로드 + 회원수/저장건수 복붙 → 비교 분석 + 노션 보고 텍스트
          </p>
        </div>

        {/* STEP 1: Two file uploads */}
        <Card step="1" color="var(--accent)" title="PV Count 파일 업로드">
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
            <FileDropZone
              label="📂 전월 파일"
              fileName={prevFileName}
              result={prevPv}
              dragOver={dragOverPrev}
              setDragOver={setDragOverPrev}
              fileRef={prevFileRef}
              onFile={(f) => loadFile(f, setPrevPv, setPrevFileName)}
              accent="#a78bfa"
            />
            <FileDropZone
              label="📊 당월 파일"
              fileName={curFileName}
              result={curPv}
              dragOver={dragOverCur}
              setDragOver={setDragOverCur}
              fileRef={curFileRef}
              onFile={(f) => loadFile(f, setCurPv, setCurFileName)}
              accent="var(--accent)"
            />
          </div>

          {/* Summary cards */}
          {curPv && (
            <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 10, marginTop: 16 }}>
              {[
                { label: "PV 합계", cur: curPv.pvTotal, prev: prevPv?.pvTotal },
                { label: "PC PV", cur: curPv.pcPvSum, prev: prevPv?.pcPvSum, avg: curPv.pcPvAvg },
                { label: "MO PV", cur: curPv.moPvSum, prev: prevPv?.moPvSum, avg: curPv.moPvAvg },
              ].map((it, i) => {
                const diff = it.prev != null ? it.cur - it.prev : null;
                return (
                  <div key={i} style={{ background: "var(--bg)", borderRadius: 8, padding: "12px 14px" }}>
                    <div style={{ fontSize: 11, color: "var(--sub)", marginBottom: 4 }}>{it.label}</div>
                    <div style={{ fontSize: 18, fontWeight: 700 }}>{fmt(it.cur)}</div>
                    {diff != null && (
                      <div style={{ fontSize: 12, color: diff >= 0 ? "var(--accent2)" : "var(--red)" }}>
                        전월 {signFmt(diff)} ({pctStr(it.cur, it.prev)})
                      </div>
                    )}
                    {it.avg != null && (
                      <div style={{ fontSize: 11, color: "var(--sub)", marginTop: 2 }}>일평균 {fmt(it.avg)}</div>
                    )}
                  </div>
                );
              })}
            </div>
          )}
        </Card>

        {/* CHART */}
        {chartData.length > 0 && (
          <Card step="" color="transparent" title="📈 PC / MO PV 일별 추이" noStep>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
              <div>
                <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 10, color: "#60a5fa" }}>PC PV</div>
                <ResponsiveContainer width="100%" height={220}>
                  <AreaChart data={chartData} margin={{ top: 5, right: 10, left: 0, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" />
                    <XAxis dataKey="date" tick={{ fontSize: 10, fill: "#666" }} interval="preserveStartEnd" />
                    <YAxis tick={{ fontSize: 10, fill: "#666" }} tickFormatter={v => (v / 1000).toFixed(0) + "k"} width={40} />
                    <Tooltip content={<CustomTooltip />} />
                    {prevPv && <Area type="monotone" dataKey="prevPc" name={prevMonthStr} stroke="#6b7280" fill="#6b728022" strokeWidth={1.5} dot={false} />}
                    {curPv && <Area type="monotone" dataKey="curPc" name={curMonthStr} stroke="#60a5fa" fill="#60a5fa18" strokeWidth={2} dot={false} />}
                    <Legend wrapperStyle={{ fontSize: 11 }} />
                  </AreaChart>
                </ResponsiveContainer>
              </div>
              <div>
                <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 10, color: "#f472b6" }}>MO PV</div>
                <ResponsiveContainer width="100%" height={220}>
                  <AreaChart data={chartData} margin={{ top: 5, right: 10, left: 0, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" />
                    <XAxis dataKey="date" tick={{ fontSize: 10, fill: "#666" }} interval="preserveStartEnd" />
                    <YAxis tick={{ fontSize: 10, fill: "#666" }} tickFormatter={v => (v / 1000).toFixed(0) + "k"} width={40} />
                    <Tooltip content={<CustomTooltip />} />
                    {prevPv && <Area type="monotone" dataKey="prevMo" name={prevMonthStr} stroke="#6b7280" fill="#6b728022" strokeWidth={1.5} dot={false} />}
                    {curPv && <Area type="monotone" dataKey="curMo" name={curMonthStr} stroke="#f472b6" fill="#f472b618" strokeWidth={2} dot={false} />}
                    <Legend wrapperStyle={{ fontSize: 11 }} />
                  </AreaChart>
                </ResponsiveContainer>
              </div>
            </div>

            {/* Combined bar chart */}
            {prevPv && curPv && (
              <div style={{ marginTop: 20 }}>
                <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 10, color: "var(--text)" }}>전월 vs 당월 총합 비교</div>
                <ResponsiveContainer width="100%" height={180}>
                  <BarChart data={[
                    { name: "PC PV", prev: prevPv.pcPvSum, cur: curPv.pcPvSum },
                    { name: "MO PV", prev: prevPv.moPvSum, cur: curPv.moPvSum },
                    { name: "PV 합계", prev: prevPv.pvTotal, cur: curPv.pvTotal },
                  ]} margin={{ top: 5, right: 20, left: 10, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" />
                    <XAxis dataKey="name" tick={{ fontSize: 12, fill: "#999" }} />
                    <YAxis tick={{ fontSize: 10, fill: "#666" }} tickFormatter={v => (v / 1000).toFixed(0) + "k"} />
                    <Tooltip content={<CustomTooltip />} />
                    <Bar dataKey="prev" name={prevMonthStr} fill="#6b7280" radius={[4, 4, 0, 0]} />
                    <Bar dataKey="cur" name={curMonthStr} fill="var(--accent)" radius={[4, 4, 0, 0]} />
                    <Legend wrapperStyle={{ fontSize: 11 }} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            )}
          </Card>
        )}

        {/* STEP 2: Members & Saves */}
        <Card step="2" color="var(--accent2)" title="저장 회원수 / 저장건수 붙여넣기">
          <p style={{ fontSize: 12, color: "var(--sub)", marginBottom: 12, lineHeight: 1.6 }}>
            DB에서 일별 데이터를 복사해서 각각 붙여넣으세요.
            <code style={{ background: "var(--bg)", padding: "2px 6px", borderRadius: 4, marginLeft: 4 }}>
              2026-03-01	5
            </code> 형식. 합계/전체 행도 자동 인식.
            <span style={{ color: "var(--sub)", marginLeft: 4 }}>
              (전월 기준: {prevMembersSaves.month.replace("-", "년 ")})
            </span>
          </p>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
            <PasteBox label="👤 회원수" value={membersText} onChange={setMembersText}
              placeholder={"2026-03-01\t5\n2026-03-02\t16\n...\n합계\t415\n저장된 전체 회원수\t9103"}
              parsed={membersData} diff={membersDiff} prevLabel={`전월(${prevMembersSaves.members})`}
            />
            <PasteBox label="📌 저장건수 (VS검색)" value={savesText} onChange={setSavesText}
              placeholder={"2026-03-01\t9\n2026-03-02\t21\n...\n합계\t581\n저장된 전체 VS검색수\t16718"}
              parsed={savesData} diff={savesDiff} prevLabel={`전월(${prevMembersSaves.saves})`}
            />
          </div>
        </Card>

        {/* STEP 3: Report */}
        {curPv && (
          <Card step="3" color="#ff9f43" title="노션 보고용 결과"
            right={
              <button onClick={copyReport} style={{
                padding: "7px 16px", borderRadius: 8, border: "none",
                background: copied ? "var(--accent2)" : "var(--accent)",
                color: "#fff", fontSize: 13, fontWeight: 600, cursor: "pointer",
                transition: "background 0.2s",
              }}>
                {copied ? "✓ 복사됨" : "📋 마크다운 복사"}
              </button>
            }>
            <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 8 }}>
              1. VS검색페이지 PV (PC/MO 구분, uxlog 기준)
            </div>
            <div style={{ overflowX: "auto", marginBottom: 20 }}>
              <table style={{ width: "100%", borderCollapse: "collapse" }}>
                <thead><tr>
                  {["날짜","PV 합계","PC PV 합계","PC PV 전월비교","PC PV 일평균","MO PV 합계","MO PV 전월비교","MO PV 일평균"].map((h,i) =>
                    <th key={i} style={thS}>{h}</th>
                  )}
                </tr></thead>
                <tbody>
                  {prevPv && (
                    <tr>
                      <td style={tdS}>{prevMonthStr}</td>
                      <td style={tdS}>{fmt(prevPv.pvTotal)}</td>
                      <td style={tdS}>{fmt(prevPv.pcPvSum)}</td>
                      <td style={tdS}></td>
                      <td style={tdS}>{fmt(prevPv.pcPvAvg)}</td>
                      <td style={tdS}>{fmt(prevPv.moPvSum)}</td>
                      <td style={tdS}></td>
                      <td style={tdS}>{fmt(prevPv.moPvAvg)}</td>
                    </tr>
                  )}
                  <tr style={{ fontWeight: 700 }}>
                    <td style={tdS}>{curMonthStr}</td>
                    <td style={tdS}>{fmt(curPv.pvTotal)}</td>
                    <td style={tdS}>{fmt(curPv.pcPvSum)}</td>
                    <td style={{ ...tdS, color: pcDiff != null ? (pcDiff >= 0 ? "var(--accent2)" : "var(--red)") : "inherit" }}>
                      {pcDiff != null ? signFmt(pcDiff) : ""}
                    </td>
                    <td style={tdS}>{fmt(curPv.pcPvAvg)}</td>
                    <td style={tdS}>{fmt(curPv.moPvSum)}</td>
                    <td style={{ ...tdS, color: moDiff != null ? (moDiff >= 0 ? "var(--accent2)" : "var(--red)") : "inherit" }}>
                      {moDiff != null ? signFmt(moDiff) : ""}
                    </td>
                    <td style={tdS}>{fmt(curPv.moPvAvg)}</td>
                  </tr>
                </tbody>
              </table>
            </div>

            <div style={{ fontSize: 13, fontWeight: 600, marginBottom: 8 }}>
              2. VS검색 저장 회원수 및 저장 건수 (DB 기준)
            </div>
            <div style={{ overflowX: "auto" }}>
              <table style={{ width: "100%", borderCollapse: "collapse" }}>
                <thead><tr>
                  {["날짜","저장 회원수","전월비교","저장건수","전월비교"].map((h,i) =>
                    <th key={i} style={thS}>{h}</th>
                  )}
                </tr></thead>
                <tbody>
                  <tr>
                    <td style={tdS}>{prevMembersSaves.month.replace("-","년 ")}</td>
                    <td style={tdS}>{fmt(prevMembersSaves.members)}</td>
                    <td style={tdS}></td>
                    <td style={tdS}>{fmt(prevMembersSaves.saves)}</td>
                    <td style={tdS}></td>
                  </tr>
                  <tr style={{ fontWeight: 700 }}>
                    <td style={tdS}>{curMonthStr}</td>
                    <td style={tdS}>{membersTotal != null ? fmt(membersTotal) : "-"}</td>
                    <td style={{ ...tdS, color: membersDiff != null ? (membersDiff >= 0 ? "var(--accent2)" : "var(--red)") : "inherit" }}>
                      {membersDiff != null ? signFmt(membersDiff) : "-"}
                    </td>
                    <td style={tdS}>{savesTotal != null ? fmt(savesTotal) : "-"}</td>
                    <td style={{ ...tdS, color: savesDiff != null ? (savesDiff >= 0 ? "var(--accent2)" : "var(--red)") : "inherit" }}>
                      {savesDiff != null ? signFmt(savesDiff) : "-"}
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
          </Card>
        )}
      </div>
    </div>
  );
}

function Card({ step, color, title, right, children, noStep }) {
  return (
    <div style={{
      background: "var(--card)", border: "1px solid var(--border)",
      borderRadius: 12, padding: 24, marginBottom: 14,
    }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          {!noStep && (
            <span style={{
              background: color, color: ["var(--accent2)", "#ff9f43"].includes(color) ? "#000" : "#fff",
              fontSize: 11, fontWeight: 700, padding: "2px 8px", borderRadius: 6,
            }}>STEP {step}</span>
          )}
          <span style={{ fontWeight: 600, fontSize: 15 }}>{title}</span>
        </div>
        {right}
      </div>
      {children}
    </div>
  );
}

function FileDropZone({ label, fileName, result, dragOver, setDragOver, fileRef, onFile, accent }) {
  return (
    <div
      onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
      onDragLeave={() => setDragOver(false)}
      onDrop={(e) => { e.preventDefault(); setDragOver(false); onFile(e.dataTransfer.files[0]); }}
      onClick={() => fileRef.current?.click()}
      style={{
        border: `2px dashed ${dragOver ? accent : "var(--border)"}`,
        borderRadius: 10, padding: "24px 16px", textAlign: "center",
        cursor: "pointer", transition: "all 0.2s",
        background: dragOver ? `${accent}08` : "transparent",
      }}
    >
      <input ref={fileRef} type="file" accept=".xlsx,.xls" style={{ display: "none" }}
        onChange={(e) => onFile(e.target.files[0])} />
      {fileName ? (
        <>
          <div style={{ fontSize: 13, fontWeight: 600 }}>{label}</div>
          <div style={{ fontSize: 12, color: "var(--sub)", marginTop: 4 }}>{fileName}</div>
          <div style={{ fontSize: 12, color: "var(--accent2)", marginTop: 2 }}>
            ✓ {result?.numDays}일 · PV {fmt(result?.pvTotal)}
          </div>
        </>
      ) : (
        <>
          <div style={{ fontSize: 13, fontWeight: 600 }}>{label}</div>
          <div style={{ fontSize: 12, color: "var(--sub)", marginTop: 4 }}>클릭 또는 드래그</div>
        </>
      )}
    </div>
  );
}

function PasteBox({ label, value, onChange, placeholder, parsed, diff, prevLabel }) {
  return (
    <div>
      <label style={{ fontSize: 12, fontWeight: 600, display: "block", marginBottom: 6 }}>{label}</label>
      <textarea value={value} onChange={(e) => onChange(e.target.value)} placeholder={placeholder}
        style={{
          width: "100%", minHeight: 150, padding: "10px 12px",
          borderRadius: 8, border: "1px solid var(--border)",
          background: "var(--bg)", color: "var(--text)",
          fontSize: 12, fontFamily: "'JetBrains Mono', monospace", lineHeight: 1.5,
        }}
      />
      {parsed && (
        <div style={{
          marginTop: 8, padding: "8px 12px",
          background: "var(--bg)", borderRadius: 8, fontSize: 12,
          display: "flex", flexWrap: "wrap", gap: 10, alignItems: "center",
        }}>
          <span style={{ fontWeight: 600 }}>합계: {fmt(parsed.total)}</span>
          {parsed.count > 0 && <span style={{ color: "var(--sub)" }}>{parsed.count}일</span>}
          {parsed.cumulative != null && <span style={{ color: "var(--sub)" }}>누적: {fmt(parsed.cumulative)}</span>}
          {diff != null && (
            <span style={{ color: diff >= 0 ? "var(--accent2)" : "var(--red)" }}>
              {prevLabel} 대비 {signFmt(diff)}
            </span>
          )}
        </div>
      )}
    </div>
  );
}
