import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

/* ═══════════════════════════════════════════
   TOPIC DATA
   ═══════════════════════════════════════════ */
const TOPICS = {
  6: {
    grade: "Grade 6", totalMarks: 40,
    topics: [
      { id: "g6-gl", name: "GL Questions", marks: 10 },
      { id: "g6-equiv-frac", name: "Equivalent Fractions & Simplifying", marks: 2 },
      { id: "g6-mixed-improper", name: "Mixed & Improper Fractions", marks: 2 },
      { id: "g6-frac-dec-perc", name: "Fractions, Decimals & Percentages", marks: 3 },
      { id: "g6-frac-amounts", name: "Fractions of Amounts", marks: 3 },
      { id: "g6-probability", name: "Probability", marks: 3 },
      { id: "g6-ratios", name: "Writing & Simplifying Ratios", marks: 3 },
      { id: "g6-sequences", name: "Sequences", marks: 2 },
      { id: "g6-proportion", name: "Proportion", marks: 3 },
      { id: "g6-sharing-ratio", name: "Sharing in a Ratio", marks: 3 },
      { id: "g6-exp-probability", name: "Expected Outcomes & Experimental Probability", marks: 3 },
      { id: "g6-coordinates", name: "Coordinates, Midpoints & Straight Line Graphs", marks: 3 },
    ],
  },
  7: {
    grade: "Grade 7", totalMarks: 42,
    topics: [
      { id: "g7-gl", name: "GL Questions", marks: 10 },
      { id: "g7-conversion-graphs", name: "Conversion Graphs", marks: 3 },
      { id: "g7-distance-time", name: "Distance Time Graph", marks: 3 },
      { id: "g7-ordering-decimals", name: "Ordering Decimals & Rounding", marks: 2 },
      { id: "g7-place-value", name: "Place Value Calculations", marks: 2 },
      { id: "g7-calc-decimals", name: "Calculations with Decimals", marks: 3 },
      { id: "g7-ratio-proportion", name: "Ratio & Proportion", marks: 3 },
      { id: "g7-parallel-angles", name: "Angles in Parallel Lines", marks: 3 },
      { id: "g7-ext-int-angles", name: "Exterior & Interior Angles", marks: 3 },
      { id: "g7-ordering-fractions", name: "Ordering Fractions", marks: 2 },
      { id: "g7-add-sub-fractions", name: "Adding & Subtracting Fractions", marks: 2 },
      { id: "g7-multiply-fractions", name: "Multiplying Fractions", marks: 2 },
      { id: "g7-divide-fractions", name: "Dividing Fractions", marks: 2 },
      { id: "g7-mixed-numbers", name: "Calculating with Mixed Numbers", marks: 2 },
    ],
  },
  8: {
    grade: "Grade 8", totalMarks: 40,
    topics: [
      { id: "g8-gl", name: "GL Questions", marks: 10 },
      { id: "g8-rounding", name: "Calculator Use & Rounding to Decimal Places", marks: 3 },
      { id: "g8-bounds", name: "Upper & Lower Bounds (1 d.p.)", marks: 2 },
      { id: "g8-probability", name: "Probability", marks: 4 },
      { id: "g8-triple-brackets", name: "Expanding Triple Brackets", marks: 3 },
      { id: "g8-trig", name: "Trigonometry (SOHCAHTOA)", marks: 3 },
      { id: "g8-quadratics", name: "Algebra: Factorising & Solving Quadratics", marks: 3 },
      { id: "g8-quad-graphs", name: "Quadratic Graphs (Table of Values & Plotting)", marks: 4 },
      { id: "g8-pythag-trig", name: "Multi-step Pythagoras & Trigonometry", marks: 4 },
      { id: "g8-quad-geometric", name: "Quadratic Equations from Geometric Context", marks: 4 },
    ],
  },
};

const TEACHER_PIN = "1234";
const genId = () => Math.random().toString(36).substr(2, 9);
function shuffle(arr) { const a = [...arr]; for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); [a[i], a[j]] = [a[j], a[i]]; } return a; }

/* ═══════════════════════════════════════════
   MATH RENDERING
   ═══════════════════════════════════════════ */
function latexToHtml(text) {
  if (!text) return "";
  let h = text;
  const syms = [[/\\times/g,"×"],[/\\div/g,"÷"],[/\\pm/g,"±"],[/\\pi/g,"π"],[/\\degree/g,"°"],[/\\circ/g,"°"],[/\\leq/g,"≤"],[/\\geq/g,"≥"],[/\\neq/g,"≠"],[/\\approx/g,"≈"],[/\\infty/g,"∞"],[/\\theta/g,"θ"],[/\\alpha/g,"α"],[/\\beta/g,"β"],[/\\angle/g,"∠"],[/\\triangle/g,"△"],[/\\rightarrow/g,"→"],[/\\left/g,""],[/\\right/g,""],[/\\quad/g,"&ensp;"],[/\\ /g," "],[/\\cdot/g,"·"],[/\\ldots/g,"…"]];
  syms.forEach(([rx, rep]) => { h = h.replace(rx, rep); });
  h = h.replace(/\\text\{([^}]*)\}/g, '<span style="font-style:normal">$1</span>');
  h = h.replace(/\\textbf\{([^}]*)\}/g, '<b>$1</b>');
  for (let i = 0; i < 3; i++) h = h.replace(/\\sqrt\{([^{}]*(?:\{[^{}]*\}[^{}]*)*)\}/, '<span class="msqrt"><span class="msqrt-sym">√</span><span class="msqrt-body">$1</span></span>');
  for (let i = 0; i < 5; i++) h = h.replace(/\\frac\{([^{}]*(?:<[^>]*>[^{}]*)*)\}\{([^{}]*(?:<[^>]*>[^{}]*)*)\}/, '<span class="mfrac"><span class="mfrac-num">$1</span><span class="mfrac-den">$2</span></span>');
  h = h.replace(/\^{([^{}]*)}/g, "<sup>$1</sup>"); h = h.replace(/\^(\w)/g, "<sup>$1</sup>");
  h = h.replace(/_{([^{}]*)}/g, "<sub>$1</sub>"); h = h.replace(/_(\w)/g, "<sub>$1</sub>");
  h = h.replace(/\$/g, ""); h = h.replace(/\\\\/g, "");
  return h;
}
function MathText({ text, className, style }) {
  return <span className={className} style={style} dangerouslySetInnerHTML={{ __html: latexToHtml(text || "") }} />;
}

/* ═══════════════════════════════════════════
   SVG VISUALS
   ═══════════════════════════════════════════ */
function CoordinateGrid({ data }) {
  const r = data.range || 6, size = 280, pad = 32, inner = size - 2 * pad, scale = inner / (2 * r);
  const cx = pad + inner / 2, cy = pad + inner / 2, toX = (x) => cx + x * scale, toY = (y) => cy - y * scale;
  return (
    <svg viewBox={`0 0 ${size} ${size}`} width="100%" style={{ maxWidth: size, display: "block", margin: "12px auto", background: "var(--card)", borderRadius: 10, border: "1px solid var(--border)" }}>
      {Array.from({ length: 2 * r + 1 }, (_, i) => i - r).map((v) => (<g key={v}><line x1={toX(v)} y1={pad} x2={toX(v)} y2={size - pad} stroke="var(--border)" strokeWidth={v === 0 ? 0 : 0.5} /><line x1={pad} y1={toY(v)} x2={size - pad} y2={toY(v)} stroke="var(--border)" strokeWidth={v === 0 ? 0 : 0.5} /></g>))}
      <line x1={pad} y1={cy} x2={size - pad} y2={cy} stroke="var(--text)" strokeWidth={1.5} /><line x1={cx} y1={size - pad} x2={cx} y2={pad} stroke="var(--text)" strokeWidth={1.5} />
      {Array.from({ length: 2 * r + 1 }, (_, i) => i - r).filter((v) => v !== 0).map((v) => (<g key={`l${v}`}><text x={toX(v)} y={cy + 14} textAnchor="middle" fontSize="9" fill="var(--muted)" fontFamily="var(--font-body)">{v}</text><text x={cx - 12} y={toY(v) + 3} textAnchor="end" fontSize="9" fill="var(--muted)" fontFamily="var(--font-body)">{v}</text></g>))}
      {(data.lines || []).map((l, i) => <line key={`ln${i}`} x1={toX(l.x1)} y1={toY(l.y1)} x2={toX(l.x2)} y2={toY(l.y2)} stroke="var(--accent)" strokeWidth={2.5} strokeLinecap="round" opacity="0.8" />)}
      {(data.points || []).map((p, i) => (<g key={`pt${i}`}><circle cx={toX(p.x)} cy={toY(p.y)} r={5} fill="var(--accent)" stroke="var(--card)" strokeWidth="2" />{p.label && <text x={toX(p.x) + 10} y={toY(p.y) - 8} fontSize="12" fontWeight="bold" fill="var(--accent)" fontFamily="var(--font-body)">{p.label}</text>}</g>))}
    </svg>
  );
}
function AngleDiagram({ data }) {
  const desc = (data.description || "").toLowerCase(), angles = data.angles || [];
  if (desc.includes("parallel")) return (
    <svg viewBox="0 0 260 200" width="100%" style={{ maxWidth: 260, display: "block", margin: "12px auto", background: "var(--card)", borderRadius: 10, border: "1px solid var(--border)", padding: 8 }}>
      <line x1="20" y1="60" x2="240" y2="60" stroke="var(--text)" strokeWidth="2" /><line x1="20" y1="140" x2="240" y2="140" stroke="var(--text)" strokeWidth="2" /><line x1="80" y1="15" x2="180" y2="185" stroke="var(--accent)" strokeWidth="2" />
      <polygon points="115,48 120,44 125,48" fill="var(--text)" /><polygon points="135,128 140,124 145,128" fill="var(--text)" />
      {angles.map((a, i) => <text key={i} x={i === 0 ? 118 : 148} y={i === 0 ? 50 : 132} fontSize="14" fontWeight="bold" fill={a.known === false ? "var(--accent)" : "var(--text)"} fontFamily="var(--font-body)">{a.known === false ? a.label : `${a.value}°`}</text>)}
    </svg>
  );
  return (
    <svg viewBox="0 0 240 200" width="100%" style={{ maxWidth: 240, display: "block", margin: "12px auto", background: "var(--card)", borderRadius: 10, border: "1px solid var(--border)", padding: 8 }}>
      <polygon points="120,25 25,175 215,175" fill="none" stroke="var(--text)" strokeWidth="2" strokeLinejoin="round" />
      {angles[0] && <text x="112" y="55" fontSize="13" fontWeight="bold" fill={angles[0].known === false ? "var(--accent)" : "var(--text)"} fontFamily="var(--font-body)">{angles[0].known === false ? angles[0].label : `${angles[0].value}°`}</text>}
      {angles[1] && <text x="38" y="168" fontSize="13" fontWeight="bold" fill={angles[1].known === false ? "var(--accent)" : "var(--text)"} fontFamily="var(--font-body)">{angles[1].known === false ? angles[1].label : `${angles[1].value}°`}</text>}
      {angles[2] && <text x="175" y="168" fontSize="13" fontWeight="bold" fill={angles[2].known === false ? "var(--accent)" : "var(--text)"} fontFamily="var(--font-body)">{angles[2].known === false ? angles[2].label : `${angles[2].value}°`}</text>}
    </svg>
  );
}
function BarChart({ data }) {
  const bars = data.bars || []; if (!bars.length) return null;
  const maxVal = Math.max(...bars.map((b) => b.value), 1), w = 280, h = 170, pad = 38;
  const barW = Math.min(32, (w - 2 * pad) / bars.length - 6);
  const colors = ["var(--accent)", "#E67E22", "#3498DB", "#9B59B6", "#E74C3C", "#1ABC9C"];
  return (
    <svg viewBox={`0 0 ${w} ${h}`} width="100%" style={{ maxWidth: w, display: "block", margin: "12px auto", background: "var(--card)", borderRadius: 10, border: "1px solid var(--border)" }}>
      <line x1={pad} y1={h - pad} x2={w - 10} y2={h - pad} stroke="var(--text)" strokeWidth="1" /><line x1={pad} y1={12} x2={pad} y2={h - pad} stroke="var(--text)" strokeWidth="1" />
      {bars.map((b, i) => { const bh = ((h - pad - 20) * b.value) / maxVal, bx = pad + 8 + i * ((w - 2 * pad) / bars.length); return (<g key={i}><rect x={bx} y={h - pad - bh} width={barW} height={bh} fill={colors[i % colors.length]} rx="3" opacity="0.85" /><text x={bx + barW / 2} y={h - pad + 13} textAnchor="middle" fontSize="8" fill="var(--muted)" fontFamily="var(--font-body)">{b.label}</text><text x={bx + barW / 2} y={h - pad - bh - 5} textAnchor="middle" fontSize="10" fontWeight="600" fill="var(--text)" fontFamily="var(--font-body)">{b.value}</text></g>); })}
      {data.title && <text x={w / 2} y={14} textAnchor="middle" fontSize="11" fontWeight="600" fill="var(--muted)" fontFamily="var(--font-body)">{data.title}</text>}
    </svg>
  );
}
function LineGraph({ data }) {
  const pts = data.points || []; if (!pts.length) return null;
  const w = 300, h = 200, pad = 44;
  const xMin = Math.min(...pts.map((p) => p.x)), xMax = Math.max(...pts.map((p) => p.x)), yMin = Math.min(...pts.map((p) => p.y), 0), yMax = Math.max(...pts.map((p) => p.y));
  const xR = xMax - xMin || 1, yR = yMax - yMin || 1, toX = (x) => pad + ((x - xMin) / xR) * (w - 2 * pad), toY = (y) => h - pad - ((y - yMin) / yR) * (h - 2 * pad);
  const pathD = pts.map((p, i) => `${i === 0 ? "M" : "L"}${toX(p.x).toFixed(1)},${toY(p.y).toFixed(1)}`).join(" ");
  return (
    <svg viewBox={`0 0 ${w} ${h}`} width="100%" style={{ maxWidth: w, display: "block", margin: "12px auto", background: "var(--card)", borderRadius: 10, border: "1px solid var(--border)" }}>
      <line x1={pad} y1={h - pad} x2={w - 12} y2={h - pad} stroke="var(--text)" strokeWidth="1.2" /><line x1={pad} y1={12} x2={pad} y2={h - pad} stroke="var(--text)" strokeWidth="1.2" />
      <path d={pathD} fill="none" stroke="var(--accent)" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" />
      {pts.map((p, i) => <circle key={i} cx={toX(p.x)} cy={toY(p.y)} r="3.5" fill="var(--accent)" stroke="var(--card)" strokeWidth="1.5" />)}
      {data.xLabel && <text x={w / 2} y={h - 4} textAnchor="middle" fontSize="10" fill="var(--muted)" fontFamily="var(--font-body)">{data.xLabel}</text>}
      {data.yLabel && <text x={12} y={h / 2} textAnchor="middle" fontSize="10" fill="var(--muted)" fontFamily="var(--font-body)" transform={`rotate(-90,12,${h / 2})`}>{data.yLabel}</text>}
      {data.title && <text x={w / 2} y={14} textAnchor="middle" fontSize="11" fontWeight="600" fill="var(--muted)" fontFamily="var(--font-body)">{data.title}</text>}
    </svg>
  );
}
function QuestionVisual({ visual }) {
  if (!visual?.type) return null;
  try { switch (visual.type) { case "coordinate_grid": return <CoordinateGrid data={visual} />; case "angle_diagram": return <AngleDiagram data={visual} />; case "bar_chart": return <BarChart data={visual} />; case "line_graph": return <LineGraph data={visual} />; default: return null; } } catch (e) { return null; }
}

/* ═══════════════════════════════════════════
   AI QUESTION GENERATION
   ═══════════════════════════════════════════ */
async function generateQuestions(grade, topicName, count, context = "") {
  const isGL = topicName.includes("GL");
  const allTopics = TOPICS[grade].topics.filter((t) => !t.id.endsWith("-gl")).map((t) => t.name);
  const visualKW = ["Coordinates","Midpoints","Straight Line","Conversion Graphs","Distance Time","Angles","Parallel","Exterior","Interior","Probability","Expected","Experimental","Quadratic Graphs","Trigonometry","Pythagoras"];
  const needsVisuals = visualKW.some((k) => topicName.toLowerCase().includes(k.toLowerCase()));

  const prompt = `You are a KS3 Mathematics exam question generator for Grade ${grade}.
Generate exactly ${count} multiple-choice questions on ${isGL ? `General Level (GL) questions covering a MIX of: ${allTopics.join(", ")}. Each question tests a DIFFERENT topic.` : `"${topicName}"`}.

${context ? `ADAPTIVE: Student struggled with: ${context}. Target these weaknesses.` : ""}

MATH FORMATTING (MANDATORY):
- ALL fractions: \\frac{num}{den} — NEVER "1/2" or "3/4"
- Powers: x^{2}  Square roots: \\sqrt{25}  Multiplication: \\times  Division: \\div  Degrees: \\degree
- Mixed numbers: 2\\frac{3}{4}

${needsVisuals ? `VISUALS: Include "visual" for diagram questions. Types:
coordinate_grid: {"type":"coordinate_grid","points":[{"x":2,"y":3,"label":"A"}],"lines":[{"x1":0,"y1":1,"x2":4,"y2":9}],"range":6}
angle_diagram: {"type":"angle_diagram","description":"Two parallel lines cut by a transversal","angles":[{"value":65,"label":"65°","known":true},{"value":65,"label":"x","known":false}]}
bar_chart: {"type":"bar_chart","bars":[{"label":"Red","value":5}],"title":"Results"}
line_graph: {"type":"line_graph","points":[{"x":0,"y":0},{"x":2,"y":30}],"xLabel":"Time","yLabel":"Distance"}
Include visuals for at least ${Math.max(1, Math.floor(count / 3))} questions.` : 'Set "visual": null for all.'}

Respond ONLY with a JSON array. No markdown, no backticks.
[{"question":"...","options":["A) ...","B) ...","C) ...","D) ..."],"correct":0,"explanation":"...","subtopic":"...","visual":null}]`;

  try {
    const res = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ model: "claude-sonnet-4-20250514", max_tokens: Math.max(3000, count * 450), messages: [{ role: "user", content: prompt }] }),
    });
    const data = await res.json();
    const text = data.content?.map((c) => c.text || "").filter(Boolean).join("") || "";
    return JSON.parse(text.replace(/```json|```/g, "").trim());
  } catch (e) { console.error("API error:", e); return null; }
}

/* ═══════════════════════════════════════════
   PERSISTENT STORAGE
   ═══════════════════════════════════════════ */
const sKey = (name) => `stu:${name.toLowerCase().replace(/\s+/g, "_")}`;
const qKey = (name) => `quiz:${name.toLowerCase().replace(/\s+/g, "_")}`;

async function saveStudent(name, data) { try { await window.storage.set(sKey(name), JSON.stringify(data)); } catch (e) {} }
async function getStudent(name) { try { const r = await window.storage.get(sKey(name)); return r ? JSON.parse(r.value) : null; } catch (e) { return null; } }
async function getAllStudents() {
  try { const keys = await window.storage.list("stu:"); if (!keys?.keys) return []; const out = []; for (const k of keys.keys) { try { const r = await window.storage.get(k); if (r) out.push(JSON.parse(r.value)); } catch (e) {} } return out; } catch (e) { return []; }
}

// Active quiz session (for resume)
async function saveQuizState(name, state) { try { await window.storage.set(qKey(name), JSON.stringify(state)); } catch (e) {} }
async function getQuizState(name) { try { const r = await window.storage.get(qKey(name)); return r ? JSON.parse(r.value) : null; } catch (e) { return null; } }
async function clearQuizState(name) { try { await window.storage.delete(qKey(name)); } catch (e) {} }

/* ═══════════════════════════════════════════
   EXCEL EXPORT
   ═══════════════════════════════════════════ */
function exportToExcel(students) {
  const wb = XLSX.utils.book_new();

  // Sheet 1: Student Overview
  const overviewData = students.map((s) => {
    const ss = s.sessions || [];
    const avg = ss.length ? Math.round(ss.reduce((sum, x) => sum + x.percentage, 0) / ss.length) : 0;
    const weakMap = {};
    ss.forEach((ses) => (ses.weakAreas || []).forEach((w) => { weakMap[w] = (weakMap[w] || 0) + 1; }));
    const topWeak = Object.entries(weakMap).sort((a, b) => b[1] - a[1]).slice(0, 5).map(([w, c]) => `${w} (${c}×)`).join(", ");
    return { Student: s.name, Grade: s.grade, Sessions: ss.length, "Avg Score %": avg, "Last Active": s.lastActive ? new Date(s.lastActive).toLocaleDateString("en-GB") : "—", "Top Weak Areas": topWeak || "None yet" };
  });
  const ws1 = XLSX.utils.json_to_sheet(overviewData);
  ws1["!cols"] = [{ wch: 20 }, { wch: 8 }, { wch: 10 }, { wch: 12 }, { wch: 14 }, { wch: 50 }];
  XLSX.utils.book_append_sheet(wb, ws1, "Student Overview");

  // Sheet 2: All Sessions
  const sessionData = [];
  students.forEach((s) => {
    (s.sessions || []).forEach((ses) => {
      sessionData.push({ Student: s.name, Grade: s.grade, Date: new Date(ses.date).toLocaleDateString("en-GB", { day: "numeric", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" }), Topic: ses.topic, Score: ses.score, Total: ses.total, "Percentage %": ses.percentage, "Weak Areas": (ses.weakAreas || []).join(", ") || "None" });
    });
  });
  const ws2 = XLSX.utils.json_to_sheet(sessionData);
  ws2["!cols"] = [{ wch: 20 }, { wch: 8 }, { wch: 22 }, { wch: 30 }, { wch: 7 }, { wch: 7 }, { wch: 12 }, { wch: 50 }];
  XLSX.utils.book_append_sheet(wb, ws2, "All Sessions");

  // Sheet 3: Topic Progress
  const topicData = [];
  students.forEach((s) => {
    const prog = s.topicProgress || {};
    Object.entries(prog).forEach(([tid, p]) => {
      topicData.push({ Student: s.name, Grade: s.grade, Topic: p.topicName || tid, Attempts: p.attempts, "Best Score %": p.bestPct, Completed: p.completed ? "YES" : "No", "Needs Work": p.bestPct < 60 ? "YES" : (p.bestPct < 80 ? "Review" : "OK") });
    });
  });
  const ws3 = XLSX.utils.json_to_sheet(topicData);
  ws3["!cols"] = [{ wch: 20 }, { wch: 8 }, { wch: 40 }, { wch: 10 }, { wch: 12 }, { wch: 10 }, { wch: 12 }];
  XLSX.utils.book_append_sheet(wb, ws3, "Topic Progress");

  // Sheet 4: AI-Ready Analysis
  const aiData = [["STUDENT PERFORMANCE DATA FOR AI ANALYSIS"], ["Generated: " + new Date().toLocaleDateString("en-GB")], [""], ["INSTRUCTIONS: Feed this sheet to an AI assistant and ask it to identify which topics each student needs more practice on and suggest specific question types."], [""]];
  students.forEach((s) => {
    const ss = s.sessions || [];
    const prog = s.topicProgress || {};
    const avg = ss.length ? Math.round(ss.reduce((sum, x) => sum + x.percentage, 0) / ss.length) : 0;
    const weakMap = {};
    ss.forEach((ses) => (ses.weakAreas || []).forEach((w) => { weakMap[w] = (weakMap[w] || 0) + 1; }));
    aiData.push([`--- ${s.name} (Grade ${s.grade}) --- Average: ${avg}% ---`]);
    const topWeak = Object.entries(weakMap).sort((a, b) => b[1] - a[1]);
    if (topWeak.length) {
      aiData.push(["Recurring weak areas (subtopic: times wrong):"]);
      topWeak.forEach(([w, c]) => aiData.push([`  • ${w}: ${c} times`]));
    }
    const incomplete = Object.entries(prog).filter(([, p]) => !p.completed && p.bestPct < 80);
    if (incomplete.length) {
      aiData.push(["Topics needing improvement:"]);
      incomplete.forEach(([, p]) => aiData.push([`  • ${p.topicName}: best ${p.bestPct}% (${p.attempts} attempts)`]));
    }
    aiData.push([""]);
  });
  const ws4 = XLSX.utils.aoa_to_sheet(aiData);
  ws4["!cols"] = [{ wch: 80 }];
  XLSX.utils.book_append_sheet(wb, ws4, "AI Analysis");

  XLSX.writeFile(wb, `MathsPrep_Report_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

/* ═══════════════════════════════════════════
   UI COMPONENTS
   ═══════════════════════════════════════════ */
function Loader({ message, progress }) {
  return (
    <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 16, padding: 60 }}>
      <div className="loader-ring" />
      <p style={{ color: "var(--muted)", fontSize: 15, fontFamily: "var(--font-body)", textAlign: "center", maxWidth: 280 }}>{message}</p>
      {progress !== undefined && <div style={{ width: 200 }}><div className="progress-bar"><div className="progress-fill" style={{ width: `${progress}%`, background: "var(--accent)" }} /></div><p style={{ textAlign: "center", fontSize: 12, color: "var(--muted)", fontFamily: "var(--font-body)", marginTop: 6 }}>{Math.round(progress)}%</p></div>}
    </div>
  );
}

function Btn({ children, variant = "primary", disabled, onClick, style: s }) {
  const base = { padding: "13px 24px", borderRadius: 10, fontFamily: "var(--font-body)", fontSize: 15, fontWeight: 600, cursor: disabled ? "default" : "pointer", transition: "all 0.2s", border: "none", width: "100%", textAlign: "center", ...s };
  const v = {
    primary: { ...base, background: disabled ? "var(--border)" : "var(--accent)", color: disabled ? "var(--muted)" : "#fff" },
    secondary: { ...base, background: "var(--card)", color: "var(--text)", border: "1px solid var(--border)" },
    accent: { ...base, background: "var(--accent-light)", color: "var(--accent)", border: "2px solid var(--accent)" },
    ghost: { ...base, background: "none", color: "var(--muted)", padding: "8px 14px", fontSize: 13, width: "auto" },
    danger: { ...base, background: "#FEF2F2", color: "var(--error)", border: "1px solid var(--error)" },
  };
  return <button onClick={onClick} disabled={disabled} style={v[variant]}>{children}</button>;
}

/* ═══════════════════════════════════════════
   LOGIN
   ═══════════════════════════════════════════ */
function LoginScreen({ onLogin, onTeacher }) {
  const [name, setName] = useState("");
  const [grade, setGrade] = useState(null);
  const [returning, setReturning] = useState(null);
  const [checking, setChecking] = useState(false);

  const handleLogin = async () => {
    if (!name.trim() || !grade) return;
    setChecking(true);
    const existing = await getStudent(name.trim());
    if (existing) {
      // Returning student — check for active quiz
      const activeQuiz = await getQuizState(name.trim());
      setReturning({ student: existing, activeQuiz });
    } else {
      // New student
      const stu = { name: name.trim(), grade, sessions: [], topicProgress: {}, created: new Date().toISOString(), lastActive: new Date().toISOString() };
      await saveStudent(name.trim(), stu);
      onLogin(name.trim(), grade, null);
    }
    setChecking(false);
  };

  if (returning) {
    const { student, activeQuiz } = returning;
    return (
      <div className="screen-center">
        <div style={{ width: "100%", maxWidth: 440, padding: "0 20px" }}>
          <div style={{ textAlign: "center", marginBottom: 32 }}>
            <div style={{ fontSize: 44, marginBottom: 8 }}>👋</div>
            <h2 className="page-title">Welcome back, {student.name}!</h2>
            <p className="page-subtitle" style={{ marginBottom: 0 }}>Grade {student.grade} · {student.sessions.length} session{student.sessions.length !== 1 ? "s" : ""} completed</p>
          </div>

          {activeQuiz && (
            <div className="card" style={{ padding: 20, marginBottom: 12, borderLeft: "3px solid var(--accent)" }}>
              <p style={{ fontFamily: "var(--font-body)", fontSize: 14, fontWeight: 600, color: "var(--accent)", marginBottom: 4 }}>📝 You have an unfinished quiz</p>
              <p style={{ fontFamily: "var(--font-body)", fontSize: 13, color: "var(--muted)" }}>
                {activeQuiz.topicName} — Question {activeQuiz.currentIndex + 1} of {activeQuiz.questions.length}
              </p>
              <div style={{ display: "flex", gap: 8, marginTop: 12 }}>
                <Btn variant="accent" onClick={() => onLogin(student.name, student.grade, activeQuiz)} style={{ flex: 1 }}>Continue Quiz</Btn>
                <Btn variant="secondary" onClick={async () => { await clearQuizState(student.name); onLogin(student.name, student.grade, null); }} style={{ flex: 1 }}>Start Fresh</Btn>
              </div>
            </div>
          )}

          {!activeQuiz && <Btn onClick={() => onLogin(student.name, student.grade, null)}>Continue to Topics →</Btn>}

          <button onClick={() => setReturning(null)} className="teacher-link" style={{ marginTop: 16 }}>Not {student.name}? Go back</button>
        </div>
      </div>
    );
  }

  return (
    <div className="screen-center">
      <div style={{ width: "100%", maxWidth: 440, padding: "0 20px" }}>
        <div style={{ textAlign: "center", marginBottom: 44 }}>
          <div className="logo-icon">∑</div>
          <h1 className="logo-text">MathsPrep</h1>
          <p className="subtitle">KS3 Exam Practice · Powered by AI</p>
        </div>
        <div className="card" style={{ padding: 32 }}>
          <label className="field-label">Your Name</label>
          <input value={name} onChange={(e) => setName(e.target.value)} placeholder="Enter your full name" className="input-field" onKeyDown={(e) => e.key === "Enter" && handleLogin()} />
          <label className="field-label" style={{ marginTop: 20 }}>Your Grade</label>
          <div style={{ display: "flex", gap: 8, marginBottom: 28 }}>
            {[6, 7, 8].map((g) => <button key={g} onClick={() => setGrade(g)} className={`grade-btn ${grade === g ? "grade-btn-active" : ""}`}>Grade {g}</button>)}
          </div>
          <Btn onClick={handleLogin} disabled={!name.trim() || !grade || checking}>{checking ? "Checking..." : "Log In →"}</Btn>
        </div>
        <button onClick={onTeacher} className="teacher-link">Teacher Dashboard</button>
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════
   TOPIC SELECT (with progress)
   ═══════════════════════════════════════════ */
function TopicSelect({ grade, onSelect, onBack, studentName }) {
  const gd = TOPICS[grade];
  const [progress, setProgress] = useState({});

  useEffect(() => {
    (async () => {
      const stu = await getStudent(studentName);
      if (stu?.topicProgress) setProgress(stu.topicProgress);
    })();
  }, [studentName]);

  return (
    <div className="page"><div className="container">
      <div className="topbar">
        <Btn variant="ghost" onClick={onBack}>← Log Out</Btn>
        <span className="topbar-info">Hi, {studentName} 👋</span>
      </div>
      <h2 className="page-title">{gd.grade} Topics</h2>
      <p className="page-subtitle">Your progress is saved. Pick up where you left off.</p>
      <div className="topic-list">
        <button onClick={() => onSelect("all")} className="topic-btn topic-btn-exam">
          <div>
            <div className="topic-btn-title" style={{ color: "var(--accent)" }}>🎯 Full Exam Practice</div>
            <div style={{ fontFamily: "var(--font-body)", fontSize: 12, color: "var(--accent)", opacity: 0.8, marginTop: 2 }}>All topics · {gd.totalMarks} questions · Shuffled</div>
          </div>
          <span style={{ fontSize: 20, color: "var(--accent)" }}>→</span>
        </button>
        {gd.topics.map((t) => {
          const p = progress[t.id];
          const pct = p ? p.bestPct : 0;
          const done = p?.completed;
          return (
            <button key={t.id} onClick={() => onSelect(t)} className="topic-btn" style={{ flexDirection: "column", alignItems: "stretch", gap: 8 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  {done && <span style={{ fontSize: 16 }}>✅</span>}
                  <span className="topic-btn-title" style={{ color: done ? "var(--success)" : "var(--text)" }}>{t.name}</span>
                </div>
                <span className="topic-badge">{t.marks} marks</span>
              </div>
              {p && (
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <div style={{ flex: 1, height: 4, background: "var(--border)", borderRadius: 2, overflow: "hidden" }}>
                    <div style={{ height: "100%", width: `${pct}%`, background: done ? "var(--success)" : pct >= 60 ? "var(--accent)" : "var(--error)", borderRadius: 2, transition: "width 0.4s" }} />
                  </div>
                  <span style={{ fontFamily: "var(--font-body)", fontSize: 11, fontWeight: 700, color: done ? "var(--success)" : "var(--muted)", minWidth: 40, textAlign: "right" }}>
                    {done ? "100%" : `${pct}%`}
                  </span>
                  <span style={{ fontFamily: "var(--font-body)", fontSize: 10, color: "var(--muted)" }}>
                    ({p.attempts}×)
                  </span>
                </div>
              )}
            </button>
          );
        })}
      </div>
    </div></div>
  );
}

/* ═══════════════════════════════════════════
   QUIZ (with auto-save for resume)
   ═══════════════════════════════════════════ */
function QuizScreen({ grade, topic, studentName, onFinish, onBack, resumeState }) {
  const [questions, setQuestions] = useState(resumeState?.questions || []);
  const [cur, setCur] = useState(resumeState?.currentIndex || 0);
  const [selected, setSelected] = useState(null);
  const [confirmed, setConfirmed] = useState(false);
  const [answers, setAnswers] = useState(resumeState?.answers || []);
  const [loading, setLoading] = useState(!resumeState);
  const [loadMsg, setLoadMsg] = useState("Preparing your questions...");
  const [loadProg, setLoadProg] = useState(0);
  const [finished, setFinished] = useState(false);
  const [adaptiveLoading, setAdaptiveLoading] = useState(false);
  const [adaptiveDone, setAdaptiveDone] = useState(false);

  const isFullExam = topic === "all";
  const topicName = isFullExam ? "Full Exam" : topic.name;
  const topicId = isFullExam ? "full-exam" : topic.id;

  // Generate questions on mount (unless resuming)
  useEffect(() => {
    if (resumeState) return;
    (async () => {
      setLoading(true);
      if (isFullExam) {
        const topics = TOPICS[grade].topics;
        let all = [];
        for (let i = 0; i < topics.length; i++) {
          setLoadMsg(`Generating: ${topics[i].name}...`);
          setLoadProg(Math.round((i / topics.length) * 100));
          const qs = await generateQuestions(grade, topics[i].name, topics[i].marks);
          if (qs) all.push(...qs.map((q) => ({ ...q, topicSource: topics[i].name })));
        }
        setQuestions(shuffle(all));
        setLoadProg(100);
      } else {
        const qs = await generateQuestions(grade, topicName, Math.max(topic.marks, 3));
        if (qs) setQuestions(qs);
      }
      setLoading(false);
    })();
  }, []);

  // Auto-save quiz state after each answer
  useEffect(() => {
    if (questions.length > 0 && !finished) {
      saveQuizState(studentName, { questions, currentIndex: cur, answers, topicName, topicId, grade, isFullExam, topic });
    }
  }, [answers, cur, questions]);

  const handleConfirm = () => {
    if (selected === null) return;
    setConfirmed(true);
    setAnswers((p) => [...p, {
      question: questions[cur].question, subtopic: questions[cur].subtopic,
      topicSource: questions[cur].topicSource || topicName,
      selected, correct: questions[cur].correct,
      isCorrect: selected === questions[cur].correct,
    }]);
  };

  const handleNext = () => {
    if (cur < questions.length - 1) { setCur((p) => p + 1); setSelected(null); setConfirmed(false); }
    else setFinished(true);
  };

  const handleAdaptive = async () => {
    const wrong = answers.filter((a) => !a.isCorrect);
    if (!wrong.length) return;
    setAdaptiveLoading(true);
    const weak = [...new Set(wrong.map((a) => a.subtopic))].join(", ");
    const qs = await generateQuestions(grade, topicName, Math.min(wrong.length * 2, 8), `Weak: ${weak}`);
    if (qs) { setQuestions(qs); setCur(0); setSelected(null); setConfirmed(false); setAnswers([]); setFinished(false); setAdaptiveDone(true); }
    setAdaptiveLoading(false);
  };

  const handleSave = async () => {
    const score = answers.filter((a) => a.isCorrect).length;
    const total = answers.length;
    const pct = Math.round((score / total) * 100);
    const session = {
      id: genId(), date: new Date().toISOString(), grade, topic: topicName,
      score, total, percentage: pct,
      weakAreas: [...new Set(answers.filter((a) => !a.isCorrect).map((w) => w.subtopic))],
    };

    let stu = await getStudent(studentName);
    if (!stu) stu = { name: studentName, grade, sessions: [], topicProgress: {}, created: new Date().toISOString() };
    stu.sessions.push(session);
    stu.lastActive = new Date().toISOString();

    // Update topic progress
    if (!stu.topicProgress) stu.topicProgress = {};
    if (!isFullExam) {
      const prev = stu.topicProgress[topicId] || { attempts: 0, bestPct: 0, completed: false, topicName };
      prev.attempts++;
      prev.bestPct = Math.max(prev.bestPct, pct);
      prev.completed = prev.bestPct === 100;
      prev.topicName = topicName;
      stu.topicProgress[topicId] = prev;
    } else {
      // For full exam, update each topic's progress
      const topicBD = {};
      answers.forEach((a) => {
        const t = a.topicSource;
        if (!topicBD[t]) topicBD[t] = { c: 0, t: 0 };
        topicBD[t].t++; if (a.isCorrect) topicBD[t].c++;
      });
      const gradeTopics = TOPICS[grade].topics;
      Object.entries(topicBD).forEach(([tName, d]) => {
        const tObj = gradeTopics.find((x) => x.name === tName);
        if (!tObj) return;
        const tid = tObj.id;
        const tPct = Math.round((d.c / d.t) * 100);
        const prev = stu.topicProgress[tid] || { attempts: 0, bestPct: 0, completed: false, topicName: tName };
        prev.attempts++;
        prev.bestPct = Math.max(prev.bestPct, tPct);
        prev.completed = prev.bestPct === 100;
        prev.topicName = tName;
        stu.topicProgress[tid] = prev;
      });
    }

    await saveStudent(studentName, stu);
    await clearQuizState(studentName);
    onFinish();
  };

  if (loading) return <div className="page"><div className="container"><Loader message={loadMsg} progress={isFullExam ? loadProg : undefined} /></div></div>;
  if (!questions.length) return (
    <div className="page"><div className="container" style={{ textAlign: "center", paddingTop: 60 }}>
      <p style={{ fontFamily: "var(--font-body)", color: "var(--muted)", fontSize: 16 }}>Failed to load questions.</p>
      <Btn variant="secondary" onClick={onBack} style={{ marginTop: 16, width: "auto", display: "inline-block" }}>← Go Back</Btn>
    </div></div>
  );

  // ── RESULTS ──
  if (finished) {
    const score = answers.filter((a) => a.isCorrect).length;
    const total = answers.length;
    const pct = Math.round((score / total) * 100);
    const wrong = answers.filter((a) => !a.isCorrect);
    const uniqueWeak = [...new Set(wrong.map((w) => w.subtopic))];
    const topicBD = {};
    answers.forEach((a) => { const t = a.topicSource; if (!topicBD[t]) topicBD[t] = { c: 0, t: 0 }; topicBD[t].t++; if (a.isCorrect) topicBD[t].c++; });

    return (
      <div className="page"><div className="container">
        <div style={{ textAlign: "center", marginBottom: 24 }}>
          <div style={{ fontSize: 52, marginBottom: 4 }}>{pct === 100 ? "🏆" : pct >= 80 ? "🌟" : pct >= 60 ? "👍" : "💪"}</div>
          <h2 className="page-title">{pct === 100 ? "Perfect score!" : pct >= 80 ? "Brilliant work!" : pct >= 60 ? "Good effort!" : "Keep going!"}</h2>
          <p className="page-subtitle" style={{ marginBottom: 0 }}>{topicName}{adaptiveDone ? " · Targeted" : ""}</p>
        </div>
        <div className="card" style={{ textAlign: "center", padding: 28, marginBottom: 14 }}>
          <div style={{ fontFamily: "var(--font-display)", fontSize: 54, color: pct >= 60 ? "var(--success)" : "var(--error)", fontWeight: 800 }}>{score}<span style={{ fontSize: 28, color: "var(--muted)" }}>/{total}</span></div>
          <div className="progress-bar" style={{ marginTop: 14 }}><div className="progress-fill" style={{ width: `${pct}%`, background: pct >= 60 ? "var(--success)" : "var(--error)" }} /></div>
          {pct === 100 && <p style={{ fontFamily: "var(--font-body)", fontSize: 14, color: "var(--success)", marginTop: 8, fontWeight: 600 }}>✅ Topic completed!</p>}
        </div>
        {isFullExam && Object.keys(topicBD).length > 1 && (
          <div className="card" style={{ padding: 20, marginBottom: 14 }}>
            <h3 className="section-label">Topic Breakdown</h3>
            {Object.entries(topicBD).sort((a, b) => (a[1].c / a[1].t) - (b[1].c / b[1].t)).map(([t, d], i) => {
              const tp = Math.round((d.c / d.t) * 100);
              return (<div key={i} style={{ display: "flex", alignItems: "center", gap: 10, padding: "7px 0", borderBottom: "1px solid var(--border)" }}>
                <span style={{ flex: 1, fontFamily: "var(--font-body)", fontSize: 13, color: "var(--text)" }}>{tp === 100 && "✅ "}{t}</span>
                <div style={{ width: 60, height: 4, background: "var(--border)", borderRadius: 2, overflow: "hidden" }}><div style={{ height: "100%", width: `${tp}%`, background: tp >= 60 ? "var(--success)" : "var(--error)", borderRadius: 2 }} /></div>
                <span style={{ fontFamily: "var(--font-body)", fontSize: 13, fontWeight: 700, color: tp >= 60 ? "var(--success)" : "var(--error)", minWidth: 36, textAlign: "right" }}>{d.c}/{d.t}</span>
              </div>);
            })}
          </div>
        )}
        {uniqueWeak.length > 0 && (
          <div className="card" style={{ padding: 20, marginBottom: 14 }}>
            <h3 className="section-label" style={{ color: "var(--error)" }}>Areas to improve</h3>
            {uniqueWeak.map((w, i) => <div key={i} style={{ fontFamily: "var(--font-body)", fontSize: 14, color: "var(--text)", padding: "6px 0", borderBottom: i < uniqueWeak.length - 1 ? "1px solid var(--border)" : "none" }}><span style={{ color: "var(--error)", marginRight: 8 }}>✗</span>{w}</div>)}
          </div>
        )}
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {wrong.length > 0 && !adaptiveDone && <Btn variant="accent" onClick={handleAdaptive} disabled={adaptiveLoading}>{adaptiveLoading ? "Generating..." : "🧠 Practise My Weak Areas"}</Btn>}
          <Btn onClick={handleSave}>Save & Exit</Btn>
        </div>
      </div></div>
    );
  }

  // ── QUESTION ──
  const q = questions[cur];
  return (
    <div className="page"><div className="container">
      <div className="topbar">
        <Btn variant="ghost" onClick={async () => { await saveQuizState(studentName, { questions, currentIndex: cur, answers, topicName, topicId, grade, isFullExam, topic }); onBack(); }}>← Save & Quit</Btn>
        <span className="topbar-info">{adaptiveDone && "🧠 "}{topicName}</span>
      </div>
      <div className="progress-dots">
        {questions.length <= 50 ? questions.map((_, i) => <div key={i} className="progress-dot" style={{ background: i < cur ? (answers[i]?.isCorrect ? "var(--success)" : "var(--error)") : i === cur ? "var(--accent)" : "var(--border)" }} />) : <div className="progress-bar" style={{ flex: 1 }}><div className="progress-fill" style={{ width: `${((cur + 1) / questions.length) * 100}%`, background: "var(--accent)" }} /></div>}
      </div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
        <span className="q-counter">Question {cur + 1} of {questions.length}</span>
        {(q.subtopic || q.topicSource) && <span className="q-subtopic">{q.topicSource || q.subtopic}</span>}
      </div>
      <div className="card question-card"><MathText text={q.question} className="question-text" /></div>
      <QuestionVisual visual={q.visual} />
      <div className="options-list">
        {q.options?.map((opt, i) => {
          let cls = "option-btn";
          if (confirmed) { if (i === q.correct) cls += " option-correct"; else if (i === selected && i !== q.correct) cls += " option-wrong"; } else if (i === selected) cls += " option-selected";
          return <button key={i} onClick={() => !confirmed && setSelected(i)} disabled={confirmed} className={cls}><MathText text={opt} /></button>;
        })}
      </div>
      {confirmed && q.explanation && <div className="explanation-box"><strong>Explanation: </strong><MathText text={q.explanation} /></div>}
      {!confirmed ? <Btn onClick={handleConfirm} disabled={selected === null}>Check Answer</Btn> : <Btn onClick={handleNext}>{cur < questions.length - 1 ? "Next Question →" : "See Results"}</Btn>}
    </div></div>
  );
}

/* ═══════════════════════════════════════════
   TEACHER
   ═══════════════════════════════════════════ */
function TeacherLogin({ onAuth, onBack }) {
  const [pin, setPin] = useState(""); const [err, setErr] = useState(false);
  const go = () => pin === TEACHER_PIN ? onAuth() : setErr(true);
  return (
    <div className="screen-center"><div style={{ width: "100%", maxWidth: 380, padding: "0 20px" }}>
      <Btn variant="ghost" onClick={onBack} style={{ marginBottom: 24 }}>← Back</Btn>
      <h2 className="page-title">Teacher Access</h2>
      <p className="page-subtitle">Enter PIN to view student results</p>
      <input type="password" value={pin} onChange={(e) => { setPin(e.target.value); setErr(false); }} placeholder="• • • •" className="input-field pin-input" onKeyDown={(e) => e.key === "Enter" && go()} />
      {err && <p style={{ fontFamily: "var(--font-body)", color: "var(--error)", fontSize: 13, textAlign: "center", margin: "8px 0" }}>Incorrect PIN</p>}
      <p style={{ fontFamily: "var(--font-body)", color: "var(--muted)", fontSize: 12, textAlign: "center", margin: "8px 0 16px", opacity: 0.6 }}>Default PIN: 1234</p>
      <Btn onClick={go}>Access Dashboard</Btn>
    </div></div>
  );
}

function TeacherDashboard({ onBack }) {
  const [students, setStudents] = useState([]); const [loading, setLoading] = useState(true); const [sel, setSel] = useState(null);
  useEffect(() => { (async () => { setStudents((await getAllStudents()).sort((a, b) => (b.lastActive || "").localeCompare(a.lastActive || ""))); setLoading(false); })(); }, []);

  if (loading) return <div className="page"><div className="container"><Loader message="Loading student data..." /></div></div>;

  if (sel) {
    const ss = sel.sessions || []; const prog = sel.topicProgress || {};
    const weakMap = {}; ss.forEach((s) => (s.weakAreas || []).forEach((w) => { weakMap[w] = (weakMap[w] || 0) + 1; }));
    const topWeak = Object.entries(weakMap).sort((a, b) => b[1] - a[1]).slice(0, 8);

    return (
      <div className="page"><div className="container">
        <Btn variant="ghost" onClick={() => setSel(null)} style={{ marginBottom: 16 }}>← All Students</Btn>
        <h2 className="page-title">{sel.name}</h2>
        <p className="page-subtitle">Grade {sel.grade} · {ss.length} session{ss.length !== 1 ? "s" : ""}</p>

        {/* Topic progress */}
        {Object.keys(prog).length > 0 && (
          <div className="card" style={{ padding: 20, marginBottom: 14 }}>
            <h3 className="section-label">Topic Mastery</h3>
            {Object.entries(prog).sort((a, b) => a[1].bestPct - b[1].bestPct).map(([tid, p], i) => (
              <div key={i} style={{ display: "flex", alignItems: "center", gap: 8, padding: "6px 0", borderBottom: "1px solid var(--border)" }}>
                {p.completed && <span style={{ fontSize: 14 }}>✅</span>}
                <span style={{ flex: 1, fontFamily: "var(--font-body)", fontSize: 13, color: "var(--text)" }}>{p.topicName}</span>
                <div style={{ width: 60, height: 4, background: "var(--border)", borderRadius: 2, overflow: "hidden" }}><div style={{ height: "100%", width: `${p.bestPct}%`, background: p.completed ? "var(--success)" : p.bestPct >= 60 ? "var(--accent)" : "var(--error)", borderRadius: 2 }} /></div>
                <span style={{ fontFamily: "var(--font-body)", fontSize: 12, fontWeight: 700, color: p.completed ? "var(--success)" : "var(--muted)", minWidth: 32, textAlign: "right" }}>{p.bestPct}%</span>
              </div>
            ))}
          </div>
        )}

        {topWeak.length > 0 && (
          <div className="card" style={{ padding: 20, marginBottom: 14 }}>
            <h3 className="section-label" style={{ color: "var(--error)" }}>Recurring Weak Areas</h3>
            {topWeak.map(([area, count], i) => <div key={i} style={{ display: "flex", justifyContent: "space-between", padding: "7px 0", borderBottom: i < topWeak.length - 1 ? "1px solid var(--border)" : "none" }}><span style={{ fontFamily: "var(--font-body)", fontSize: 14, color: "var(--text)" }}>{area}</span><span style={{ fontFamily: "var(--font-body)", fontSize: 13, color: "var(--error)", fontWeight: 600 }}>{count}×</span></div>)}
          </div>
        )}

        <h3 className="section-label" style={{ marginBottom: 8 }}>Session History</h3>
        {ss.slice().reverse().map((s, i) => (
          <div key={i} className="card session-card">
            <div>
              <div style={{ fontFamily: "var(--font-body)", fontSize: 15, fontWeight: 600, color: "var(--text)" }}>{s.percentage === 100 && "✅ "}{s.topic}</div>
              <div style={{ fontFamily: "var(--font-body)", fontSize: 12, color: "var(--muted)", marginTop: 2 }}>{new Date(s.date).toLocaleDateString("en-GB", { day: "numeric", month: "short", hour: "2-digit", minute: "2-digit" })}</div>
            </div>
            <div style={{ textAlign: "right" }}><div style={{ fontFamily: "var(--font-display)", fontSize: 22, fontWeight: 800, color: s.percentage >= 60 ? "var(--success)" : "var(--error)" }}>{s.score}/{s.total}</div></div>
          </div>
        ))}
      </div></div>
    );
  }

  return (
    <div className="page"><div className="container">
      <Btn variant="ghost" onClick={onBack} style={{ marginBottom: 16 }}>← Back to Login</Btn>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 4 }}>
        <h2 className="page-title">Teacher Dashboard</h2>
        {students.length > 0 && <Btn variant="secondary" onClick={() => exportToExcel(students)} style={{ width: "auto", fontSize: 13, padding: "8px 16px" }}>📥 Download Excel</Btn>}
      </div>
      <p className="page-subtitle">{students.length} student{students.length !== 1 ? "s" : ""} registered</p>

      {!students.length ? (
        <div className="card" style={{ padding: 40, textAlign: "center" }}><div style={{ fontSize: 40, marginBottom: 8 }}>📭</div><p className="page-subtitle" style={{ marginBottom: 0 }}>No students have completed a quiz yet.</p></div>
      ) : (
        <div className="topic-list">
          {students.map((s, i) => {
            const ss = s.sessions || []; const avg = ss.length ? Math.round(ss.reduce((sum, x) => sum + x.percentage, 0) / ss.length) : 0;
            const prog = s.topicProgress || {};
            const completed = Object.values(prog).filter((p) => p.completed).length;
            const totalTopics = TOPICS[s.grade]?.topics.length || 0;
            return (
              <button key={i} onClick={() => setSel(s)} className="topic-btn" style={{ flexDirection: "column", alignItems: "stretch", gap: 6 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <div>
                    <div className="topic-btn-title">{s.name}</div>
                    <div style={{ fontFamily: "var(--font-body)", fontSize: 12, color: "var(--muted)", marginTop: 2 }}>
                      Grade {s.grade} · {ss.length} session{ss.length !== 1 ? "s" : ""} · {completed}/{totalTopics} topics ✅
                    </div>
                  </div>
                  <span style={{ fontFamily: "var(--font-display)", fontSize: 20, fontWeight: 800, color: avg >= 60 ? "var(--success)" : "var(--error)" }}>{avg}%</span>
                </div>
                {totalTopics > 0 && <div style={{ height: 4, background: "var(--border)", borderRadius: 2, overflow: "hidden" }}><div style={{ height: "100%", width: `${(completed / totalTopics) * 100}%`, background: "var(--success)", borderRadius: 2 }} /></div>}
              </button>
            );
          })}
        </div>
      )}
    </div></div>
  );
}

/* ═══════════════════════════════════════════
   APP ROOT
   ═══════════════════════════════════════════ */
export default function App() {
  const [view, setView] = useState("login");
  const [studentName, setStudentName] = useState("");
  const [grade, setGrade] = useState(null);
  const [topic, setTopic] = useState(null);
  const [resumeState, setResumeState] = useState(null);

  const handleLogin = (name, g, activeQuiz) => {
    setStudentName(name); setGrade(g);
    if (activeQuiz) {
      setTopic(activeQuiz.topic);
      setResumeState(activeQuiz);
      setView("quiz");
    } else {
      setResumeState(null);
      setView("topics");
    }
  };

  return (
    <div>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:ital,opsz,wght@0,9..40,400;0,9..40,500;0,9..40,600;0,9..40,700&family=Playfair+Display:wght@400;700;800&display=swap');
        :root { --font-display: 'Playfair Display', serif; --font-body: 'DM Sans', sans-serif; --bg: #F4F1EC; --card: #FFFFFF; --text: #1C1917; --muted: #78716C; --border: #E7E0D8; --accent: #1D4E3E; --accent-light: #E8F0ED; --success: #1D6B45; --success-bg: #E8F5EE; --error: #B91C1C; --error-bg: #FEF2F2; }
        @media (prefers-color-scheme: dark) { :root { --bg: #0F0E0C; --card: #1A1816; --text: #E7E0D8; --muted: #8A8480; --border: #2A2622; --accent: #6EBF8B; --accent-light: #162420; --success: #6EBF8B; --success-bg: #162420; --error: #EF6B6B; --error-bg: #2A1616; } }
        * { box-sizing: border-box; margin: 0; padding: 0; } body { background: var(--bg); }
        .mfrac { display: inline-flex; flex-direction: column; align-items: center; vertical-align: middle; margin: 0 3px; line-height: 1; }
        .mfrac-num { border-bottom: 1.8px solid currentColor; padding: 1px 5px 3px; font-size: 0.88em; line-height: 1.3; text-align: center; }
        .mfrac-den { padding: 3px 5px 1px; font-size: 0.88em; line-height: 1.3; text-align: center; }
        .msqrt { display: inline-flex; align-items: stretch; vertical-align: middle; } .msqrt-sym { font-size: 1.15em; margin-right: 1px; line-height: 1; } .msqrt-body { border-top: 1.5px solid currentColor; padding: 1px 4px 0; line-height: 1.3; }
        .screen-center { min-height: 100vh; display: flex; align-items: center; justify-content: center; background: var(--bg); }
        .page { min-height: 100vh; background: var(--bg); padding: 20px 16px 40px; } .container { max-width: 580px; margin: 0 auto; }
        .card { background: var(--card); border-radius: 14px; border: 1px solid var(--border); padding: 24px; }
        .logo-icon { width: 56px; height: 56px; margin: 0 auto 10px; display: flex; align-items: center; justify-content: center; background: var(--accent); color: #fff; font-family: var(--font-display); font-size: 28px; font-weight: 800; border-radius: 14px; }
        .logo-text { font-family: var(--font-display); font-size: 34px; font-weight: 800; color: var(--text); letter-spacing: -0.5px; }
        .subtitle { font-family: var(--font-body); color: var(--muted); font-size: 15px; margin-top: 4px; }
        .page-title { font-family: var(--font-display); font-size: 26px; font-weight: 800; color: var(--text); margin-bottom: 4px; }
        .page-subtitle { font-family: var(--font-body); color: var(--muted); font-size: 14px; margin-bottom: 20px; }
        .section-label { font-family: var(--font-body); font-size: 12px; font-weight: 700; color: var(--muted); text-transform: uppercase; letter-spacing: 0.6px; margin-bottom: 10px; }
        .field-label { display: block; font-family: var(--font-body); font-size: 12px; font-weight: 600; color: var(--muted); text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 8px; }
        .input-field { width: 100%; padding: 12px 16px; border-radius: 10px; border: 1px solid var(--border); font-family: var(--font-body); font-size: 16px; background: var(--bg); color: var(--text); outline: none; }
        .input-field:focus { border-color: var(--accent); } .pin-input { text-align: center; font-size: 22px; letter-spacing: 10px; }
        .grade-btn { flex: 1; padding: 14px 0; border-radius: 10px; border: 1px solid var(--border); background: var(--card); color: var(--text); font-family: var(--font-body); font-size: 15px; font-weight: 600; cursor: pointer; transition: all 0.15s; }
        .grade-btn-active { border: 2px solid var(--accent); background: var(--accent-light); color: var(--accent); }
        .topic-list { display: flex; flex-direction: column; gap: 6px; }
        .topic-btn { display: flex; align-items: center; justify-content: space-between; padding: 14px 18px; border-radius: 12px; border: 1px solid var(--border); background: var(--card); cursor: pointer; text-align: left; transition: all 0.15s; font-family: var(--font-body); }
        .topic-btn:hover { border-color: var(--accent); transform: translateX(3px); }
        .topic-btn-exam { border: 2px solid var(--accent); background: var(--accent-light); }
        .topic-btn-title { font-family: var(--font-body); font-size: 15px; font-weight: 600; color: var(--text); }
        .topic-badge { font-family: var(--font-body); font-size: 11px; font-weight: 700; color: var(--muted); background: var(--bg); padding: 4px 10px; border-radius: 20px; white-space: nowrap; }
        .topbar { display: flex; align-items: center; justify-content: space-between; margin-bottom: 16px; }
        .topbar-info { font-family: var(--font-body); font-size: 13px; color: var(--muted); }
        .progress-dots { display: flex; gap: 2px; margin-bottom: 18px; flex-wrap: wrap; }
        .progress-dot { flex: 1; min-width: 4px; height: 4px; border-radius: 2px; transition: background 0.3s; }
        .progress-bar { height: 6px; background: var(--border); border-radius: 3px; overflow: hidden; }
        .progress-fill { height: 100%; border-radius: 3px; transition: width 0.6s ease; }
        .q-counter { font-family: var(--font-body); font-size: 13px; color: var(--muted); }
        .q-subtopic { font-family: var(--font-body); font-size: 12px; color: var(--accent); font-weight: 600; max-width: 200px; text-align: right; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
        .question-card { margin-bottom: 14px; } .question-text { font-family: var(--font-body); font-size: 17px; color: var(--text); line-height: 1.8; font-weight: 500; display: block; }
        .options-list { display: flex; flex-direction: column; gap: 7px; margin-bottom: 16px; }
        .option-btn { padding: 14px 18px; border-radius: 10px; border: 1px solid var(--border); background: var(--card); color: var(--text); font-family: var(--font-body); font-size: 15px; cursor: pointer; text-align: left; transition: all 0.15s; line-height: 1.7; }
        .option-btn:hover:not(:disabled) { border-color: var(--accent); }
        .option-selected { border: 2px solid var(--accent) !important; background: var(--accent-light) !important; color: var(--accent) !important; font-weight: 600; }
        .option-correct { border: 2px solid var(--success) !important; background: var(--success-bg) !important; color: var(--success) !important; font-weight: 600; }
        .option-wrong { border: 2px solid var(--error) !important; background: var(--error-bg) !important; color: var(--error) !important; font-weight: 600; }
        .explanation-box { background: var(--card); border-radius: 10px; padding: 16px; border: 1px solid var(--border); border-left: 3px solid var(--accent); margin-bottom: 16px; font-family: var(--font-body); font-size: 14px; color: var(--text); line-height: 1.7; }
        .session-card { display: flex; justify-content: space-between; align-items: center; padding: 16px 20px; margin-bottom: 6px; }
        .teacher-link { display: block; margin: 20px auto 0; background: none; border: none; color: var(--muted); font-family: var(--font-body); font-size: 13px; cursor: pointer; text-decoration: underline; opacity: 0.6; } .teacher-link:hover { opacity: 1; }
        .loader-ring { width: 44px; height: 44px; border: 3.5px solid var(--border); border-top-color: var(--accent); border-radius: 50%; animation: spin 0.9s linear infinite; }
        @keyframes spin { to { transform: rotate(360deg); } }
        @keyframes fadeIn { from { opacity: 0; transform: translateY(6px); } to { opacity: 1; transform: translateY(0); } }
        .card, .topic-btn { animation: fadeIn 0.25s ease both; }
      `}</style>

      {view === "login" && <LoginScreen onLogin={handleLogin} onTeacher={() => setView("teacher-login")} />}
      {view === "topics" && <TopicSelect grade={grade} studentName={studentName} onSelect={(t) => { setTopic(t); setResumeState(null); setView("quiz"); }} onBack={() => setView("login")} />}
      {view === "quiz" && <QuizScreen grade={grade} topic={topic} studentName={studentName} onFinish={() => setView("topics")} onBack={() => setView("topics")} resumeState={resumeState} />}
      {view === "teacher-login" && <TeacherLogin onAuth={() => setView("teacher")} onBack={() => setView("login")} />}
      {view === "teacher" && <TeacherDashboard onBack={() => setView("login")} />}
    </div>
  );
}
ReactDOM.createRoot(document.getElementById("root")).render(<App />);
