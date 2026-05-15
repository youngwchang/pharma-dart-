/**
 * 한국 제약/바이오 DART 재무 분석 Explorer
 * 
 * 수집 방식: Claude API 호출 없음 → Render.com DART 프록시 직접 호출
 * 비용:  토큰 0 / 수집비용 0 (DART API 무료)
 * 속도:  기업당 ~4–5초 → 150개사 약 12분
 */

import { useState, useMemo, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  LineChart, Line, BarChart, Bar, XAxis, YAxis, CartesianGrid,
  Tooltip, Legend, ResponsiveContainer, Cell, ReferenceLine
} from "recharts";
import _ from "lodash";

/* ─────────────────────────────────────────
   CONFIG  (Render.com 서버 URL)
───────────────────────────────────────── */
const PROXY = "/dart";

/* ─────────────────────────────────────────
   제약/바이오 상장사 목록 (~200개)
   corp_code는 프록시 /corplist 에서 이름으로 매칭
───────────────────────────────────────── */
const PHARMA_LIST = [
  // 대형 제약
  {name:"유한양행",type:"의약품·제약",market:"KOSPI"},
  {name:"한미약품",type:"의약품·제약",market:"KOSPI"},
  {name:"GC녹십자",type:"의약품·제약",market:"KOSPI"},
  {name:"대웅제약",type:"의약품·제약",market:"KOSPI"},
  {name:"종근당",type:"의약품·제약",market:"KOSPI"},
  {name:"동아에스티",type:"의약품·제약",market:"KOSPI"},
  {name:"보령",type:"의약품·제약",market:"KOSPI"},
  {name:"일동제약",type:"의약품·제약",market:"KOSPI"},
  {name:"광동제약",type:"의약품·제약",market:"KOSPI"},
  {name:"동화약품",type:"의약품·제약",market:"KOSPI"},
  {name:"한독",type:"의약품·제약",market:"KOSPI"},
  {name:"JW중외제약",type:"의약품·제약",market:"KOSPI"},
  {name:"현대약품",type:"의약품·제약",market:"KOSPI"},
  {name:"신풍제약",type:"의약품·제약",market:"KOSPI"},
  {name:"부광약품",type:"의약품·제약",market:"KOSPI"},
  {name:"삼진제약",type:"의약품·제약",market:"KOSPI"},
  {name:"유나이티드제약",type:"의약품·제약",market:"KOSPI"},
  {name:"경동제약",type:"의약품·제약",market:"KOSPI"},
  {name:"동성제약",type:"의약품·제약",market:"KOSPI"},
  {name:"안국약품",type:"의약품·제약",market:"KOSPI"},
  {name:"대원제약",type:"의약품·제약",market:"KOSPI"},
  {name:"명인제약",type:"의약품·제약",market:"KOSPI"},
  {name:"조아제약",type:"의약품·제약",market:"KOSDAQ"},
  {name:"마더스제약",type:"의약품·제약",market:"KOSDAQ"},
  {name:"국제약품",type:"의약품·제약",market:"KOSDAQ"},
  {name:"비씨월드제약",type:"의약품·제약",market:"KOSDAQ"},
  {name:"환인제약",type:"의약품·제약",market:"KOSPI"},
  {name:"동국제약",type:"의약품·제약",market:"KOSPI"},
  {name:"제일약품",type:"의약품·제약",market:"KOSPI"},
  {name:"영진약품",type:"의약품·제약",market:"KOSDAQ"},
  {name:"한국유나이티드제약",type:"의약품·제약",market:"KOSPI"},
  {name:"일양약품",type:"의약품·제약",market:"KOSPI"},
  {name:"동아제약",type:"의약품·제약",market:"KOSPI"},
  {name:"태극제약",type:"의약품·제약",market:"KOSDAQ"},
  {name:"진양제약",type:"의약품·제약",market:"KOSDAQ"},
  {name:"코오롱제약",type:"의약품·제약",market:"KOSPI"},
  {name:"한미정밀화학",type:"의약품·제약",market:"KOSPI"},
  {name:"대한뉴팜",type:"의약품·제약",market:"KOSDAQ"},
  {name:"동방메디칼",type:"의약품·제약",market:"KOSDAQ"},
  {name:"팜젠사이언스",type:"의약품·제약",market:"KOSDAQ"},
  // 바이오·제약
  {name:"삼성바이오로직스",type:"바이오·제약",market:"KOSPI"},
  {name:"셀트리온",type:"바이오·제약",market:"KOSPI"},
  {name:"셀트리온제약",type:"바이오·제약",market:"KOSDAQ"},
  {name:"에스티팜",type:"바이오·제약",market:"KOSDAQ"},
  {name:"메디톡스",type:"바이오·제약",market:"KOSDAQ"},
  {name:"휴젤",type:"바이오·제약",market:"KOSDAQ"},
  {name:"알테오젠",type:"바이오·제약",market:"KOSDAQ"},
  {name:"레고켐바이오사이언스",type:"바이오·제약",market:"KOSDAQ"},
  {name:"에이비엘바이오",type:"바이오·제약",market:"KOSDAQ"},
  {name:"에이치엘비",type:"바이오·제약",market:"KOSDAQ"},
  {name:"한올바이오파마",type:"바이오·제약",market:"KOSDAQ"},
  {name:"바이넥스",type:"바이오·제약",market:"KOSDAQ"},
  {name:"파마리서치",type:"바이오·제약",market:"KOSDAQ"},
  {name:"동국생명과학",type:"바이오·제약",market:"KOSDAQ"},
  {name:"차바이오텍",type:"바이오·제약",market:"KOSDAQ"},
  {name:"크리스탈지노믹스",type:"바이오·제약",market:"KOSDAQ"},
  {name:"제넥신",type:"바이오·제약",market:"KOSDAQ"},
  {name:"유틸렉스",type:"바이오·제약",market:"KOSDAQ"},
  {name:"지놈앤컴퍼니",type:"바이오·제약",market:"KOSDAQ"},
  {name:"압타바이오",type:"바이오·제약",market:"KOSDAQ"},
  {name:"보로노이",type:"바이오·제약",market:"KOSDAQ"},
  {name:"오스코텍",type:"바이오·제약",market:"KOSDAQ"},
  {name:"HK이노엔",type:"바이오·제약",market:"KOSPI"},
  {name:"SK바이오팜",type:"바이오·제약",market:"KOSPI"},
  {name:"SK바이오사이언스",type:"바이오·제약",market:"KOSPI"},
  {name:"GC셀",type:"바이오·제약",market:"KOSDAQ"},
  {name:"큐리언트",type:"바이오·제약",market:"KOSDAQ"},
  {name:"리가켐바이오사이언스",type:"바이오·제약",market:"KOSDAQ"},
  {name:"에이프릴바이오",type:"바이오·제약",market:"KOSDAQ"},
  {name:"ABL바이오",type:"바이오·제약",market:"KOSDAQ"},
  {name:"지씨씨엘",type:"바이오·제약",market:"KOSDAQ"},
  {name:"한국비엔씨",type:"바이오·제약",market:"KOSDAQ"},
  {name:"펩트론",type:"바이오·제약",market:"KOSDAQ"},
  {name:"에이비온",type:"바이오·제약",market:"KOSDAQ"},
  {name:"에스바이오메딕스",type:"바이오·제약",market:"KOSDAQ"},
  {name:"큐로셀",type:"바이오·제약",market:"KOSDAQ"},
  {name:"이뮤노텍",type:"바이오·제약",market:"KOSDAQ"},
  {name:"GC녹십자셀",type:"바이오·제약",market:"KOSDAQ"},
  {name:"코아스템켐온",type:"바이오·제약",market:"KOSDAQ"},
  // 의료기기
  {name:"오스템임플란트",type:"의료기기",market:"KOSDAQ"},
  {name:"덴티움",type:"의료기기",market:"KOSDAQ"},
  {name:"인바디",type:"의료기기",market:"KOSDAQ"},
  {name:"씨젠",type:"의료기기",market:"KOSDAQ"},
  {name:"루닛",type:"의료기기",market:"KOSDAQ"},
  {name:"뷰노",type:"의료기기",market:"KOSDAQ"},
  {name:"바텍",type:"의료기기",market:"KOSDAQ"},
  {name:"나노엔텍",type:"의료기기",market:"KOSDAQ"},
  {name:"오리엔트바이오",type:"의료기기",market:"KOSDAQ"},
  {name:"레이",type:"의료기기",market:"KOSDAQ"},
  {name:"바이오니아",type:"의료기기",market:"KOSDAQ"},
  {name:"메디아나",type:"의료기기",market:"KOSDAQ"},
  {name:"솔고바이오",type:"의료기기",market:"KOSDAQ"},
];

/* ─────────────────────────────────────────
   CONSTANTS
───────────────────────────────────────── */
const THIS_YEAR = new Date().getFullYear();
const START_YEAR = 2010; // 선택 가능한 최초 연도

const REPORT_TYPES = [
  { code:"11011", label:"연간 (사업보고서)",        short:"연간" },
  { code:"11013", label:"1분기 (1~3월)",            short:"1Q"  },
  { code:"11012", label:"반기 (4~6월)",         short:"2Q"  },
  { code:"11014", label:"3분기 (7~9월)",        short:"3Q"  },
];

const makeYears = (from, to) =>
  Array.from({ length: to - from + 1 }, (_, i) => from + i);

const TYPE_COLORS = {
  "의약품·제약":"#38bdf8",
  "바이오·제약":"#34d399",
  "의료기기":   "#a78bfa",
  "기타":       "#94a3b8",
};
const wait   = ms => new Promise(r => setTimeout(r, ms));
const fmtP   = v  => v!=null&&!isNaN(v)?`${Number(v).toFixed(1)}%`:"—";
const fmtB   = v  => v!=null&&!isNaN(v)?`${Math.round(v/1e8).toLocaleString()}억`:"—";
const fmtT   = v  => v!=null&&!isNaN(v)?`${(v/1e12).toFixed(2)}조`:"—";
const fmtN   = v  => v!=null&&!isNaN(v)?Number(v).toLocaleString():"—";

/* ─────────────────────────────────────────
   ANALYTICS
───────────────────────────────────────── */
function computeMetrics(rawData, YEARS) {
  return Object.entries(rawData).map(([name, d]) => {
    if (!d?.years) return null;
    const rows = YEARS.map(y => {
      const yd = d.years[String(y)];
      return yd?.revenue > 0 ? { year:y, ...yd } : null;
    }).filter(Boolean);
    if (rows.length < 2) return null;
    const first = rows[0], last = rows[rows.length-1];
    // 연도 차이: 문자열도 안전하게 숫자로 변환
    const yrs = Number(String(last.year).replace(/\D.*$/,"")) -
                Number(String(first.year).replace(/\D.*$/,""));
    const cagr = yrs > 0
      ? (Math.pow(last.revenue/first.revenue, 1/yrs)-1)*100
      : null; // 같은 연도면 CAGR 계산 불가 → "—"
    const avg = (arr, fn) => arr.length ? _.mean(arr.map(fn)) : null;
    return {
      name, type:d.type, market:d.market||"",
      cagr:       cagr!=null ? +cagr.toFixed(1) : null,
      latestRev:  last.revenue,
      latestRevB: +(last.revenue/1e8).toFixed(0),
      opMargin:   avg(rows.filter(r=>r.op_profit!=null&&r.revenue>0), r=>r.op_profit/r.revenue*100),
      netMargin:  avg(rows.filter(r=>r.net_income!=null&&r.revenue>0), r=>r.net_income/r.revenue*100),
      sgaRatio:   avg(rows.filter(r=>r.sga>0&&r.revenue>0),           r=>r.sga/r.revenue*100),
      latestSgaB: last.sga>0 ? +(last.sga/1e8).toFixed(0) : null,  // 최근 판관비 금액(억)
      cogsRatio:  avg(rows.filter(r=>r.cogs>0&&r.revenue>0),          r=>r.cogs/r.revenue*100),
      debtRatio:  avg(rows.filter(r=>r.assets>0&&r.liabilities>0),    r=>r.liabilities/r.assets*100),
      latestEmp:  last.employees||null,
      dataYears:  rows.length,
      rows,
    };
  }).filter(Boolean).sort((a,b)=>(b.cagr??-999)-(a.cagr??-999));
}

/* ─────────────────────────────────────────
   EXCEL EXPORT
───────────────────────────────────────── */
function exportExcel(metrics, rawData, filtered, filterInfo, YEARS) {
  const wb = XLSX.utils.book_new();

  // Sheet 0: 현재뷰
  const h0 = ["기업명","업종","시장","CAGR(%)","최신매출(억)","영업이익률(%)","순이익률(%)","부채비율(%)","데이터(개)"];
  const ws0 = XLSX.utils.aoa_to_sheet([
    [`생성: ${new Date().toLocaleString("ko-KR")} | 필터: ${filterInfo.type} | 정렬: ${filterInfo.sortKey} | ${filtered.length}개사`],
    [], h0,
    ...filtered.map(m=>[
      m.name, m.type, m.market,
      m.cagr!=null?+m.cagr.toFixed(1):null,
      m.latestRevB,
      m.opMargin!=null?+m.opMargin.toFixed(1):null,
      m.netMargin!=null?+m.netMargin.toFixed(1):null,
      m.debtRatio!=null?+m.debtRatio.toFixed(1):null,
      m.dataYears,
    ])
  ]);
  ws0["!cols"] = [16,12,8,8,12,12,10,10,8].map(w=>({wch:w}));
  ws0["!merges"] = [{s:{r:0,c:0},e:{r:0,c:8}}];
  XLSX.utils.book_append_sheet(wb, ws0, "📌현재뷰");

  // Sheet 1: 전체 기업 요약
  const h1 = ["기업명","업종","시장","CAGR(%)","최신매출(억)","영업이익률(%)","순이익률(%)","부채비율(%)","데이터(개)"];
  const ws1 = XLSX.utils.aoa_to_sheet([h1, ...metrics.map(m=>[
    m.name, m.type, m.market,
    m.cagr!=null?+m.cagr.toFixed(1):null,
    m.latestRevB,
    m.opMargin!=null?+m.opMargin.toFixed(1):null,
    m.netMargin!=null?+m.netMargin.toFixed(1):null,
    m.debtRatio!=null?+m.debtRatio.toFixed(1):null,
    m.dataYears,
  ])]);
  ws1["!cols"] = [16,12,8,8,12,12,10,10,8].map(w=>({wch:w}));
  XLSX.utils.book_append_sheet(wb, ws1, "전체기업요약");

  // 연도별 시트
  const makeSheet = (field, label, div=1) => {
    if (!YEARS||!YEARS.length) return;
    const header = ["기업명","업종",...YEARS];
    const rows = metrics.map(m=>[
      m.name, m.type,
      ...YEARS.map(y=>{
        const v = rawData[m.name]?.years?.[String(y)]?.[field];
        return v!=null ? +(v/div).toFixed(0) : null;
      })
    ]);
    const ws = XLSX.utils.aoa_to_sheet([header,...rows]);
    ws["!cols"] = [16,12,...YEARS.map(()=>({wch:10}))];
    XLSX.utils.book_append_sheet(wb, ws, label);
  };
  const makeRatioSheet = (num, den, label) => {
    if (!YEARS||!YEARS.length) return;
    const header = ["기업명","업종",...YEARS];
    const rows = metrics.map(m=>[
      m.name, m.type,
      ...YEARS.map(y=>{
        const yd = rawData[m.name]?.years?.[String(y)];
        return yd?.[num]!=null&&yd?.[den]>0 ? +(yd[num]/yd[den]*100).toFixed(1) : null;
      })
    ]);
    const ws = XLSX.utils.aoa_to_sheet([header,...rows]);
    ws["!cols"] = [16,12,...YEARS.map(()=>({wch:8}))];
    XLSX.utils.book_append_sheet(wb, ws, label);
  };

  makeSheet("revenue",    "매출_억원",   1e8);
  makeSheet("op_profit",  "영업이익_억", 1e8);
  makeSheet("assets",     "자산_억",     1e8);
  makeSheet("liabilities","부채_억",     1e8);
  makeRatioSheet("op_profit","revenue",  "영업이익률_%");
  makeRatioSheet("net_income","revenue", "순이익률_%");

  XLSX.writeFile(wb, `한국제약바이오_DART_${new Date().toISOString().slice(0,10)}.xlsx`);
}

/* ─────────────────────────────────────────
   HTML SNAPSHOT EXPORT
   현재 필터/정렬 상태를 그대로 자체완결 HTML로
───────────────────────────────────────── */
function exportHTML(metrics, rawData, filtered, filterInfo, trend) {
  const YEARS_ALL = [2016,2017,2018,2019,2020,2021,2022,2023,2024,2025];
  const dateStr   = new Date().toLocaleString("ko-KR");
  const TC = {"의약품·제약":"#38bdf8","바이오·제약":"#34d399","의료기기":"#a78bfa","기타":"#94a3b8"};
  const fp = (v,d=1) => v!=null&&!isNaN(v)?`${Number(v).toFixed(d)}%`:"—";
  const fb = v => v!=null&&!isNaN(v)?`${Math.round(v/1e8).toLocaleString()}억`:"—";

  /* SVG 스파크라인 생성 */
  const sparkline = (values, color="#38bdf8", w=120, h=32) => {
    const nums = values.filter(v=>v!=null&&!isNaN(v));
    if (nums.length < 2) return `<svg width="${w}" height="${h}"><text x="4" y="20" fill="#334155" font-size="9">—</text></svg>`;
    const min=Math.min(...nums), max=Math.max(...nums), range=max-min||1;
    const pts = values.map((v,i)=>{
      if(v==null||isNaN(v)) return null;
      const x = (i/(values.length-1))*(w-8)+4;
      const y = h-4-((v-min)/range)*(h-8);
      return `${x},${y}`;
    }).filter(Boolean).join(" ");
    return `<svg width="${w}" height="${h}" style="overflow:visible">
      <polyline points="${pts}" fill="none" stroke="${color}" stroke-width="1.5" stroke-linejoin="round"/>
      <circle cx="${pts.split(" ").at(-1)?.split(",")[0]}" cy="${pts.split(" ").at(-1)?.split(",")[1]}" r="2.5" fill="${color}"/>
    </svg>`;
  };

  /* 업종별 평균 */
  const typeStats = Object.keys(TC).map(t=>{
    const ms = metrics.filter(m=>m.type===t);
    if(!ms.length) return null;
    const avg = arr => arr.length?arr.reduce((s,v)=>s+v,0)/arr.length:null;
    return {
      type:t, count:ms.length,
      cagr:    avg(ms.map(m=>m.cagr).filter(v=>v!=null)),
      opM:     avg(ms.filter(m=>m.opMargin!=null).map(m=>m.opMargin)),
      sgaR:    avg(ms.filter(m=>m.sgaRatio!=null).map(m=>m.sgaRatio)),
      cogsR:   avg(ms.filter(m=>m.cogsRatio!=null).map(m=>m.cogsRatio)),
    };
  }).filter(Boolean);

  /* 산업 트렌드 스파크 데이터 */
  const trendRevs = trend.map(t=>t.totalRev);
  const trendOp   = trend.map(t=>t.opMargin);

  const html = `<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>한국 제약/바이오 재무분석 — ${dateStr}</title>
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  ::-webkit-scrollbar{width:5px;height:5px}
  ::-webkit-scrollbar-thumb{background:#d1d5db;border-radius:3px}
  body{background:#ffffff;color:#111827;font-family:'Malgun Gothic','Apple SD Gothic Neo',sans-serif;font-size:12px;line-height:1.6;padding:28px 24px}
  h1{font-size:22px;font-weight:900;letter-spacing:-0.03em;margin-bottom:4px;color:#0f172a}
  h2{font-size:13px;font-weight:700;color:#374151;margin-bottom:14px;padding-bottom:8px;border-bottom:2px solid #e5e7eb}
  .tag{display:inline-block;padding:2px 8px;border-radius:4px;font-size:10px;font-weight:700}
  .section{margin-bottom:32px}
  /* KPI */
  .kpi-row{display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin-bottom:24px}
  .kpi{background:#ffffff;border-radius:8px;padding:14px 16px;border:1px solid #e5e7eb;border-top:3px solid var(--c);box-shadow:0 1px 3px rgba(0,0,0,0.06)}
  .kpi-label{font-size:9px;color:#6b7280;letter-spacing:.1em;text-transform:uppercase;margin-bottom:6px}
  .kpi-val{font-size:20px;font-weight:800;color:#0f172a;font-family:monospace;line-height:1}
  /* TYPE CARDS */
  .type-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:24px}
  .type-card{background:#ffffff;border-radius:8px;padding:14px 16px;border:1px solid #e5e7eb;border-top:3px solid var(--c);box-shadow:0 1px 3px rgba(0,0,0,0.06)}
  .type-name{font-size:13px;font-weight:700;color:var(--c);margin-bottom:10px}
  .type-row{display:flex;justify-content:space-between;font-size:11px;padding:5px 0;border-bottom:1px solid #f3f4f6}
  .type-key{color:#6b7280}.type-val{font-family:monospace;color:#111827;font-weight:600}
  /* TREND */
  .trend-wrap{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:24px}
  .trend-card{background:#f9fafb;border-radius:8px;padding:14px 16px;border:1px solid #e5e7eb}
  .trend-label{font-size:10px;color:#6b7280;margin-bottom:8px;font-weight:600}
  /* MAIN TABLE */
  .tbl-wrap{background:#ffffff;border-radius:10px;border:1px solid #e5e7eb;overflow-x:auto;box-shadow:0 1px 4px rgba(0,0,0,0.06)}
  table{width:100%;border-collapse:collapse;font-size:11px}
  thead tr{border-bottom:2px solid #e5e7eb}
  th{padding:10px 12px;text-align:right;color:#111827;font-weight:700;white-space:nowrap;background:#f9fafb;cursor:pointer;user-select:none}
  th.left{text-align:left}
  th:hover{background:#f3f4f6}
  td{padding:8px 12px;text-align:right;color:#111827;border-bottom:1px solid #f3f4f6}
  td.left{text-align:left}
  tr:nth-child(even) td{background:#f9fafb}
  tr:hover td{background:#eff6ff!important}
  /* COMPANY SECTIONS */
  .co-section{background:#ffffff;border-radius:10px;border:1px solid #e5e7eb;margin-bottom:12px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,0.05)}
  .co-header{padding:14px 18px;display:flex;justify-content:space-between;align-items:center;cursor:pointer;border-bottom:1px solid #f3f4f6;background:#f9fafb}
  .co-name{font-size:14px;font-weight:700;color:#0f172a}
  .co-kpis{display:flex;gap:8px;flex-wrap:wrap;padding:12px 18px;background:#ffffff;border-bottom:1px solid #f3f4f6}
  .co-k{background:#f9fafb;border-radius:6px;padding:6px 12px;font-size:10px;border:1px solid #e5e7eb}
  .co-k-label{color:#6b7280;display:block;margin-bottom:2px}
  .co-k-val{font-family:monospace;font-weight:700;color:#0f172a}
  .co-body{padding:0 18px 16px}
  .co-tbl table{font-size:10px;margin-top:10px}
  .co-tbl td,.co-tbl th{padding:6px 10px;color:#111827}
  .co-tbl thead tr{background:#f9fafb}
  .co-tbl tr:nth-child(even) td{background:#f9fafb}
  /* FILTER INFO */
  .filter-bar{background:#f0f7ff;border:1px solid #bfdbfe;border-radius:8px;padding:10px 16px;margin-bottom:20px;font-size:11px;color:#374151;display:flex;gap:16px;flex-wrap:wrap}
  .filter-item span{color:#2563eb;font-weight:700}
  /* TOC */
  .toc-btn{position:fixed;bottom:24px;right:24px;width:48px;height:48px;border-radius:50%;background:#2563eb;color:#fff;border:none;font-size:20px;cursor:pointer;box-shadow:0 4px 16px rgba(37,99,235,0.4);z-index:1000;display:flex;align-items:center;justify-content:center;transition:all 0.2s}
  .toc-btn:hover{background:#1d4ed8;transform:scale(1.08)}
  .toc-panel{position:fixed;bottom:84px;right:24px;width:240px;max-height:420px;overflow-y:auto;background:#ffffff;border:1px solid #e5e7eb;border-radius:12px;box-shadow:0 8px 32px rgba(0,0,0,0.15);z-index:999;display:none;flex-direction:column}
  .toc-panel.open{display:flex}
  .toc-header{padding:10px 14px;font-size:11px;font-weight:700;color:#374151;border-bottom:1px solid #e5e7eb;background:#f9fafb;border-radius:12px 12px 0 0;display:flex;justify-content:space-between;align-items:center}
  .toc-search{width:100%;padding:7px 12px;border:none;border-bottom:1px solid #e5e7eb;font-size:11px;outline:none;color:#111827}
  .toc-search::placeholder{color:#9ca3af}
  .toc-list{overflow-y:auto;max-height:300px}
  .toc-item{padding:8px 14px;font-size:11px;color:#111827;cursor:pointer;display:flex;justify-content:space-between;align-items:center;text-decoration:none;border-bottom:1px solid #f3f4f6}
  .toc-item:hover{background:#eff6ff;color:#2563eb}
  .toc-item .toc-type{font-size:9px;padding:1px 6px;border-radius:3px;font-weight:700}
  .toc-item .toc-cagr{font-size:10px;font-family:monospace;color:#6b7280}
  footer{margin-top:32px;padding-top:16px;border-top:1px solid #e5e7eb;font-size:10px;color:#6b7280}
  @media print{.co-body{display:block!important}body{padding:12px}th,td{font-size:9px;padding:5px 7px}}
</style>
</head>
<body>

<div style="margin-bottom:20px">
  <div style="font-size:10px;color:#2563eb;letter-spacing:.1em;margin-bottom:6px;font-weight:600">DART OPENAPI · 한국 제약/바이오 상장사 분석 리포트</div>
  <h1>한국 제약/바이오 상장사 재무 분석</h1>
  <p style="color:#6b7280;font-size:11px;margin-top:6px">생성일시: ${dateStr}</p>
</div>

<!-- FILTER INFO -->
<div class="filter-bar">
  <div class="filter-item">📌 현재 뷰: <span>${filterInfo.type}</span></div>
  <div class="filter-item">🔍 검색: <span>"${filterInfo.search||"전체"}"</span></div>
  <div class="filter-item">↕ 정렬: <span>${filterInfo.sortKey} ${filterInfo.sortDir==="asc"?"오름차순":"내림차순"}</span></div>
  <div class="filter-item">📊 표시 기업: <span>${filtered.length}개사</span> / 전체 ${metrics.length}개사</div>
</div>

<!-- KPI -->
<div class="section">
<h2>📊 종합 지표 요약</h2>
<div class="kpi-row">
${[
  {label:"분석 기업",val:`${metrics.length}개사`,c:"#2563eb"},
  {label:"표시 기업 (현재뷰)",val:`${filtered.length}개사`,c:"#3b82f6"},
  {label:"평균 CAGR",val:fp(filtered.reduce((s,m)=>s+m.cagr,0)/filtered.length),c:"#16a34a"},
  {label:"합산 최신매출",val:(()=>{const s=metrics.reduce((a,m)=>a+(m.latestRev||0),0);return s>0?`${(s/1e12).toFixed(2)}조`:"—"})(),c:"#7c3aed"},
  {label:"평균 영업이익률",val:fp(metrics.filter(m=>m.opMargin!=null).reduce((s,m)=>s+m.opMargin,0)/Math.max(metrics.filter(m=>m.opMargin!=null).length,1)),c:"#ea580c"},
].map(k=>`<div class="kpi" style="--c:${k.c}"><div class="kpi-label">${k.label}</div><div class="kpi-val">${k.val}</div></div>`).join("")}
</div>
</div>

<!-- TREND SPARKLINES -->
<div class="section">
<h2>📈 산업 트렌드 (2016–2025)</h2>
<div class="trend-wrap">
  <div class="trend-card">
    <div class="trend-label">합산 매출 추이 (조원) — ${trend.map(t=>t.year).join("·")}</div>
    <div style="margin-bottom:6px">${sparkline(trendRevs,"#2563eb",580,60)}</div>
    <div style="display:flex;justify-content:space-between;font-size:9px;font-family:monospace;color:#475569">
      <span>2016: ${trendRevs[0]?.toFixed(2)||"—"}조</span>
      <span>2025: ${trendRevs.at(-1)?.toFixed(2)||"—"}조</span>
    </div>
  </div>
  <div class="trend-card">
    <div class="trend-label">평균 영업이익률 추이 (%)</div>
    <div style="margin-bottom:6px">${sparkline(trendOp,"#16a34a",580,60)}</div>
    <div style="display:flex;justify-content:space-between;font-size:9px;font-family:monospace;color:#475569">
      <span>2016: ${trendOp[0]?.toFixed(1)||"—"}%</span>
      <span>2025: ${trendOp.at(-1)?.toFixed(1)||"—"}%</span>
    </div>
  </div>
</div>
</div>

<!-- TYPE CARDS -->
<div class="section">
<h2>🏷 업종별 평균 지표</h2>
<div class="type-grid">
${typeStats.map(t=>`
  <div class="type-card" style="--c:${TC[t.type]||"#64748b"}">
    <div class="type-name">${t.type} <span style="font-size:9px;color:#334155;font-weight:400">(${t.count}개사)</span></div>
    ${[["평균 CAGR",fp(t.cagr)],["영업이익률",fp(t.opM)],["순이익률",fp(t.netM)||"—"]].map(([k,v])=>`
    <div class="type-row"><span class="type-key">${k}</span><span class="type-val">${v}</span></div>`).join("")}
  </div>`).join("")}
</div>
</div>

<!-- MAIN TABLE -->
<div class="section">
<h2>📋 기업 비교표 (현재 뷰: ${filterInfo.type}${filterInfo.search?` · "${filterInfo.search}"`:""} · ${filtered.length}개사)</h2>
<div class="tbl-wrap">
<table id="mainTable">
<thead>
<tr>
  <th class="left" onclick="sortTable(0)">순위</th>
  <th class="left" onclick="sortTable(1)">기업명</th>
  <th class="left" onclick="sortTable(2)">업종</th>
  <th class="left" onclick="sortTable(3)">시장</th>
  <th onclick="sortTable(4)">CAGR</th>
  <th onclick="sortTable(5)">최신매출(억)</th>
  <th onclick="sortTable(6)">영업이익률</th>
  <th onclick="sortTable(7)">순이익률</th>

  <th onclick="sortTable(10)">부채비율</th>

  <th>매출추이</th>
</tr>
</thead>
<tbody>
${filtered.map((m,i)=>{
  const tc = TC[m.type]||"#64748b";
  const revSpark = sparkline(YEARS_ALL.map(y=>rawData[m.name]?.years?.[String(y)]?.revenue!=null?rawData[m.name].years[String(y)].revenue/1e8:null),"#2563eb",100,24);
  return `<tr data-cagr="${m.cagr||0}" data-rev="${m.latestRevB||0}">
  <td style="color:#6b7280">${i+1}</td>
  <td class="left" style="font-weight:700;color:#111827">${m.name}</td>
  <td class="left"><span class="tag" style="background:${tc}22;color:${tc}">${m.type}</span></td>
  <td class="left" style="color:#6b7280;font-size:10px">${m.market||""}</td>
  <td style="font-weight:700;color:${m.cagr>=20?"#16a34a":m.cagr>=5?"#111827":m.cagr<0?"#dc2626":"#6b7280"}">${fp(m.cagr)}</td>
  <td>${m.latestRevB!=null?m.latestRevB.toLocaleString()+"억":"—"}</td>
  <td style="color:${m.opMargin>15?"#16a34a":m.opMargin<0?"#dc2626":"#111827"}">${fp(m.opMargin)}</td>
  <td>${fp(m.netMargin)}</td>
  <td>${fp(m.debtRatio)}</td>
  <td>${revSpark}</td>
</tr>`;}).join("")}
</tbody>
</table>
</div>
</div>

<!-- COMPANY DETAIL SECTIONS -->
<div class="section">
<h2>🏢 기업별 상세 데이터 (현재 뷰 기준 · 클릭하여 펼치기)</h2>
${filtered.map((m,_ci)=>{
  const tc = TC[m.type]||"#64748b";
  const rows = YEARS_ALL.map(y=>{
    const yd = rawData[m.name]?.years?.[String(y)];
    return yd?.revenue>0?{y,...yd}:null;
  }).filter(Boolean);
  if(!rows.length) return "";
  const revSpark = sparkline(rows.map(r=>r.revenue/1e8),"#2563eb",300,40);
  const opSpark  = sparkline(rows.map(r=>r.op_profit!=null&&r.revenue>0?r.op_profit/r.revenue*100:null),"#16a34a",300,40);
  return `
<div class="co-section" id="co-${_ci}">
  <div class="co-header" onclick="toggle(this)">
    <div>
      <span class="co-name">${m.name}</span>
      <span class="tag" style="background:${tc}22;color:${tc};margin-left:8px">${m.type}</span>
      <span style="font-size:9px;color:#6b7280;margin-left:6px">${m.market} · ${m.dataYears}개년</span>
    </div>
    <div style="display:flex;gap:16px;align-items:center">
      <span style="font-family:monospace;font-size:11px;font-weight:700;color:${m.cagr>=20?"#16a34a":m.cagr<0?"#dc2626":"#374151"}">CAGR ${fp(m.cagr)}</span>
      <span style="font-family:monospace;font-size:11px;color:#2563eb">${fb(m.latestRev)}</span>
      <span style="color:#374151;font-size:14px" class="toggle-arrow">▾</span>
    </div>
  </div>
  <div class="co-body">
    <div class="co-kpis">
      ${[["CAGR",fp(m.cagr),m.cagr>=15?"#16a34a":m.cagr<0?"#dc2626":"#2563eb"],["최신매출",fb(m.latestRev),"#2563eb"],["영업이익률",fp(m.opMargin),"#16a34a"],["순이익률",fp(m.netMargin),"#7c3aed"],["부채비율",fp(m.debtRatio),"#ea580c"]].map(([l,v,c])=>`<div class="co-k"><span class="co-k-label">${l}</span><span class="co-k-val" style="color:${c}">${v}</span></div>`).join("")}
    </div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin:0 0 10px;padding:0 0 6px">
      <div style="background:#f8fafc;border-radius:6px;padding:10px 12px">
        <div style="font-size:9px;color:#475569;margin-bottom:6px">매출 추이 (억원)</div>
        ${revSpark}
      </div>
      <div style="background:#f8fafc;border-radius:6px;padding:10px 12px">
        <div style="font-size:9px;color:#475569;margin-bottom:6px">영업이익률 추이 (%)</div>
        ${opSpark}
      </div>
    </div>
    <div class="co-tbl">
      <table>
        <thead><tr>
          <th class="left">연도</th><th>매출(억)</th><th>영업이익</th><th>순이익</th>
          <th>자산</th><th>부채</th>
          <th>영업이익률</th><th>순이익률</th><th>부채비율</th>
        </tr></thead>
        <tbody>
        ${rows.map((r,i)=>`<tr style="background:${i%2?"#f8fafc":"transparent"}">
          <td class="left" style="color:#2563eb;font-weight:700">${r.y}</td>
          ${[r.revenue,r.op_profit,r.net_income,r.assets,r.liabilities].map(v=>`<td>${v!=null?Math.round(v/1e8).toLocaleString():"—"}</td>`).join("")}
          <td style="color:${r.op_profit!=null&&r.revenue>0?(r.op_profit/r.revenue>0.15?"#16a34a":r.op_profit/r.revenue<0?"#dc2626":"#111827"):"#111827"}">${r.op_profit!=null&&r.revenue>0?`${(r.op_profit/r.revenue*100).toFixed(1)}%`:"—"}</td>
        </tr>`).join("")}
        </tbody>
      </table>
    </div>
  </div>
</div>`;}).join("")}
</div>

<footer>
  <div style="font-weight:600;color:#374151;margin-bottom:4px">한국 제약/바이오 상장사 재무분석</div>
  <div style="margin-top:4px">출처: DART OpenAPI (금융감독원) · 필터: ${filterInfo.type} · ${filtered.length}/${metrics.length}개사</div>
</footer>

<!-- TOC 플로팅 버튼 -->
<button class="toc-btn" onclick="toggleTOC()" title="기업 목차">☰</button>
<div class="toc-panel" id="tocPanel">
  <div class="toc-header">
    <span>📋 기업 목차 (${filtered.length}개사)</span>
    <span onclick="collapseAll()" style="cursor:pointer;font-size:10px;color:#6b7280;font-weight:400">모두 접기</span>
  </div>
  <input class="toc-search" id="tocSearch" placeholder="기업명 검색..." oninput="filterTOC(this.value)"/>
  <div class="toc-list" id="tocList">
    ${filtered.map(m=>`
    <a class="toc-item" href="#co-${i}" onclick="return tocJump(this)" data-name="${m.name}" data-idx="${i}">
      <span>${m.name}</span>
      <span style="display:flex;gap:6px;align-items:center">
        <span class="toc-type" style="background:${({"의약품·제약":"#dbeafe","바이오·제약":"#dcfce7","의료기기":"#ede9fe"})[m.type]||"#f3f4f6"};color:${({"의약품·제약":"#1d4ed8","바이오·제약":"#16a34a","의료기기":"#6d28d9"})[m.type]||"#374151"}">${m.type.split("·")[0]}</span>
        <span class="toc-cagr">${m.cagr!=null?(m.cagr>=0?"+":"")+m.cagr.toFixed(1)+"%":"—"}</span>
      </span>
    </a>`).join("")}
  </div>
</div>

<script>
// 펼치기/닫기
function toggle(el) {
  const body  = el.nextElementSibling;
  const arrow = el.querySelector('.toggle-arrow');
  const isOpen = body.style.display !== 'none';
  body.style.display = isOpen ? 'none' : 'block';
  if (arrow) arrow.textContent = isOpen ? '▸' : '▾';
}
// 모두 접기
function collapseAll() {
  document.querySelectorAll('.co-body').forEach(b=>{ b.style.display='none'; });
  document.querySelectorAll('.toggle-arrow').forEach(a=>{ a.textContent='▸'; });
}
// TOC 토글
function toggleTOC() {
  const p = document.getElementById('tocPanel');
  p.classList.toggle('open');
  if (p.classList.contains('open')) document.getElementById('tocSearch').focus();
}
// TOC 클릭 → 섹션으로 이동 + 펼치기
function tocJump(el) {
  // 패널 먼저 닫기
  document.getElementById('tocPanel').classList.remove('open');
  // 인덱스 기반으로 섹션 찾기 (한국어 ID 문제 우회)
  const idx = el.getAttribute('data-idx');
  const sec = document.getElementById('co-' + idx);
  if (sec) {
    // 섹션 펼치기
    const body  = sec.querySelector('.co-body');
    const arrow = sec.querySelector('.toggle-arrow');
    if (body)  { body.style.display = 'block'; }
    if (arrow) { arrow.textContent = '▾'; }
    // 패널 닫힌 후 스크롤 (setTimeout으로 렌더링 대기)
    setTimeout(function() {
      const top = sec.getBoundingClientRect().top + window.pageYOffset - 20;
      window.scrollTo({ top: top, behavior: 'smooth' });
    }, 50);
  }
  return false;
}
// TOC 검색 필터
function filterTOC(q) {
  const items = document.querySelectorAll('.toc-item');
  const lq = q.toLowerCase();
  items.forEach(item=>{
    item.style.display = item.dataset.name.toLowerCase().includes(lq) ? 'flex' : 'none';
  });
}
// 테이블 정렬
let sortCol=-1, sortAsc=true;
function sortTable(col) {
  const tbl = document.getElementById('mainTable');
  const ths  = tbl.querySelectorAll('th');
  ths.forEach((th,i)=>{ th.style.color=i===col?"#2563eb":""; });
  const tbody = tbl.querySelector('tbody');
  const rows  = Array.from(tbody.querySelectorAll('tr'));
  sortAsc = sortCol===col ? !sortAsc : true;
  sortCol = col;
  rows.sort((a,b)=>{
    const av = a.querySelectorAll('td')[col]?.textContent.replace(/[%억명,조]/g,'').trim()||'';
    const bv = b.querySelectorAll('td')[col]?.textContent.replace(/[%억명,조]/g,'').trim()||'';
    const an=parseFloat(av),bn=parseFloat(bv);
    if(!isNaN(an)&&!isNaN(bn)) return sortAsc?an-bn:bn-an;
    return sortAsc?av.localeCompare(bv,'ko'):bv.localeCompare(av,'ko');
  });
  rows.forEach(r=>tbody.appendChild(r));
  rows.forEach((r,i)=>{ r.querySelector('td').textContent=i+1; });
}
// 외부 클릭 시 TOC 닫기
document.addEventListener('click', e=>{
  const panel = document.getElementById('tocPanel');
  const btn   = document.querySelector('.toc-btn');
  if (!panel.contains(e.target) && !btn.contains(e.target)) panel.classList.remove('open');
});
</script>
</body>
</html>`;

  const blob = new Blob([html], {type:"text/html;charset=utf-8"});
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement("a");
  a.href     = url;
  a.download = `한국제약바이오_분석보고서_${filterInfo.type}_${new Date().toISOString().slice(0,10)}.html`;
  a.click();
  URL.revokeObjectURL(url);
}

/* ─────────────────────────────────────────
   UI ATOMS
───────────────────────────────────────── */
const Tip = ({ active, payload, label }) => {
  if (!active||!payload?.length) return null;
  return (
    <div style={{background:"#ffffff",border:"1px solid #e2e8f0",borderRadius:8,padding:"10px 14px",fontSize:11,boxShadow:"0 4px 16px rgba(0,0,0,0.08)"}}>
      <div style={{color:"#1f2937",marginBottom:6,fontWeight:600}}>{label}</div>
      {payload.map((p,i)=>(
        <div key={i} style={{display:"flex",justifyContent:"space-between",gap:12}}>
          <span style={{color:"#1f2937"}}>{p.name}</span>
          <b style={{color:p.color||"#0f172a"}}>{typeof p.value==="number"?p.value.toFixed(1):p.value}</b>
        </div>
      ))}
    </div>
  );
};

const Kpi = ({label, value, color="#2563eb", sub}) => (
  <div style={{background:"#ffffff",borderRadius:8,padding:"13px 16px",flex:1,minWidth:110,
    boxShadow:"0 1px 4px rgba(0,0,0,0.08)",border:"1px solid #e2e8f0",borderTop:`3px solid ${color}`}}>
    <div style={{fontSize:9,color:"#1f2937",letterSpacing:"0.1em",textTransform:"uppercase",marginBottom:6}}>{label}</div>
    <div style={{fontSize:20,fontWeight:800,color:"#0f172a",fontFamily:"'IBM Plex Mono',monospace",lineHeight:1}}>{value}</div>
    {sub&&<div style={{fontSize:9,color:"#1f2937",marginTop:4}}>{sub}</div>}
  </div>
);

/* ─────────────────────────────────────────
   COMPANY DETAIL
───────────────────────────────────────── */
function CompanyDetail({name, d, metrics}) {
  const m = metrics.find(x=>x.name===name);
  if (!m) return <div style={{color:"#1f2937",padding:20}}>데이터 없음</div>;
  const tc = TYPE_COLORS[m.type]||"#64748b";
  const revData  = m.rows.map(r=>({year:r.year,"매출(억)":+(r.revenue/1e8).toFixed(0)}));
  const margData = m.rows.map(r=>({year:r.year,
    "영업이익률": r.op_profit!=null&&r.revenue>0?+(r.op_profit/r.revenue*100).toFixed(1):null,
    "순이익률":   r.net_income!=null&&r.revenue>0?+(r.net_income/r.revenue*100).toFixed(1):null,
  }));

  return (
    <div>
      <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:14}}>
        <div>
          <div style={{fontSize:18,fontWeight:900,color:"#0f172a"}}>{name}</div>
          <div style={{display:"flex",gap:6,marginTop:4}}>
            {[m.type,m.market,`데이터 ${m.dataYears}개`].map(t=>(
              <span key={t} style={{background:"#e5e7eb",color:t===m.type?tc:"#64748b",borderRadius:4,padding:"2px 7px",fontSize:9,fontWeight:700}}>{t}</span>
            ))}
          </div>
        </div>
      </div>

      {/* KPI 카드 */}
      <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:14}}>
        <Kpi label="CAGR"      value={m.cagr!=null?fmtP(m.cagr):"—"} color={m.cagr>=15?"#16a34a":m.cagr<0?"#dc2626":"#2563eb"}/>
        <Kpi label="최신매출"   value={fmtB(m.latestRev)}              color="#2563eb"/>
        <Kpi label="영업이익률" value={fmtP(m.opMargin)}               color="#16a34a"/>
        <Kpi label="순이익률"   value={fmtP(m.netMargin)}              color="#7c3aed"/>

        <Kpi label="부채비율"   value={fmtP(m.debtRatio)}              color="#ea580c"/>
      </div>

      {/* 차트 2개 */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}>
        {[
          {data:revData,  lines:[{k:"매출(억)",  c:"#2563eb"}], title:"매출 추이 (억원)", u:"억"},
          {data:margData, lines:[{k:"영업이익률",c:"#16a34a"},{k:"순이익률",c:"#7c3aed"}], title:"이익률 추이 (%)", u:"%"},
        ].map(({data,lines,title,u})=>(
          <div key={title} style={{background:"#f8fafc",borderRadius:8,padding:"12px 14px",border:"1px solid #e2e8f0"}}>
            <div style={{fontSize:11,fontWeight:700,color:"#1f2937",marginBottom:10}}>{title}</div>
            <ResponsiveContainer width="100%" height={160}>
              <LineChart data={data}>
                <CartesianGrid strokeDasharray="3 3" stroke="#e5e7eb"/>
                <XAxis dataKey="year" stroke="#e5e7eb" tick={{fontSize:9,fill:"#4b5563"}}/>
                <YAxis stroke="#e5e7eb" tick={{fontSize:9,fill:"#4b5563"}} unit={u} width={44}/>
                <Tooltip content={<Tip/>}/>
                <Legend wrapperStyle={{fontSize:10,color:"#1f2937"}}/>
                {lines.map(l=><Line key={l.k} type="monotone" dataKey={l.k} stroke={l.c} strokeWidth={2.5} dot={{r:3,fill:l.c}} connectNulls/>)}
              </LineChart>
            </ResponsiveContainer>
          </div>
        ))}
      </div>

      {/* 연도별 원시 데이터 */}
      <div style={{background:"#f8fafc",borderRadius:8,border:"1px solid #e2e8f0",overflowX:"auto"}}>
        <div style={{fontSize:11,fontWeight:700,color:"#1f2937",padding:"10px 14px",borderBottom:"1px solid #e2e8f0"}}>연도별 상세 데이터</div>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
          <thead>
            <tr>{["연도","매출(억)","영업이익(억)","순이익(억)","자산(억)","부채(억)","영업이익률","순이익률","부채비율"].map(h=>(
              <th key={h} style={{padding:"7px 10px",textAlign:"right",color:"#111827",fontWeight:700,borderBottom:"1px solid #e2e8f0",fontSize:9,whiteSpace:"nowrap"}}>{h}</th>
            ))}</tr>
          </thead>
          <tbody>
            {m.rows.map((r,i)=>(
              <tr key={r.year} style={{background:i%2?"#f9fafb":"#ffffff", borderBottom:"1px solid #e5e7eb"}}>
                <td style={{padding:"7px 10px", fontWeight:700, color:"#2563eb", fontFamily:"monospace", textAlign:"right"}}>{r.year}</td>
                {[r.revenue,r.op_profit,r.net_income,r.assets,r.liabilities].map((v,j)=>(
                  <td key={j} style={{padding:"7px 10px", fontFamily:"monospace", textAlign:"right", color:"#111827"}}>
                    {v!=null?Math.round(v/1e8).toLocaleString():"—"}
                  </td>
                ))}
                <td style={{padding:"7px 10px", fontFamily:"monospace", textAlign:"right", color:"#111827"}}>
                  {r.op_profit!=null&&r.revenue>0?`${(r.op_profit/r.revenue*100).toFixed(1)}%`:"—"}
                </td>
                <td style={{padding:"7px 10px", fontFamily:"monospace", textAlign:"right", color:"#111827"}}>
                  {r.net_income!=null&&r.revenue>0?`${(r.net_income/r.revenue*100).toFixed(1)}%`:"—"}
                </td>
                <td style={{padding:"7px 10px", fontFamily:"monospace", textAlign:"right", color:"#111827"}}>
                  {r.assets>0&&r.liabilities>0?`${(r.liabilities/r.assets*100).toFixed(1)}%`:"—"}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ─────────────────────────────────────────
   DETAIL LAYOUT (searchable sidebar)
───────────────────────────────────────── */
function DetailLayout({metrics, rawData, selCompany, setSelCo}) {
  const [q,      setQ]      = useState("");
  const [fType,  setFType]  = useState("전체");
  const [sortBy, setSortBy] = useState("cagr");

  const visible = useMemo(()=>{
    let list = metrics;
    if (fType!=="전체") list = list.filter(m=>m.type===fType);
    if (q.trim())       list = list.filter(m=>m.name.includes(q.trim()));
    return [...list].sort((a,b)=>{
      if (sortBy==="name") return a.name.localeCompare(b.name,"ko");
      if (sortBy==="rev")  return (b.latestRevB||0)-(a.latestRevB||0);
      return (b.cagr||0)-(a.cagr||0);
    });
  },[metrics,fType,q,sortBy]);

  const selected = visible.find(m=>m.name===selCompany)?selCompany:visible[0]?.name;

  return (
    <div style={{display:"grid",gridTemplateColumns:"200px 1fr",gap:10,alignItems:"start"}}>
      {/* sidebar */}
      <div style={{background:"#f1f5f9",borderRadius:10,border:"1px solid #d1d5db",overflow:"hidden",position:"sticky",top:8}}>
        {/* search */}
        <div style={{padding:"8px 8px 6px",borderBottom:"1px solid #d1d5db"}}>
          <div style={{position:"relative"}}>
            <span style={{position:"absolute",left:7,top:"50%",transform:"translateY(-50%)",fontSize:10,color:"#1f2937"}}>🔍</span>
            <input value={q} onChange={e=>setQ(e.target.value)} placeholder="기업명..."
              style={{width:"100%",padding:"5px 22px",background:"#ffffff",border:"1px solid #d1d5db",borderRadius:5,color:"#0f172a",fontSize:11,outline:"none",boxSizing:"border-box"}}/>
            {q&&<button onClick={()=>setQ("")} style={{position:"absolute",right:5,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",color:"#1f2937",cursor:"pointer",fontSize:11,padding:0}}>✕</button>}
          </div>
        </div>
        {/* filters */}
        <div style={{padding:"6px 8px",borderBottom:"1px solid #d1d5db",display:"flex",flexDirection:"column",gap:5}}>
          <div style={{display:"flex",gap:3,flexWrap:"wrap"}}>
            {["전체",...Object.keys(TYPE_COLORS)].map(t=>(
              <button key={t} onClick={()=>setFType(t)} style={{
                padding:"2px 6px",fontSize:8,fontWeight:600,borderRadius:8,cursor:"pointer",
                background:fType===t?(TYPE_COLORS[t]||"#38bdf8")+"33":"transparent",
                color:fType===t?(TYPE_COLORS[t]||"#38bdf8"):"#334155",
                border:`1px solid ${fType===t?(TYPE_COLORS[t]||"#38bdf8")+"55":"#d1d5db"}`
              }}>{t==="전체"?"ALL":t.split("·")[0]}</button>
            ))}
          </div>
          <select value={sortBy} onChange={e=>setSortBy(e.target.value)} style={{
            background:"#ffffff",border:"1px solid #d1d5db",borderRadius:4,color:"#1f2937",
            fontSize:9,padding:"2px 4px",outline:"none",cursor:"pointer",width:"100%"
          }}>
            <option value="cagr">CAGR 높은순</option>
            <option value="rev">매출 높은순</option>
            <option value="name">이름순</option>
          </select>
        </div>
        {/* count */}
        <div style={{padding:"4px 10px",borderBottom:"1px solid #d1d5db",fontSize:9,color:"#1f2937",fontFamily:"monospace"}}>
          {visible.length}개 / 전체 {metrics.length}개
        </div>
        {/* list */}
        <div style={{height:500,overflowY:"auto"}}>
          {visible.length===0&&<div style={{padding:16,textAlign:"center",color:"#1f2937",fontSize:11}}>검색 결과 없음</div>}
          {visible.map((m,i)=>{
            const isActive = m.name===selected;
            const tc = TYPE_COLORS[m.type]||"#94a3b8";
            return (
              <div key={m.name} onClick={()=>setSelCo(m.name)}
                style={{padding:"8px 10px",background:isActive?tc+"20":i%2?"#f8fafc":"#ffffff",
                  borderLeft:`3px solid ${isActive?tc:"transparent"}`,cursor:"pointer",
                  borderBottom:"1px solid #0d1526",transition:"all 0.1s"}}
                onMouseEnter={e=>{if(!isActive)e.currentTarget.style.background="#eff6ff"}}
                onMouseLeave={e=>{if(!isActive)e.currentTarget.style.background=i%2?"#f8fafc":"#ffffff"}}
              >
                <div style={{fontSize:11,fontWeight:isActive?700:500,color:isActive?tc:"#1f2937",
                  whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis",marginBottom:2}}>{m.name}</div>
                <div style={{display:"flex",justifyContent:"space-between"}}>
                  <span style={{fontSize:8,color:tc,fontWeight:600}}>{m.type?.split("·")[0]}</span>
                  <span style={{fontSize:9,fontFamily:"monospace",color:m.cagr>=20?"#4ade80":m.cagr<0?"#f87171":"#475569",fontWeight:600}}>
                    {m.cagr!=null?`${m.cagr>=0?"+":""}${m.cagr.toFixed(1)}%`:"—"}
                  </span>
                </div>
              </div>
            );
          })}
        </div>
      </div>
      {/* detail panel */}
      <div>
        {selected
          ? <CompanyDetail name={selected} d={rawData[selected]} metrics={metrics}/>
          : <div style={{padding:40,textAlign:"center",color:"#1f2937",fontSize:12}}>← 기업을 선택하세요</div>
        }
      </div>
    </div>
  );
}

/* ─────────────────────────────────────────
   SORTABLE TABLE
───────────────────────────────────────── */
const COLS = [
  {key:"name",       label:"기업명",     w:130, align:"left", fmt:v=>v},
  {key:"type",       label:"업종",       w:90,  align:"left", fmt:v=>v},
  {key:"market",     label:"시장",       w:65,  align:"left", fmt:v=>v||"—"},
  {key:"cagr",       label:"CAGR",       w:75,  fmt:v=>v!=null?fmtP(v):"—",
    color:v=>v>=20?"#16a34a":v>=5?"#0f172a":v<0?"#dc2626":"#64748b"},
  {key:"latestRevB", label:"최신매출(억)",w:105, fmt:v=>v!=null?`${v.toLocaleString()}억`:"—"},
  {key:"opMargin",   label:"영업이익률", w:85,
    fmt:v=>fmtP(v), color:v=>v>15?"#16a34a":v<0?"#dc2626":"#0f172a"},
  {key:"netMargin",  label:"순이익률",   w:80,  fmt:v=>fmtP(v)},

  {key:"debtRatio",  label:"부채비율",   w:80,  fmt:v=>fmtP(v)},
  {key:"dataYears",  label:"데이터",     w:65,  fmt:v=>`${v}개`},
];

/* ─────────────────────────────────────────
   MAIN
───────────────────────────────────────── */
export default function PharmaDART() {
  const [rawData,   setRawData]   = useState({});
  const [phase,     setPhase]     = useState("idle");
  const [progress,  setProgress]  = useState({n:0,total:0,cur:"",ok:0,fail:0});
  const [log,       setLog]       = useState([]);
  const [errors,    setErrors]    = useState([]);
  const [tab,       setTab]       = useState("table");
  const [selCo,     setSelCo]     = useState("");
  const [sortKey,   setSortKey]   = useState("cagr");
  const [sortDir,   setSortDir]   = useState("desc");
  const [fType,     setFType]     = useState("전체");
  const [search,    setSearch]    = useState("");
  // 기간 설정
  const [periodMode, setPeriodMode] = useState("annual");   // annual | quarterly
  const [yearFrom,   setYearFrom]   = useState(2016);
  const [yearTo,     setYearTo]     = useState(THIS_YEAR);
  const [reprtCode,  setReprtCode]  = useState("11011");     // 연간 모드용
  const [quarters,   setQuarters]   = useState(["Q1","Q2","Q3","Q4"]); // 분기별 모드용
  const [collectedPeriod, setCollectedPeriod] = useState(null);
  const abortRef = useRef(false);
  const addLog = msg => setLog(p=>[...p.slice(-100), `[${new Date().toLocaleTimeString("ko-KR")}] ${msg}`]);

  // 현재 설정으로 기간 키 목록 생성
  const buildPeriodKeys = (from, to, mode, qts, rCode) => {
    if (mode === "quarterly") {
      const keys = [];
      for (let y = from; y <= to; y++)
        for (const q of qts) keys.push(`${y}_${q}`);
      return keys;
    }
    return makeYears(from, to).map(String);
  };

  const periodKeys = collectedPeriod
    ? buildPeriodKeys(collectedPeriod.yearFrom, collectedPeriod.yearTo,
                      collectedPeriod.mode, collectedPeriod.quarters, collectedPeriod.reprtCode)
    : buildPeriodKeys(yearFrom, yearTo, periodMode, quarters, reprtCode);

  const periodLabel = k => periodMode==="quarterly" || (collectedPeriod?.mode==="quarterly")
    ? k.replace("_", " ")           // "2024_Q1" → "2024 Q1"
    : k;                            // "2024" 그대로

  /* ── COLLECT ── */
  const startCollection = async () => {
    abortRef.current = false;
    setPhase("resolving");
    setRawData({});
    setLog([]);
    setErrors([]);

    const meta = { yearFrom, yearTo, mode:periodMode, quarters, reprtCode };
    setCollectedPeriod(meta);

    const modeDesc = periodMode==="quarterly"
      ? `분기별 (${quarters.join(",")})`
      : REPORT_TYPES.find(r=>r.code===reprtCode)?.label||"연간";
    addLog(`기간: ${yearFrom}~${yearTo} · ${modeDesc}`);
    addLog("Step 1: DART corplist 조회 중...");

    let nameMap = {};
    try {
      const r = await fetch(`${PROXY}/corplist`);
      const j = await r.json();
      if (j.list) {
        for (const c of j.list) nameMap[c.corp_name] = c.corp_code;
        addLog(`✅ ${j.total.toLocaleString()}개 corp_code 로드`);
      }
    } catch(e) { addLog(`⚠ corplist 실패: ${e.message}`); setPhase("idle"); return; }

    const resolved = PHARMA_LIST.map(c => {
      const corp_code = nameMap[c.name];
      if (!corp_code) addLog(`⚠ 미발견: ${c.name}`);
      return { ...c, corp_code };
    }).filter(c => c.corp_code);

    addLog(`Step 2: ${resolved.length}개사 매칭`);
    setPhase("collecting");
    setProgress({n:0, total:resolved.length, cur:"", ok:0, fail:0});

    for (let i=0; i<resolved.length; i++) {
      if (abortRef.current) { addLog("⏹ 중단"); break; }
      const c = resolved[i];
      setProgress(p=>({...p, n:i+1, cur:c.name}));

      try {
        let url;
        if (periodMode === "quarterly") {
          url = `${PROXY}/batch?corp_code=${c.corp_code}&mode=quarterly`
              + `&year_from=${yearFrom}&year_to=${yearTo}&quarters=${quarters.join(",")}`;
        } else {
          const yrs = makeYears(yearFrom, yearTo).join(",");
          url = `${PROXY}/batch?corp_code=${c.corp_code}&mode=annual&years=${yrs}&reprt_code=${reprtCode}`;
        }
        const r = await fetch(url);
        if (!r.ok) throw new Error(`HTTP ${r.status}`);
        const j = await r.json();
        const cnt = Object.values(j.years||{}).filter(y=>y?.revenue>0).length;
        setRawData(prev=>({...prev, [c.name]:{type:c.type, market:c.market, years:j.years||{}}}));
        setProgress(p=>({...p, ok:p.ok+1}));
        addLog(`✓ ${c.name}: ${cnt}개 기간`);
      } catch(e) {
        setErrors(p=>[...p, `${c.name}: ${e.message}`]);
        setProgress(p=>({...p, fail:p.fail+1}));
        addLog(`✗ ${c.name}: ${e.message.slice(0,50)}`);
      }
    }
    setPhase("done");
    addLog(`🎉 완료`);
  };

  const YEARS = collectedPeriod ? buildPeriodKeys(
    collectedPeriod.yearFrom, collectedPeriod.yearTo,
    collectedPeriod.mode, collectedPeriod.quarters, collectedPeriod.reprtCode
  ) : buildPeriodKeys(yearFrom, yearTo, periodMode, quarters, reprtCode);

  const metrics = useMemo(()=>computeMetrics(rawData, YEARS),[rawData, YEARS]);
  const trend   = useMemo(()=>YEARS.map(y=>{
    const ys=String(y);
    const vs=Object.values(rawData).filter(d=>d?.years?.[ys]?.revenue>0).map(d=>d.years[ys]);
    if(!vs.length) return null;
    const sg=vs.filter(v=>v.sga>0&&v.revenue>0),cg=vs.filter(v=>v.cogs>0&&v.revenue>0),op=vs.filter(v=>v.op_profit!=null&&v.revenue>0);
    return {year:y,companies:vs.length,
      totalRev:+(_.sum(vs.map(v=>v.revenue))/1e12).toFixed(2),
      opMargin: op.length?+(_.mean(op.map(v=>v.op_profit/v.revenue*100))).toFixed(1):null,
      sgaRatio: sg.length?+(_.mean(sg.map(v=>v.sga/v.revenue*100))).toFixed(1):null,
      cogsRatio:cg.length?+(_.mean(cg.map(v=>v.cogs/v.revenue*100))).toFixed(1):null,
    };
  }).filter(Boolean),[rawData]);

  const filtered = useMemo(()=>{
    let list = metrics;
    if (fType!=="전체") list=list.filter(m=>m.type===fType);
    if (search.trim()) list=list.filter(m=>m.name.includes(search.trim()));
    return [...list].sort((a,b)=>{
      const av=a[sortKey],bv=b[sortKey];
      if(av==null&&bv==null) return 0; if(av==null) return 1; if(bv==null) return -1;
      const cmp=typeof av==="string"?av.localeCompare(bv,"ko"):av-bv;
      return sortDir==="asc"?cmp:-cmp;
    });
  },[metrics,sortKey,sortDir,fType,search]);

  const handleSort = k => {
    if(sortKey===k) setSortDir(d=>d==="asc"?"desc":"asc");
    else {setSortKey(k);setSortDir("desc");}
  };

  const pct = progress.total>0 ? Math.round(progress.n/progress.total*100) : 0;
  const TYPES = ["전체",...Object.keys(TYPE_COLORS)];
  const TABS_DEF = [{id:"table",label:"📋 비교표"},{id:"detail",label:"🏢 기업 상세"},{id:"trend",label:"📈 산업 트렌드"}];

  return (
    <div style={{background:"#ffffff",minHeight:"100vh",color:"#0f172a",
      fontFamily:"'Pretendard','Apple SD Gothic Neo',sans-serif",padding:"20px 18px",maxWidth:1140,margin:"0 auto"}}>
      <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;700&display=swap"/>
      <style>{`*{box-sizing:border-box}*::-webkit-scrollbar{width:5px;height:5px}*::-webkit-scrollbar-track{background:#f1f5f9}*::-webkit-scrollbar-thumb{background:#d1d5db;border-radius:3px}.th{cursor:pointer;user-select:none}.th:hover{color:#e2e8f0!important}`}</style>

      {/* HEADER */}
      <div style={{marginBottom:18}}>
        <div style={{fontSize:10,color:"#38bdf8",fontFamily:"monospace",letterSpacing:"0.15em",marginBottom:4,color:"#374151"}}>
          DART OPENAPI DIRECT · 토큰 0 · 비용 0
        </div>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
          <div>
            <h1 style={{fontSize:17,fontWeight:900,margin:0,letterSpacing:"-0.03em"}}>
              한국 제약/바이오 상장사 재무 분석
            </h1>
            <p style={{color:"#1f2937",margin:"4px 0 0",fontSize:11}}>
              DART API 직접 호출 · {PHARMA_LIST.length}개사 대상
              {collectedPeriod && ` · ${collectedPeriod.yearFrom}~${collectedPeriod.yearTo} ${
                collectedPeriod.mode==="quarterly"
                  ? `분기(${(collectedPeriod.quarters||[]).join(",")})`
                  : REPORT_TYPES.find(r=>r.code===collectedPeriod.reprtCode)?.short||"연간"
              }`}
            </p>
          </div>
          <div style={{display:"flex",gap:8}}>
            <button
              onClick={()=>exportExcel(metrics,rawData,filtered,{type:fType,search,sortKey,sortDir},YEARS)}
              disabled={!metrics.length}
              style={{padding:"8px 14px",background:metrics.length?"#2563eb":"#e5e7eb",color:metrics.length?"#ffffff":"#9ca3af",
                border:`1px solid ${metrics.length?"#1e40af":"#d1d5db"}`,borderRadius:7,cursor:metrics.length?"pointer":"default",
                fontSize:11,fontWeight:700,boxShadow:metrics.length?"0 0 12px #1d4ed840":"none"}}>
              📊 Excel ({filtered.length}/{metrics.length})
            </button>
            <button
              onClick={()=>exportHTML(metrics,rawData,filtered,{type:fType,search,sortKey,sortDir},trend)}
              disabled={!metrics.length}
              style={{padding:"8px 14px",background:metrics.length?"#7c3aed":"#e5e7eb",color:metrics.length?"#ffffff":"#9ca3af",
                border:`1px solid ${metrics.length?"#6d28d9":"#d1d5db"}`,borderRadius:7,cursor:metrics.length?"pointer":"default",
                fontSize:11,fontWeight:700,boxShadow:metrics.length?"0 0 12px #6d28d940":"none"}}>
              🖨 HTML 보고서
            </button>
          </div>
        </div>
      </div>

      {/* COLLECTION PANEL */}
      <div style={{background:"#f1f5f9",borderRadius:10,padding:16,marginBottom:16,border:"1px solid #d1d5db"}}>

        {/* 1행: 데이터 단위 */}
        <div style={{display:"flex",gap:8,alignItems:"center",marginBottom:12}}>
          <div style={{fontSize:9,color:"#1f2937",letterSpacing:"0.1em",whiteSpace:"nowrap"}}>데이터 단위</div>
          {[{v:"annual",l:"연간"},{v:"quarterly",l:"분기별"}].map(m=>(
            <button key={m.v} onClick={()=>setPeriodMode(m.v)}
              disabled={phase==="resolving"||phase==="collecting"}
              style={{padding:"5px 14px",fontSize:11,fontWeight:600,borderRadius:6,cursor:"pointer",
                background:periodMode===m.v?"#1e3a5f":"#f1f5f9",
                color:periodMode===m.v?"#38bdf8":"#475569",
                border:`1px solid ${periodMode===m.v?"#1d4ed8":"#d1d5db"}`,transition:"all 0.15s"}}>
              {m.l}
            </button>
          ))}
          {/* 연간 모드: 보고서 종류 */}
          {periodMode==="annual" && REPORT_TYPES.map(r=>(
            <button key={r.code} onClick={()=>setReprtCode(r.code)}
              disabled={phase==="resolving"||phase==="collecting"}
              style={{padding:"5px 12px",fontSize:10,fontWeight:600,borderRadius:6,cursor:"pointer",
                background:reprtCode===r.code?"#0c2340":"#f1f5f9",
                color:reprtCode===r.code?"#7dd3fc":"#334155",
                border:`1px solid ${reprtCode===r.code?"#0369a1":"#d1d5db"}`,transition:"all 0.15s"}}>
              {r.short}
            </button>
          ))}
          {/* 분기별 모드: 분기 선택 */}
          {periodMode==="quarterly" && [
            {q:"Q1", label:"1Q (1~3월)"},
            {q:"Q2", label:"2Q (4~6월)"},
            {q:"Q3", label:"3Q (7~9월)"},
            {q:"Q4", label:"4Q (연간 전체)"},
          ].map(({q, label})=>(
            <button key={q} onClick={()=>setQuarters(prev=>
              prev.includes(q) ? prev.filter(x=>x!==q) : [...prev,q].sort()
            )}
              disabled={phase==="resolving"||phase==="collecting"}
              style={{padding:"5px 12px",fontSize:10,fontWeight:600,borderRadius:6,cursor:"pointer",
                background:quarters.includes(q)?"#1a2e1a":"#f1f5f9",
                color:quarters.includes(q)?"#4ade80":"#334155",
                border:`1px solid ${quarters.includes(q)?"#166534":"#d1d5db"}`,transition:"all 0.15s"}}>
              {label}
            </button>
          ))}
        </div>

        {/* 2행: 연도 범위 + 수집 버튼 */}
        <div style={{display:"flex",gap:10,alignItems:"flex-end",flexWrap:"wrap"}}>
          <div>
            <div style={{fontSize:9,color:"#1f2937",marginBottom:5,letterSpacing:"0.1em"}}>시작 연도</div>
            <select value={yearFrom} onChange={e=>setYearFrom(+e.target.value)}
              disabled={phase==="resolving"||phase==="collecting"}
              style={{background:"#ffffff",border:"1px solid #d1d5db",borderRadius:6,color:"#0f172a",
                fontSize:12,padding:"7px 10px",outline:"none",cursor:"pointer"}}>
              {Array.from({length: THIS_YEAR - START_YEAR + 1}, (_,i)=>START_YEAR+i).map(y=>(
                <option key={y} value={y}>{y}년</option>
              ))}
            </select>
          </div>
          <div style={{fontSize:14,color:"#1f2937",paddingBottom:8}}>~</div>
          <div>
            <div style={{fontSize:9,color:"#1f2937",marginBottom:5,letterSpacing:"0.1em"}}>종료 연도</div>
            <select value={yearTo} onChange={e=>setYearTo(+e.target.value)}
              disabled={phase==="resolving"||phase==="collecting"}
              style={{background:"#ffffff",border:"1px solid #d1d5db",borderRadius:6,color:"#0f172a",
                fontSize:12,padding:"7px 10px",outline:"none",cursor:"pointer"}}>
              {Array.from({length: THIS_YEAR - START_YEAR + 1}, (_,i)=>START_YEAR+i)
                .filter(y=>y>=yearFrom).map(y=>(
                <option key={y} value={y}>{y}년</option>
              ))}
            </select>
          </div>
          <div style={{fontSize:10,color:"#1f2937",paddingBottom:8,whiteSpace:"nowrap"}}>
            {periodMode==="quarterly"
              ? `${(yearTo-yearFrom+1)*quarters.length}개 기간`
              : `${yearTo-yearFrom+1}개년`}
          </div>
          <div style={{flex:1}}/>
          <button onClick={phase==="collecting"||phase==="resolving"?()=>{abortRef.current=true}:startCollection}
            style={{padding:"9px 20px",background:phase==="collecting"||phase==="resolving"?"#7f1d1d":phase==="done"?"#1e3a5f":"#14532d",
              color:"#fff",border:"none",borderRadius:7,cursor:"pointer",fontWeight:700,fontSize:12,whiteSpace:"nowrap"}}>
            {phase==="collecting"||phase==="resolving"?"⏹ 중단":phase==="done"?"↺ 재수집":"▶ 수집 시작"}
          </button>
        </div>

        {phase!=="idle"&&(
          <div style={{marginTop:12}}>
            <div style={{display:"flex",justifyContent:"space-between",fontSize:10,color:"#1f2937",marginBottom:4,fontFamily:"monospace"}}>
              <span>
                {phase==="resolving"&&"corp_code 해석 중..."}
                {phase==="collecting"&&`[${progress.n}/${progress.total}] ${progress.cur} · 성공 ${progress.ok} / 실패 ${progress.fail}`}
                {phase==="done"&&`완료 · ${metrics.length}개사 데이터 확보`}
              </span>
              <span style={{color:phase==="done"?"#4ade80":"#64748b"}}>{phase==="collecting"?`${pct}%`:phase==="done"?"✓":""}</span>
            </div>
            <div style={{background:"#ffffff",borderRadius:4,height:3,overflow:"hidden",marginBottom:8}}>
              <div style={{width:phase==="resolving"?"20%":phase==="done"?"100%":`${pct}%`,height:"100%",
                background:phase==="done"?"#16a34a":"#1d4ed8",transition:"width 0.4s ease",borderRadius:4}}/>
            </div>
            <div style={{height:56,overflowY:"auto",background:"#ffffff",borderRadius:5,padding:"4px 8px"}}>
              {log.slice(-6).map((l,i)=>(
                <div key={i} style={{fontSize:9,fontFamily:"monospace",lineHeight:1.8,
                  color:l.includes("✓")||l.includes("🎉")||l.includes("✅")?"#4ade80":l.includes("✗")||l.includes("⚠")?"#f59e0b":"#334155"}}>{l}</div>
              ))}
            </div>
          </div>
        )}
      </div>

      {/* TABS */}
      <div style={{display:"flex",gap:2,marginBottom:16,background:"#f1f5f9",padding:4,borderRadius:8,border:"1px solid #d1d5db"}}>
        {TABS_DEF.map(t=>(
          <button key={t.id} onClick={()=>setTab(t.id)}
            style={{flex:1,padding:"8px",background:tab===t.id?"#2563eb":"transparent",color:tab===t.id?"#ffffff":"#374151",
              border:tab===t.id?"1px solid #2563eb":"1px solid transparent",borderRadius:6,cursor:"pointer",fontSize:12,
              fontWeight:tab===t.id?700:400,transition:"all 0.15s"}}>{t.label}</button>
        ))}
      </div>

      {/* TABLE */}
      {tab==="table"&&(
        <div>
          <div style={{display:"flex",gap:8,marginBottom:12,flexWrap:"wrap",alignItems:"center"}}>
            <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="기업명 검색..."
              style={{padding:"6px 11px",background:"#f1f5f9",border:"1px solid #d1d5db",borderRadius:6,color:"#0f172a",fontSize:11,outline:"none",width:140}}/>
            <div style={{display:"flex",gap:4,flexWrap:"wrap"}}>
              {TYPES.map(t=>(
                <button key={t} onClick={()=>setFType(t)} style={{
                  padding:"4px 11px",fontSize:10,fontWeight:600,borderRadius:20,cursor:"pointer",
                  background:fType===t?(TYPE_COLORS[t]||"#38bdf8")+"33":"#0f172a",
                  color:fType===t?(TYPE_COLORS[t]||"#38bdf8"):"#475569",
                  border:`1px solid ${fType===t?(TYPE_COLORS[t]||"#38bdf8")+"55":"#d1d5db"}`
                }}>{t}</button>
              ))}
            </div>
            <div style={{marginLeft:"auto",fontSize:10,color:"#1f2937",fontFamily:"monospace"}}>{filtered.length} / {metrics.length}개사</div>
          </div>
          <div style={{background:"#f1f5f9",borderRadius:10,border:"1px solid #d1d5db",overflowX:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
              <thead>
                <tr style={{borderBottom:"2px solid #e2e8f0"}}>
                  {COLS.map(c=>(
                    <th key={c.key} className="th" onClick={()=>handleSort(c.key)}
                      style={{padding:"8px 10px",textAlign:c.align||"right",color:sortKey===c.key?"#2563eb":"#111827",fontWeight:700,
                        whiteSpace:"nowrap",fontSize:9,letterSpacing:"0.05em",fontFamily:"monospace",minWidth:c.w}}>
                      {c.label}{" "}{sortKey===c.key?sortDir==="asc"?"▲":"▼":"⇅"}
                    </th>
                  ))}
                  <th style={{padding:"8px 9px",color:"#1f2937",fontSize:9}}>상세</th>
                </tr>
              </thead>
              <tbody>
                {filtered.map((m,i)=>(
                  <tr key={m.name}
                    style={{borderBottom:"1px solid #e5e7eb", background:i%2?"#f9fafb":"#ffffff", transition:"background 0.1s", cursor:"pointer"}}
                    onMouseEnter={e=>e.currentTarget.style.background="#f0f4ff"}
                    onMouseLeave={e=>e.currentTarget.style.background=i%2?"#f9fafb":"#ffffff"}>
                    {COLS.map(c=>{
                      const v=m[c.key];
                      return (
                        <td key={c.key} style={{padding:"8px 10px", textAlign:c.align||"right",
                          fontFamily:c.align==="left"?"inherit":"'IBM Plex Mono',monospace",
                          color:"#111827", fontWeight:c.key==="name"?700:400, fontSize:12}}>
                          {c.key==="type"
                            ?<span style={{background:(TYPE_COLORS[v]||"#64748b")+"18", color:TYPE_COLORS[v]||"#64748b", borderRadius:4, padding:"2px 8px", fontSize:10, fontWeight:700}}>{v}</span>
                            :c.fmt(v)}
                        </td>
                      );
                    })}
                    <td style={{padding:"8px 10px", textAlign:"center"}}>
                      <button onClick={()=>{setSelCo(m.name);setTab("detail");}}
                        style={{background:"#e0e7ff", border:"none", borderRadius:4, color:"#2563eb", fontSize:11, cursor:"pointer", padding:"3px 10px", fontWeight:600}}>→</button>
                    </td>
                  </tr>
                ))}
                {!filtered.length&&(
                  <tr><td colSpan={COLS.length+1} style={{padding:30,textAlign:"center",color:"#1f2937",fontSize:11}}>
                    {metrics.length?"조건에 맞는 기업 없음":"수집 후 데이터가 표시됩니다"}
                  </td></tr>
                )}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* DETAIL */}
      {tab==="detail"&&(
        <DetailLayout metrics={metrics} rawData={rawData} selCompany={selCo} setSelCo={setSelCo}/>
      )}

      {/* TREND */}
      {tab==="trend"&&(
        <div>
          <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:14}}>
            <Kpi label="분석 기업" value={`${metrics.length}개`} color="#38bdf8"/>
            <Kpi label="평균 CAGR" value={fmtP(_.mean(metrics.map(m=>m.cagr)))} color="#4ade80"/>
            <Kpi label="합산 최신매출" value={fmtT(_.sum(metrics.map(m=>m.latestRev)))} color="#a78bfa"/>
            <Kpi label="평균 영업이익률" value={fmtP(_.mean(metrics.filter(m=>m.opMargin!=null).map(m=>m.opMargin)))} color="#fb923c"/>
          </div>
          {[
            {keys:[{k:"totalRev",c:"#38bdf8",n:"합산매출(조)"}], title:"합산 매출 추이 (조원)", u:"조"},
            {keys:[{k:"opMargin",c:"#16a34a",n:"영업이익률"},{k:"netMargin",c:"#7c3aed",n:"순이익률"}], title:"평균 이익률 (%)", u:"%"},
          ].map(({keys,title,u})=>(
            <div key={title} style={{background:"#f1f5f9",borderRadius:10,padding:"14px 16px",marginBottom:12,border:"1px solid #d1d5db"}}>
              <div style={{fontSize:11,fontWeight:700,marginBottom:12}}>{title}</div>
              <ResponsiveContainer width="100%" height={190}>
                <LineChart data={trend}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#e5e7eb"/>
                  <XAxis dataKey="year" stroke="#e5e7eb" tick={{fontSize:10,fill:"#1f2937"}}/>
                  <YAxis stroke="#e5e7eb" tick={{fontSize:10,fill:"#1f2937"}} unit={u} width={42}/>
                  <Tooltip content={<Tip/>}/>
                  <Legend wrapperStyle={{fontSize:10,color:"#1f2937"}}/>
                  {keys.map(k=><Line key={k.k} type="monotone" dataKey={k.k} name={k.n} stroke={k.c} strokeWidth={2.5} dot={{r:3,fill:k.c}} connectNulls/>)}
                  <ReferenceLine y={0} stroke="#334155" strokeDasharray="3 3"/>
                </LineChart>
              </ResponsiveContainer>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
