/**
 * routes/dart.js
 * DART OpenAPI 프록시 — hdp-bd-report Render.com 서버에 추가
 *
 * 설치:
 *   npm install adm-zip xml2js axios
 *
 * 환경변수 (Render Dashboard > Environment):
 *   DART_API_KEY=your_key_here
 *
 * server.js 또는 app.js에 추가:
 *   app.use('/dart', require('./routes/dart'));
 */

const express = require('express');
const router  = express.Router();
const AdmZip  = require('adm-zip');
const xml2js  = require('xml2js');
const axios   = require('axios');

const BASE    = 'https://opendart.fss.or.kr/api';
const KEY     = () => process.env.DART_API_KEY;
const DELAY   = 220; // ms between DART calls (안전마진 포함, 한도 초과 방지)

/* ──────── helpers ──────── */
const wait   = ms => new Promise(r => setTimeout(r, ms));
const toInt  = s  => { const n = parseInt((s||'').replace(/,/g,''),10); return isNaN(n)?null:n; };

// 계정 매칭: account_id 우선, 없으면 account_nm 포함 검색
const firstMatch = (list, ids, nms) => {
  for (const x of list) if (ids.includes(x.account_id))                    return toInt(x.thstrm_amount);
  for (const x of list) if (nms.some(n => (x.account_nm||'').includes(n))) return toInt(x.thstrm_amount);
  return null;
};

// 판관비: 여러 계정명 합산 (제조원가 + 판관비 분리 공시 대응)
const sumMatch = (list, nmsList) => {
  let total = null;
  for (const nms of nmsList) {
    const v = firstMatch(list, [], nms);
    if (v != null) total = (total || 0) + v;
  }
  return total;
};

// DART 보고서 코드
const REPRT = { annual:'11011', q1:'11013', q2:'11012', q3:'11014' };

// 재무제표 파싱 공통 함수
function parseFinancial(list) {
  const hasCFS = list.some(i => i.fs_div==='CFS');
  const fsDiv  = hasCFS ? 'CFS' : 'OFS';
  const IS     = list.filter(i => i.sj_div==='IS' && i.fs_div===fsDiv);
  const BS     = list.filter(i => i.sj_div==='BS' && i.fs_div===fsDiv);

  // 판관비: 표준 계정 → 없으면 여러 대체 명칭 합산
  const sga = firstMatch(IS,
    ['ifrs-full_SellingGeneralAndAdministrativeExpense','dart_Sga','dart_SellingGeneralAdministrativeExpenses'],
    ['판매비와관리비','판매비 및 관리비','판관비','판매관리비']
  ) ?? sumMatch(IS, [
    ['판매비'],['관리비'],['영업관리비']
  ]);

  // 매출원가: 제약사는 제품+상품 합산인 경우 있음
  const cogs = firstMatch(IS,
    ['ifrs-full_CostOfSales','dart_CostOfSales'],
    ['매출원가','제품매출원가']
  ) ?? sumMatch(IS, [
    ['제품매출원가'],['상품매출원가'],['용역매출원가']
  ]);

  return {
    revenue:     firstMatch(IS, ['ifrs-full_Revenue','dart_Revenue','dart_OperatingRevenue'], ['매출액','수익(매출액)','영업수익','매출']),
    op_profit:   firstMatch(IS, ['dart_OperatingIncomeLoss','ifrs-full_ProfitLossFromOperatingActivities'], ['영업이익','영업손익','영업이익(손실)']),
    net_income:  firstMatch(IS, ['ifrs-full_ProfitLoss','ifrs-full_ProfitLossAttributableToOwnersOfParent'], ['당기순이익','당기순손익','당기순이익(손실)']),
    assets:      firstMatch(BS, ['ifrs-full_Assets'], ['자산총계']),
    liabilities: firstMatch(BS, ['ifrs-full_Liabilities'], ['부채총계']),
    sga,
    cogs,
  };
}

/* ──────────────────────────────────────────────────────────
   GET /dart/corplist
   상장사 전체 목록 (corp_code, corp_name, stock_code) 반환
   캐시: 서버 메모리 (24시간)
────────────────────────────────────────────────────────── */
let _corpCache = null, _corpCacheAt = 0;

router.get('/corplist', async (req, res) => {
  try {
    if (_corpCache && Date.now() - _corpCacheAt < 24*3600*1000) {
      return res.json({ status:'ok', total:_corpCache.length, list:_corpCache });
    }

    const resp = await axios.get(`${BASE}/corpCode.xml`, {
      params: { crtfc_key: KEY() },
      responseType: 'arraybuffer',
      timeout: 20000,
    });

    const zip  = new AdmZip(Buffer.from(resp.data));
    const xml  = zip.getEntry('CORPCODE.xml').getData().toString('utf-8');
    const parsed = await xml2js.parseStringPromise(xml, { explicitArray:true });

    _corpCache = (parsed.result.list || [])
      .filter(i => (i.stock_code?.[0]||'').trim())          // 상장사만
      .map(i => ({
        corp_code:  i.corp_code?.[0]?.trim(),
        corp_name:  i.corp_name?.[0]?.trim(),
        stock_code: i.stock_code?.[0]?.trim(),
        modify_date:i.modify_date?.[0]?.trim(),
      }));
    _corpCacheAt = Date.now();

    res.json({ status:'ok', total:_corpCache.length, list:_corpCache });
  } catch(e) {
    console.error('[DART /corplist]', e.message);
    res.status(500).json({ error: e.message });
  }
});

/* ──────────────────────────────────────────────────────────
   GET /dart/financial?corp_code=XXXXXX[&years=2016,2017,...,2025]
   단일 기업 10개년 재무제표 (IS + BS 핵심 항목)
────────────────────────────────────────────────────────── */
router.get('/financial', async (req, res) => {
  const { corp_code, years='2016,2017,2018,2019,2020,2021,2022,2023,2024,2025' } = req.query;
  if (!corp_code) return res.status(400).json({ error:'corp_code required' });
  if (!KEY())     return res.status(500).json({ error:'DART_API_KEY not set' });

  const yearList = years.split(',').map(s=>s.trim()).filter(Boolean);
  const result   = {};

  for (const year of yearList) {
    await wait(DELAY);
    try {
      const r = await axios.get(`${BASE}/fnlttSinglAcnt.json`, {
        params: { crtfc_key:KEY(), corp_code, bsns_year:year, reprt_code:'11011' },
        timeout: 10000,
      });

      if (r.data.status !== '000' || !r.data.list?.length) { result[year]=null; continue; }

      // 연결재무제표 우선, 없으면 별도
      const hasCFS = r.data.list.some(i => i.fs_div==='CFS');
      const fsDiv  = hasCFS ? 'CFS' : 'OFS';
      const IS     = r.data.list.filter(i => i.sj_div==='IS' && i.fs_div===fsDiv);
      const BS     = r.data.list.filter(i => i.sj_div==='BS' && i.fs_div===fsDiv);

      result[year] = {
        revenue:     firstMatch(IS, ['ifrs-full_Revenue','dart_Revenue'],
                                    ['매출액','수익(매출액)','영업수익']),
        op_profit:   firstMatch(IS, ['dart_OperatingIncomeLoss','ifrs-full_ProfitLossFromOperatingActivities'],
                                    ['영업이익','영업손익']),
        net_income:  firstMatch(IS, ['ifrs-full_ProfitLoss','ifrs-full_ProfitLossAttributableToOwnersOfParent'],
                                    ['당기순이익','당기순손익']),
        assets:      firstMatch(BS, ['ifrs-full_Assets'],
                                    ['자산총계']),
        liabilities: firstMatch(BS, ['ifrs-full_Liabilities'],
                                    ['부채총계']),
        sga:         firstMatch(IS, ['ifrs-full_SellingGeneralAndAdministrativeExpense','dart_Sga'],
                                    ['판매비와관리비','판매비 및 관리비','판관비']),
        cogs:        firstMatch(IS, ['ifrs-full_CostOfSales'],
                                    ['매출원가']),
        fs_div: fsDiv,
      };
    } catch(e) {
      result[year] = { _error: e.message };
    }
  }

  res.json({ corp_code, years: result });
});

/* ──────────────────────────────────────────────────────────
   GET /dart/employees?corp_code=XXXXXX[&years=...]
   직원 현황 (연도별 총원)
────────────────────────────────────────────────────────── */
router.get('/employees', async (req, res) => {
  const { corp_code, years='2016,2017,2018,2019,2020,2021,2022,2023,2024,2025' } = req.query;
  if (!corp_code) return res.status(400).json({ error:'corp_code required' });

  const yearList = years.split(',').map(s=>s.trim()).filter(Boolean);
  const result   = {};

  for (const year of yearList) {
    await wait(DELAY);
    try {
      const r = await axios.get(`${BASE}/empSttus.json`, {
        params: { crtfc_key:KEY(), corp_code, bsns_year:year, reprt_code:'11011' },
        timeout: 8000,
      });

      if (r.data.status !== '000' || !r.data.list?.length) { result[year]=null; continue; }

      // 합계 행 또는 전체 합산
      const totRows = r.data.list.filter(e =>
        !e.fo_bbm || e.fo_bbm.includes('합계') || e.fo_bbm.trim()===''
      );
      const total = totRows.length
        ? Math.max(...totRows.map(e => toInt(e.tot_cnt)||0))
        : r.data.list.reduce((s,e) => s+(toInt(e.tot_cnt)||0), 0);

      result[year] = total || null;
    } catch {
      result[year] = null;
    }
  }

  res.json({ corp_code, employees: result });
});

/* ──────────────────────────────────────────────────────────
   GET /dart/batch?corp_code=XXXXXX
   financial + employees 한 번에 (아티팩트 호출 최소화)
────────────────────────────────────────────────────────── */
/* ──────────────────────────────────────────────────────────
   GET /dart/batch
   financial + employees 한 번에
   
   연간 모드:   ?corp_code=X&years=2016,2017,...&reprt_code=11011
   분기별 모드: ?corp_code=X&mode=quarterly&year_from=2024&year_to=2026&quarters=Q1,Q2,Q3,Q4
────────────────────────────────────────────────────────── */
/* ──────────────────────────────────────────────────────────
   GET /dart/batch
   financial + employees 한 번에

   연간 모드:   ?corp_code=X&mode=annual&years=2016,...&reprt_code=11011
   분기별 모드: ?corp_code=X&mode=quarterly&year_from=2024&year_to=2026&quarters=Q1,Q2,Q3,Q4
               → 독립 분기값 (누적 차분) 반환
────────────────────────────────────────────────────────── */
router.get('/batch', async (req, res) => {
  const { corp_code, mode='annual' } = req.query;
  if (!corp_code) return res.status(400).json({ error:'corp_code required' });
  if (!KEY())     return res.status(500).json({ error:'DART_API_KEY not set' });

  const result = {};

  /* ── 연간 모드 ── */
  if (mode !== 'quarterly') {
    const years      = (req.query.years || '2016,2017,2018,2019,2020,2021,2022,2023,2024,2025').split(',').map(s=>s.trim());
    const reprt_code = req.query.reprt_code || REPRT.annual;

    for (const year of years) {
      await wait(DELAY);
      try {
        const r = await axios.get(`${BASE}/fnlttSinglAcnt.json`, {
          params: { crtfc_key:KEY(), corp_code, bsns_year:year, reprt_code },
          timeout: 12000,
        });
        result[year] = (r.data.status==='000' && r.data.list?.length)
          ? parseFinancial(r.data.list) : null;
      } catch { result[year] = null; }

      // 직원수
      await wait(DELAY);
      try {
        const r = await axios.get(`${BASE}/empSttus.json`, {
          params: { crtfc_key:KEY(), corp_code, bsns_year:year, reprt_code },
          timeout: 8000,
        });
        if (r.data.status==='000' && r.data.list?.length) {
          const rows  = r.data.list.filter(e => !e.fo_bbm||e.fo_bbm.includes('합계')||e.fo_bbm.trim()===''||e.fo_bbm.trim()==='-');
          const total = (rows.length ? rows : r.data.list).reduce((s,e)=>s+(toInt(e.tot_cnt)||0),0);
          if (result[year]) result[year].employees = total||null;
        }
      } catch { /* 직원수 없어도 계속 */ }
    }
    return res.json({ corp_code, mode, years: result });
  }

  /* ── 분기별 모드: 독립 분기값 = 누적 차분 ── */
  const year_from = parseInt(req.query.year_from || new Date().getFullYear());
  const year_to   = parseInt(req.query.year_to   || new Date().getFullYear());
  const wantQ     = new Set((req.query.quarters||'Q1,Q2,Q3,Q4').split(',').map(s=>s.trim().toUpperCase()));

  // P&L 항목 (누적→독립 차분 대상)
  const PL_FIELDS = ['revenue','op_profit','net_income','sga','cogs'];
  // B/S 항목 (기말 시점값, 차분 불필요)
  const BS_FIELDS = ['assets','liabilities'];

  const sub = (a, b) => {
    // a - b, null 안전 처리
    if (a == null) return null;
    if (b == null) return a;
    return a - b;
  };

  for (let y = year_from; y <= year_to; y++) {
    const yr = String(y);

    // 4개 보고서 순서대로 fetch
    const fetched = {};
    for (const [qKey, reprt_code] of [['Q1',REPRT.q1],['Q2',REPRT.q2],['Q3',REPRT.q3],['Q4',REPRT.annual]]) {
      await wait(DELAY);
      try {
        const r = await axios.get(`${BASE}/fnlttSinglAcnt.json`, {
          params: { crtfc_key:KEY(), corp_code, bsns_year:yr, reprt_code },
          timeout: 12000,
        });
        fetched[qKey] = (r.data.status==='000' && r.data.list?.length)
          ? parseFinancial(r.data.list) : null;
      } catch { fetched[qKey] = null; }
    }

    // 독립 분기값 계산
    const standalone = {
      Q1: fetched.Q1,                                      // Q1 단독 = Q1 누적
      Q2: (() => {                                          // Q2 단독 = 반기 - Q1
        if (!fetched.Q2) return null;
        const out = { ...fetched.Q2 };
        PL_FIELDS.forEach(f => { out[f] = sub(fetched.Q2?.[f], fetched.Q1?.[f]); });
        return out;
      })(),
      Q3: (() => {                                          // Q3 단독 = 3Q누적 - 반기
        if (!fetched.Q3) return null;
        const out = { ...fetched.Q3 };
        PL_FIELDS.forEach(f => { out[f] = sub(fetched.Q3?.[f], fetched.Q2?.[f]); });
        return out;
      })(),
      Q4: (() => {                                          // Q4 단독 = 연간 - 3Q누적
        if (!fetched.Q4) return null;
        const out = { ...fetched.Q4 };
        PL_FIELDS.forEach(f => { out[f] = sub(fetched.Q4?.[f], fetched.Q3?.[f]); });
        return out;
      })(),
    };

    // 원하는 분기만 결과에 포함
    for (const q of ['Q1','Q2','Q3','Q4']) {
      if (wantQ.has(q)) result[`${yr}_${q}`] = standalone[q];
    }
  }

  res.json({ corp_code, mode, years: result });
});

/* health */
router.get('/health', (_req, res) => res.json({ ok:true, keySet:!!KEY() }));

module.exports = router;
