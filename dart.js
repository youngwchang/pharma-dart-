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
const wait       = ms => new Promise(r => setTimeout(r, ms));
const toInt      = s  => { const n = parseInt((s||'').replace(/,/g,''),10); return isNaN(n)?null:n; };
const firstMatch = (list, ids, nms) => {
  for (const x of list) if (ids.includes(x.account_id))                        return toInt(x.thstrm_amount);
  for (const x of list) if (nms.some(n => (x.account_nm||'').includes(n)))     return toInt(x.thstrm_amount);
  return null;
};

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
router.get('/batch', async (req, res) => {
  const { corp_code, years='2016,2017,2018,2019,2020,2021,2022,2023,2024,2025' } = req.query;
  if (!corp_code) return res.status(400).json({ error:'corp_code required' });

  const yearList = years.split(',').map(s=>s.trim()).filter(Boolean);
  const fin      = {};
  const emp      = {};

  for (const year of yearList) {
    // ── financial
    await wait(DELAY);
    try {
      const r = await axios.get(`${BASE}/fnlttSinglAcnt.json`, {
        params: { crtfc_key:KEY(), corp_code, bsns_year:year, reprt_code:'11011' },
        timeout: 10000,
      });
      if (r.data.status==='000' && r.data.list?.length) {
        const hasCFS = r.data.list.some(i=>i.fs_div==='CFS');
        const fsDiv  = hasCFS?'CFS':'OFS';
        const IS     = r.data.list.filter(i=>i.sj_div==='IS'&&i.fs_div===fsDiv);
        const BS     = r.data.list.filter(i=>i.sj_div==='BS'&&i.fs_div===fsDiv);
        fin[year] = {
          revenue:     firstMatch(IS,['ifrs-full_Revenue','dart_Revenue'],['매출액','수익(매출액)','영업수익']),
          op_profit:   firstMatch(IS,['dart_OperatingIncomeLoss','ifrs-full_ProfitLossFromOperatingActivities'],['영업이익','영업손익']),
          net_income:  firstMatch(IS,['ifrs-full_ProfitLoss','ifrs-full_ProfitLossAttributableToOwnersOfParent'],['당기순이익','당기순손익']),
          assets:      firstMatch(BS,['ifrs-full_Assets'],['자산총계']),
          liabilities: firstMatch(BS,['ifrs-full_Liabilities'],['부채총계']),
          sga:         firstMatch(IS,['ifrs-full_SellingGeneralAndAdministrativeExpense','dart_Sga'],['판매비와관리비','판매비 및 관리비']),
          cogs:        firstMatch(IS,['ifrs-full_CostOfSales'],['매출원가']),
        };
      } else fin[year]=null;
    } catch { fin[year]=null; }

    // ── employees
    await wait(DELAY);
    try {
      const r = await axios.get(`${BASE}/empSttus.json`, {
        params: { crtfc_key:KEY(), corp_code, bsns_year:year, reprt_code:'11011' },
        timeout: 8000,
      });
      if (r.data.status==='000' && r.data.list?.length) {
        const rows  = r.data.list.filter(e=>!e.fo_bbm||e.fo_bbm.includes('합계')||e.fo_bbm.trim()==='');
        const total = rows.length
          ? Math.max(...rows.map(e=>toInt(e.tot_cnt)||0))
          : r.data.list.reduce((s,e)=>s+(toInt(e.tot_cnt)||0),0);
        emp[year] = total||null;
      } else emp[year]=null;
    } catch { emp[year]=null; }
  }

  // Merge employees into financial
  const merged = {};
  for (const y of yearList) {
    merged[y] = fin[y] ? { ...fin[y], employees: emp[y] } : (emp[y] ? { employees:emp[y] } : null);
  }

  res.json({ corp_code, years: merged });
});

/* health */
router.get('/health', (_req, res) => res.json({ ok:true, keySet:!!KEY() }));

module.exports = router;