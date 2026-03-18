import React, { useState, useCallback } from 'react';
import * as XLSX from 'xlsx';
import {
  ComposedChart, BarChart, LineChart, AreaChart,
  Bar, Line, Area, XAxis, YAxis, CartesianGrid, Tooltip, Legend,
  ResponsiveContainer
} from 'recharts';

// ─── Color Palette ──────────────────────────────────────────────────────────
const C = {
  bg: '#0f1117',
  card: '#1a1d27',
  border: '#2a2e3d',
  text: '#e8eaed',
  sub: '#8b8fa3',
  accent1: '#4ecdc4',
  accent2: '#ff6b6b',
  accent3: '#ffd93d',
  accent4: '#6c5ce7',
  up: '#00c853',
  down: '#ff5252',
};

const COOP_NAMES = { 2: 'アイチョイス', 6: '岐阜', 7: '一宮' };
const COOP_COLORS = { 2: '#4ecdc4', 6: '#ffd93d', 7: '#6c5ce7' };

// ─── Formatters ─────────────────────────────────────────────────────────────
const fmtNum = (v) => {
  if (v == null || isNaN(v)) return '-';
  const abs = Math.abs(v);
  if (abs >= 1e8) return (v / 1e8).toFixed(1) + '億';
  if (abs >= 1e4) return Math.round(v / 1e4) + '万';
  return v.toLocaleString('ja-JP');
};

const fmtYen = (v) => '¥' + Math.round(v).toLocaleString('ja-JP');

const fmtPct = (v) => (v != null ? v.toFixed(1) + '%' : '-');

const fmtRatio = (curr, prev) => {
  if (prev == null || prev === 0 || curr == null) return null;
  return ((curr - prev) / prev) * 100;
};

const RatioBadge = ({ ratio }) => {
  if (ratio == null) return null;
  const color = ratio >= 0 ? C.up : C.down;
  const sign = ratio >= 0 ? '+' : '';
  return <span style={{ color, fontWeight: 600, fontSize: 12 }}>{sign}{ratio.toFixed(1)}%</span>;
};

const PtBadge = ({ diff }) => {
  if (diff == null) return null;
  const color = diff >= 0 ? C.up : C.down;
  const sign = diff >= 0 ? '+' : '';
  return <span style={{ color, fontWeight: 600, fontSize: 12 }}>{sign}{diff.toFixed(1)}pt</span>;
};

// ─── Data Processing ─────────────────────────────────────────────────────────
const processData = (rows) => {
  const filtered = rows.filter((r) => Number(r['区分']) === 1);

  // yearlyData
  const yearMap = {};
  const coopYearMap = {};

  filtered.forEach((r) => {
    const yr = Number(r['対象年度']);
    const coop = Number(r['生協コード']);

    const addTo = (map, key) => {
      if (!map[key]) {
        map[key] = {
          年度: yr, 生協コード: coop,
          食材供給高_sum: 0, 食材供給高_cnt: 0,
          食材利用人数_sum: 0, 食材利用人数_cnt: 0,
          食材金額割合_sum: 0, 食材金額割合_cnt: 0,
          食材受注人数割合_sum: 0, 食材受注人数割合_cnt: 0,
          食材一人当利用高_sum: 0, 食材一人当利用高_cnt: 0,
          食材一点当単価_sum: 0, 食材一点当単価_cnt: 0,
          食材一人当点数_sum: 0, 食材一人当点数_cnt: 0,
          実質GP_sum: 0, 実質GP_cnt: 0,
          _count: 0,
        };
      }
      const d = map[key];
      d._count++;
      const add = (col) => {
        const val = Number(r[col]);
        if (!isNaN(val)) { d[col + '_sum'] += val; d[col + '_cnt']++; }
      };
      d['食材供給高_sum'] += Number(r['食材供給高']) || 0; d['食材供給高_cnt']++;
      d['実質GP_sum'] += Number(r['実質GP']) || 0; d['実質GP_cnt']++;
      add('食材利用人数');
      add('食材金額割合');
      add('食材受注人数割合');
      add('食材一人当利用高');
      add('食材一点当単価');
      add('食材一人当点数');
    };

    addTo(yearMap, yr);
    addTo(coopYearMap, `${yr}_${coop}`);
  });

  const finalize = (d) => ({
    年度: d.年度,
    生協コード: d.生協コード,
    食材供給高: d['食材供給高_sum'],
    実質GP: d['実質GP_sum'],
    食材利用人数: d['食材利用人数_cnt'] ? d['食材利用人数_sum'] / d['食材利用人数_cnt'] : 0,
    食材金額割合: d['食材金額割合_cnt'] ? d['食材金額割合_sum'] / d['食材金額割合_cnt'] : 0,
    食材受注人数割合: d['食材受注人数割合_cnt'] ? d['食材受注人数割合_sum'] / d['食材受注人数割合_cnt'] : 0,
    食材一人当利用高: d['食材一人当利用高_cnt'] ? d['食材一人当利用高_sum'] / d['食材一人当利用高_cnt'] : 0,
    食材一点当単価: d['食材一点当単価_cnt'] ? d['食材一点当単価_sum'] / d['食材一点当単価_cnt'] : 0,
    食材一人当点数: d['食材一人当点数_cnt'] ? d['食材一人当点数_sum'] / d['食材一人当点数_cnt'] : 0,
    _count: d._count,
  });

  const yearlyData = Object.fromEntries(
    Object.entries(yearMap).map(([k, v]) => [k, finalize(v)])
  );

  const coopYearlyData = Object.fromEntries(
    Object.entries(coopYearMap).map(([k, v]) => [k, finalize(v)])
  );

  // weeklyData
  const weeklyData = filtered.map((r) => ({
    対象年度: Number(r['対象年度']),
    生協コード: Number(r['生協コード']),
    SEQ: Number(r['SEQ']),
    企画号数: r['企画号数'],
    配送週: r['配送週'],
    食材供給高: Number(r['食材供給高']) || 0,
    食材利用人数: Number(r['食材利用人数']) || 0,
    食材金額割合: Number(r['食材金額割合']) || 0,
    食材受注人数割合: Number(r['食材受注人数割合']) || 0,
    食材一人当利用高: Number(r['食材一人当利用高']) || 0,
    食材一点当単価: Number(r['食材一点当単価']) || 0,
    食材一人当点数: Number(r['食材一人当点数']) || 0,
    実質GP: Number(r['実質GP']) || 0,
    SKU: Number(r['SKU']) || 0,
    メニュー数: Number(r['メニュー数']) || 0,
    新規商品数: Number(r['新規商品数']) || 0,
  }));

  return { yearlyData, coopYearlyData, weeklyData };
};

// ─── Latest Full Year Detection ──────────────────────────────────────────────
const getLatestFullYear = (yearlyData) => {
  const years = Object.keys(yearlyData).map(Number).sort((a, b) => a - b);
  if (years.length === 0) return null;
  const maxCount = Math.max(...years.map((y) => yearlyData[y]._count));
  const latestYear = years[years.length - 1];
  const latestCount = yearlyData[latestYear]._count;
  const isIncomplete = latestCount < maxCount / 2;
  const latestFullYear = isIncomplete && years.length >= 2
    ? years[years.length - 2]
    : latestYear;
  const prevYear = years[years.indexOf(latestFullYear) - 1] ?? null;
  return { latestFullYear, prevYear, years, isIncomplete };
};

// ─── Tooltip ─────────────────────────────────────────────────────────────────
const CustomTooltip = ({ active, payload, label, formatter }) => {
  if (!active || !payload || !payload.length) return null;
  return (
    <div style={{
      background: '#1e2130', border: `1px solid ${C.border}`, borderRadius: 8,
      padding: '10px 14px', fontSize: 12, color: C.text,
      boxShadow: '0 8px 32px rgba(0,0,0,0.4)'
    }}>
      <div style={{ marginBottom: 6, color: C.sub, fontWeight: 600 }}>{label}</div>
      {payload.map((p, i) => (
        <div key={i} style={{ color: p.color, marginBottom: 2 }}>
          {p.name}: {formatter ? formatter(p.value, p.name) : p.value}
        </div>
      ))}
    </div>
  );
};

// ─── KPI Card ─────────────────────────────────────────────────────────────────
const KpiCard = ({ label, value, sub, color, badge }) => (
  <div style={{
    background: C.card, borderRadius: 12, padding: '20px 22px', flex: '1 1 200px',
    borderTop: `3px solid ${color}`, minWidth: 180
  }}>
    <div style={{ fontSize: 12, color: C.sub, marginBottom: 8 }}>{label}</div>
    <div style={{ fontSize: 26, fontWeight: 700, color: C.text, marginBottom: 4 }}>{value}</div>
    <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
      {sub && <span style={{ fontSize: 11, color: C.sub }}>{sub}</span>}
      {badge}
    </div>
  </div>
);

// ─── Chart Card ──────────────────────────────────────────────────────────────
const ChartCard = ({ title, children, style }) => (
  <div style={{
    background: C.card, borderRadius: 12, border: `1px solid ${C.border}`,
    padding: 20, ...style
  }}>
    {title && <div style={{ fontSize: 13, fontWeight: 600, color: C.sub, marginBottom: 16 }}>{title}</div>}
    {children}
  </div>
);

// ─── Section Title ────────────────────────────────────────────────────────────
const SectionTitle = ({ icon, children }) => (
  <div style={{ fontSize: 17, fontWeight: 700, color: C.text, margin: '32px 0 16px', display: 'flex', alignItems: 'center', gap: 8 }}>
    <span>{icon}</span>{children}
  </div>
);

// ─── Upload Screen ────────────────────────────────────────────────────────────
const UploadScreen = ({ onLoad }) => {
  const [dragging, setDragging] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');

  const handleFile = useCallback((file) => {
    if (!file) return;
    setLoading(true);
    setError('');
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws);
        const data = processData(rows);
        onLoad(data, file.name);
      } catch (err) {
        setError('ファイルの読み込みに失敗しました: ' + err.message);
      } finally {
        setLoading(false);
      }
    };
    reader.readAsArrayBuffer(file);
  }, [onLoad]);

  const onDrop = (e) => {
    e.preventDefault();
    setDragging(false);
    const f = e.dataTransfer.files[0];
    if (f) handleFile(f);
  };

  return (
    <div style={{ minHeight: '100vh', background: C.bg, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', padding: 24 }}>
      <div style={{ textAlign: 'center', marginBottom: 40 }}>
        <h1 style={{ fontSize: 28, fontWeight: 700, color: C.text, marginBottom: 8 }}>
          <span style={{ color: C.accent1 }}>食材セット</span> マーケティングダッシュボード
        </h1>
        <p style={{ color: C.sub, fontSize: 14 }}>週次販売データを分析・可視化します</p>
      </div>

      <div
        onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
        onDragLeave={() => setDragging(false)}
        onDrop={onDrop}
        style={{
          width: '100%', maxWidth: 520,
          border: `2px dashed ${dragging ? C.accent1 : C.border}`,
          borderRadius: 16, padding: '60px 40px', textAlign: 'center',
          background: dragging ? 'rgba(78,205,196,0.05)' : C.card,
          transition: 'all 0.2s', cursor: 'pointer',
        }}
      >
        {loading ? (
          <div style={{ color: C.accent1, fontSize: 16 }}>読み込み中...</div>
        ) : (
          <>
            <div style={{ fontSize: 48, marginBottom: 16 }}>📊</div>
            <div style={{ color: C.text, fontSize: 16, fontWeight: 600, marginBottom: 8 }}>
              syokuzai.xlsx をアップロードしてください
            </div>
            <div style={{ color: C.sub, fontSize: 13, marginBottom: 24 }}>
              ここにファイルをドラッグ＆ドロップ
            </div>
            <label style={{
              display: 'inline-block', padding: '10px 24px',
              background: C.accent1, color: '#000', borderRadius: 8,
              fontWeight: 600, fontSize: 14, cursor: 'pointer'
            }}>
              ファイルを選択
              <input type="file" accept=".xlsx,.xls" style={{ display: 'none' }}
                onChange={(e) => handleFile(e.target.files[0])} />
            </label>
            <div style={{ color: C.sub, fontSize: 11, marginTop: 20 }}>
              参照パス: C:\Users\n-harada\Desktop\syokuzai.xlsx
            </div>
          </>
        )}
      </div>
      {error && <div style={{ color: C.down, marginTop: 16, fontSize: 13 }}>{error}</div>}
    </div>
  );
};

// ─── Tab 1: 概況 ─────────────────────────────────────────────────────────────
const TabOverview = ({ yearlyData, yearInfo }) => {
  const { latestFullYear, prevYear, years } = yearInfo;
  const cur = yearlyData[latestFullYear] || {};
  const prev = prevYear ? yearlyData[prevYear] || {} : {};
  const firstYear = years[0];
  const first = yearlyData[firstYear] || {};

  const supplyRatio = fmtRatio(cur.食材供給高, prev.食材供給高);
  const peopleRatio = fmtRatio(cur.食材利用人数, prev.食材利用人数);
  const pctDiff = cur.食材金額割合 != null && prev.食材金額割合 != null
    ? cur.食材金額割合 - prev.食材金額割合 : null;
  const gpRatio = fmtRatio(cur.実質GP, prev.実質GP);

  const growthFactor = first.食材金額割合 > 0
    ? (cur.食材金額割合 / first.食材金額割合).toFixed(1) : '-';
  const peopleChange = first.食材利用人数 > 0
    ? ((cur.食材利用人数 - first.食材利用人数) / first.食材利用人数 * 100).toFixed(1) : '-';
  const unitChange = first.食材一人当利用高 > 0
    ? ((cur.食材一人当利用高 - first.食材一人当利用高) / first.食材一人当利用高 * 100).toFixed(1) : '-';
  const gpGrowth = first.実質GP > 0
    ? (cur.実質GP / first.実質GP).toFixed(1) : '-';

  return (
    <div>
      <SectionTitle icon="📊">主要KPI（{latestFullYear}年度）</SectionTitle>

      {/* KPI 上段 */}
      <div style={{ display: 'flex', gap: 16, flexWrap: 'wrap', marginBottom: 16 }}>
        <KpiCard label="食材供給高" value={fmtNum(cur.食材供給高)} sub="年間合計" color={C.accent1}
          badge={<RatioBadge ratio={supplyRatio} />} />
        <KpiCard label="食材利用人数（週平均）" value={Math.round(cur.食材利用人数 || 0).toLocaleString('ja-JP') + '人'}
          sub="週平均" color={C.accent3} badge={<RatioBadge ratio={peopleRatio} />} />
        <KpiCard label="食材金額割合" value={fmtPct(cur.食材金額割合)} sub="総受注高に占める比率"
          color={C.accent4} badge={<PtBadge diff={pctDiff} />} />
        <KpiCard label="実質GP" value={fmtNum(cur.実質GP)} sub="年間合計" color={C.accent2}
          badge={<RatioBadge ratio={gpRatio} />} />
      </div>

      {/* KPI 下段 */}
      <div style={{ display: 'flex', gap: 16, flexWrap: 'wrap', marginBottom: 32 }}>
        <KpiCard label="一人当利用高" value={fmtYen(cur.食材一人当利用高 || 0)} color="#a8e6cf" />
        <KpiCard label="一点当単価" value={fmtYen(cur.食材一点当単価 || 0)} color={C.accent1} />
        <KpiCard label="一人当点数" value={(cur.食材一人当点数 || 0).toFixed(2) + '点'} color={C.accent3} />
        <KpiCard label="食材受注人数割合" value={fmtPct(cur.食材受注人数割合)} color={C.accent2} />
      </div>

      {/* Insight Box */}
      <div style={{
        background: 'linear-gradient(135deg, rgba(78,205,196,0.08), rgba(108,92,231,0.08))',
        border: `1px solid ${C.border}`, borderRadius: 12, padding: '20px 24px'
      }}>
        <div style={{ fontSize: 14, fontWeight: 700, color: C.text, marginBottom: 16 }}>
          💡 マーケター注目ポイント
        </div>
        {[
          {
            icon: '📈',
            label: '浸透率が急上昇中',
            text: `食材金額割合が${firstYear}年度→${latestFullYear}年度で${growthFactor}倍成長。伸びしろがある状況。`
          },
          {
            icon: '👥',
            label: '利用人数の拡大',
            text: `週平均利用人数が${firstYear}年度比+${peopleChange}%増加。継続的な利用者獲得が奏功。`
          },
          {
            icon: '💰',
            label: '客単価も改善傾向',
            text: `一人当利用高は${firstYear}年度比${unitChange > 0 ? '+' : ''}${unitChange}%変化。品揃え強化の効果が反映。`
          },
          {
            icon: '🎯',
            label: '次のアクション',
            text: '生協別浸透率差の横展開、季節企画強化、SKU拡大効果測定を推進。'
          },
        ].map((p, i) => (
          <div key={i} style={{ display: 'flex', gap: 12, marginBottom: i < 3 ? 12 : 0 }}>
            <span style={{ fontSize: 18 }}>{p.icon}</span>
            <div>
              <span style={{ fontWeight: 700, color: C.text, fontSize: 13 }}>{p.label}</span>
              <span style={{ color: C.sub, fontSize: 13 }}> — {p.text}</span>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
};

// ─── Tab 2: 成長推移 ──────────────────────────────────────────────────────────
const TabGrowth = ({ yearlyData, yearInfo }) => {
  const { years } = yearInfo;
  const chartData = years.map((y) => ({
    name: y + '年度',
    食材供給高: yearlyData[y]?.食材供給高 || 0,
    実質GP: yearlyData[y]?.実質GP || 0,
    食材金額割合: yearlyData[y]?.食材金額割合 || 0,
    食材受注人数割合: yearlyData[y]?.食材受注人数割合 || 0,
  }));

  const yFmt = (v) => {
    if (v >= 1e8) return (v / 1e8).toFixed(0) + '億';
    if (v >= 1e4) return (v / 1e4).toFixed(0) + '万';
    return v;
  };

  return (
    <div>
      <SectionTitle icon="📈">食材供給高 &amp; 実質GP 年度推移</SectionTitle>
      <ChartCard>
        <ResponsiveContainer width="100%" height={300}>
          <ComposedChart data={chartData} margin={{ top: 10, right: 20, left: 10, bottom: 0 }}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
            <XAxis dataKey="name" tick={{ fill: C.sub, fontSize: 12 }} axisLine={false} tickLine={false} />
            <YAxis tickFormatter={yFmt} tick={{ fill: C.sub, fontSize: 11 }} axisLine={false} tickLine={false} />
            <Tooltip content={<CustomTooltip formatter={(v) => fmtNum(v)} />} />
            <Legend wrapperStyle={{ color: C.sub, fontSize: 12 }} />
            <Bar dataKey="食材供給高" fill={C.accent1} radius={[4, 4, 0, 0]} name="食材供給高" />
            <Bar dataKey="実質GP" fill={C.accent4} radius={[4, 4, 0, 0]} name="実質GP" />
          </ComposedChart>
        </ResponsiveContainer>
      </ChartCard>

      <SectionTitle icon="📈">浸透率（金額・人数割合）推移</SectionTitle>
      <ChartCard>
        <ResponsiveContainer width="100%" height={280}>
          <LineChart data={chartData} margin={{ top: 10, right: 20, left: 10, bottom: 0 }}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
            <XAxis dataKey="name" tick={{ fill: C.sub, fontSize: 12 }} axisLine={false} tickLine={false} />
            <YAxis tickFormatter={(v) => v + '%'} tick={{ fill: C.sub, fontSize: 11 }} axisLine={false} tickLine={false} />
            <Tooltip content={<CustomTooltip formatter={(v) => v.toFixed(1) + '%'} />} />
            <Legend wrapperStyle={{ color: C.sub, fontSize: 12 }} />
            <Line dataKey="食材金額割合" stroke={C.accent1} strokeWidth={3} dot={{ r: 5, fill: C.accent1 }} name="食材金額割合" />
            <Line dataKey="食材受注人数割合" stroke={C.accent3} strokeWidth={3} dot={{ r: 5, fill: C.accent3 }} name="食材受注人数割合" />
          </LineChart>
        </ResponsiveContainer>
      </ChartCard>
    </div>
  );
};

// ─── Tab 3: 生協比較 ──────────────────────────────────────────────────────────
const TabCoop = ({ coopYearlyData, yearInfo }) => {
  const { latestFullYear, years } = yearInfo;
  const coops = [2, 6, 7];

  // Which coops have data (供給高>0) in latestFullYear
  const activeCoop = coops.filter((c) => {
    const d = coopYearlyData[`${latestFullYear}_${c}`];
    return d && d.食材供給高 > 0;
  });

  // Build chart data
  const supplyChartData = years.map((y) => {
    const row = { name: y + '年度' };
    coops.forEach((c) => {
      const d = coopYearlyData[`${y}_${c}`];
      row[COOP_NAMES[c]] = d ? d.食材供給高 : 0;
    });
    return row;
  });

  const rateChartData = years.map((y) => {
    const row = { name: y + '年度' };
    coops.forEach((c) => {
      const d = coopYearlyData[`${y}_${c}`];
      row[COOP_NAMES[c]] = d && d.食材供給高 > 0 ? d.食材金額割合 : null;
    });
    return row;
  });

  const yFmt = (v) => {
    if (v >= 1e8) return (v / 1e8).toFixed(0) + '億';
    if (v >= 1e4) return (v / 1e4).toFixed(0) + '万';
    return v;
  };

  return (
    <div>
      <SectionTitle icon="🏢">生協別サマリー（{latestFullYear}年度）</SectionTitle>
      <div style={{ display: 'flex', gap: 16, flexWrap: 'wrap', marginBottom: 32 }}>
        {activeCoop.map((c) => {
          const d = coopYearlyData[`${latestFullYear}_${c}`] || {};
          return (
            <div key={c} style={{
              background: C.card, borderRadius: 12, padding: '20px 24px',
              borderTop: `3px solid ${COOP_COLORS[c]}`, flex: '1 1 200px', minWidth: 200
            }}>
              <div style={{ fontSize: 16, fontWeight: 700, color: COOP_COLORS[c], marginBottom: 16 }}>
                {COOP_NAMES[c]}
              </div>
              {[
                ['供給高', fmtNum(d.食材供給高)],
                ['利用人数（週平均）', Math.round(d.食材利用人数 || 0).toLocaleString('ja-JP') + '人'],
                ['金額割合', fmtPct(d.食材金額割合)],
                ['一人当利用高', fmtYen(d.食材一人当利用高 || 0)],
                ['実質GP', fmtNum(d.実質GP)],
              ].map(([lbl, val]) => (
                <div key={lbl} style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 8 }}>
                  <span style={{ color: C.sub, fontSize: 12 }}>{lbl}</span>
                  <span style={{ color: C.text, fontSize: 13, fontWeight: 600 }}>{val}</span>
                </div>
              ))}
            </div>
          );
        })}
      </div>

      <SectionTitle icon="📊">生協別 食材供給高 年度推移</SectionTitle>
      <ChartCard>
        <ResponsiveContainer width="100%" height={300}>
          <BarChart data={supplyChartData} margin={{ top: 10, right: 20, left: 10, bottom: 0 }}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
            <XAxis dataKey="name" tick={{ fill: C.sub, fontSize: 12 }} axisLine={false} tickLine={false} />
            <YAxis tickFormatter={yFmt} tick={{ fill: C.sub, fontSize: 11 }} axisLine={false} tickLine={false} />
            <Tooltip content={<CustomTooltip formatter={(v) => fmtNum(v)} />} />
            <Legend wrapperStyle={{ color: C.sub, fontSize: 12 }} />
            {coops.map((c) => (
              <Bar key={c} dataKey={COOP_NAMES[c]} fill={COOP_COLORS[c]} radius={[4, 4, 0, 0]} />
            ))}
          </BarChart>
        </ResponsiveContainer>
      </ChartCard>

      <SectionTitle icon="📈">生協別 浸透率（金額割合）推移</SectionTitle>
      <ChartCard>
        <ResponsiveContainer width="100%" height={280}>
          <LineChart data={rateChartData} margin={{ top: 10, right: 20, left: 10, bottom: 0 }}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
            <XAxis dataKey="name" tick={{ fill: C.sub, fontSize: 12 }} axisLine={false} tickLine={false} />
            <YAxis tickFormatter={(v) => v + '%'} tick={{ fill: C.sub, fontSize: 11 }} axisLine={false} tickLine={false} />
            <Tooltip content={<CustomTooltip formatter={(v) => v != null ? v.toFixed(1) + '%' : '-'} />} />
            <Legend wrapperStyle={{ color: C.sub, fontSize: 12 }} />
            {coops.map((c) => (
              <Line key={c} dataKey={COOP_NAMES[c]} stroke={COOP_COLORS[c]} strokeWidth={3}
                dot={{ r: 5, fill: COOP_COLORS[c] }} connectNulls={false} />
            ))}
          </LineChart>
        </ResponsiveContainer>
      </ChartCard>
    </div>
  );
};

// ─── Tab 4: 週次分析 ──────────────────────────────────────────────────────────
const TabWeekly = ({ weeklyData, yearInfo }) => {
  const { years } = yearInfo;
  const [selYear, setSelYear] = useState('全年度');
  const [selCoop, setSelCoop] = useState(2);

  const filtered = weeklyData.filter((r) => {
    const yearOk = selYear === '全年度' || r.対象年度 === Number(selYear);
    return yearOk && r.生協コード === selCoop;
  });

  // Aggregate by SEQ
  const seqMap = {};
  filtered.forEach((r) => {
    const k = r.SEQ;
    if (!seqMap[k]) {
      seqMap[k] = {
        SEQ: k, label: r.企画号数 || k + '号',
        食材供給高: 0, 実質GP: 0,
        食材利用人数_sum: 0, 食材利用人数_cnt: 0,
        食材金額割合_sum: 0, 食材金額割合_cnt: 0,
        食材受注人数割合_sum: 0, 食材受注人数割合_cnt: 0,
        cnt: 0,
      };
    }
    const d = seqMap[k];
    d.食材供給高 += r.食材供給高;
    d.実質GP += r.実質GP;
    d.食材利用人数_sum += r.食材利用人数; d.食材利用人数_cnt++;
    d.食材金額割合_sum += r.食材金額割合; d.食材金額割合_cnt++;
    d.食材受注人数割合_sum += r.食材受注人数割合; d.食材受注人数割合_cnt++;
    d.cnt++;
  });

  const chartData = Object.values(seqMap)
    .sort((a, b) => a.SEQ - b.SEQ)
    .map((d) => ({
      name: d.label,
      食材供給高: d.食材供給高,
      実質GP: d.実質GP,
      食材利用人数: d.食材利用人数_cnt ? d.食材利用人数_sum / d.食材利用人数_cnt : 0,
      食材金額割合: d.食材金額割合_cnt ? d.食材金額割合_sum / d.食材金額割合_cnt : 0,
      食材受注人数割合: d.食材受注人数割合_cnt ? d.食材受注人数割合_sum / d.食材受注人数割合_cnt : 0,
    }));

  const yFmt = (v) => {
    if (v >= 1e8) return (v / 1e8).toFixed(0) + '億';
    if (v >= 1e6) return (v / 1e6).toFixed(0) + '百万';
    if (v >= 1e4) return (v / 1e4).toFixed(0) + '万';
    return v;
  };

  const BtnStyle = (active) => ({
    padding: '6px 14px', borderRadius: 6, border: 'none', cursor: 'pointer',
    fontSize: 13, fontWeight: active ? 700 : 400,
    background: active ? C.accent1 : C.card,
    color: active ? '#000' : C.sub,
    transition: 'all 0.15s',
  });

  return (
    <div>
      {/* Filters */}
      <div style={{ marginBottom: 16 }}>
        <div style={{ fontSize: 12, color: C.sub, marginBottom: 8 }}>年度フィルター</div>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', marginBottom: 12 }}>
          {['全年度', ...years.map(String)].map((y) => (
            <button key={y} style={BtnStyle(selYear === y)} onClick={() => setSelYear(y)}>
              {y === '全年度' ? '全年度' : y + '年度'}
            </button>
          ))}
        </div>
        <div style={{ fontSize: 12, color: C.sub, marginBottom: 8 }}>生協フィルター</div>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
          {[2, 6, 7].map((c) => (
            <button key={c} style={BtnStyle(selCoop === c)} onClick={() => setSelCoop(c)}>
              {COOP_NAMES[c]}
            </button>
          ))}
        </div>
      </div>

      <SectionTitle icon="📅">食材供給高 週次推移</SectionTitle>
      <ChartCard>
        <ResponsiveContainer width="100%" height={280}>
          <ComposedChart data={chartData} margin={{ top: 10, right: 20, left: 10, bottom: 0 }}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
            <XAxis dataKey="name" tick={{ fill: C.sub, fontSize: 10 }} axisLine={false} tickLine={false}
              interval={Math.max(0, Math.floor(chartData.length / 15) - 1)} />
            <YAxis tickFormatter={yFmt} tick={{ fill: C.sub, fontSize: 11 }} axisLine={false} tickLine={false} />
            <Tooltip content={<CustomTooltip formatter={(v, n) => n === '実質GP' ? fmtNum(v) : fmtNum(v)} />} />
            <Legend wrapperStyle={{ color: C.sub, fontSize: 12 }} />
            <Area dataKey="食材供給高" fill={C.accent1} stroke={C.accent1} fillOpacity={0.15} name="食材供給高" />
            <Line dataKey="実質GP" stroke={C.accent4} strokeWidth={2} dot={false} name="実質GP" />
          </ComposedChart>
        </ResponsiveContainer>
      </ChartCard>

      <div style={{ display: 'flex', gap: 16, marginTop: 16 }}>
        <ChartCard title="食材利用人数 週次推移" style={{ flex: 1 }}>
          <ResponsiveContainer width="100%" height={220}>
            <LineChart data={chartData} margin={{ top: 5, right: 10, left: 0, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
              <XAxis dataKey="name" tick={{ fill: C.sub, fontSize: 10 }} axisLine={false} tickLine={false}
                interval={Math.max(0, Math.floor(chartData.length / 10) - 1)} />
              <YAxis tick={{ fill: C.sub, fontSize: 10 }} axisLine={false} tickLine={false} />
              <Tooltip content={<CustomTooltip formatter={(v) => Math.round(v).toLocaleString('ja-JP') + '人'} />} />
              <Line dataKey="食材利用人数" stroke={C.accent3} strokeWidth={2} dot={false} name="食材利用人数" />
            </LineChart>
          </ResponsiveContainer>
        </ChartCard>

        <ChartCard title="金額割合 &amp; 人数割合 週次推移" style={{ flex: 1 }}>
          <ResponsiveContainer width="100%" height={220}>
            <LineChart data={chartData} margin={{ top: 5, right: 10, left: 0, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
              <XAxis dataKey="name" tick={{ fill: C.sub, fontSize: 10 }} axisLine={false} tickLine={false}
                interval={Math.max(0, Math.floor(chartData.length / 10) - 1)} />
              <YAxis tickFormatter={(v) => v + '%'} tick={{ fill: C.sub, fontSize: 10 }} axisLine={false} tickLine={false} />
              <Tooltip content={<CustomTooltip formatter={(v) => v.toFixed(1) + '%'} />} />
              <Legend wrapperStyle={{ color: C.sub, fontSize: 11 }} />
              <Line dataKey="食材金額割合" stroke={C.accent1} strokeWidth={2} dot={false} name="食材金額割合" />
              <Line dataKey="食材受注人数割合" stroke={C.accent2} strokeWidth={2} dot={false} name="食材受注人数割合" />
            </LineChart>
          </ResponsiveContainer>
        </ChartCard>
      </div>
    </div>
  );
};

// ─── Tab 5: KPI深掘り ─────────────────────────────────────────────────────────
const TabKpi = ({ yearlyData, yearInfo }) => {
  const { years, latestFullYear } = yearInfo;
  const firstYear = years[0];
  const first = yearlyData[firstYear] || {};
  const cur = yearlyData[latestFullYear] || {};

  const chartData = years.map((y) => ({
    name: y + '年度',
    一人当利用高: yearlyData[y]?.食材一人当利用高 || 0,
    一点当単価: yearlyData[y]?.食材一点当単価 || 0,
    一人当点数: yearlyData[y]?.食材一人当点数 || 0,
  }));

  const unitPriceChange = first.食材一点当単価 > 0
    ? ((cur.食材一点当単価 - first.食材一点当単価) / first.食材一点当単価 * 100).toFixed(1) : '-';
  const pointsChange = first.食材一人当点数 > 0
    ? ((cur.食材一人当点数 - first.食材一人当点数) / first.食材一人当点数 * 100).toFixed(1) : '-';
  const unitSpendChange = first.食材一人当利用高 > 0
    ? ((cur.食材一人当利用高 - first.食材一人当利用高) / first.食材一人当利用高 * 100).toFixed(1) : '-';
  const gpFactor = first.実質GP > 0 ? (cur.実質GP / first.実質GP).toFixed(1) : '-';

  const insights = [
    {
      color: C.accent1,
      text: `一点当単価: ${fmtYen(first.食材一点当単価 || 0)}（${firstYear}年度）→ ${fmtYen(cur.食材一点当単価 || 0)}（${latestFullYear}年度）、変化率: ${unitPriceChange > 0 ? '+' : ''}${unitPriceChange}%`
    },
    {
      color: C.accent3,
      text: `一人当点数: ${(first.食材一人当点数 || 0).toFixed(2)}点 → ${(cur.食材一人当点数 || 0).toFixed(2)}点（+${pointsChange}%）。クロスセル効果が継続的に拡大中。`
    },
    {
      color: C.accent2,
      text: `一人当利用高: ${fmtYen(first.食材一人当利用高 || 0)} → ${fmtYen(cur.食材一人当利用高 || 0)}、成長率: ${unitSpendChange > 0 ? '+' : ''}${unitSpendChange}%`
    },
    {
      color: C.accent4,
      text: `実質GP: ${fmtNum(first.実質GP)} → ${fmtNum(cur.実質GP)}（${gpFactor}倍）。収益基盤が着実に拡大。`
    },
  ];

  const tblRows = years.map((y) => {
    const d = yearlyData[y] || {};
    return {
      年度: y,
      供給高: fmtNum(d.食材供給高),
      利用人数: Math.round(d.食材利用人数 || 0).toLocaleString('ja-JP'),
      金額割合: fmtPct(d.食材金額割合),
      人数割合: fmtPct(d.食材受注人数割合),
      一人当利用高: fmtYen(d.食材一人当利用高 || 0),
      一点当単価: fmtYen(d.食材一点当単価 || 0),
      一人当点数: (d.食材一人当点数 || 0).toFixed(2),
      実質GP: fmtNum(d.実質GP),
    };
  });

  return (
    <div>
      {/* Decomposition box */}
      <div style={{
        background: C.card, border: `1px solid ${C.border}`, borderRadius: 12,
        padding: '16px 20px', marginBottom: 8, fontSize: 14, color: C.text
      }}>
        客単価（一人当利用高）={' '}
        <span style={{ color: C.accent1, fontWeight: 700 }}>一点当単価</span>
        {' '}×{' '}
        <span style={{ color: C.accent3, fontWeight: 700 }}>一人当点数</span>
        {' '}→ 両方の改善が利用高アップに直結
      </div>

      <SectionTitle icon="🎯">一人当利用高 &amp; 構成要素の年度推移</SectionTitle>
      <ChartCard>
        <ResponsiveContainer width="100%" height={300}>
          <ComposedChart data={chartData} margin={{ top: 10, right: 40, left: 10, bottom: 0 }}>
            <CartesianGrid strokeDasharray="3 3" stroke={C.border} />
            <XAxis dataKey="name" tick={{ fill: C.sub, fontSize: 12 }} axisLine={false} tickLine={false} />
            <YAxis yAxisId="left" tickFormatter={(v) => '¥' + v.toLocaleString()} tick={{ fill: C.sub, fontSize: 11 }} axisLine={false} tickLine={false} />
            <YAxis yAxisId="right" orientation="right" tickFormatter={(v) => v.toFixed(1)} tick={{ fill: C.sub, fontSize: 11 }} axisLine={false} tickLine={false} />
            <Tooltip content={<CustomTooltip formatter={(v, n) => n === '一人当点数' ? v.toFixed(2) + '点' : fmtYen(v)} />} />
            <Legend wrapperStyle={{ color: C.sub, fontSize: 12 }} />
            <Bar yAxisId="left" dataKey="一人当利用高" fill={C.accent1} radius={[4, 4, 0, 0]} name="一人当利用高" />
            <Line yAxisId="left" dataKey="一点当単価" stroke={C.accent3} strokeWidth={3} dot={{ r: 5, fill: C.accent3 }} name="一点当単価" />
            <Line yAxisId="right" dataKey="一人当点数" stroke={C.accent2} strokeWidth={3} dot={{ r: 5, fill: C.accent2 }} name="一人当点数" />
          </ComposedChart>
        </ResponsiveContainer>
      </ChartCard>

      <SectionTitle icon="📋">年度別KPI一覧</SectionTitle>
      <div style={{ background: C.card, borderRadius: 12, border: `1px solid ${C.border}`, overflow: 'auto' }}>
        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
          <thead>
            <tr style={{ borderBottom: `2px solid ${C.accent1}` }}>
              {['年度', '供給高', '利用人数', '金額割合', '人数割合', '一人当利用高', '一点当単価', '一人当点数', '実質GP'].map((h) => (
                <th key={h} style={{ padding: '12px 14px', color: C.sub, fontWeight: 600, textAlign: 'right', whiteSpace: 'nowrap' }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {tblRows.map((r, i) => (
              <tr key={r.年度} style={{ background: i % 2 === 0 ? 'rgba(255,255,255,0.01)' : 'rgba(255,255,255,0.03)' }}>
                <td style={{ padding: '10px 14px', color: C.accent1, fontWeight: 700 }}>{r.年度}年度</td>
                {['供給高', '利用人数', '金額割合', '人数割合', '一人当利用高', '一点当単価', '一人当点数', '実質GP'].map((k) => (
                  <td key={k} style={{ padding: '10px 14px', color: C.text, textAlign: 'right' }}>{r[k]}</td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <SectionTitle icon="💡">分析インサイト</SectionTitle>
      <div style={{
        background: 'linear-gradient(135deg, rgba(78,205,196,0.08), rgba(108,92,231,0.08))',
        border: `1px solid ${C.border}`, borderRadius: 12, padding: '20px 24px'
      }}>
        {insights.map((ins, i) => (
          <div key={i} style={{ display: 'flex', gap: 12, alignItems: 'flex-start', marginBottom: i < insights.length - 1 ? 12 : 0 }}>
            <span style={{ color: ins.color, fontSize: 18, lineHeight: 1.4 }}>●</span>
            <span style={{ color: C.sub, fontSize: 13, lineHeight: 1.6 }}>{ins.text}</span>
          </div>
        ))}
      </div>
    </div>
  );
};

// ─── Main Dashboard ────────────────────────────────────────────────────────────
const Dashboard = ({ data, fileName, loadedAt, onReload }) => {
  const [activeTab, setActiveTab] = useState(0);
  const { yearlyData, coopYearlyData, weeklyData } = data;
  const yearInfo = getLatestFullYear(yearlyData);

  const tabs = [
    { icon: '📊', label: '概況' },
    { icon: '📈', label: '成長推移' },
    { icon: '🏢', label: '生協比較' },
    { icon: '📅', label: '週次分析' },
    { icon: '🎯', label: 'KPI深掘り' },
  ];

  if (!yearInfo) return <div style={{ color: C.text, padding: 40 }}>データが不足しています</div>;

  return (
    <div style={{ minHeight: '100vh', background: C.bg, color: C.text, fontFamily: "'Noto Sans JP', 'Hiragino Sans', sans-serif" }}>
      {/* Header */}
      <div style={{
        background: C.card, borderBottom: `1px solid ${C.border}`,
        padding: '16px 24px', display: 'flex', alignItems: 'center', justifyContent: 'space-between',
        position: 'sticky', top: 0, zIndex: 100
      }}>
        <div>
          <h1 style={{ fontSize: 20, fontWeight: 700, margin: 0, marginBottom: 4 }}>
            <span style={{ color: C.accent1 }}>食材セット</span> マーケティングダッシュボード
          </h1>
          <div style={{ fontSize: 11, color: C.sub }}>
            最終更新: {loadedAt} ｜ {fileName}
          </div>
        </div>
        <label style={{
          padding: '8px 16px', background: C.accent1, color: '#000',
          borderRadius: 8, fontWeight: 600, fontSize: 13, cursor: 'pointer',
          border: 'none', whiteSpace: 'nowrap'
        }}>
          📂 ファイル更新
          <input type="file" accept=".xlsx,.xls" style={{ display: 'none' }}
            onChange={(e) => { if (e.target.files[0]) onReload(e.target.files[0]); }} />
        </label>
      </div>

      {/* Tabs */}
      <div style={{
        background: C.card, borderBottom: `1px solid ${C.border}`,
        display: 'flex', padding: '0 24px', gap: 4
      }}>
        {tabs.map((t, i) => (
          <button key={i} onClick={() => setActiveTab(i)} style={{
            padding: '12px 18px', background: 'none', border: 'none', cursor: 'pointer',
            fontSize: 14, fontWeight: activeTab === i ? 700 : 400,
            color: activeTab === i ? C.accent1 : C.sub,
            borderBottom: activeTab === i ? `2px solid ${C.accent1}` : '2px solid transparent',
            transition: 'all 0.15s', whiteSpace: 'nowrap'
          }}>
            {t.icon} {t.label}
          </button>
        ))}
      </div>

      {/* Content */}
      <div style={{ padding: '24px', maxWidth: 1200, margin: '0 auto' }}>
        {activeTab === 0 && <TabOverview yearlyData={yearlyData} yearInfo={yearInfo} />}
        {activeTab === 1 && <TabGrowth yearlyData={yearlyData} yearInfo={yearInfo} />}
        {activeTab === 2 && <TabCoop coopYearlyData={coopYearlyData} yearInfo={yearInfo} />}
        {activeTab === 3 && <TabWeekly weeklyData={weeklyData} yearInfo={yearInfo} />}
        {activeTab === 4 && <TabKpi yearlyData={yearlyData} yearInfo={yearInfo} />}
      </div>
    </div>
  );
};

// ─── App Root ─────────────────────────────────────────────────────────────────
export default function App() {
  const [appData, setAppData] = useState(null);
  const [fileName, setFileName] = useState('');
  const [loadedAt, setLoadedAt] = useState('');

  const handleLoad = useCallback((data, name) => {
    setAppData(data);
    setFileName(name);
    setLoadedAt(new Date().toLocaleString('ja-JP'));
  }, []);

  const handleReload = useCallback((file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws);
        const data = processData(rows);
        setAppData(data);
        setFileName(file.name);
        setLoadedAt(new Date().toLocaleString('ja-JP'));
      } catch (err) {
        alert('ファイルの読み込みに失敗しました: ' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  }, []);

  if (!appData) return <UploadScreen onLoad={handleLoad} />;

  return (
    <Dashboard
      data={appData}
      fileName={fileName}
      loadedAt={loadedAt}
      onReload={handleReload}
    />
  );
}
