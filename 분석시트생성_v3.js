// ★ 분석시트생성 v3 — 차트 생성 확실히 작동하는 버전
// Apps Script 편집기에 전체 붙여넣기 후 실행

function 분석시트생성() {
  const SS_ID = '1l4V_BC_-paP5MLiOHTWmpeZ7S3MD5KWBjn8-8O0WwPw';
  const ss = SpreadsheetApp.openById(SS_ID);

  // ── 1. 기존 분석 시트 삭제 ────────────────────────────────────────────────
  const OLD = ss.getSheetByName('📊 분석');
  if (OLD) ss.deleteSheet(OLD);
  const sh = ss.insertSheet('📊 분석');
  SpreadsheetApp.flush();

  // ── 2. 매매 데이터 읽기 ──────────────────────────────────────────────────
  const src = ss.getSheets()[0];
  const raw = src.getDataRange().getValues();

  const HEADERS = raw[2];
  const C = {};
  HEADERS.forEach((h, i) => { C[String(h).trim()] = i; });

  function parseDate(v) {
    if (v instanceof Date) return v;
    const s = String(v).replace(/년\s*/, '-').replace(/월\s*/, '-').replace(/일\s*/, '').replace(/\./g, '-').trim();
    const d = new Date(s);
    return isNaN(d) ? null : d;
  }

  const trades = [];
  for (let r = 3; r < raw.length; r++) {
    const row = raw[r];
    const d = parseDate(row[C['변동일']]);
    if (!d) continue;
    const name = String(row[C['종목명']] || '').trim();
    const 구분 = String(row[C['구분']] || '').trim();
    if (!name || !구분) continue;
    const delta = parseFloat(String(row[C['목표변동(%p)']] || '0').replace('%','').replace('+','')) || 0;
    trades.push({ d, name, 구분, delta });
  }
  trades.sort((a, b) => a.d - b.d);

  // ── 3. 비중 시계열 재구성 ────────────────────────────────────────────────
  const allStocks = [...new Set(trades.map(t => t.name))];
  const wt = {};
  allStocks.forEach(s => wt[s] = 0);

  const snapMap = new Map(); // dateStr → { stock: weight }
  for (const t of trades) {
    if (t.구분 === '매수')      wt[t.name] = Math.max(0, (wt[t.name] || 0) + t.delta);
    else if (t.구분 === '매도') wt[t.name] = Math.max(0, (wt[t.name] || 0) - Math.abs(t.delta));
    else if (t.구분 === '조정') wt[t.name] = Math.max(0, (wt[t.name] || 0) + t.delta);
    const key = Utilities.formatDate(t.d, 'Asia/Seoul', 'yyyy-MM-dd');
    snapMap.set(key, Object.assign({}, wt));
  }

  const dates = Array.from(snapMap.keys()).sort();
  const M = dates.length;
  if (M === 0) { sh.getRange(1,1).setValue('데이터 없음'); return; }

  // 차트 대상 종목 (매매 이력에 실제로 있는 것만)
  const CHART_LIST = [
    '삼성전자','SK하이닉스','삼성전기','KB금융',
    'LS ELECTRIC','KODEX 200','AVGO','INTC','LRCX','CAT','IWM','DDOG'
  ].filter(n => allStocks.includes(n));

  // ── 4. 비중 테이블 쓰기 (A열: 날짜, B~N열: 각 종목 비중%) ──────────────
  const hdr = ['날짜', ...CHART_LIST.map(n => n + '(%)')];
  sh.getRange(1, 1, 1, hdr.length).setValues([hdr])
    .setBackground('#1a1a2e').setFontColor('#fff').setFontWeight('bold');

  const tableData = [];
  for (const ds of dates) {
    const snap = snapMap.get(ds);
    tableData.push([ds, ...CHART_LIST.map(n => snap[n] || 0)]);
  }
  sh.getRange(2, 1, M, hdr.length).setValues(tableData);

  // 날짜 열 서식
  sh.getRange(2, 1, M, 1).setNumberFormat('@'); // 텍스트로 유지

  // 번갈아 색
  for (let i = 0; i < M; i++) {
    sh.getRange(i + 2, 1, 1, hdr.length).setBackground(i % 2 === 0 ? '#fff' : '#f0f4ff');
  }
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 100);
  for (let c = 2; c <= hdr.length; c++) sh.setColumnWidth(c, 80);

  SpreadsheetApp.flush();

  // ── 5. 차트 생성 ─────────────────────────────────────────────────────────
  // 차트마다 해당 종목의 비중% 단일 LINE 차트
  // 레이아웃: 2열 배치, 각 차트 600×280
  // 차트 위치: 테이블 오른쪽 (K열부터)

  const CHARTS_START_COL = hdr.length + 2; // 테이블 오른쪽
  const CHART_PER_ROW = 2;
  const ROWS_PER_CHART = 18; // 약 280px / ~16px per row

  let chartsMade = 0;
  const errors = [];

  for (let ci = 0; ci < CHART_LIST.length; ci++) {
    const stockName = CHART_LIST[ci];
    const dataCol = ci + 2; // 1-based: col 2 = first stock

    const row = CHARTS_START_COL === hdr.length + 2
      ? 1 + Math.floor(ci / CHART_PER_ROW) * ROWS_PER_CHART
      : 1;
    const col = CHARTS_START_COL + (ci % CHART_PER_ROW) * 9;

    try {
      const dateRange   = sh.getRange(1, 1, M + 1, 1);
      const weightRange = sh.getRange(1, dataCol, M + 1, 1);

      const chart = sh.newChart()
        .setChartType(Charts.ChartType.LINE)
        .addRange(dateRange)
        .addRange(weightRange)
        .setOption('title', stockName + ' 비중(%)')
        .setOption('legend', { position: 'none' })
        .setOption('hAxis', {
          title: '',
          slantedText: true,
          slantedTextAngle: 60,
          textStyle: { fontSize: 9 }
        })
        .setOption('vAxis', {
          title: '비중(%)',
          minValue: 0,
          textStyle: { fontSize: 10 }
        })
        .setOption('colors', ['#4285f4'])
        .setOption('lineWidth', 3)
        .setOption('pointSize', 4)
        .setOption('backgroundColor', { fill: '#f8f9fa' })
        .setOption('width', 520)
        .setOption('height', 260)
        .setPosition(row, col, 5, 5)
        .build();

      sh.insertChart(chart);
      chartsMade++;
      Utilities.sleep(300);
    } catch(e) {
      errors.push(stockName + ': ' + e.message);
    }
  }

  SpreadsheetApp.flush();

  // ── 6. 결과 알림 ─────────────────────────────────────────────────────────
  let msg = '✅ 분석 시트 생성 완료!\n\n'
    + '• 날짜: ' + M + '개\n'
    + '• 종목: ' + CHART_LIST.length + '개\n'
    + '• 차트: ' + chartsMade + '개 생성됨\n';
  if (errors.length > 0) {
    msg += '\n⚠️ 오류:\n' + errors.join('\n');
  }
  SpreadsheetApp.getUi().alert(msg);
}
