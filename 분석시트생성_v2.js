// 분석시트생성 v2 — LINE charts, normalized price index=100
// Paste this entire script into Apps Script editor and run 분석시트생성

function 분석시트생성() {
  const SS_ID = '1l4V_BC_-paP5MLiOHTWmpeZ7S3MD5KWBjn8-8O0WwPw';
  const ss = SpreadsheetApp.openById(SS_ID);

  // ── 1. Delete old 분석 sheet if exists ──────────────────────────────────────
  const OLD = ss.getSheetByName('📊 분석');
  if (OLD) ss.deleteSheet(OLD);
  const sh = ss.insertSheet('📊 분석');

  // ── 2. Read trade data ──────────────────────────────────────────────────────
  const src = ss.getSheets()[0];
  const raw = src.getDataRange().getValues();

  // Row 0 = 원금, Row 1 = blank, Row 2 = headers, Row 3+ = data
  const HEADERS = raw[2]; // 0-indexed
  const COL = {};
  HEADERS.forEach((h, i) => { COL[String(h).trim()] = i; });

  const trades = [];
  for (let r = 3; r < raw.length; r++) {
    const row = raw[r];
    const dateRaw = row[COL['변동일']];
    if (!dateRaw) continue;
    let d;
    if (dateRaw instanceof Date) {
      d = dateRaw;
    } else {
      const s = String(dateRaw).replace(/\./g, '-').replace(/년\s*/,'-').replace(/월\s*/,'-').replace(/일\s*/,'').trim();
      d = new Date(s);
    }
    if (isNaN(d)) continue;

    const name   = String(row[COL['종목명']] || '').trim();
    const ticker = String(row[COL['티커']]   || '').trim();
    const 구분   = String(row[COL['구분']]   || '').trim();
    const 목표변동Raw = row[COL['목표변동(%p)']];
    const 목표변동 = parseFloat(String(목표변동Raw).replace('%','').replace('+','')) || 0;

    if (!name || !구분) continue;
    trades.push({ d, name, ticker, 구분, 목표변동 });
  }

  if (trades.length === 0) {
    sh.getRange(1,1).setValue('매매 데이터가 없습니다.');
    return;
  }

  // ── 3. Reconstruct position snapshots ──────────────────────────────────────
  const stockSet = new Set();
  trades.forEach(t => stockSet.add(t.name));
  const stocks = Array.from(stockSet);

  // Sort trades by date
  trades.sort((a, b) => a.d - b.d);

  // Track cumulative weight per stock
  const weight = {};
  stocks.forEach(s => weight[s] = 0);

  // Build snapshots: map dateStr → { stock → weight }
  const snapMap = new Map();
  for (const t of trades) {
    if (t.구분 === '매수') {
      weight[t.name] = (weight[t.name] || 0) + t.목표변동;
    } else if (t.구분 === '매도') {
      weight[t.name] = Math.max(0, (weight[t.name] || 0) - Math.abs(t.목표변동));
    } else if (t.구분 === '조정') {
      weight[t.name] = Math.max(0, (weight[t.name] || 0) + t.목표변동);
    }
    const key = Utilities.formatDate(t.d, 'Asia/Seoul', 'yyyy-MM-dd');
    snapMap.set(key, Object.assign({}, weight));
  }

  // Get sorted unique dates
  const dates = Array.from(snapMap.keys()).sort();
  const M = dates.length;

  // Key stocks to chart (must have GOOGLEFINANCE-compatible tickers)
  const CHART_STOCKS = [
    { name: '삼성전자',   ticker: 'KRX:005930' },
    { name: 'SK하이닉스', ticker: 'KRX:000660' },
    { name: '삼성전기',   ticker: 'KRX:009150' },
    { name: 'KB금융',     ticker: 'KRX:105560' },
    { name: 'LS ELECTRIC', ticker: 'KRX:010120' },
    { name: 'KODEX 200', ticker: 'KRX:069500' },
    { name: 'AVGO',  ticker: 'NASDAQ:AVGO' },
    { name: 'INTC',  ticker: 'NASDAQ:INTC' },
    { name: 'LRCX',  ticker: 'NASDAQ:LRCX' },
    { name: 'CAT',   ticker: 'NYSE:CAT' },
    { name: 'IWM',   ticker: 'NYSEARCA:IWM' },
    { name: 'DDOG',  ticker: 'NASDAQ:DDOG' },
  ];

  // Only chart stocks that exist in our trade data
  const validChartStocks = CHART_STOCKS.filter(cs => stockSet.has(cs.name));

  // ── 4. Write weight time-series table ─────────────────────────────────────
  // Section A: weight table starting at row 1, col 1
  // Layout: col1=날짜, col2..N = each chart stock weight

  const wHeaders = ['날짜', ...validChartStocks.map(s => s.name + ' 비중(%)')];
  sh.getRange(1, 1, 1, wHeaders.length).setValues([wHeaders]);
  sh.getRange(1, 1, 1, wHeaders.length)
    .setBackground('#1a1a2e').setFontColor('#ffffff').setFontWeight('bold');

  const wRows = [];
  for (const dateStr of dates) {
    const snap = snapMap.get(dateStr);
    const row = [dateStr];
    for (const cs of validChartStocks) {
      row.push(snap[cs.name] || 0);
    }
    wRows.push(row);
  }
  if (wRows.length > 0) {
    sh.getRange(2, 1, wRows.length, wHeaders.length).setValues(wRows);
  }

  // ── 5. Write price table (GOOGLEFINANCE) ──────────────────────────────────
  // Section B: price table starting at col = wHeaders.length + 2
  const priceStartCol = wHeaders.length + 2;
  const pHeaders = ['날짜', ...validChartStocks.map(s => s.name + ' 가격')];
  sh.getRange(1, priceStartCol, 1, pHeaders.length).setValues([pHeaders]);
  sh.getRange(1, priceStartCol, 1, pHeaders.length)
    .setBackground('#0f3460').setFontColor('#ffffff').setFontWeight('bold');

  // Write GOOGLEFINANCE formulas for each stock
  const startDate = dates[0];
  const endDate   = dates[dates.length - 1];

  // Date value column (col priceStartCol) — same date list
  for (let i = 0; i < dates.length; i++) {
    sh.getRange(i + 2, priceStartCol).setValue(dates[i]);
  }

  // Price formulas — use IFERROR(GOOGLEFINANCE(...,"close",date,date),""
  for (let si = 0; si < validChartStocks.length; si++) {
    const cs = validChartStocks[si];
    const pCol = priceStartCol + 1 + si;
    for (let i = 0; i < dates.length; i++) {
      const dateCell = sh.getRange(i + 2, priceStartCol).getA1Notation();
      const formula = `=IFERROR(INDEX(GOOGLEFINANCE("${cs.ticker}","close",DATE(LEFT(${dateCell},4)*1,MID(${dateCell},6,2)*1,RIGHT(${dateCell},2)*1),DATE(LEFT(${dateCell},4)*1,MID(${dateCell},6,2)*1,RIGHT(${dateCell},2)*1)),2,2),"")`;
      sh.getRange(i + 2, pCol).setFormula(formula);
    }
  }

  SpreadsheetApp.flush();
  Utilities.sleep(3000); // wait for formulas to partially resolve

  // ── 6. Create charts ───────────────────────────────────────────────────────
  // Chart layout: 2 per row
  // Each chart: 650 wide × 320 tall
  // Anchor positions (in terms of sheet rows/cols):
  //   Data ends at row M+1 (1-indexed), so charts start at row M+3
  //   Use pixels offset via setPosition(row, col, offsetX, offsetY)

  const CHART_ROW_START = M + 3; // anchor row for first chart row
  const CHART_COLS = [1, 10]; // two charts per row, cols 1 and 10

  for (let ci = 0; ci < validChartStocks.length; ci++) {
    const cs = validChartStocks[ci];
    const chartRow = CHART_ROW_START + Math.floor(ci / 2) * 22;
    const chartCol = CHART_COLS[ci % 2];

    const wCol = ci + 2; // weight col index (1-based: col 2 = first stock weight)
    const pCol = priceStartCol + 1 + ci; // price col index

    try {
      // We create TWO separate line charts overlaid by using combo approach
      // But since COMBO is unreliable, we use a LINE chart with dual Y axes
      // using the EmbeddedChartBuilder's setOption API

      const chartBuilder = sh.newChart()
        .setChartType(Charts.ChartType.LINE)
        .addRange(sh.getRange(1, 1, M + 1, 1))       // dates (col A, rows 1..M+1)
        .addRange(sh.getRange(1, wCol, M + 1, 1))     // weight
        .addRange(sh.getRange(1, pCol, M + 1, 1))     // price
        .setOption('title', cs.name + ' — 비중(%) vs 주가')
        .setOption('legend', { position: 'bottom' })
        .setOption('series', {
          0: { targetAxisIndex: 0, color: '#4285f4', lineWidth: 3 },
          1: { targetAxisIndex: 1, color: '#ea4335', lineWidth: 2, lineDashStyle: 'SOLID' }
        })
        .setOption('vAxes', {
          0: { title: '비중(%)', minValue: 0, textStyle: { color: '#4285f4' } },
          1: { title: '가격',   textStyle: { color: '#ea4335' } }
        })
        .setOption('hAxis', { title: '날짜', slantedText: true, slantedTextAngle: 45 })
        .setOption('backgroundColor', '#f8f9fa')
        .setOption('chartArea', { left: 60, right: 60, top: 40, bottom: 60, width: '80%', height: '65%' })
        .setOption('width', 650)
        .setOption('height', 320)
        .setPosition(chartRow, chartCol, 0, 0);

      sh.insertChart(chartBuilder.build());
      Utilities.sleep(400);
    } catch (e) {
      sh.getRange(chartRow, chartCol).setValue('차트 오류: ' + cs.name + ' — ' + e.message);
    }
  }

  // ── 7. Format the sheet ────────────────────────────────────────────────────
  sh.setColumnWidth(1, 100);
  for (let i = 2; i <= wHeaders.length; i++) sh.setColumnWidth(i, 90);
  sh.setColumnWidth(priceStartCol, 100);
  for (let i = priceStartCol + 1; i <= priceStartCol + validChartStocks.length; i++) {
    sh.setColumnWidth(i, 90);
  }

  // Freeze header row
  sh.setFrozenRows(1);

  // Alternate row colors for weight table
  for (let r = 0; r < wRows.length; r++) {
    const bg = r % 2 === 0 ? '#ffffff' : '#f0f4ff';
    sh.getRange(r + 2, 1, 1, wHeaders.length).setBackground(bg);
  }

  SpreadsheetApp.flush();

  const ui = SpreadsheetApp.getUi();
  ui.alert('✅ 분석 시트 생성 완료!\n\n'
    + '종목 수: ' + validChartStocks.length + '\n'
    + '날짜 수: ' + M + '\n'
    + '차트 수: ' + validChartStocks.length + '\n\n'
    + '가격 데이터는 GOOGLEFINANCE로 자동 조회됩니다.\n'
    + '(시장 휴일 데이터는 빈칸으로 표시될 수 있습니다)');
}
