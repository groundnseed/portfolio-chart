import yfinance as yf
import json, sys
from datetime import datetime, timezone

STOCKS = {
    '삼성전기':  {'ticker':'009150.KS','market':'kr','flag':'🇰🇷','weights':[['3/10',2],['3/17',4],['3/27',6],['4/14',6],['4/27',6],['6/16',6],['6/19',12]]},
    'SNDK':      {'ticker':'SNDK',     'market':'us','flag':'🇺🇸','weights':[['3/10',1],['3/17',2],['3/27',3],['4/27',3],['6/16',3]]},
    'INTC':      {'ticker':'INTC',     'market':'us','flag':'🇺🇸','weights':[['3/10',1],['3/17',2],['4/14',4],['4/27',5],['6/16',5]]},
    'SK하이닉스': {'ticker':'000660.KS','market':'kr','flag':'🇰🇷','weights':[['3/10',2],['3/17',4],['4/14',4],['4/27',4],['6/16',5.5],['6/17',6.5],['6/19',9]]},
    'LRCX':      {'ticker':'LRCX',     'market':'us','flag':'🇺🇸','weights':[['3/10',2],['3/17',4],['4/8',6],['4/27',6],['6/16',6]]},
    '삼성전자':   {'ticker':'005930.KS','market':'kr','flag':'🇰🇷','weights':[['3/10',2],['3/24',4],['4/27',6],['6/16',6],['6/19',8]]},
    'AMAT':      {'ticker':'AMAT',     'market':'us','flag':'🇺🇸','weights':[['4/14',2],['4/27',2],['6/16',2],['6/19',3]]},
    'ARM':       {'ticker':'ARM',      'market':'us','flag':'🇺🇸','weights':[['3/17',2],['4/27',3],['6/16',3]]},
    'IWM':       {'ticker':'IWM',      'market':'us','flag':'🇺🇸','weights':[['4/8',5],['4/27',5],['6/16',5],['6/19',6]]},
    'KB금융':    {'ticker':'105560.KS','market':'kr','flag':'🇰🇷','weights':[['3/10',2],['4/14',4],['6/16',4]]},
    'MRVL':      {'ticker':'MRVL',     'market':'us','flag':'🇺🇸','weights':[['4/27',3],['6/16',3]]},
    '현대차':    {'ticker':'005380.KS','market':'kr','flag':'🇰🇷','weights':[['3/17',2],['4/27',3],['6/16',3]]},
}

result = {'built': datetime.now(timezone.utc).isoformat(), 'stocks': {}}

for name, meta in STOCKS.items():
    try:
        df = yf.download(meta['ticker'], start='2026-03-09', progress=False, auto_adjust=True)
        if df.empty:
            print(f"SKIP {name}: empty", file=sys.stderr)
            continue
        prices = {}
        for dt, row in df.iterrows():
            c = float(row['Close'])
            if c != c: continue  # NaN check
            prices[dt.strftime('%-m/%-d')] = round(c, 2)
        
        plist = sorted(prices.items(), key=lambda x: tuple(int(v) for v in x[0].split('/')))
        basis = plist[0][1] if plist else 0
        current = plist[-1][1] if plist else 0
        ret = round((current - basis) / basis * 100, 1) if basis else 0

        result['stocks'][name] = {
            **meta,
            'prices': prices,
            'basis': basis,
            'currentPrice': current,
            'returnPct': ret,
        }
        print(f"OK {name}: {len(prices)} days, ret={ret}%", file=sys.stderr)
    except Exception as e:
        print(f"ERR {name}: {e}", file=sys.stderr)

with open('prices.json', 'w', encoding='utf-8') as f:
    json.dump(result, f, ensure_ascii=False)
print("Done:", len(result['stocks']), "stocks")
