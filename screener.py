#!/usr/bin/env python3
"""日本株 候補銘柄スクリーナー v1

流れ:
  1. JPXの上場銘柄一覧を取得（内国株式のみ）
  2. 直近20日の売買代金で流動性の低い銘柄を除外
  3. 上場5年以内の銘柄だけ残す
  4. 財務（増収増益・長期借入金・負債比率）を取得してスコア化
  5. docs/index.html と docs/results.csv に出力
"""
import datetime as dt
import html
import io
import os
import time
from urllib.parse import quote

import pandas as pd
import requests
import yfinance as yf

# ---------------- 設定（ここを変えれば条件が変わる） ----------------
MAX_LISTING_YEARS = 5      # 上場からこの年数以内だけ対象（足切り）
MCAP_MAX = 300e8           # 時価総額の上限（円）
DEBT_RATIO_MAX = 0.8       # 負債比率の上限（負債合計 ÷ 純資産合計）
LIQ_PREFILTER = 2e7        # 売買代金(20日平均)がこれ未満は最初に除外（円）
LIQ_GOOD = 1e8             # 売買代金がこれ以上なら流動性 +2点
LIQ_OK = 5e7               # これ以上なら +1点
MAX_SCORE = 13

JPX_URL = ("https://www.jpx.co.jp/markets/statistics-equities/misc/"
           "tvdivq0000001vg2-att/data_j.xlsx")
OUT_DIR = "docs"
JST = dt.timezone(dt.timedelta(hours=9))


# ---------------- 1. 銘柄一覧 ----------------
def load_universe():
    r = requests.get(JPX_URL, timeout=60, headers={"User-Agent": "Mozilla/5.0"})
    r.raise_for_status()
    df = pd.read_excel(io.BytesIO(r.content))
    df = df[df["市場・商品区分"].astype(str).str.contains("内国株式", na=False)].copy()
    df["code"] = df["コード"].astype(str).str.strip()
    df["name"] = df["銘柄名"].astype(str).str.strip()
    df["ticker"] = df["code"] + ".T"
    return df[["code", "name", "ticker"]].reset_index(drop=True)


# ---------------- 2. 流動性（売買代金・出来高の20日平均） ----------------
def liquidity(tickers, chunk_size=200):
    out = {}
    for i in range(0, len(tickers), chunk_size):
        chunk = tickers[i:i + chunk_size]
        try:
            data = yf.download(chunk, period="2mo", interval="1d",
                               group_by="ticker", auto_adjust=False,
                               threads=True, progress=False)
        except Exception as e:
            print("download error:", e)
            continue
        for t in chunk:
            try:
                d = data[t] if isinstance(data.columns, pd.MultiIndex) else data
                d = d.dropna(subset=["Close"]).tail(20)
            except Exception:
                continue
            if len(d) < 10:
                continue
            out[t] = {
                "price": float(d["Close"].iloc[-1]),
                "vol": float(d["Volume"].mean()),
                "value": float((d["Close"] * d["Volume"]).mean()),
            }
        print(f"liquidity {min(i + chunk_size, len(tickers))}/{len(tickers)}")
        time.sleep(1)
    return out


# ---------------- 3. 銘柄情報（上場日・時価総額・事業内容） ----------------
def get_info(t):
    for _ in range(3):
        try:
            info = yf.Ticker(t).info
            if info:
                return info
        except Exception:
            pass
        time.sleep(2)
    return {}


def first_trade(info):
    ft = info.get("firstTradeDateEpochUtc")
    if ft is None and info.get("firstTradeDateMilliseconds"):
        ft = info["firstTradeDateMilliseconds"] / 1000
    if ft is None:
        return None
    return dt.datetime.fromtimestamp(ft, tz=dt.timezone.utc)


def business_line(info):
    parts = [p for p in (info.get("sector"), info.get("industry")) if p]
    head = " / ".join(parts)
    summ = (info.get("longBusinessSummary") or "").strip()
    first = summ.split(". ")[0][:90] if summ else ""
    return (head + "｜" + first) if head and first else (head or first or "不明")


# ---------------- 4. 財務 ----------------
def pick(df, names):
    if df is None or getattr(df, "empty", True):
        return None
    for n in names:
        if n in df.index:
            return df.loc[n]
    return None


def latest(s):
    if s is None:
        return None
    s = s.dropna()
    if s.empty:
        return None
    return float(s.sort_index().iloc[-1])


def consecutive_increase(s):
    """直近から数えて何期連続で前期を上回ったか"""
    if s is None:
        return None
    vals = list(s.dropna().sort_index().values)
    if len(vals) < 2:
        return None
    n = 0
    for i in range(len(vals) - 1, 0, -1):
        if vals[i] > vals[i - 1]:
            n += 1
        else:
            break
    return n


def get_fin(t):
    tk = yf.Ticker(t)
    try:
        inc = tk.financials
    except Exception:
        inc = None
    try:
        bs = tk.balance_sheet
    except Exception:
        bs = None

    rev_n = consecutive_increase(pick(inc, ["Total Revenue", "Operating Revenue"]))
    prof = pick(inc, ["Operating Income"])
    if prof is None:
        prof = pick(inc, ["Net Income"])
    prof_n = consecutive_increase(prof)
    growth = None if (rev_n is None or prof_n is None) else min(rev_n, prof_n)

    ltd = latest(pick(bs, ["Long Term Debt"]))
    if ltd is None:
        td = latest(pick(bs, ["Total Debt"]))
        ltd = 0.0 if td == 0 else None

    liab = latest(pick(bs, ["Total Liabilities Net Minority Interest",
                            "Total Liabilities"]))
    eq = latest(pick(bs, ["Stockholders Equity", "Common Stock Equity",
                          "Total Equity Gross Minority Interest"]))
    ratio = (liab / eq) if (liab is not None and eq and eq > 0) else None
    return {"growth": growth, "ltd": ltd, "debt_ratio": ratio}


# ---------------- 5. スコア ----------------
def band(years):
    if years < 0.5:
        return "半年未満"
    if years < 2:
        return "半年〜2年"
    return "2〜5年"


def score(r):
    s = 0
    g = r.get("growth")
    if g is not None:
        s += 3 if g >= 3 else (1 if g == 2 else 0)
    if r.get("ltd") == 0:
        s += 2
    dr = r.get("debt_ratio")
    if dr is not None and dr <= DEBT_RATIO_MAX:
        s += 2
    if r.get("mcap") is not None and r["mcap"] <= MCAP_MAX:
        s += 2
    if r["value"] >= LIQ_GOOD:
        s += 2
    elif r["value"] >= LIQ_OK:
        s += 1
    s += 2 if r["years"] < 2 else 1
    return s


# ---------------- 出力 ----------------
def yen(x):
    if x is None or pd.isna(x):
        return "不明"
    if x == 0:
        return "0"
    return f"{x / 1e8:,.1f}億円"


def build_html(df, updated):
    e = html.escape
    rows = []
    for _, r in df.iterrows():
        code = r["code"]
        ir = "https://www.google.com/search?q=" + quote(f"{r['name']} IR情報")
        kabutan = f"https://kabutan.jp/stock/?code={code}"
        yahoo = f"https://finance.yahoo.co.jp/quote/{code}.T"
        g = "不明" if pd.isna(r["growth"]) else f"{int(r['growth'])}期"
        dr = "不明" if pd.isna(r["debt_ratio"]) else f"{r['debt_ratio']:.2f}"
        rows.append(
            f'<tr data-band="{e(r["band"])}">'
            f'<td data-v="{r["score"]}">{r["score"]}/{MAX_SCORE}</td>'
            f'<td>{e(r["name"])}</td>'
            f'<td class="biz">{e(r["business"])}</td>'
            f'<td data-v="{r["years"]:.2f}">{r["years"]:.1f}年</td>'
            f'<td>{e(r["band"])}</td>'
            f'<td data-v="{r["mcap"] if pd.notna(r["mcap"]) else 0}">{yen(r["mcap"])}</td>'
            f'<td data-v="{-1 if pd.isna(r["growth"]) else r["growth"]}">{g}</td>'
            f'<td>{yen(r["ltd"])}</td>'
            f'<td data-v="{r["debt_ratio"] if pd.notna(r["debt_ratio"]) else 99}">{dr}</td>'
            f'<td data-v="{r["value"]:.0f}">{r["vol"] / 1e4:,.1f}万株<br>{yen(r["value"])}</td>'
            f'<td>{e(code)}</td>'
            f'<td><a href="{ir}" target="_blank">IR</a></td>'
            f'<td><a href="{kabutan}" target="_blank">株探</a></td>'
            f'<td data-v="{r["price"]}"><a href="{yahoo}" target="_blank">{r["price"]:,.0f}円</a></td>'
            f'</tr>'
        )
    heads = ["スコア", "銘柄名", "事業内容", "上場年数", "上場区分", "時価総額", "増収増益",
             "長期借入金", "負債比率", "出来高・売買代金(20日平均)", "コード", "IR情報",
             "株探", "今日の株価"]
    th = "".join(f'<th onclick="sortBy({i})">{h}</th>' for i, h in enumerate(heads))
    return f"""<!doctype html>
<html lang="ja"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>候補銘柄スクリーナー</title>
<style>
body{{font-family:-apple-system,sans-serif;margin:8px;font-size:13px}}
.bar button{{margin:2px;padding:6px 10px;font-size:13px}}
.wrap{{overflow-x:auto}}
table{{border-collapse:collapse;white-space:nowrap}}
th,td{{border:1px solid #ccc;padding:4px 6px;text-align:left}}
th{{background:#f2f2f2;position:sticky;top:0;cursor:pointer}}
td.biz{{white-space:normal;min-width:220px;max-width:320px}}
</style></head><body>
<h3>候補銘柄スクリーナー（日本株）</h3>
<p>更新: {updated}　/ 表示 {len(df)} 銘柄　/ 見出しタップで並べ替え</p>
<div class="bar">
<button onclick="flt('')">全て</button>
<button onclick="flt('2y')">2年以内</button>
<button onclick="flt('半年〜2年')">半年〜2年</button>
<button onclick="flt('2〜5年')">2〜5年</button>
</div>
<div class="wrap"><table id="t"><thead><tr>{th}</tr></thead>
<tbody>{''.join(rows)}</tbody></table></div>
<p>※ 約定回数は取得できないため、売買代金で代用しています。最終確認は証券会社アプリで。</p>
<script>
function flt(b){{
  document.querySelectorAll('#t tbody tr').forEach(r=>{{
    var x=r.dataset.band, show=true;
    if(b==='2y') show=(x==='半年未満'||x==='半年〜2年');
    else if(b) show=(x===b);
    r.style.display=show?'':'none';
  }});
}}
var dir={{}};
function sortBy(i){{
  var tb=document.querySelector('#t tbody'), rs=Array.from(tb.rows);
  dir[i]=!dir[i];
  rs.sort(function(a,b){{
    var x=a.cells[i], y=b.cells[i];
    var xv=x.dataset.v!==undefined?parseFloat(x.dataset.v):x.textContent;
    var yv=y.dataset.v!==undefined?parseFloat(y.dataset.v):y.textContent;
    var c=(xv>yv)-(xv<yv);
    return dir[i]?c:-c;
  }});
  rs.forEach(r=>tb.appendChild(r));
}}
</script></body></html>"""


# ---------------- main ----------------
def main():
    uni = load_universe()
    print("universe:", len(uni))
    liq = liquidity(list(uni["ticker"]))
    ok = {t for t, v in liq.items() if v["value"] >= LIQ_PREFILTER}
    cands = uni[uni["ticker"].isin(ok)]
    print("after liquidity prefilter:", len(cands))

    now = dt.datetime.now(dt.timezone.utc)
    rows = []
    for n, (_, r) in enumerate(cands.iterrows(), 1):
        if n % 100 == 0:
            print(f"info {n}/{len(cands)}")
        info = get_info(r["ticker"])
        ft = first_trade(info)
        if ft is None:
            continue
        years = (now - ft).days / 365.25
        if years > MAX_LISTING_YEARS:
            continue
        fin = get_fin(r["ticker"])
        rec = {
            "code": r["code"], "name": r["name"],
            "business": business_line(info),
            "years": years, "band": band(years),
            "mcap": info.get("marketCap"),
            **fin, **liq[r["ticker"]],
        }
        rec["score"] = score(rec)
        rows.append(rec)
        time.sleep(0.3)

    df = pd.DataFrame(rows)
    os.makedirs(OUT_DIR, exist_ok=True)
    updated = dt.datetime.now(JST).strftime("%Y-%m-%d %H:%M JST")
    if df.empty:
        df = pd.DataFrame(columns=["code", "name", "business", "years", "band", "mcap",
                                   "growth", "ltd", "debt_ratio", "price", "vol",
                                   "value", "score"])
    else:
        df = df.sort_values(["score", "value"], ascending=[False, False])
    df.to_csv(os.path.join(OUT_DIR, "results.csv"), index=False, encoding="utf-8-sig")
    with open(os.path.join(OUT_DIR, "index.html"), "w", encoding="utf-8") as f:
        f.write(build_html(df, updated))
    print("done:", len(df))


if __name__ == "__main__":
    main()
