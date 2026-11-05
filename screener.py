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
import importlib
import io
import os
import re
import subprocess
import sys
import time
import unicodedata
from urllib.parse import quote

import pandas as pd
import requests
import yfinance as yf



def _import_optional(module, pip_name):
    """ライブラリが無ければ自動でインストールして読み込む。ダメならNone"""
    try:
        return importlib.import_module(module)
    except Exception:
        try:
            subprocess.run([sys.executable, "-m", "pip", "install", "-q", pip_name],
                           check=True, timeout=300)
            importlib.invalidate_caches()
            return importlib.import_module(module)
        except Exception as e:
            print(f"WARN: {pip_name} を使えません: {e}")
            return None


_dt_mod = _import_optional("deep_translator", "deep-translator")
GoogleTranslator = getattr(_dt_mod, "GoogleTranslator", None) if _dt_mod else None
_bs4_mod = _import_optional("bs4", "beautifulsoup4")
BeautifulSoup = getattr(_bs4_mod, "BeautifulSoup", None) if _bs4_mod else None

# ---------------- 設定（ここを変えれば条件が変わる） ----------------
MAX_LISTING_YEARS = 5      # 上場からこの年数以内だけ対象（足切り）
MCAP_MAX = 300e8           # 時価総額の上限（円）
DEBT_RATIO_MAX = 0.8       # 負債比率の上限（負債合計 ÷ 純資産合計）
LIQ_PREFILTER = 2e7        # 売買代金(20日平均)がこれ未満は最初に除外（円）
LIQ_GOOD = 1e8             # 売買代金がこれ以上なら流動性 +2点
LIQ_OK = 5e7               # これ以上なら +1点
MAX_SCORE = 13
USE_YAHOO_PROFILE = True   # Yahoo!ファイナンスの企業情報から日本語の「特色」「連結事業」を取る（使わないならFalse）
SUMMARY_CHARS = 450        # 事業内容に使う英語説明の最大文字数（翻訳前）。増やすほど詳しくなる

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


def has_ja(s):
    return any("぀" <= ch <= "ヿ" or "一" <= ch <= "鿿" for ch in s)


def _gtx(text):
    """予備の翻訳ルート（ライブラリを使わず直接問い合わせる）"""
    r = requests.get("https://translate.googleapis.com/translate_a/single",
                     params={"client": "gtx", "sl": "en", "tl": "ja", "dt": "t", "q": text},
                     timeout=30, headers={"User-Agent": "Mozilla/5.0"})
    r.raise_for_status()
    return "".join(seg[0] for seg in r.json()[0] if seg and seg[0])


def to_ja(text):
    """英語 → 日本語（無料の機械翻訳）。失敗したら原文のまま返す"""
    if not text or has_ja(text):
        return text
    routes = []
    if GoogleTranslator is not None:
        routes.append(lambda t: GoogleTranslator(source="en", target="ja").translate(t))
    routes.append(_gtx)
    for fn in routes:
        for _ in range(2):
            try:
                out = fn(text)
                if out and has_ja(out):
                    return out
            except Exception:
                time.sleep(2)
    return text


EMPTY_PROFILE = {"feature": "", "segments": "", "ceo": ""}


def yahoo_profile(code):
    """Yahoo!ファイナンス企業情報ページから 特色・連結事業・代表者名 を取る。
    ページの作りが変わると取れなくなるので、取れなければ空で返す。"""
    if BeautifulSoup is None:
        return dict(EMPTY_PROFILE)
    url = f"https://finance.yahoo.co.jp/quote/{code}.T/profile"
    try:
        r = requests.get(url, timeout=30, headers={
            "User-Agent": "Mozilla/5.0 (compatible; personal-screener)",
            "Accept-Language": "ja"})
        if r.status_code != 200:
            print("profile", code, "status", r.status_code)
            return dict(EMPTY_PROFILE)
        soup = BeautifulSoup(r.text, "html.parser")
        for t in soup(["script", "style"]):
            t.decompose()
        lines = [t.strip() for t in soup.get_text("\n").split("\n") if t.strip()]
    except Exception as e:
        print("profile error", code, e)
        return dict(EMPTY_PROFILE)

    labels = {"特色", "連結事業", "代表者名", "業種分類", "英文社名", "設立年月日", "市場名",
              "上場年月日", "決算", "本社所在地", "単元株数", "従業員数（単独）", "従業員数（連結）",
              "平均年齢", "平均年収", "最近見た銘柄"}

    def after(label):
        for i, t in enumerate(lines[:-1]):
            if t == label and lines[i + 1] not in labels and len(lines[i + 1]) >= 2:
                return lines[i + 1]
        return ""

    out = {"feature": after("特色"), "segments": after("連結事業"), "ceo": after("代表者名")}
    if not out["feature"] and not out["segments"]:
        print("profile not parsed", code, lines[:5])
    return out


def parse_segments(text):
    """「Ａ事業95、Ｂ事業5、他0(2025.12)」→ ([(名前, %), ...], '2025.12')"""
    t = unicodedata.normalize("NFKC", text or "").strip()
    asof = ""
    m = re.search(r"\((\d{4}[./]\d{1,2})\)\s*$", t)
    if m:
        asof = m.group(1)
        t = t[:m.start()].strip()
    items = []
    for part in re.split(r"[、,]", t):
        part = part.strip()
        if not part:
            continue
        m = re.match(r"^(.*?)\s*(\d+(?:\.\d+)?)$", part)
        if m and m.group(1):
            items.append((m.group(1), float(m.group(2))))
        else:
            items.append((part, None))
    nums = [x for x in items if x[1] is not None]
    if len(nums) == len(items):
        items = [x for x in sorted(items, key=lambda x: -x[1]) if x[1] > 0]
    return items, asof


def trim_sentences(s, limit):
    s = " ".join(s.split())
    if len(s) <= limit:
        return s
    cut = s[:limit]
    i = cut.rfind(". ")
    return cut[:i + 1] if i > 80 else cut


def business_fields(info):
    ind = info.get("industry") or info.get("sector") or ""
    summ = (info.get("longBusinessSummary") or "").strip()
    return {
        "industry": to_ja(ind),
        "business": to_ja(trim_sentences(summ, SUMMARY_CHARS)) if summ else "",
    }


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
        return "0円"
    return f"{x / 1e8:,.1f}億円"


def shares(x):
    if x is None or pd.isna(x):
        return "不明"
    if x >= 1e4:
        return f"{x / 1e4:,.1f}万株"
    return f"{x:,.0f}株"


def txt(x):
    return "" if (x is None or (not isinstance(x, str) and pd.isna(x))) else str(x).strip()


def num_attr(x):
    return "" if (x is None or pd.isna(x)) else f"{x}"


TEMPLATE = """<!doctype html>
<html lang="ja"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>候補銘柄スクリーナー</title>
<style>
:root{--bg:#f5f6f8;--card:#fff;--text:#1c1e21;--sub:#667085;--line:#e4e7ec;--accent:#2563eb;--good:#067647;--chip:#eef2ff}
@media (prefers-color-scheme:dark){:root{--bg:#0f1115;--card:#181b21;--text:#eceef2;--sub:#98a2b3;--line:#2a2f38;--accent:#7aa7ff;--good:#5fd38d;--chip:#1f2a4d}}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--text);font:15px/1.55 -apple-system,"Hiragino Sans",sans-serif}
header.top{padding:12px 12px 6px}
h1{font-size:17px;margin:0 0 2px}
.meta{font-size:12px;color:var(--sub)}
.ctl{position:sticky;top:0;z-index:5;background:var(--bg);padding:6px 12px 8px;border-bottom:1px solid var(--line)}
.chips{display:flex;gap:6px;margin-bottom:6px}
.chips button{flex:1;padding:8px 2px;font-size:13px;border:1px solid var(--line);border-radius:999px;background:var(--card);color:var(--text)}
.chips button.on{background:var(--accent);border-color:var(--accent);color:#fff;font-weight:600}
.sortrow{display:flex;align-items:center;gap:8px;font-size:12px;color:var(--sub)}
select{flex:1;font-size:14px;padding:7px;border:1px solid var(--line);border-radius:8px;background:var(--card);color:var(--text)}
#list{padding:10px 12px 40px;display:grid;gap:10px}
.card{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:12px}
.hd{display:flex;gap:10px;align-items:center}
.score{flex:none;min-width:56px;text-align:center;font-weight:700;font-size:24px;line-height:1;padding:9px 6px;border-radius:12px;background:var(--chip);color:var(--accent)}
.score small{display:block;font-size:11px;font-weight:500;color:var(--sub);margin-top:4px}
.hdr{min-width:0;flex:1}
.tags{display:flex;flex-wrap:wrap;gap:6px;align-items:center;font-size:12px;color:var(--sub)}
.code{font-weight:700;color:var(--text);font-size:15px}
.flag{background:var(--chip);color:var(--accent);padding:1px 8px;border-radius:999px;font-weight:600}
.name{font-size:12px;color:var(--sub);margin-top:2px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.biz{margin:10px 0;font-size:14px}
.biz b{display:inline-block;font-size:12px;color:var(--sub);font-weight:600;margin-right:6px}
.biz2{margin:-4px 0 10px;font-size:13px;color:var(--sub)}
.segs{display:flex;flex-wrap:wrap;gap:6px;align-items:center;margin:0 0 10px;font-size:12px}
.cap{color:var(--sub)}
.seg{border:1px solid var(--line);border-radius:999px;padding:2px 9px}
.seg.top{background:var(--chip);color:var(--accent);border-color:transparent;font-weight:600}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:6px 14px;margin:0 0 10px}
.grid div{border-top:1px solid var(--line);padding-top:4px}
dt{font-size:11px;color:var(--sub)}
dd{margin:0;font-weight:600;font-size:14px}
.ok{color:var(--good)}
.links{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.links a{flex:1;text-align:center;padding:11px 4px;border:1px solid var(--line);border-radius:10px;text-decoration:none;color:var(--accent);font-size:14px;font-weight:600}
.foot{padding:0 12px 30px;font-size:12px;color:var(--sub)}
</style></head><body>
<header class="top">
<h1>候補銘柄スクリーナー（日本株）</h1>
<div class="meta">更新 __UPDATED__ ／ <span id="cnt">__COUNT__</span>銘柄 ／ __MAX__点満点</div>
</header>
<div class="ctl">
<div class="chips">
<button class="on" onclick="setBand(this,'')">全て</button>
<button onclick="setBand(this,'2y')">2年以内</button>
<button onclick="setBand(this,'半年〜2年')">半年〜2年</button>
<button onclick="setBand(this,'2〜5年')">2〜5年</button>
</div>
<div class="sortrow">並び順
<select onchange="setSort(this.value)">
<option value="score">スコアが高い順</option>
<option value="value">売買代金が大きい順</option>
<option value="mcap">時価総額が小さい順</option>
<option value="years">上場が新しい順</option>
<option value="dr">負債比率が低い順</option>
</select></div>
</div>
<div id="list">__CARDS__</div>
<p class="foot">※ 約定回数は取得できないため、売買代金で代用しています。最終確認は証券会社アプリで。「特色」「事業構成」はYahoo!ファイナンス、取れない銘柄の事業内容は英語の説明の機械翻訳です（事業構成の数字は売上の構成比（%）と思われます）。</p>
<script>
var cur={band:'',sort:'score'};
function num(c,k,empty){var v=parseFloat(c.dataset[k]);return isNaN(v)?empty:v;}
function apply(){
  var box=document.getElementById('list');
  var cs=Array.prototype.slice.call(box.children);
  cs.sort(function(a,b){
    var s=cur.sort;
    if(s==='score'){return (num(b,'score',-1)-num(a,'score',-1))||(num(b,'value',0)-num(a,'value',0));}
    if(s==='value'){return num(b,'value',0)-num(a,'value',0);}
    if(s==='mcap'){return num(a,'mcap',1e30)-num(b,'mcap',1e30);}
    if(s==='years'){return num(a,'years',1e30)-num(b,'years',1e30);}
    if(s==='dr'){return num(a,'dr',1e30)-num(b,'dr',1e30);}
    return 0;
  });
  var n=0;
  cs.forEach(function(c){
    var x=c.dataset.band,show=true;
    if(cur.band==='2y')show=(x==='半年未満'||x==='半年〜2年');
    else if(cur.band)show=(x===cur.band);
    c.style.display=show?'':'none';
    if(show)n++;
    box.appendChild(c);
  });
  document.getElementById('cnt').textContent=n;
}
function setBand(btn,b){
  cur.band=b;
  Array.prototype.forEach.call(btn.parentNode.children,function(x){x.classList.remove('on');});
  btn.classList.add('on');apply();
}
function setSort(s){cur.sort=s;apply();}
</script></body></html>"""


def build_html(df, updated):
    e = html.escape
    cards = []
    for _, r in df.iterrows():
        code = str(r["code"])
        ir = "https://www.google.com/search?q=" + quote(f"{r['name']} IR情報")
        kabutan = f"https://kabutan.jp/stock/?code={code}"
        yahoo = f"https://finance.yahoo.co.jp/quote/{code}.T"
        profile = f"https://finance.yahoo.co.jp/quote/{code}.T/profile"

        g_ok = pd.notna(r["growth"]) and r["growth"] >= 3
        g = "不明" if pd.isna(r["growth"]) else f"{int(r['growth'])}期連続"
        ltd_ok = pd.notna(r["ltd"]) and r["ltd"] == 0
        ltd = "なし" if ltd_ok else yen(r["ltd"])
        dr_ok = pd.notna(r["debt_ratio"]) and r["debt_ratio"] <= DEBT_RATIO_MAX
        dr = "不明" if pd.isna(r["debt_ratio"]) else f"{r['debt_ratio']:.2f}"
        mc_ok = pd.notna(r["mcap"]) and r["mcap"] <= MCAP_MAX

        def cls(ok):
            return ' class="ok"' if ok else ""

        industry = txt(r.get("industry"))
        feature = txt(r.get("feature"))
        summary = txt(r.get("business"))
        ceo = txt(r.get("ceo"))
        if feature:
            lead = feature
            sub = summary if has_ja(summary) else ""   # 英語のままの説明は出さない
        else:
            lead = summary or "（事業内容を取得できませんでした）"
            sub = ""
        ind_html = f"<b>{e(industry)}</b>" if industry else ""
        sub_html = f'<p class="biz2">{e(sub)}</p>' if sub else ""
        ceo_html = f"　代表 {e(ceo)}" if ceo else ""
        seg_items, seg_asof = parse_segments(txt(r.get("segments")))
        seg_block = ""
        if seg_items:
            chips = "".join(
                '<span class="seg{}">{}</span>'.format(
                    " top" if i == 0 else "",
                    e(f"{n} {p:g}%") if p is not None else e(n))
                for i, (n, p) in enumerate(seg_items))
            cap = "事業構成" + (f"（{seg_asof}）" if seg_asof else "")
            seg_block = f'<div class="segs"><span class="cap">{e(cap)}</span>{chips}</div>'

        cards.append(
            f'<article class="card" data-band="{e(r["band"])}" data-score="{r["score"]}" '
            f'data-value="{r["value"]:.0f}" data-mcap="{num_attr(r["mcap"])}" '
            f'data-years="{r["years"]:.2f}" data-dr="{num_attr(r["debt_ratio"])}">'
            f'<div class="hd"><div class="score">{r["score"]}<small>/ {MAX_SCORE}</small></div>'
            f'<div class="hdr"><div class="tags"><span class="code">{e(code)}</span>'
            f'<span class="flag">{e(r["band"])}</span><span>上場{r["years"]:.1f}年</span></div>'
            f'<div class="name">{e(str(r["name"]))}{ceo_html}</div></div></div>'
            f'<p class="biz">{ind_html}{e(lead)}</p>{sub_html}{seg_block}'
            f'<dl class="grid">'
            f'<div><dt>時価総額</dt><dd{cls(mc_ok)}>{yen(r["mcap"])}</dd></div>'
            f'<div><dt>増収増益</dt><dd{cls(g_ok)}>{g}</dd></div>'
            f'<div><dt>長期借入金</dt><dd{cls(ltd_ok)}>{ltd}</dd></div>'
            f'<div><dt>負債比率</dt><dd{cls(dr_ok)}>{dr}</dd></div>'
            f'<div><dt>売買代金(20日平均)</dt><dd>{yen(r["value"])}</dd></div>'
            f'<div><dt>出来高(20日平均)</dt><dd>{shares(r["vol"])}</dd></div>'
            f'</dl>'
            f'<div class="links"><a href="{kabutan}" target="_blank">株探</a>'
            f'<a href="{profile}" target="_blank">企業情報</a>'
            f'<a href="{ir}" target="_blank">IR情報</a>'
            f'<a href="{yahoo}" target="_blank">株価 {r["price"]:,.0f}円</a></div>'
            f'</article>'
        )
    return (TEMPLATE.replace("__UPDATED__", updated)
            .replace("__COUNT__", str(len(df)))
            .replace("__MAX__", str(MAX_SCORE))
            .replace("__CARDS__", "".join(cards)))


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
            **business_fields(info),
            **(yahoo_profile(r["code"]) if USE_YAHOO_PROFILE else EMPTY_PROFILE),
            "years": years, "band": band(years),
            "mcap": info.get("marketCap"),
            **fin, **liq[r["ticker"]],
        }
        rec["score"] = score(rec)
        rows.append(rec)
        time.sleep(1.0 if USE_YAHOO_PROFILE else 0.3)

    df = pd.DataFrame(rows)
    os.makedirs(OUT_DIR, exist_ok=True)
    updated = dt.datetime.now(JST).strftime("%Y-%m-%d %H:%M JST")
    if df.empty:
        df = pd.DataFrame(columns=["code", "name", "industry", "business", "feature", "segments", "ceo", "years", "band", "mcap",
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
