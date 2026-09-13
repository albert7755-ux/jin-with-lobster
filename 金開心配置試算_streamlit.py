# -*- coding: utf-8 -*-
"""
金開心配置試算（Streamlit 版）
================================
把 Excel 版的核心——「每年配息時程表」——搬上網頁，並串接龍蝦的報價 API，
報價永遠是最新的，不用每次手動更新券單。

功能：
  1. 從龍蝦 API 取得可申購債券（也可手動新增基金/SI）
  2. 選 1~8 個標的、輸入投資金額
  3. 自動算：資金配置比例、各標的年化配息、整包年化配息率
  4. 每年配息時程表（12 個月，依配息頻率與到期月推算）
  5. 現金流長條圖 + 明細表，可下載 Excel

執行：streamlit run 金開心配置試算_streamlit.py
環境變數（或在側邊欄輸入）：
  LOBSTER_API   例 https://eln-bot.onrender.com/api/bonds
  LOBSTER_TOKEN 即 BOND_UPLOAD_TOKEN
"""
import os
import io
from datetime import date, datetime

import pandas as pd
import requests
import streamlit as st

st.set_page_config(page_title="金開心配置試算", page_icon="💰", layout="wide",
                   initial_sidebar_state="expanded")

# ---------- 自訂樣式:富邦藍專業金融風,擺脫 Streamlit 預設外觀 ----------
st.markdown("""<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;700&display=swap');
html, body, [class*="css"] { font-family:'Noto Sans TC', 'Microsoft JhengHei', sans-serif; }
.stApp { background:#eef2f7; }
.block-container { padding-top:1.4rem; max-width:1400px; }
#MainMenu, footer, header { visibility:hidden; }

/* 頁首橫幅 */
.hero { background:linear-gradient(115deg,#0B2A4A 0%,#1F8AC0 100%); color:#fff;
        padding:22px 26px; border-radius:14px; margin-bottom:18px;
        box-shadow:0 4px 18px rgba(11,42,74,.18); }
.hero h1 { margin:0; font-size:25px; font-weight:700; letter-spacing:1px; }
.hero p  { margin:6px 0 0; font-size:13px; opacity:.9; }

/* 區塊標題 */
h2, h3 { color:#0B2A4A !important; font-weight:700 !important;
         border-left:5px solid #C9A227; padding-left:11px; margin-top:1.4rem !important; }

/* 卡片容器 */
div[data-testid="stVerticalBlockBorderWrapper"] { background:#fff; border-radius:12px; }

/* 指標卡 */
div[data-testid="stMetric"] { background:#fff; border:1px solid #e3e9f0; border-top:3px solid #1F8AC0;
    border-radius:10px; padding:14px 16px; box-shadow:0 2px 8px rgba(0,0,0,.05); }
div[data-testid="stMetricLabel"] p { font-size:13px !important; color:#5a6b7d !important; }
div[data-testid="stMetricValue"] { font-size:26px !important; color:#0B2A4A !important; font-weight:700; }

/* 輸入元件:明確可見的框線與底色,讓人一看就知道可以輸入 */
.stNumberInput div[data-baseweb="input"],
.stTextInput div[data-baseweb="input"],
.stSelectbox div[data-baseweb="select"] > div {
    border:1.6px solid #9db4c9 !important; border-radius:9px !important;
    background:#fff !important; box-shadow:0 1px 3px rgba(11,42,74,.07) !important; }
.stNumberInput div[data-baseweb="input"]:hover,
.stTextInput div[data-baseweb="input"]:hover,
.stSelectbox div[data-baseweb="select"] > div:hover { border-color:#1F8AC0 !important; }
.stNumberInput div[data-baseweb="input"]:focus-within,
.stTextInput div[data-baseweb="input"]:focus-within {
    border-color:#1F8AC0 !important; box-shadow:0 0 0 3px rgba(31,138,192,.16) !important; }
.stNumberInput input, .stTextInput input {
    background:#fff !important; color:#0B2A4A !important; font-weight:600; font-size:15px !important; }
/* 輸入框標籤 */
.stNumberInput label p, .stTextInput label p, .stSelectbox label p {
    font-weight:600 !important; color:#2c3f52 !important; font-size:13.5px !important; }
/* 數字加減鈕 */
.stNumberInput button { background:#f1f5f9 !important; border-left:1px solid #dbe4ec !important; }
.stNumberInput button:hover { background:#e2ebf3 !important; }
/* 多選(標的選擇)框 */
.stMultiSelect div[data-baseweb="select"] > div {
    border:1.6px solid #9db4c9 !important; border-radius:9px !important;
    background:#fff !important; min-height:46px; }
.stMultiSelect div[data-baseweb="select"] > div:hover { border-color:#1F8AC0 !important; }
.stButton button { background:#1F8AC0; color:#fff; border:0; border-radius:8px;
    padding:.45rem 1.3rem; font-weight:600; transition:.15s; }
.stButton button:hover { background:#0B2A4A; transform:translateY(-1px); }
.stDownloadButton button { background:#C9A227; color:#fff; border:0; border-radius:8px;
    font-weight:600; padding:.5rem 1.4rem; }
.stDownloadButton button:hover { background:#a8871f; }

/* 表格 */
div[data-testid="stDataFrame"] { border:1px solid #e3e9f0; border-radius:10px; overflow:hidden; }

/* 側邊欄 */
section[data-testid="stSidebar"] { background:#0B2A4A; }
section[data-testid="stSidebar"] * { color:#dbe6f0 !important; }
section[data-testid="stSidebar"] input { background:#163a5f !important; border-color:#2b5680 !important; }

/* 展開區 */
details { background:#fff !important; border:1px solid #e3e9f0 !important; border-radius:10px !important; }
</style>""", unsafe_allow_html=True)

MONTHS = ["一月", "二月", "三月", "四月", "五月", "六月",
          "七月", "八月", "九月", "十月", "十一月", "十二月"]
FREQ_MONTHS = {"每月": 1, "每季": 3, "每半年": 6, "每年": 12,
               "月配": 1, "季配": 3, "半年配": 6, "年配": 12}


# ---------- 取得報價 ----------
@st.cache_data(ttl=600, show_spinner=False)
def load_bonds(api, token):
    r = requests.get(api, params={"token": token}, timeout=30)
    r.raise_for_status()
    d = r.json()
    return d.get("bonds", []), d.get("updated_at", "")


def payment_months(freq, maturity: date):
    """
    依配息頻率與到期月推算「哪幾個月份會配息」。
    半年配：到期月與其 +6 個月；季配：到期月起每 3 個月；月配：全部月份。
    """
    step = FREQ_MONTHS.get(str(freq).strip(), None)
    if not step:
        return []
    if step == 1:
        return list(range(1, 13))
    base = maturity.month
    return sorted({((base - 1 + k * step) % 12) + 1 for k in range(12 // step)})


def build_rows(picked, amounts):
    """計算每個標的的年化配息、每期金額與配息月份"""
    rows = []
    for item in picked:
        amt = float(amounts.get(item["_key"], 0) or 0)
        if amt <= 0:
            continue
        cy = float(item.get("cy") or 0) / 100.0          # 當期收益率(小數)
        annual = amt * cy
        if item["_type"] == "債券":
            mat = datetime.fromisoformat(item["maturity"]).date()
            pm = payment_months(item.get("freq"), mat)
        else:
            pm = list(range(1, 13))                       # 基金預設月配
        per_pay = annual / len(pm) if pm else 0
        rows.append({
            "標的": item["name"], "代碼": item.get("code", ""), "類型": item["_type"],
            "產業": item.get("sector", "—"), "較美債bp": item.get("spread_bp"),
            "投資金額": amt, "當期收益率%": cy * 100, "配息頻率": item.get("freq", "月配"),
            "年化配息": annual, "每期配息": per_pay, "配息月份": pm,
            "到期日": item.get("maturity", "-"), "評等": item.get("ratings", ""),
            "資格": item.get("tag", ""),
        })
    return rows


def monthly_table(rows):
    """組出 12 個月 × 各標的 的配息時程表"""
    data = {}
    for r in rows:
        col = [0.0] * 12
        for m in r["配息月份"]:
            col[m - 1] += r["每期配息"]
        data[r["標的"]] = col
    df = pd.DataFrame(data, index=MONTHS)
    if not df.empty:
        df["每月合計"] = df.sum(axis=1)
    return df


# ---------- 側邊欄：資料來源 ----------
st.sidebar.header("資料來源")
api = st.sidebar.text_input("龍蝦報價 API",
                            os.getenv("LOBSTER_API", "https://eln-bot.onrender.com/api/bonds"))
token = st.sidebar.text_input("存取密碼", os.getenv("LOBSTER_TOKEN", ""), type="password")

bonds, updated = [], ""
if token:
    try:
        bonds, updated = load_bonds(api, token)
        st.sidebar.success(f"已載入 {len(bonds)} 檔\n報價檔：{updated}")
    except Exception as e:
        st.sidebar.error(f"載入失敗：{str(e)[:120]}")
else:
    st.sidebar.info("請輸入存取密碼以載入報價")

st.sidebar.markdown("---")
st.sidebar.caption("基金／SI 等非報價檔商品，可在下方「自行新增標的」手動輸入配息率。")

# ---------- 主畫面 ----------
st.markdown(f"""<div class="hero">
<h1>金開心配置試算</h1>
<p>選標的、填金額，立刻看到客戶每個月的現金流
{f"　·　報價檔更新：{updated}" if updated else ""}</p></div>""", unsafe_allow_html=True)

cap1, cap2 = st.columns([1, 2])
with cap1:
    total_capital = st.number_input("💵 投資本金（總額）", min_value=0.0, value=0.0,
                                    step=100000.0, format="%.0f",
                                    placeholder="請輸入金額，例如 10000000",
                                    help="選填。填了之後會自動比對各標的金額加總是否相符。")
with cap2:
    st.markdown("<div style='padding-top:34px;color:#5a6b7d;font-size:13px'>"
                "選填欄位。先填總額，下面配置完會自動幫你核對加總是否吻合。</div>",
                unsafe_allow_html=True)

# 標的來源：報價檔債券 + 自行新增
opts = {}
for b in bonds:
    _sec = b.get("sector") or ""
    _sec_s = f'｜{_sec}' if _sec and _sec != "未分類" else ""
    label = (f'{b["name"]}（{b.get("code","")}）｜{b["ccy"]} {b.get("coupon","-")}%'
             f'｜當期 {b["cy"]:.2f}%{_sec_s}' if b.get("cy")
             else f'{b["name"]}（{b.get("code","")}）{_sec_s}')
    opts[label] = dict(b, _type="債券", _key=b.get("code") or b["name"])

st.subheader("① 選擇標的")
picked_labels = st.multiselect("可搜尋名稱或代碼（最多 8 個）", list(opts.keys()), max_selections=8)
picked = [opts[l] for l in picked_labels]

with st.expander("➕ 自行新增標的（基金、SI、定存等）"):
    c1, c2, c3, c4 = st.columns([3, 1.2, 1.2, 1.2])
    n_name = c1.text_input("名稱", key="mn")
    n_cy = c2.number_input("配息率 %", min_value=0.0, max_value=30.0, value=0.0, step=0.1, key="mc")
    n_freq = c3.selectbox("配息頻率", ["每月", "每季", "每半年", "每年"], key="mf")
    n_type = c4.selectbox("類型", ["基金", "SI", "其他"], key="mt")
    if st.button("加入") and n_name and n_cy > 0:
        _ms = st.session_state.setdefault("manual", [])
        _key = f"m_{n_name}"
        if any(x["_key"] == _key for x in _ms):
            st.warning(f"「{n_name}」已在清單中，請改用不同名稱。")
        else:
            _ms.append({"name": n_name, "cy": n_cy, "freq": n_freq, "_type": n_type,
                        "_key": _key, "code": "", "maturity": "-", "ratings": "",
                        "tag": "", "sector": "—", "spread_bp": None})
            st.rerun()
    _ms = st.session_state.get("manual", [])
    if _ms:
        st.caption("已自行新增：")
        for i_, m_ in enumerate(list(_ms)):
            cA, cB = st.columns([5, 1])
            cA.write(f"・{m_['name']}（{m_['_type']}　{m_['cy']:g}%　{m_['freq']}）")
            if cB.button("移除", key=f"del_{m_['_key']}"):
                _ms.pop(i_)
                st.rerun()

# 合併已選標的與手動新增標的,並以 _key 去重(避免 widget key 重複)
picked = list({x["_key"]: x for x in (picked + st.session_state.get("manual", []))}.values())

if not picked:
    st.info("請先選擇至少一個標的。")
    st.stop()

# ---------- 輸入金額 ----------
st.subheader("② 輸入各標的投資金額")
amounts = {}
cols = st.columns(min(4, len(picked)))
for i, item in enumerate(picked):
    with cols[i % len(cols)]:
        amounts[item["_key"]] = st.number_input(
            item["name"][:18], min_value=0.0, value=0.0, step=10000.0, format="%.0f",
            key=f"amt_{i}_{item['_key']}")

rows = build_rows(picked, amounts)
if not rows:
    st.warning("請至少輸入一個標的的投資金額。")
    st.stop()

total_amt = sum(r["投資金額"] for r in rows)
total_annual = sum(r["年化配息"] for r in rows)
blended = (total_annual / total_amt * 100) if total_amt else 0

# ---------- 摘要 ----------
st.subheader("③ 試算結果")
m1, m2, m3, m4 = st.columns(4)
m1.metric("投入總額", f"{total_amt:,.0f}")
m2.metric("預估每年領息", f"{total_annual:,.0f}")
m3.metric("預估每月平均", f"{total_annual/12:,.0f}")
m4.metric("整包年化配息率", f"{blended:.2f}%")
if total_capital and abs(total_capital - total_amt) > 1:
    st.warning(f"各標的金額加總 {total_amt:,.0f}，與投資本金 {total_capital:,.0f} 相差 "
               f"{total_capital - total_amt:,.0f}")

df_detail = pd.DataFrame([{
    "標的": r["標的"], "代碼": r["代碼"], "類型": r["類型"], "產業": r.get("產業", "—"),
    "較美債bp": r.get("較美債bp"),
    "投資金額": r["投資金額"], "配置比例%": r["投資金額"] / total_amt * 100,
    "當期收益率%": r["當期收益率%"], "配息頻率": r["配息頻率"],
    "年化配息": r["年化配息"], "每期配息": r["每期配息"],
    "配息月份": "、".join(MONTHS[m - 1] for m in r["配息月份"]),
    "到期日": r["到期日"], "評等": r["評等"], "資格": r["資格"],
} for r in rows])
st.dataframe(df_detail.style.format({
    "投資金額": "{:,.0f}", "配置比例%": "{:.1f}", "當期收益率%": "{:.2f}",
    "年化配息": "{:,.0f}", "每期配息": "{:,.0f}"}), use_container_width=True, hide_index=True)

# ---------- 配息時程表 ----------
st.subheader("④ 每年配息時程表")
df_m = monthly_table(rows)
st.bar_chart(df_m["每月合計"], height=260)
# 不用 background_gradient(需 matplotlib),改用內建長條顯示每月合計的相對大小
st.dataframe(
    df_m.style.format("{:,.0f}"),
    use_container_width=True,
    column_config={"每月合計": st.column_config.ProgressColumn(
        "每月合計", format="%.0f", min_value=0,
        max_value=float(df_m["每月合計"].max() or 1))},
)

c1, c2, c3 = st.columns(3)
c1.metric("配息最多的月份", f"{df_m['每月合計'].idxmax()}　{df_m['每月合計'].max():,.0f}")
c2.metric("配息最少的月份", f"{df_m['每月合計'].idxmin()}　{df_m['每月合計'].min():,.0f}")
zero_m = [m for m in MONTHS if df_m.loc[m, "每月合計"] == 0]
c3.metric("沒有配息的月份", f"{len(zero_m)} 個月" + (f"（{'、'.join(zero_m)}）" if zero_m else ""))

# ---------- 產業分布（這包組合） ----------
st.subheader("⑤ 產業分布")
df_sec = (df_detail[df_detail["類型"] == "債券"]
          .groupby("產業", as_index=False)
          .agg(檔數=("標的", "count"), 投資金額=("投資金額", "sum"),
               年化配息=("標的", lambda x: 0)))
if not df_sec.empty:
    # 重算年化配息(groupby lambda 無法直接取另一欄,改用 merge)
    tmp = pd.DataFrame([{"產業": r.get("產業", "—"), "年化配息": r["年化配息"]}
                        for r in rows if r["類型"] == "債券"])
    df_sec = df_sec.drop(columns=["年化配息"]).merge(
        tmp.groupby("產業", as_index=False).sum(), on="產業", how="left")
    df_sec["配置比例%"] = df_sec["投資金額"] / df_sec["投資金額"].sum() * 100
    known = df_sec[~df_sec["產業"].isin(["—", "未分類", ""])]
    if known.empty:
        st.caption("這包組合的債券尚未取得產業分類。"
                   "請在 LINE 打一次 /sector 建立分類，並確認 Render 已部署最新版 main.py。")
    else:
        c1, c2 = st.columns([1.1, 1])
        with c1:
            st.dataframe(df_sec.sort_values("投資金額", ascending=False).style.format(
                {"投資金額": "{:,.0f}", "年化配息": "{:,.0f}", "配置比例%": "{:.1f}"}),
                use_container_width=True, hide_index=True)
        with c2:
            st.bar_chart(known.set_index("產業")["投資金額"], height=260)
        if len(known) == 1:
            st.info(f"債券部位全部集中在「{known.iloc[0]['產業']}」，可考慮分散到其他產業。")
else:
    st.caption("目前組合沒有債券部位（或報價來源尚未建立產業分類）。")

# ---------- 產業利差（市場面，全架上） ----------
with st.expander("📊 全架上產業利差概況（點開看市場行情）"):
    df_all = pd.DataFrame([b for b in bonds
                           if b.get("ytm") and b.get("sector") not in (None, "", "未分類")])
    if df_all.empty:
        has_field = any("sector" in (b or {}) for b in bonds)
        st.caption(
            "尚無產業分類資料。請在 LINE 打一次 /sector 建立分類快取。"
            if has_field else
            "報價 API 尚未回傳產業欄位，請確認 Render 已部署最新版 main.py（需含 sector 欄位）。")
    else:
        df_all = df_all[(df_all["ytm"] > 0) & (df_all["ytm"] <= 25)]
        ccy_pick = st.selectbox("幣別", sorted(df_all["ccy"].dropna().unique()),
                                index=0, key="sec_ccy")
        d = df_all[df_all["ccy"] == ccy_pick]
        g = d.groupby("sector").agg(
            檔數=("name", "count"), YTM中位=("ytm", "median"),
            當期中位=("cy", "median"), 利差中位bp=("spread_bp", "median"),
            平均年期=("years", "mean")).reset_index().rename(columns={"sector": "產業"})
        g = g[g["檔數"] >= 3].sort_values("利差中位bp", na_position="last")
        st.dataframe(g.style.format({"YTM中位": "{:.2f}", "當期中位": "{:.2f}",
                                     "利差中位bp": "{:+.0f}", "平均年期": "{:.1f}"}),
                     use_container_width=True, hide_index=True)
        st.caption("利差＝每檔 YTM 減同剩餘年期的美債殖利率（曲線內插）後取中位數；"
                   "各產業平均年期不同，比較僅供參考。")

# ---------- 匯出 ----------
st.subheader("⑥ 匯出")
cA, cB = st.columns(2)
client_name = cA.text_input("客戶稱謂（選填，會印在PDF標題）", "")
pdf_note = cB.text_input("備註（選填）", "")

try:
    from jkx_pdf import build_pdf
    import tempfile as _tf
    _p = _tf.NamedTemporaryFile(suffix=".pdf", delete=False)
    build_pdf(_p.name, rows, df_m, total_amt, total_annual, blended,
              client_name=client_name, note=pdf_note, today=date.today())
    with open(_p.name, "rb") as f:
        st.download_button("📄 下載試算報告 PDF", f.read(),
                           file_name=f"金開心配置試算_{client_name or '試算'}_{date.today():%Y%m%d}.pdf",
                           mime="application/pdf")
    os.remove(_p.name)
except Exception as e:
    st.warning(f"PDF 產生失敗：{str(e)[:150]}")

# ---------- 下載 ----------
# 下載:優先輸出 Excel(兩個分頁);若環境未安裝 openpyxl 則自動退回 CSV
try:
    import openpyxl  # noqa: F401
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df_detail.to_excel(w, sheet_name="配置明細", index=False)
        df_m.to_excel(w, sheet_name="配息時程表")
    st.download_button("📥 下載試算結果 Excel", buf.getvalue(),
                       file_name=f"金開心配置試算_{date.today():%Y%m%d}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
except ModuleNotFoundError:
    csv_txt = ("【配置明細】\n" + df_detail.to_csv(index=False)
               + "\n【配息時程表】\n" + df_m.to_csv())
    st.download_button("📥 下載試算結果 CSV", csv_txt.encode("utf-8-sig"),
                       file_name=f"金開心配置試算_{date.today():%Y%m%d}.csv", mime="text/csv")
    st.caption("（環境未安裝 openpyxl，改提供 CSV；在 requirements.txt 加入 openpyxl>=3.1 即可輸出 Excel）")

st.markdown("---")
st.markdown("""
**名詞解釋：當期收益率＝票息／offer price  
**基金配息率＝年化配息率（僅概算，以基金公司公告為主）  
**下一配息日僅為推算，一切以產品說明書為主  
**整包投資組合的年化配息率僅試算，不代表真實結果  
**網行銀可否申購以總行公佈為主  
**永續債到期日預設第一次可贖回日
""")
st.markdown("""<div style="background:#FDECEC;border:1.5px solid #B23A2E;border-radius:9px;
padding:12px 16px;text-align:center;color:#B23A2E;font-weight:700;font-size:15px;margin-top:10px">
**本文件僅供內部試算使用　請勿外流**</div>""", unsafe_allow_html=True)
