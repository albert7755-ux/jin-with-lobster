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

st.set_page_config(page_title="金開心配置試算", page_icon="💰", layout="wide")

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
st.title("💰 金開心配置試算")
st.caption("選標的、填金額，立刻看到每個月的現金流　"
           + (f"｜報價檔更新：{updated}" if updated else ""))

total_capital = st.number_input("投資本金（總額）", min_value=0.0, value=0.0, step=100000.0,
                                format="%.0f", help="僅供比對各標的金額加總是否相符")

# 標的來源：報價檔債券 + 自行新增
opts = {}
for b in bonds:
    label = (f'{b["name"]}（{b.get("code","")}）｜{b["ccy"]} {b.get("coupon","-")}%'
             f'｜當期 {b["cy"]:.2f}%' if b.get("cy") else f'{b["name"]}（{b.get("code","")}）')
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
        st.session_state.setdefault("manual", []).append(
            {"name": n_name, "cy": n_cy, "freq": n_freq, "_type": n_type,
             "_key": f"m_{n_name}", "code": "", "maturity": "-", "ratings": "", "tag": ""})
        st.rerun()

for m in st.session_state.get("manual", []):
    picked.append(m)

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
            key=f"amt_{item['_key']}")

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
    "標的": r["標的"], "代碼": r["代碼"], "類型": r["類型"],
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

st.caption(
    "配息月份係依配息頻率與到期月推算，少數債券實際付息日可能不同號，請以產品說明書為準。｜"
    "當期收益率＝票面÷Offer，未計入前手息、信託管理費與匯率影響。｜"
    "報價來自總行報價檔，實際成交與可否承作以總行系統為準。本試算僅供內部參考，非投資建議。")
