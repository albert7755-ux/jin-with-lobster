# -*- coding: utf-8 -*-
"""
jkx_pdf.py — 金開心配置試算 PDF 產生器
沿用 bond_sheet 的字型搜尋邏輯,輸出直式 A4 一頁(必要時兩頁)
"""
import os
import glob
import tempfile
from datetime import date

MONTHS = ["一月", "二月", "三月", "四月", "五月", "六月",
          "七月", "八月", "九月", "十月", "十一月", "十二月"]


def _cjk_font():
    base = os.path.dirname(os.path.abspath(__file__))
    cands = [os.getenv("BOND_SHEET_FONT", "")]
    for n in ("NotoSansTC-Regular.ttf", "NotoSansTC.ttf", "msjh.ttf"):
        cands += [os.path.join(base, "fonts", n), os.path.join(base, n)]
    cands += ["/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
              "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
              "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc"]
    cands += sorted(glob.glob(os.path.join(base, "fonts", "*.tt*")))
    for c in cands:
        try:
            if c and os.path.exists(c) and os.path.getsize(c) > 50 * 1024:
                return c
        except Exception:
            pass
    return None


def build_chart(df_m, out_png):
    """每月配息長條圖"""
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        from matplotlib import font_manager
        fp = _cjk_font()
        PROP = None
        if fp:
            try:
                PROP = font_manager.FontProperties(fname=fp)
                font_manager.fontManager.addfont(fp)
                matplotlib.rcParams["font.family"] = PROP.get_name()
            except Exception:
                PROP = None
        matplotlib.rcParams["axes.unicode_minus"] = False
        vals = list(df_m["每月合計"])
        fig, ax = plt.subplots(figsize=(9.2, 2.9), dpi=170)
        bars = ax.bar(MONTHS, vals, color="#1F8AC0", width=0.62)
        mx = max(vals) if vals else 0
        for b, v in zip(bars, vals):
            if v > 0:
                ax.annotate(f"{v:,.0f}", (b.get_x() + b.get_width() / 2, v),
                            ha="center", va="bottom", fontsize=8.5, fontproperties=PROP)
            if mx and v == mx:
                b.set_color("#0B2A4A")
        ax.set_ylim(0, mx * 1.18 if mx else 1)
        ax.spines[["top", "right"]].set_visible(False)
        ax.tick_params(axis="y", labelsize=8)
        for lb in ax.get_xticklabels():
            lb.set_fontproperties(PROP); lb.set_fontsize(9)
        ax.set_title("每年配息時程（各月合計）", fontproperties=PROP, fontsize=11.5, color="#0B2A4A", pad=10)
        fig.tight_layout(pad=0.4)
        fig.savefig(out_png, bbox_inches="tight", facecolor="white")
        plt.close(fig)
        return out_png
    except Exception as e:
        print(f"[JKX] chart fail: {e}")
        return None


def build_pdf(out_path, rows, df_m, total_amt, total_annual, blended,
              client_name="", note="", today=None):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm
    from reportlab.lib import colors
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                    TableStyle, Image as RLImage)
    from reportlab.lib.styles import ParagraphStyle

    today = today or date.today()
    fp = _cjk_font()
    FN = "MSung-Light"
    if fp:
        try:
            pdfmetrics.registerFont(TTFont("CJK", fp, subfontIndex=0)); FN = "CJK"
        except Exception:
            pdfmetrics.registerFont(UnicodeCIDFont("MSung-Light"))
    else:
        pdfmetrics.registerFont(UnicodeCIDFont("MSung-Light"))

    NAVY = colors.HexColor("#0B2A4A"); BLUE = colors.HexColor("#1F8AC0")
    GOLD = colors.HexColor("#C9A227"); GRAY = colors.HexColor("#666666")
    LIGHT = colors.HexColor("#F2F6FA"); CREAM = colors.HexColor("#FFF9E6")
    W = 18.0 * cm

    st_title = ParagraphStyle("t", fontName=FN, fontSize=19, leading=24, textColor=NAVY)
    st_sub = ParagraphStyle("s", fontName=FN, fontSize=9.5, leading=13, textColor=GRAY)
    st_h = ParagraphStyle("h", fontName=FN, fontSize=12.5, leading=16, textColor=NAVY,
                          spaceBefore=9, spaceAfter=3)
    st_c = ParagraphStyle("c", fontName=FN, fontSize=8.5, leading=11.5)
    st_small = ParagraphStyle("sm", fontName=FN, fontSize=7.8, leading=11, textColor=GRAY)

    def sec(t):
        tb = Table([[Paragraph(f"<b>{t}</b>", st_h)]], colWidths=[W])
        tb.setStyle(TableStyle([("LINEBELOW", (0, 0), (-1, -1), 1.2, BLUE),
                                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                                ("BOTTOMPADDING", (0, 0), (-1, -1), 2)]))
        return tb

    doc = SimpleDocTemplate(out_path, pagesize=A4, leftMargin=1.5 * cm, rightMargin=1.5 * cm,
                            topMargin=1.2 * cm, bottomMargin=1.2 * cm)
    el = []
    el.append(Paragraph("金開心配置試算", st_title))
    sub = f"配息現金流規劃　{today:%Y/%m/%d}"
    if client_name:
        sub = f"{client_name}　|　" + sub
    el.append(Paragraph(sub, st_sub))
    el.append(Spacer(1, 0.3 * cm))

    # 摘要
    snap = [["投入總額", "預估每年領息", "預估每月平均", "整包年化配息率"],
            [f"{total_amt:,.0f}", f"{total_annual:,.0f}",
             f"{total_annual/12:,.0f}", f"{blended:.2f}%"]]
    t0 = Table(snap, colWidths=[W / 4] * 4)
    t0.setStyle(TableStyle([("FONTNAME", (0, 0), (-1, -1), FN),
                            ("FONTSIZE", (0, 0), (-1, 0), 9), ("FONTSIZE", (0, 1), (-1, 1), 14),
                            ("BACKGROUND", (0, 0), (-1, 0), NAVY), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                            ("BACKGROUND", (0, 1), (-1, 1), LIGHT), ("TEXTCOLOR", (0, 1), (-1, 1), NAVY),
                            ("ALIGN", (0, 0), (-1, -1), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                            ("TOPPADDING", (0, 0), (-1, -1), 7), ("BOTTOMPADDING", (0, 0), (-1, -1), 7)]))
    el.append(t0)

    # 配置明細
    el.append(sec("配置明細"))
    head = ["標的", "類型", "投資金額", "占比", "配息率", "頻率", "年化配息", "每期配息", "配息月份"]
    data = [head]
    for r in rows:
        data.append([
            Paragraph(str(r["標的"])[:22], st_c), r["類型"],
            f'{r["投資金額"]:,.0f}', f'{r["投資金額"]/total_amt*100:.1f}%',
            f'{r["當期收益率%"]:.2f}%', r["配息頻率"],
            f'{r["年化配息"]:,.0f}', f'{r["每期配息"]:,.0f}',
            Paragraph(("每月" if len(r["配息月份"]) == 12
                       else "、".join(MONTHS[m - 1].replace("月", "") for m in r["配息月份"]) + "月"), st_c)])
    tb = Table(data, colWidths=[W*0.20, W*0.07, W*0.12, W*0.06, W*0.08, W*0.08, W*0.11, W*0.10, W*0.18])
    tb.setStyle(TableStyle([("FONTNAME", (0, 0), (-1, -1), FN), ("FONTSIZE", (0, 0), (-1, -1), 8),
                            ("BACKGROUND", (0, 0), (-1, 0), NAVY), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D9D9D9")),
                            ("ALIGN", (1, 1), (-2, -1), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                            ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#FAFBFC")])]))
    el.append(tb)

    # 圖表
    png = None
    try:
        f = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
        png = build_chart(df_m, f.name)
        if png:
            from PIL import Image as _P
            iw, ih = _P.open(png).size
            el.append(Spacer(1, 0.25 * cm))
            el.append(RLImage(png, width=W, height=W * ih / iw))
    except Exception as e:
        print(f"[JKX] chart embed: {e}")

    # 時程表
    el.append(sec("每年配息時程表"))
    cols = [c for c in df_m.columns if c != "每月合計"]
    hdr = ["月份"] + [Paragraph(str(c)[:10], st_c) for c in cols] + ["合計"]
    rows_m = [hdr]
    for m in MONTHS:
        rows_m.append([m] + [f'{df_m.loc[m, c]:,.0f}' if df_m.loc[m, c] else "-" for c in cols]
                      + [f'{df_m.loc[m, "每月合計"]:,.0f}'])
    cw = [W * 0.10] + [(W * 0.72) / max(1, len(cols))] * len(cols) + [W * 0.18]
    tm = Table(rows_m, colWidths=cw, repeatRows=1)
    mx_month = df_m["每月合計"].idxmax() if len(df_m) else None
    style_m = [("FONTNAME", (0, 0), (-1, -1), FN), ("FONTSIZE", (0, 0), (-1, -1), 7.8),
               ("BACKGROUND", (0, 0), (-1, 0), NAVY), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
               ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D9D9D9")),
               ("ALIGN", (1, 1), (-1, -1), "RIGHT"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
               ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
               ("BACKGROUND", (-1, 1), (-1, -1), LIGHT),
               ("TEXTCOLOR", (-1, 1), (-1, -1), NAVY)]
    if mx_month in MONTHS:
        r_ = MONTHS.index(mx_month) + 1
        style_m.append(("BACKGROUND", (0, r_), (-1, r_), CREAM))
    tm.setStyle(TableStyle(style_m))
    el.append(tm)

    if note:
        el.append(Spacer(1, 0.2 * cm))
        nb = Table([[Paragraph(f"<b>備註：</b>{note}", st_c)]], colWidths=[W])
        nb.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.8, GOLD),
                                ("BACKGROUND", (0, 0), (-1, -1), CREAM),
                                ("TOPPADDING", (0, 0), (-1, -1), 7), ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
                                ("LEFTPADDING", (0, 0), (-1, -1), 8)]))
        el.append(nb)

    el.append(Spacer(1, 0.3 * cm))
    el.append(Paragraph(
        "配息月份係依配息頻率與到期月推算，少數債券實際付息日可能不同號，請以產品說明書為準。｜"
        "當期收益率＝票面÷Offer，未計入前手息、信託管理費與匯率影響；基金配息率為輸入值，"
        "配息可能由本金支付。｜報價來自總行報價檔，實際成交與可否承作以總行系統為準。｜"
        "本試算僅供內部參考，非投資建議，不保證未來績效。", st_small))

    def _footer(canv, doc_):
        canv.saveState(); canv.setFont(FN, 7.5)
        canv.setFillColor(colors.HexColor("#B23A2E"))
        canv.drawString(1.5 * cm, 0.85 * cm, "僅限內部參考使用，非投資建議")
        canv.setFillColor(GRAY)
        canv.drawRightString(A4[0] - 1.5 * cm, 0.85 * cm,
                             f"金開心配置試算 · {today:%Y/%m/%d} · 第 {doc_.page} 頁")
        canv.restoreState()

    doc.build(el, onFirstPage=_footer, onLaterPages=_footer)
    if png:
        try:
            os.remove(png)
        except Exception:
            pass
    return out_path
