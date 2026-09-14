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
    """
    搜尋可用的中文字型。涵蓋 fonts/ 子資料夾、專案根目錄，
    以及常見的檔名誤植(例如 NotoSansTC-Regular.tff.ttf)。
    """
    base = os.path.dirname(os.path.abspath(__file__))
    cands = [os.getenv("BOND_SHEET_FONT", "")]
    # 優先序:微軟正黑體(若自行放入) > Noto Sans TC > 其他
    names = ("msjh.ttc", "msjh.ttf", "msjhbd.ttc", "MSJH.TTC",
             "NotoSansTC-Regular.ttf", "NotoSansTC-Regular.tff.ttf",
             "NotoSansTC.ttf", "NotoSansTC-Regular.otf",
             "NotoSansCJKtc-Regular.otf", "SourceHanSansTC-Regular.otf")
    # 常見誤植副檔名也一併嘗試(.tff / .ttf. / 大小寫)
    names = names + tuple(n.replace(".ttf", ".tff") for n in names if n.endswith(".ttf"))
    for n in names:
        cands += [os.path.join(base, "fonts", n), os.path.join(base, n),
                  os.path.join(os.getcwd(), "fonts", n), os.path.join(os.getcwd(), n)]
    # 掃描 fonts/ 與根目錄下所有看起來像字型的檔案(含副檔名打錯的情況)
    for d in (os.path.join(base, "fonts"), base,
              os.path.join(os.getcwd(), "fonts"), os.getcwd()):
        for pat in ("*.tt*", "*.TT*", "*.otf", "*.OTF", "*.tff", "*.TFF",
                    "*Noto*", "*noto*", "*msjh*", "*MSJH*"):
            cands += sorted(glob.glob(os.path.join(d, pat)))
    cands += ["/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
              "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
              "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
              "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc"]
    seen = set()
    for c in cands:
        try:
            if not c or c in seen or not os.path.isfile(c):
                continue
            seen.add(c)
            if os.path.getsize(c) <= 50 * 1024:
                continue
            # 驗證檔案真的是可用字型(避免抓到非字型檔或損壞檔)
            from matplotlib import font_manager as _fm
            _fm.FontProperties(fname=c).get_name()
            print(f"[JKX] 使用字型:{c}")
            return c
        except Exception:
            continue
    print("[JKX] 警告:找不到可用的中文字型,圖表中文將顯示為方框。"
          "請將字型檔放到 repo 的 fonts/ 資料夾,且副檔名須為 .ttf/.ttc/.otf,"
          "例如 fonts/NotoSansTC-Regular.ttf(注意是 ttf 不是 tff)")
    return None


FREQ_N = {"每月": 12, "每季": 4, "每半年": 2, "每年": 1,
          "月配": 12, "季配": 4, "半年配": 2, "年配": 1}


def modified_duration(coupon_pct, ytm_pct, years, freq_label):
    """
    以現金流折現法計算修正存續期間(Modified Duration)與凸性(Convexity)。
    coupon_pct/ytm_pct 為年化百分比,years 為剩餘年期。
    回傳 (修正存續期間, 凸性) ;資料不足回 (None, None)
    """
    try:
        m = FREQ_N.get(str(freq_label).strip(), 2)
        y = float(ytm_pct) / 100.0
        c = float(coupon_pct) / 100.0
        n = int(round(float(years) * m))
        if n < 1 or y <= -0.99:
            return None, None
        cpn = 100.0 * c / m
        yq = y / m
        pv_sum, wt_sum, cx_sum = 0.0, 0.0, 0.0
        for t in range(1, n + 1):
            cf = cpn + (100.0 if t == n else 0.0)
            pv = cf / ((1 + yq) ** t)
            pv_sum += pv
            wt_sum += (t / m) * pv
            # 凸性分子:t(t+1)*CF/(1+y/m)^(t+2)
            cx_sum += t * (t + 1) * cf / ((1 + yq) ** (t + 2))
        if pv_sum <= 0:
            return None, None
        mod = (wt_sum / pv_sum) / (1 + yq)
        convexity = cx_sum / (pv_sum * (m ** 2))
        return mod, convexity
    except Exception:
        return None, None


def price_change_pct(mod_dur, convexity, bp):
    """
    估計價格變動% = -D×Δy + ½×C×Δy²
    第二項(凸性)使利率下跌時漲幅大於上漲時的跌幅。
    """
    dy = bp / 10000.0
    return (-mod_dur * dy + 0.5 * (convexity or 0.0) * dy * dy) * 100.0


def build_chart(df_m, out_png):
    """每月配息長條圖（中文字型以 FontProperties 直接綁定每個元素，避免 rcParams 失效）"""
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        from matplotlib import font_manager
        fp = _cjk_font()
        PROP = None
        if fp:
            # 先建立 FontProperties(這一步最關鍵,直接綁檔案,不依賴字型名稱解析)
            try:
                PROP = font_manager.FontProperties(fname=fp)
                print(f"[JKX] chart font 已綁定檔案: {fp}")
            except Exception as e:
                print(f"[JKX] FontProperties 建立失敗: {e}")
            # 再嘗試註冊到字型管理員(失敗不影響 PROP)
            try:
                font_manager.fontManager.addfont(fp)
                if PROP is not None:
                    _fam = PROP.get_name()
                    matplotlib.rcParams["font.family"] = "sans-serif"
                    matplotlib.rcParams["font.sans-serif"] = \
                        [_fam] + list(matplotlib.rcParams.get("font.sans-serif", []))
                    print(f"[JKX] chart font 已註冊: {_fam}")
            except Exception as e:
                print(f"[JKX] addfont 略過(不影響繪圖): {e}")
        else:
            print("[JKX] 找不到中文字型,圖表中文將顯示為方框")
        matplotlib.rcParams["axes.unicode_minus"] = False
        vals = list(df_m["每月合計"])
        labels = [m.replace("月", "") for m in MONTHS]   # 一、二…十二,縮短避免重疊
        fig, ax = plt.subplots(figsize=(9.2, 2.9), dpi=170)
        bars = ax.bar(labels, vals, color="#1F8AC0", width=0.62)
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
            if PROP is not None:
                lb.set_fontproperties(PROP)
            lb.set_fontsize(9)
        ax.set_title("每年配息時程（各月合計）", fontproperties=PROP, fontsize=11.5,
                     color="#0B2A4A", pad=10)
        ax.set_xlabel("月", fontproperties=PROP, fontsize=9, labelpad=2)
        fig.tight_layout(pad=0.4)
        fig.savefig(out_png, bbox_inches="tight", facecolor="white")
        plt.close(fig)
        return out_png
    except Exception as e:
        print(f"[JKX] chart fail: {e}")
        return None


def build_pdf(out_path, rows, df_m, total_amt, total_annual, blended,
              note="", today=None, client_name=""):   # client_name 已停用,保留參數相容舊呼叫
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm
    from reportlab.lib import colors
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                    TableStyle, Image as RLImage, KeepTogether,
                                    CondPageBreak)
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
    st_th = ParagraphStyle("th", fontName=FN, fontSize=8, leading=11,
                           textColor=colors.white, alignment=1)  # 表頭:Paragraph 會蓋掉表格 TEXTCOLOR

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
    el.append(Paragraph(f"配息現金流規劃　{today:%Y/%m/%d}　|　內部試算", st_sub))
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
    head = [Paragraph(f"<b>{h}</b>", st_th) for h in
            ["標的", "類型", "投資金額", "占比", "配息率", "YTM", "頻率",
             "年化配息", "到期日", "剩餘年期", "配息月份"]]
    data = [head]
    for r in rows:
        _mat = r.get("到期日") or "-"
        if _mat and _mat not in ("-", "—"):
            _mat = str(_mat)[:10].replace("-", "/")
        _yr = r.get("剩餘年期")
        _ytm = r.get("YTM")
        data.append([
            Paragraph(str(r["標的"])[:20], st_c), r["類型"],
            f'{r["投資金額"]:,.0f}', f'{r["投資金額"]/total_amt*100:.1f}%',
            f'{r["當期收益率%"]:.2f}%',
            (f'{_ytm:.2f}%' if isinstance(_ytm, (int, float)) else "-"),
            r["配息頻率"], f'{r["年化配息"]:,.0f}',
            _mat, (f'{_yr:.1f}年' if isinstance(_yr, (int, float)) else "-"),
            Paragraph(("每月" if len(r["配息月份"]) == 12
                       else "、".join(MONTHS[m - 1].replace("月", "") for m in r["配息月份"]) + "月"), st_c)])
    tb = Table(data, colWidths=[W*0.17, W*0.06, W*0.11, W*0.055, W*0.07, W*0.07,
                                W*0.07, W*0.10, W*0.09, W*0.07, W*0.12])
    tb.setStyle(TableStyle([("FONTNAME", (0, 0), (-1, -1), FN), ("FONTSIZE", (0, 0), (-1, -1), 8),
                            ("BACKGROUND", (0, 0), (-1, 0), NAVY), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D9D9D9")),
                            ("ALIGN", (1, 1), (-2, -1), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                            ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#FAFBFC")])]))
    el.append(tb)

    # 圖表:以 reportlab 原生繪製(與表格共用字型引擎,中文不受外部字型檔影響)
    try:
        from reportlab.graphics.shapes import Drawing, Rect, String, Line
        _vals = [float(df_m.loc[m, "每月合計"]) for m in MONTHS]
        _mx = max(_vals) if _vals else 0
        _cw, _ch = W, 4.6 * cm
        _pad_l, _pad_b, _pad_t = 0.2 * cm, 0.85 * cm, 0.75 * cm
        _plot_h = _ch - _pad_b - _pad_t
        _slot = (_cw - _pad_l) / 12.0
        _bw = _slot * 0.58
        d = Drawing(_cw, _ch)
        d.add(Line(_pad_l, _pad_b, _cw, _pad_b, strokeColor=colors.HexColor("#ccd6e0"),
                   strokeWidth=0.8))
        d.add(String(_cw / 2, _ch - 0.42 * cm, "每年配息時程（各月合計）",
                     fontName=FN, fontSize=10.5, fillColor=NAVY, textAnchor="middle"))
        for _i, (_m, _v) in enumerate(zip(MONTHS, _vals)):
            _x = _pad_l + _i * _slot + (_slot - _bw) / 2
            _h = (_v / _mx * _plot_h) if _mx else 0
            if _h > 0:
                d.add(Rect(_x, _pad_b, _bw, _h, fillColor=(NAVY if _v == _mx else BLUE),
                           strokeColor=None))
                d.add(String(_x + _bw / 2, _pad_b + _h + 2.5, f"{_v:,.0f}",
                             fontName=FN, fontSize=6.6, fillColor=NAVY, textAnchor="middle"))
            d.add(String(_x + _bw / 2, _pad_b - 0.42 * cm, _m.replace("月", ""),
                         fontName=FN, fontSize=7.6, fillColor=colors.HexColor("#555555"),
                         textAnchor="middle"))
        el.append(Spacer(1, 0.25 * cm))
        el.append(d)
    except Exception as e:
        print(f"[JKX] 原生圖表失敗: {e}")
    png = None

    # 時程表
    cols = [c for c in df_m.columns if c != "每月合計"]
    hdr = ["月份"] + [Paragraph(str(c)[:12], st_th) for c in cols] + ["合計"]
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
    # 標題與前幾列保持在一起(避免標題落單),表格本身可跨頁並自動重複表頭
    el.append(CondPageBreak(4.5 * cm))
    el.append(sec("每年配息時程表"))
    el.append(tm)

    # ===== 利率敏感度(僅債券部位) =====
    _bonds = [r for r in rows if r.get("類型") == "債券"
              and isinstance(r.get("剩餘年期"), (int, float))
              and isinstance(r.get("YTM"), (int, float))]
    if _bonds:
        _bond_amt = sum(r["投資金額"] for r in _bonds)
        _wd, _wc, _rows_d = 0.0, 0.0, []
        for r in _bonds:
            _md, _cx = modified_duration(r.get("票面") or r["當期收益率%"],
                                         r["YTM"], r["剩餘年期"], r["配息頻率"])
            if _md is None:
                continue
            _w = r["投資金額"] / _bond_amt
            _wd += _w * _md
            _wc += _w * (_cx or 0.0)
            _rows_d.append([Paragraph(str(r["標的"])[:22], st_c),
                            f'{r["剩餘年期"]:.1f}年', f'{_md:.2f}',
                            f'{(_cx or 0):.1f}',
                            f'{r["投資金額"]:,.0f}', f'{_w*100:.1f}%'])
        if _rows_d and _wd > 0:
            el.append(CondPageBreak(6 * cm))
            el.append(sec("利率敏感度（估算，僅債券部位）"))
            _t1 = Table([[Paragraph(f"<b>{h}</b>", st_th) for h in
                          ["標的", "剩餘年期", "修正存續期間", "凸性", "投資金額", "債券部位占比"]]] + _rows_d,
                        colWidths=[W*0.30, W*0.12, W*0.16, W*0.10, W*0.18, W*0.14])
            _t1.setStyle(TableStyle([("FONTNAME", (0,0), (-1,-1), FN), ("FONTSIZE", (0,0), (-1,-1), 8),
                                     ("BACKGROUND", (0,0), (-1,0), NAVY), ("TEXTCOLOR", (0,0), (-1,0), colors.white),
                                     ("GRID", (0,0), (-1,-1), 0.4, colors.HexColor("#D9D9D9")),
                                     ("ALIGN", (1,1), (-1,-1), "CENTER"),
                                     ("TOPPADDING", (0,0), (-1,-1), 4), ("BOTTOMPADDING", (0,0), (-1,-1), 4),
                                     ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#FAFBFC")])]))
            el.append(_t1)
            el.append(Spacer(1, 0.15 * cm))
            _sc = [[Paragraph(f"<b>{h}</b>", st_th) for h in
                    ["利率變動", "估計價格變動", "債券部位價值變動", "變動後債券部位"]]]
            for _bp in (-100, -50, 50, 100):
                _chg = price_change_pct(_wd, _wc, _bp)
                _amt = _bond_amt * _chg / 100.0
                _sc.append([f"{_bp:+d} bp", f"{_chg:+.2f}%", f"{_amt:+,.0f}",
                            f"{_bond_amt + _amt:,.0f}"])
            _t2 = Table(_sc, colWidths=[W*0.22, W*0.24, W*0.27, W*0.27])
            _t2.setStyle(TableStyle([("FONTNAME", (0,0), (-1,-1), FN), ("FONTSIZE", (0,0), (-1,-1), 9),
                                     ("BACKGROUND", (0,0), (-1,0), NAVY), ("TEXTCOLOR", (0,0), (-1,0), colors.white),
                                     ("GRID", (0,0), (-1,-1), 0.4, colors.HexColor("#D9D9D9")),
                                     ("ALIGN", (0,0), (-1,-1), "CENTER"),
                                     ("TOPPADDING", (0,0), (-1,-1), 5), ("BOTTOMPADDING", (0,0), (-1,-1), 5),
                                     ("TEXTCOLOR", (2,1), (2,2), colors.HexColor("#1B7A3D")),
                                     ("TEXTCOLOR", (2,3), (2,4), colors.HexColor("#B23A2E"))]))
            el.append(_t2)
            el.append(Spacer(1, 0.1 * cm))
            _up = price_change_pct(_wd, _wc, 100)
            _dn = price_change_pct(_wd, _wc, -100)
            el.append(Paragraph(
                f"債券部位加權平均修正存續期間約 <b>{_wd:.2f}</b>、凸性約 <b>{_wc:.1f}</b>。"
                f"已納入凸性調整，因此利率下跌 100bp 的估計漲幅（{_dn:+.2f}%）"
                f"大於上漲 100bp 的估計跌幅（{_up:+.2f}%），"
                "此為債券價格與殖利率呈凸性關係之特性，年期愈長、票面愈低者愈明顯。", st_small))
            el.append(Paragraph(
                "本表為簡化估算：以現金流折現法計算，未計入提前買回條款、信用利差變動、"
                "流動性與匯率影響；實際價格以總行報價為準。基金與結構型商品未納入計算。", st_small))

    if note:
        el.append(Spacer(1, 0.2 * cm))
        nb = Table([[Paragraph(f"<b>備註：</b>{note}", st_c)]], colWidths=[W])
        nb.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.8, GOLD),
                                ("BACKGROUND", (0, 0), (-1, -1), CREAM),
                                ("TOPPADDING", (0, 0), (-1, -1), 7), ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
                                ("LEFTPADDING", (0, 0), (-1, -1), 8)]))
        el.append(nb)

    el.append(Spacer(1, 0.3 * cm))
    for _d in [
        "**名詞解釋：當期收益率＝票息／offer price",
        "**基金配息率＝年化配息率（僅概算，以基金公司公告為主）",
        "**下一配息日僅為推算，一切以產品說明書為主",
        "**整包投資組合的年化配息率僅試算，不代表真實結果",
        "**網行銀可否申購以總行公佈為主",
        "**永續債到期日預設第一次可贖回日",
    ]:
        el.append(Paragraph(_d, st_small))
    el.append(Spacer(1, 0.12 * cm))
    _warn = Table([[Paragraph("<b>**本文件僅供內部試算使用　請勿外流**</b>",
                              ParagraphStyle("w", fontName=FN, fontSize=9.5, leading=13,
                                             textColor=colors.HexColor("#B23A2E"), alignment=1))]],
                  colWidths=[W])
    _warn.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#B23A2E")),
                               ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#FDECEC")),
                               ("TOPPADDING", (0, 0), (-1, -1), 6),
                               ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))
    el.append(_warn)

    def _footer(canv, doc_):
        # ── 浮水印(斜向淺灰,置於內容下層) ──
        canv.saveState()
        canv.setFont(FN, 44)
        canv.setFillColor(colors.Color(0.55, 0.6, 0.68, alpha=0.13))
        canv.translate(A4[0] / 2, A4[1] / 2)
        canv.rotate(38)
        for _dy in (5.5 * cm, 0, -5.5 * cm):
            canv.drawCentredString(0, _dy, "僅限內部教育訓練使用")
        canv.restoreState()
        # ── 頁尾 ──
        canv.saveState(); canv.setFont(FN, 7.5)
        canv.setFillColor(colors.HexColor("#B23A2E"))
        canv.drawString(1.5 * cm, 0.85 * cm,
                        "僅限內部教育訓練使用｜**本文件僅供內部試算使用 請勿外流**")
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
