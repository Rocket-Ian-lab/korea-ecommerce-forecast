# -*- coding: utf-8 -*-
import warnings; warnings.filterwarnings("ignore")
import io as _io
import os
import tempfile
from datetime import datetime
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from statsmodels.tsa.statespace.sarimax import SARIMAX

# ── reportlab 임포트 + 한글 폰트 전역 등록 ─────────────────────
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table,
        TableStyle, Image, PageBreak, HRFlowable
    )
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.lib.fonts import addMapping
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import matplotlib.ticker as mticker

    # 한글 폰트 전역 등록 — 함수 내부가 아닌 모듈 로드 시 1회만 실행
    if os.path.exists(r"C:\Windows\Fonts\malgun.ttf"):
        _FONT_REG  = r"C:\Windows\Fonts\malgun.ttf"
        _FONT_BOLD = r"C:\Windows\Fonts\malgunbd.ttf"
    else:
        _FONT_REG = '/usr/share/fonts/truetype/nanum/NanumGothic.ttf'
        _FONT_BOLD = '/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf'

    _registered = pdfmetrics.getRegisteredFontNames()
    if "Malgun" not in _registered:
        pdfmetrics.registerFont(TTFont("Malgun",   _FONT_REG))
    if "MalgunBd" not in _registered:
        pdfmetrics.registerFont(TTFont("MalgunBd", _FONT_BOLD))
    # 폰트 패밀리 매핑 (읽기/볼드 조합)
    addMapping("Malgun", 0, 0, "Malgun")
    addMapping("Malgun", 1, 0, "Malgun")
    addMapping("Malgun", 0, 1, "MalgunBd")
    addMapping("Malgun", 1, 1, "MalgunBd")

    PDF_FONT   = "Malgun"
    PDF_FONT_B = "MalgunBd"
    REPORTLAB_OK = True
except Exception as _rl_err:
    PDF_FONT   = "Helvetica"
    PDF_FONT_B = "Helvetica-Bold"
    REPORTLAB_OK = False


def generate_pdf_report(tbl: pd.DataFrame, fstart, fend, mult: float,
                        exp_ts, imp_ts, exp_fc_a, imp_fc_a,
                        exp_lo, exp_hi, imp_lo, imp_hi) -> bytes:
    """
    예측 결과를 바탕으로 A4 PDF 보고서를 생성하고 bytes를 반환합니다.
    - 표지 + KPI 요약
    - 수출/수입 예측 차트 (matplotlib)
    - 월별 예측 테이블
    - 시사점 및 시스템 정보
    """
    if not REPORTLAB_OK:
        raise ImportError("reportlab 또는 matplotlib이 설치되지 않았습니다.")

    # 모듈 레벨에서 등록된 전역 폰트 사용
    FONT   = PDF_FONT
    FONT_B = PDF_FONT_B

    W, H = A4
    MARGIN = 18 * mm
    BLUE   = colors.HexColor("#1A237E")
    BLUE2  = colors.HexColor("#1565C0")
    RED    = colors.HexColor("#B71C1C")
    GREY   = colors.HexColor("#F5F5F5")
    HEADER = colors.HexColor("#1A237E")

    def sty(name, **kw):
        base = getSampleStyleSheet()["Normal"]
        d = {"fontName": FONT, "fontSize": 9, "leading": 13}
        d.update(kw)
        return ParagraphStyle(name, parent=base, **d)

    S_TITLE = sty("t", fontSize=19, fontName=FONT_B, alignment=TA_CENTER,
                  leading=26, textColor=BLUE)
    S_SUB   = sty("s", fontSize=10, alignment=TA_CENTER, textColor=colors.grey)
    S_H1    = sty("h1", fontSize=13, fontName=FONT_B, textColor=BLUE2,
                  leading=18, spaceAfter=3)
    S_BODY  = sty("bd", fontSize=8.5, leading=13, textColor=colors.HexColor("#212121"))
    S_CAP   = sty("cp", fontSize=7.5, alignment=TA_CENTER, textColor=colors.grey)
    S_NOTE  = sty("nt", fontSize=7.5, textColor=colors.HexColor("#555"))

    def hr():
        return HRFlowable(width="100%", thickness=0.7, color=BLUE2,
                          spaceAfter=4, spaceBefore=4)
    def sp(n=5):
        return Spacer(1, n * mm)

    def tbl_style(rows, cw, header_bg=HEADER):
        t = Table(rows, colWidths=cw, repeatRows=1)
        t.setStyle(TableStyle([
            ("BACKGROUND",    (0,0), (-1,0), header_bg),
            ("TEXTCOLOR",     (0,0), (-1,0), colors.white),
            ("FONTNAME",      (0,0), (-1,0), FONT_B),
            ("FONTNAME",      (0,1), (-1,-1), FONT),
            ("FONTSIZE",      (0,0), (-1,-1), 8),
            ("ROWBACKGROUNDS",(0,1), (-1,-1), [colors.white, GREY]),
            ("GRID",          (0,0), (-1,-1), 0.35, colors.HexColor("#BDBDBD")),
            ("TOPPADDING",    (0,0), (-1,-1), 4),
            ("BOTTOMPADDING", (0,0), (-1,-1), 4),
            ("LEFTPADDING",   (0,0), (-1,-1), 5),
            ("RIGHTPADDING",  (0,0), (-1,-1), 5),
            ("ALIGN",         (1,1), (-1,-1), "RIGHT"),
            ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
        ]))
        return t

    def on_page(canvas, doc):
        canvas.saveState()
        canvas.setFont(FONT, 7.5)
        canvas.setFillColor(colors.grey)
        canvas.drawRightString(W - MARGIN, 11*mm, f"- {doc.page} -")
        canvas.drawString(MARGIN, 11*mm, "한국 전자상거래 수출입 AI 예측 보고서")
        canvas.setStrokeColor(colors.HexColor("#BDBDBD"))
        canvas.line(MARGIN, 13*mm, W - MARGIN, 13*mm)
        canvas.restoreState()

    # ── 차트 생성 (matplotlib → PNG → BytesIO) ────────────────
    def make_chart_image(exp_ts, imp_ts, exp_fc_a, imp_fc_a,
                         exp_lo, exp_hi, imp_lo, imp_hi, fstart):
        if os.path.exists(r"C:\Windows\Fonts\malgun.ttf"):
            plt.rcParams["font.family"] = "Malgun Gothic"
        else:
            plt.rcParams["font.family"] = "NanumGothic"
        plt.rcParams["axes.unicode_minus"] = False
        fig, axes = plt.subplots(2, 1, figsize=(13, 7))
        fig.suptitle("전자상거래 수출입 예측 (SARIMAX + 이벤트 회귀)",
                     fontsize=12, fontweight="bold")
        for ax, ts, fc, lo, hi, label, clr in [
            (axes[0], exp_ts, exp_fc_a, exp_lo, exp_hi, "수출", "#1565C0"),
            (axes[1], imp_ts, imp_fc_a, imp_lo, imp_hi, "수입", "#C62828"),
        ]:
            ax.plot(ts.index, ts.values/1e6, "o", color=clr, ms=3, alpha=.6, label="실적")
            ax.plot(fc.index, fc.values/1e6, "-", color=clr, lw=2.2, label="예측")
            ax.fill_between(fc.index, lo.values/1e6, hi.values/1e6,
                            color=clr, alpha=.15, label="90% CI")
            ax.axvline(fstart, color="gray", lw=1, ls="--", alpha=.6)
            ax.set_title(f"전자상거래 {label}액 (M USD)", fontsize=10)
            ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x,_: f"{x:,.0f}"))
            ax.legend(fontsize=7.5); ax.grid(axis="y", alpha=.3, ls="--")
        plt.tight_layout()
        buf = _io.BytesIO()
        plt.savefig(buf, format="png", dpi=140, bbox_inches="tight")
        plt.close()
        buf.seek(0)
        return buf

    chart_buf = make_chart_image(
        exp_ts, imp_ts, exp_fc_a, imp_fc_a,
        exp_lo, exp_hi, imp_lo, imp_hi, fstart
    )

    # ── PDF 조립 ─────────────────────────────────────────────
    pdf_buf = _io.BytesIO()
    doc = SimpleDocTemplate(
        pdf_buf, pagesize=A4,
        leftMargin=MARGIN, rightMargin=MARGIN,
        topMargin=MARGIN, bottomMargin=22*mm,
    )

    story = []

    # --- 표지 ---
    cover_data = [[Paragraph("한국 전자상거래 수출입\nAI 수요예측 보고서", S_TITLE)]]
    cover_tbl = Table(cover_data, colWidths=[W - 2*MARGIN])
    cover_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#E8EAF6")),
        ("BOX",        (0,0), (-1,-1), 1.5, BLUE),
        ("TOPPADDING", (0,0), (-1,-1), 16),
        ("BOTTOMPADDING", (0,0), (-1,-1), 16),
    ]))
    story += [
        sp(18), cover_tbl, sp(5),
        Paragraph("Korea E-Commerce Trade Demand Forecast Report", S_SUB), sp(3),
        Paragraph(
            f"작성일: {datetime.now().strftime('%Y년 %m월 %d일')} &nbsp;|&nbsp; "
            f"예측 기간: {fstart.strftime('%Y-%m')} ~ {fend.strftime('%Y-%m')} &nbsp;|&nbsp; "
            f"시나리오 가중치: ×{mult}",
            sty("meta", fontSize=8.5, alignment=TA_CENTER, textColor=colors.grey)
        ),
        Paragraph(
            "데이터: 관세청 전자상거래 무역통계 (bandtrass.or.kr) &nbsp;|&nbsp; 모델: SARIMAX + 한국 이커머스 이벤트 회귀",
            sty("meta2", fontSize=8, alignment=TA_CENTER, textColor=colors.grey)
        ),
        sp(10), hr(),
    ]

    # --- 1장: 분석 개요 ---
    story += [
        sp(5),
        Paragraph("1. 분석 개요", S_H1), hr(), sp(2),
        Paragraph(
            "본 보고서는 관세청 전자상거래 무역통계(2014년 1월~2026년 2월, 146개월)를 기반으로 "
            "SARIMAX(1,1,1)(1,1,1,12) 모델과 한국 이커머스 계절 이벤트 외생 회귀 변수를 결합하여 "
            f"{fstart.strftime('%Y년 %m월')}~{fend.strftime('%Y년 %m월')}의 월별 "
            "전자상거래 수출·수입 금액을 예측한 결과를 담고 있습니다.", S_BODY),
        sp(4),
    ]

    # 요약 통계 표
    ev = tbl["수출_예측(달러)"].values
    iv = tbl["수입_예측(달러)"].values
    summary_rows = [
        ["구분", "수출액", "수입액"],
        ["예측 기간 합계",
         f"$ {ev.sum()/1e9:.3f} B", f"$ {iv.sum()/1e9:.3f} B"],
        ["월 평균 예측",
         f"$ {ev.mean()/1e6:.1f} M", f"$ {iv.mean()/1e6:.1f} M"],
        ["최고치 월",
         tbl.loc[ev.argmax(), "연월"], tbl.loc[iv.argmax(), "연월"]],
        ["최고치 금액",
         f"$ {ev.max()/1e6:.1f} M", f"$ {iv.max()/1e6:.1f} M"],
        ["적용 가중치", f"×{mult}", f"×{mult}"],
    ]
    story += [
        tbl_style(summary_rows, [50*mm, 65*mm, 65*mm]), sp(4),
    ]

    # 이벤트 변수 표
    ev_rows = [
        ["이벤트", "적용 기간", "효과"],
        ["설날", "설날 해당 월", "선물·소비재 수요 증가"],
        ["추석", "추석 해당 월", "선물 + 귀성 소비"],
        ["연말 쇼핑", "11~12월", "블프·크리스마스·연말"],
        ["어린이날", "5월", "아동 상품 구매 증가"],
        ["COVID-19", "2020년 1~3월", "구조 변화 더미"],
    ]
    story += [
        Paragraph("적용 이벤트 회귀 변수 (Prophet holidays 대응)",
                  sty("eh", fontName=FONT_B, fontSize=9, textColor=BLUE2)),
        sp(2),
        tbl_style(ev_rows, [40*mm, 45*mm, W-2*MARGIN-85*mm]), sp(5),
        PageBreak(),
    ]

    # --- 2장: 예측 차트 ---
    story += [
        Paragraph("2. 수출입 예측 차트", S_H1), hr(), sp(2),
        Paragraph(
            "아래 차트는 2014~2026년 실제 실적(점)과 예측값(실선), 90% 신뢰구간(음영)을 함께 나타냅니다. "
            "세로 점선은 예측 시작 시점입니다.", S_BODY),
        sp(3),
        Image(chart_buf, width=W-2*MARGIN, height=(W-2*MARGIN)*0.53),
        sp(2),
        Paragraph("[그림] 전자상거래 수출·수입 예측 (SARIMAX + 한국 이커머스 이벤트)", S_CAP),
        sp(4), PageBreak(),
    ]

    # --- 3장: 월별 상세 테이블 ---
    story += [
        Paragraph("3. 월별 상세 예측 결과", S_H1), hr(), sp(2),
        Paragraph(
            f"예측 기간 {fstart.strftime('%Y년 %m월')}~{fend.strftime('%Y년 %m월')} 월별 "
            "수출·수입 예측값과 90% 신뢰구간 하한·상한입니다. 단위: USD.", S_BODY),
        sp(3),
    ]

    # 테이블 헤더
    det_hdr = ["연월", "수출_예측", "수출_하한", "수출_상한",
               "수입_예측", "수입_하한", "수입_상한"]
    det_rows = [det_hdr]
    for _, r in tbl.iterrows():
        row_style = colors.HexColor("#E3F2FD") if "-09" in r["연월"] else \
                    colors.HexColor("#FFEBEE") if "-11" in r["연월"] else None
        det_rows.append([
            r["연월"],
            f"{int(r['수출_예측(달러)']):,}",
            f"{int(r['수출_하한(90%)']):,}",
            f"{int(r['수출_상한(90%)']):,}",
            f"{int(r['수입_예측(달러)']):,}",
            f"{int(r['수입_하한(90%)']):,}",
            f"{int(r['수입_상한(90%)']):,}",
        ])

    det_cw = [18*mm, 30*mm, 27*mm, 27*mm, 30*mm, 27*mm, 27*mm]

    # 기본 스타일 명령 목록 구성
    det_style_cmds = [
        ("BACKGROUND",    (0,0), (-1,0), HEADER),
        ("TEXTCOLOR",     (0,0), (-1,0), colors.white),
        ("FONTNAME",      (0,0), (-1,0), FONT_B),
        ("FONTNAME",      (0,1), (-1,-1), FONT),
        ("FONTSIZE",      (0,0), (-1,-1), 8),
        ("ROWBACKGROUNDS",(0,1), (-1,-1), [colors.white, GREY]),
        ("GRID",          (0,0), (-1,-1), 0.35, colors.HexColor("#BDBDBD")),
        ("TOPPADDING",    (0,0), (-1,-1), 4),
        ("BOTTOMPADDING", (0,0), (-1,-1), 4),
        ("LEFTPADDING",   (0,0), (-1,-1), 5),
        ("RIGHTPADDING",  (0,0), (-1,-1), 5),
        ("ALIGN",         (1,1), (-1,-1), "RIGHT"),
        ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
    ]
    # 추석(9월) 파란색, 광군제(11월) 빨간색 강조 — 행 번호로 직접 추가
    for i, r in enumerate(tbl.itertuples(), 1):
        if "-09" in r.연월:
            det_style_cmds.append(("BACKGROUND", (1,i), (3,i), colors.HexColor("#E3F2FD")))
        if "-11" in r.연월:
            det_style_cmds.append(("BACKGROUND", (4,i), (6,i), colors.HexColor("#FFEBEE")))

    det_tbl = Table(det_rows, colWidths=det_cw, repeatRows=1)
    det_tbl.setStyle(TableStyle(det_style_cmds))

    story += [
        det_tbl, sp(3),
        Paragraph(
            "* 파란 배경: 수출 성수기(추석·9월) | 빨간 배경: 수입 성수기(광군제·11월)",
            S_NOTE),
        sp(8), PageBreak(),
    ]

    # --- 4장: 시사점 ---
    exp_peak_m = tbl.loc[ev.argmax(), "연월"]
    imp_peak_m = tbl.loc[iv.argmax(), "연월"]

    def insight_card(title, body, bg, border_clr):
        d = [[Paragraph(f"<b>{title}</b>",
                        sty("it", fontName=FONT_B, fontSize=9, textColor=border_clr)),
              Paragraph(body, sty("ib", fontSize=8.5))]]
        t = Table(d, colWidths=[32*mm, W-2*MARGIN-32*mm])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,-1), bg),
            ("BOX",        (0,0), (-1,-1), 0.8, border_clr),
            ("TOPPADDING", (0,0), (-1,-1), 7),
            ("BOTTOMPADDING", (0,0), (-1,-1), 7),
            ("LEFTPADDING",   (0,0), (-1,-1), 8),
        ]))
        return t

    story += [
        Paragraph("4. 주요 시사점", S_H1), hr(), sp(3),
        insight_card("수출 성수기",
            f"{exp_peak_m}에 수출 최고점(${ev.max()/1e6:,.1f}M)이 예측됩니다. "
            "해당 시기 수출 물류·재고 확보 전략이 필요합니다.",
            colors.HexColor("#E3F2FD"), BLUE2),
        sp(3),
        insight_card("수입 성수기",
            f"{imp_peak_m}에 수입 최고점(${iv.max()/1e6:,.1f}M)이 예측됩니다. "
            "광군제·블프 해외직구 대응 통관 인력 및 배송 인프라 준비가 권고됩니다.",
            colors.HexColor("#FFEBEE"), RED),
        sp(3),
        insight_card("수출입 구조",
            f"예측 기간 수입액이 수출액의 {iv.sum()/ev.sum():.1f}배 수준으로, "
            "무역 적자 구조가 지속될 전망입니다.",
            colors.HexColor("#F3E5F5"), colors.HexColor("#6A1B9A")),
        sp(3),
        insight_card("시나리오",
            f"현재 적용 가중치 ×{mult}{'(기본값)' if mult==1.0 else ' (What-if 시나리오 적용 중)'}. "
            "환율·물류비·수요 변화에 따른 민감도 분석을 사이드바 슬라이더로 조정하세요.",
            colors.HexColor("#E8F5E9"), colors.HexColor("#2E7D32")),
        sp(10),
        HRFlowable(width="100%", thickness=0.5, color=colors.HexColor("#9E9E9E"),
                   spaceAfter=4, spaceBefore=4),
        sp(2),
        Paragraph(
            f"자동 생성 | {datetime.now().strftime('%Y-%m-%d %H:%M')} | "
            "AI Model: SARIMAX + Prophet-style | Platform: Antigravity Agentic Workflow | "
            "Data: bandtrass.or.kr",
            sty("ft", fontSize=7, alignment=TA_CENTER, textColor=colors.grey)
        ),
    ]

    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
    pdf_buf.seek(0)
    return pdf_buf.read()

# ── 페이지 설정 ─────────────────────────────────────────────
st.set_page_config(
    page_title="전자상거래 AI 예측 대시보드",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS ──────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');
html,body,[class*="css"]{font-family:'Inter',sans-serif;}
.kpi-card{
    background:linear-gradient(135deg,#1E293B,#0F172A);
    border:1px solid #334155;border-radius:14px;
    padding:18px 16px;text-align:center;
    box-shadow:0 4px 20px rgba(0,0,0,.4);}
.kpi-label{color:#94A3B8;font-size:.72rem;font-weight:700;
    text-transform:uppercase;letter-spacing:.09em;margin-bottom:6px;}
.kpi-value{color:#F1F5F9;font-size:1.55rem;font-weight:800;}
.kpi-delta{font-size:.78rem;margin-top:4px;}
.up{color:#34D399;} .down{color:#F87171;}
.sec{background:linear-gradient(90deg,#1E40AF,#7C3AED);
    border-radius:8px;padding:9px 18px;color:#fff;
    font-weight:700;font-size:.95rem;margin:18px 0 10px;}
.footer{background:#1E293B;border:1px solid #334155;
    border-radius:12px;padding:16px;text-align:center;margin-top:24px;}
.badge{display:inline-block;background:#1E40AF;color:#fff;
    padding:3px 11px;border-radius:20px;font-size:.73rem;
    font-weight:600;margin:3px;}
</style>
""", unsafe_allow_html=True)

# ── 상수 ─────────────────────────────────────────────────────
FILE_PATH = r"전자상거래무역.xlsx"

# ── 데이터 로드 ──────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_data():
    xl = pd.ExcelFile(FILE_PATH)
    sheets = xl.sheet_names
    exp_sh = next((s for s in sheets if "수출" in s), sheets[1])
    imp_sh = next((s for s in sheets if "수입" in s), sheets[2])

    def parse(sheet):
        raw = pd.read_excel(FILE_PATH, sheet_name=sheet, header=None)
        hr = None
        for i, v in enumerate(raw.iloc[:, 0]):
            try:
                y = int(str(v).strip())
                if 2000 <= y <= 2030:
                    hr = i; break
            except Exception:
                pass
        d = raw.iloc[hr:].copy().reset_index(drop=True)
        d.columns = ["연","월","전자상거래_금액","전자상거래_건수","전체_금액","전체_건수"]
        for c in d.columns:
            d[c] = pd.to_numeric(
                d[c].astype(str).str.replace(",", "", regex=False).str.strip(),
                errors="coerce")
        d = d.dropna(subset=["연","월","전자상거래_금액"])
        d["연월"] = pd.to_datetime(
            d["연"].astype(int).astype(str) + "-" +
            d["월"].astype(int).astype(str).str.zfill(2),
            format="%Y-%m")
        return d.sort_values("연월").reset_index(drop=True)

    return parse(exp_sh), parse(imp_sh)

def build_exog(idx):
    sl = {2014:1,2015:2,2016:2,2017:1,2018:2,2019:2,2020:1,
          2021:2,2022:2,2023:1,2024:2,2025:1,2026:2}
    ck = {2014:9,2015:9,2016:9,2017:10,2018:9,2019:9,2020:10,
          2021:9,2022:9,2023:9,2024:9,2025:10,2026:9}
    df = pd.DataFrame(index=idx)
    df["설날"]     = [1. if d.year in sl and d.month == sl.get(d.year) else 0. for d in idx]
    df["추석"]     = [1. if d.year in ck and d.month == ck.get(d.year) else 0. for d in idx]
    df["연말쇼핑"] = [1. if d.month in (11,12) else 0. for d in idx]
    df["어린이날"] = [1. if d.month == 5   else 0. for d in idx]
    df["COVID"]   = [1. if d.year == 2020 and d.month <= 3 else 0. for d in idx]
    return df

@st.cache_data(show_spinner=False)
def run_model(s_yr, s_mo, e_yr, e_mo):
    export_df, import_df = load_data()
    FSTART = pd.Timestamp(f"{s_yr}-{s_mo:02d}-01")
    FEND   = pd.Timestamp(f"{e_yr}-{e_mo:02d}-01")
    out = {}
    for key, df in [("export", export_df), ("import", import_df)]:
        ts = df.set_index("연월")["전자상거래_금액"].asfreq("MS")
        fut = pd.date_range(ts.index[0], FEND, freq="MS")
        ex_all = build_exog(fut)
        m = SARIMAX(ts, exog=ex_all.loc[ts.index],
                    order=(1,1,1), seasonal_order=(1,1,1,12),
                    enforce_stationarity=False, enforce_invertibility=False)
        r = m.fit(disp=False, maxiter=200)
        pred_idx = pd.date_range(ts.index[-1] + pd.offsets.MonthBegin(1), FEND, freq="MS")
        fo = r.get_forecast(steps=len(pred_idx), exog=ex_all.loc[pred_idx])
        out[key] = {"ts": ts, "fitted": r.fittedvalues,
                    "fc": fo.predicted_mean, "ci": fo.conf_int(alpha=0.10), "aic": r.aic}
    return out, FSTART, FEND

# ── 헤더 ─────────────────────────────────────────────────────
st.markdown("""
<div style="text-align:center;padding:8px 0 4px">
  <h1 style="background:linear-gradient(90deg,#60A5FA,#A78BFA,#F472B6);
             -webkit-background-clip:text;-webkit-text-fill-color:transparent;
             font-size:2.1rem;font-weight:800;margin:0">
    📦 전자상거래 수출입 AI 예측 대시보드
  </h1>
  <p style="color:#94A3B8;font-size:.88rem;margin-top:5px">
    Korea E-Commerce Trade Forecast · SARIMAX + Korean Event Regression Model
  </p>
</div>
""", unsafe_allow_html=True)
st.divider()

# ── 사이드바 ─────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ 분석 설정")
    st.markdown("---")
    st.markdown("**예측 시작**")
    c1, c2 = st.columns(2)
    with c1: s_yr = st.selectbox("연도", [2026,2025,2024], key="syr")
    with c2: s_mo = st.selectbox("월",   list(range(1,13)), index=2, key="smo")
    st.markdown("**예측 종료**")
    c3, c4 = st.columns(2)
    with c3: e_yr = st.selectbox("연도", [2026,2027], key="eyr")
    with c4: e_mo = st.selectbox("월",   list(range(1,13)), index=11, key="emo")
    st.markdown("---")
    run_btn = st.button("🚀 예측 실행", use_container_width=True, type="primary")
    st.markdown("---")
    st.markdown("### 🔧 What-if 시나리오")
    fx      = st.slider("환율 변동 (%)",      -20, 20, 0, help="환율 상승 → 수출 가격 경쟁력 변화")
    logist  = st.slider("물류비 변동 (%)",    -20, 20, 0, help="물류비 상승 → 순수익 감소")
    demand  = st.slider("해외 수요 변동 (%)", -20, 20, 0, help="글로벌 수요 증감 적용")
    mult    = round(1.0 + (fx + logist + demand) / 300, 4)
    if mult != 1.0:
        color = "#34D399" if mult > 1 else "#F87171"
        st.markdown(f"<p style='color:{color};font-weight:700'>적용 가중치 ×{mult:.4f}</p>",
                    unsafe_allow_html=True)
    st.markdown("---")
    st.caption("📁 관세청 bandtrass.or.kr")

# ── 모델 실행 ────────────────────────────────────────────────
if "results" not in st.session_state or run_btn:
    with st.spinner("🤖 SARIMAX 모델 학습 중... (30~60초 소요)"):
        try:
            res, fstart, fend = run_model(s_yr, s_mo, e_yr, e_mo)
            st.session_state.update({"results": res, "fstart": fstart, "fend": fend})
        except Exception as e:
            st.error(f"오류: {e}"); st.stop()

res    = st.session_state["results"]
fstart = st.session_state["fstart"]
fend   = st.session_state["fend"]

exp_ts = res["export"]["ts"];  exp_fc = res["export"]["fc"];  exp_ci = res["export"]["ci"]
imp_ts = res["import"]["ts"];  imp_fc = res["import"]["fc"];  imp_ci = res["import"]["ci"]

exp_fc_a = exp_fc * mult;  exp_lo = exp_ci.iloc[:,0]*mult;  exp_hi = exp_ci.iloc[:,1]*mult
imp_fc_a = imp_fc * mult;  imp_lo = imp_ci.iloc[:,0]*mult;  imp_hi = imp_ci.iloc[:,1]*mult

# ── KPI 카드 ─────────────────────────────────────────────────
exp_latest = exp_ts.iloc[-1];  exp_yoy = (exp_ts.iloc[-1]/exp_ts.iloc[-13]-1)*100
imp_latest = imp_ts.iloc[-1];  imp_yoy = (imp_ts.iloc[-1]/imp_ts.iloc[-13]-1)*100

k1,k2,k3,k4 = st.columns(4)
kpis = [
    (k1, "최근 수출액",  f"{exp_latest/1e6:,.1f}M USD", f"YoY {exp_yoy:+.1f}%", exp_yoy>0),
    (k2, "최근 수입액",  f"{imp_latest/1e6:,.1f}M USD", f"YoY {imp_yoy:+.1f}%", imp_yoy>0),
    (k3, "예측 수출 합계", f"{exp_fc_a.sum()/1e9:.2f}B USD",
         f"{fstart.strftime('%Y.%m')}~{fend.strftime('%m')}", True),
    (k4, "예측 수입 합계", f"{imp_fc_a.sum()/1e9:.2f}B USD",
         f"{fstart.strftime('%Y.%m')}~{fend.strftime('%m')}", True),
]
for col, lbl, val, delta, up in kpis:
    cls = "up" if up else "down"
    col.markdown(f"""<div class="kpi-card">
        <div class="kpi-label">{lbl}</div>
        <div class="kpi-value">{val}</div>
        <div class="kpi-delta {cls}">{delta}</div>
    </div>""", unsafe_allow_html=True)

# ── 메인 차트 ────────────────────────────────────────────────
st.markdown('<div class="sec">📈 수출입 추세 및 예측</div>', unsafe_allow_html=True)
fig = make_subplots(rows=2, cols=1,
    subplot_titles=("전자상거래 수출액 (M USD)","전자상거래 수입액 (M USD)"),
    vertical_spacing=0.12)

specs = [
    (exp_ts, res["export"]["fitted"], exp_fc_a, exp_lo, exp_hi, "수출", "#60A5FA", "rgba(96,165,250,.15)", 1),
    (imp_ts, res["import"]["fitted"], imp_fc_a, imp_lo, imp_hi, "수입", "#F87171", "rgba(248,113,113,.15)", 2),
]
for ts, fitted, fc, lo, hi, nm, clr, fill, row in specs:
    fig.add_trace(go.Scatter(x=ts.index, y=ts.values/1e6, mode="markers",
        name=f"{nm} 실적", marker=dict(color=clr, size=4, opacity=.65),
        hovertemplate="%{x|%Y-%m}<br>실적: %{y:,.2f}M<extra></extra>"), row=row, col=1)
    fig.add_trace(go.Scatter(x=ts.index, y=fitted.values/1e6, mode="lines",
        name=f"{nm} 적합값", line=dict(color=clr, width=1, dash="dot"), opacity=.45,
        hovertemplate="%{x|%Y-%m}<br>적합: %{y:,.2f}M<extra></extra>"), row=row, col=1)
    fig.add_trace(go.Scatter(
        x=list(fc.index)+list(fc.index[::-1]),
        y=list(hi.values/1e6)+list(lo.values[::-1]/1e6),
        fill="toself", fillcolor=fill, line=dict(color="rgba(0,0,0,0)"),
        name=f"{nm} 90% CI", hoverinfo="skip"), row=row, col=1)
    fig.add_trace(go.Scatter(x=fc.index, y=fc.values/1e6, mode="lines+markers",
        name=f"{nm} 예측", line=dict(color=clr, width=2.5), marker=dict(size=7),
        hovertemplate="%{x|%Y-%m}<br>예측: %{y:,.2f}M<extra></extra>"), row=row, col=1)
    fig.add_vline(x=fstart, line_dash="dash", line_color="#64748B", line_width=1.2, row=row, col=1)

fig.update_layout(template="plotly_dark", height=680, hovermode="x unified",
    legend=dict(orientation="h", y=-0.06, x=.5, xanchor="center"),
    margin=dict(l=10,r=10,t=45,b=10))
fig.update_yaxes(tickformat=",.0f", title_text="M USD")
st.plotly_chart(fig, use_container_width=True)

# ── 시나리오 분석 ─────────────────────────────────────────────
if mult != 1.0:
    st.markdown('<div class="sec">🔀 What-if 시나리오 비교</div>', unsafe_allow_html=True)
    months = [d.strftime("%Y-%m") for d in exp_fc.index]
    fig2 = go.Figure()
    fig2.add_trace(go.Bar(name="수출 기본",    x=months, y=exp_fc.values/1e6,   marker_color="#3B82F6", opacity=.55))
    fig2.add_trace(go.Bar(name="수출 시나리오",x=months, y=exp_fc_a.values/1e6, marker_color="#93C5FD"))
    fig2.add_trace(go.Bar(name="수입 기본",    x=months, y=imp_fc.values/1e6,   marker_color="#EF4444", opacity=.55))
    fig2.add_trace(go.Bar(name="수입 시나리오",x=months, y=imp_fc_a.values/1e6, marker_color="#FCA5A5"))
    fig2.update_layout(template="plotly_dark", barmode="group", height=370,
        title=f"기본 예측 vs 시나리오 (가중치 ×{mult})",
        yaxis_title="M USD", hovermode="x unified",
        margin=dict(l=10,r=10,t=50,b=10))
    st.plotly_chart(fig2, use_container_width=True)

# ── 예측 테이블 ──────────────────────────────────────────────
st.markdown('<div class="sec">📋 월별 상세 예측 결과</div>', unsafe_allow_html=True)
tbl = pd.DataFrame({
    "연월":          exp_fc_a.index.strftime("%Y-%m"),
    "수출_예측(달러)": exp_fc_a.values.round(0).astype(int),
    "수출_하한(90%)": exp_lo.values.round(0).astype(int),
    "수출_상한(90%)": exp_hi.values.round(0).astype(int),
    "수입_예측(달러)": imp_fc_a.values.round(0).astype(int),
    "수입_하한(90%)": imp_lo.values.round(0).astype(int),
    "수입_상한(90%)": imp_hi.values.round(0).astype(int),
})
num_cols = [c for c in tbl.columns if c != "연월"]
st.dataframe(
    tbl.style
       .format({c: "{:,}" for c in num_cols})
       .background_gradient(subset=["수출_예측(달러)"], cmap="Blues")
       .background_gradient(subset=["수입_예측(달러)"], cmap="Reds"),
    use_container_width=True, hide_index=True)

# ── 다운로드 ─────────────────────────────────────────────────
st.markdown('<div class="sec">⬇️ 데이터 다운로드</div>', unsafe_allow_html=True)
dc1, dc2, dc3 = st.columns(3)
with dc1:
    csv = tbl.to_csv(index=False, encoding="utf-8-sig")
    st.download_button("📥 CSV 다운로드", data=csv.encode("utf-8-sig"),
        file_name="2026_ecommerce_forecast.csv", mime="text/csv",
        use_container_width=True)
with dc2:
    buf = _io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
        tbl.to_excel(w, index=False, sheet_name="예측결과")
    buf.seek(0)
    st.download_button("📥 엑셀 다운로드", data=buf.read(),
        file_name="2026_ecommerce_forecast.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True)
with dc3:
    if not REPORTLAB_OK:
        st.warning("PDF 생성을 위한 라이브러리 오류")
    else:
        pdf_bytes = generate_pdf_report(
            tbl, fstart, fend, mult,
            exp_ts, imp_ts,
            exp_fc_a, imp_fc_a,
            exp_lo, exp_hi,
            imp_lo, imp_hi,
        )
        pdf_fn = f"ecommerce_forecast_{fstart.strftime('%Y%m')}_{fend.strftime('%Y%m')}.pdf"
        st.download_button(
            label=f"📥 PDF 다운로드 ({len(pdf_bytes)//1024} KB)",
            data=pdf_bytes,
            file_name=pdf_fn,
            mime="application/pdf",
            use_container_width=True,
            type="secondary",
        )

# ── 푸터 ─────────────────────────────────────────────────────
st.markdown("""
<div class="footer">
  <p style="color:#94A3B8;font-size:.82rem;margin:0 0 8px">⚡ 시스템 개요</p>
  <span class="badge">AI Model: Prophet / SARIMAX</span>
  <span class="badge">Platform: Antigravity Agentic Workflow</span>
  <span class="badge">Data: bandtrass.or.kr</span>
  <span class="badge">Framework: Streamlit 1.56</span>
  <span class="badge">Visualization: Plotly</span>
  <p style="color:#475569;font-size:.72rem;margin:10px 0 0">
    © 2026 한국 전자상거래 수출입 AI 예측 대시보드 · 관세청 공공데이터 기반
  </p>
</div>
""", unsafe_allow_html=True)
