# -*- coding: utf-8 -*-
import sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import pandas as pd
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    Image, PageBreak, HRFlowable
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ── 한글 폰트 등록 ──────────────────────────────────────────
if os.path.exists(r'C:\Windows\Fonts\malgun.ttf'):
    FONT_PATH = r'C:\Windows\Fonts\malgun.ttf'
    FONT_BOLD  = r'C:\Windows\Fonts\malgunbd.ttf'
else:
    # Linux (Google Cloud Run) fallback using fonts-nanum
    FONT_PATH = '/usr/share/fonts/truetype/nanum/NanumGothic.ttf'
    FONT_BOLD  = '/usr/share/fonts/truetype/nanum/NanumGothicBold.ttf'

pdfmetrics.registerFont(TTFont('Malgun', FONT_PATH))
pdfmetrics.registerFont(TTFont('MalgunBd', FONT_BOLD))

W, H = A4
MARGIN = 20*mm

# ── 스타일 ───────────────────────────────────────────────────
base = getSampleStyleSheet()

def sty(name, parent='Normal', **kw):
    defaults = {'fontName': 'Malgun', 'fontSize': 10, 'leading': 14}
    defaults.update(kw)
    return ParagraphStyle(name, parent=base[parent], **defaults)

S_TITLE   = sty('title',   fontSize=22, fontName='MalgunBd', alignment=TA_CENTER, leading=30, textColor=colors.HexColor('#1A237E'))
S_SUB     = sty('sub',     fontSize=13, fontName='MalgunBd', alignment=TA_CENTER, textColor=colors.HexColor('#283593'), leading=18)
S_META    = sty('meta',    fontSize=9,  alignment=TA_CENTER, textColor=colors.grey)
S_H1      = sty('h1',      fontSize=14, fontName='MalgunBd', textColor=colors.HexColor('#1565C0'), leading=20, spaceAfter=4)
S_H2      = sty('h2',      fontSize=11, fontName='MalgunBd', textColor=colors.HexColor('#283593'), leading=16, spaceAfter=2)
S_BODY    = sty('body',    fontSize=9,  leading=14, textColor=colors.HexColor('#212121'))
S_CAPTION = sty('cap',     fontSize=8,  alignment=TA_CENTER, textColor=colors.grey, leading=12)
S_NOTE    = sty('note',    fontSize=8,  textColor=colors.HexColor('#555555'), leading=12)

BLUE_DARK  = colors.HexColor('#1565C0')
BLUE_LIGHT = colors.HexColor('#E3F2FD')
RED_DARK   = colors.HexColor('#B71C1C')
RED_LIGHT  = colors.HexColor('#FFEBEE')
GREY_BG    = colors.HexColor('#F5F5F5')
HEADER_BG  = colors.HexColor('#1A237E')

# ── 이미지 유틸 ─────────────────────────────────────────────
def img(path, width_mm, caption=None):
    items = []
    if os.path.exists(path):
        w = width_mm * mm
        items.append(Image(path, width=w, height=w*0.55))
    if caption:
        items.append(Paragraph(caption, S_CAPTION))
    return items

# ── 구분선 ────────────────────────────────────────────────────
def hr(color=BLUE_DARK, thickness=0.8):
    return HRFlowable(width='100%', thickness=thickness, color=color, spaceAfter=4, spaceBefore=4)

# ── CSV 로드 ─────────────────────────────────────────────────
CSV = '2026_ecommerce_forecast.csv'
df = pd.read_csv(CSV, encoding='utf-8-sig')

# ── 표지 헤더 박스 테이블 ────────────────────────────────────
def cover_box():
    data = [[Paragraph('한국 전자상거래 수출입 동향 및 수요 예측 보고서', S_TITLE)]]
    t = Table(data, colWidths=[W - 2*MARGIN])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#E8EAF6')),
        ('BOX',        (0,0), (-1,-1), 1.5, BLUE_DARK),
        ('TOPPADDING', (0,0), (-1,-1), 14),
        ('BOTTOMPADDING', (0,0), (-1,-1), 14),
        ('LEFTPADDING',  (0,0), (-1,-1), 10),
    ]))
    return t

# ── 예측 결과 테이블 ─────────────────────────────────────────
def forecast_table(df):
    header = ['연월', '수출액 예측\n(달러)', '수출 하한\n(90%)', '수출 상한\n(90%)',
              '수입액 예측\n(달러)', '수입 하한\n(90%)', '수입 상한\n(90%)']
    rows = [header]
    for _, r in df.iterrows():
        rows.append([
            Paragraph(r['연월'], sty('tc', fontSize=8, alignment=TA_CENTER)),
            Paragraph(f"{int(r['수출액_예측']):,}", sty('tc', fontSize=8, alignment=TA_RIGHT, textColor=BLUE_DARK, fontName='MalgunBd')),
            Paragraph(f"{int(r['수출액_하한(90%)']):,}", sty('tc', fontSize=7.5, alignment=TA_RIGHT, textColor=colors.HexColor('#1976D2'))),
            Paragraph(f"{int(r['수출액_상한(90%)']):,}", sty('tc', fontSize=7.5, alignment=TA_RIGHT, textColor=colors.HexColor('#1976D2'))),
            Paragraph(f"{int(r['수입액_예측']):,}", sty('tc', fontSize=8, alignment=TA_RIGHT, textColor=RED_DARK, fontName='MalgunBd')),
            Paragraph(f"{int(r['수입액_하한(90%)']):,}", sty('tc', fontSize=7.5, alignment=TA_RIGHT, textColor=colors.HexColor('#C62828'))),
            Paragraph(f"{int(r['수입액_상한(90%)']):,}", sty('tc', fontSize=7.5, alignment=TA_RIGHT, textColor=colors.HexColor('#C62828'))),
        ])

    cw = [18*mm, 32*mm, 29*mm, 29*mm, 32*mm, 29*mm, 29*mm]
    t = Table(rows, colWidths=cw, repeatRows=1)
    style = TableStyle([
        ('BACKGROUND',   (0,0), (-1,0), HEADER_BG),
        ('TEXTCOLOR',    (0,0), (-1,0), colors.white),
        ('FONTNAME',     (0,0), (-1,0), 'MalgunBd'),
        ('FONTSIZE',     (0,0), (-1,0), 8),
        ('ALIGN',        (0,0), (-1,0), 'CENTER'),
        ('VALIGN',       (0,0), (-1,-1), 'MIDDLE'),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, GREY_BG]),
        ('GRID',         (0,0), (-1,-1), 0.4, colors.HexColor('#BDBDBD')),
        ('TOPPADDING',   (0,0), (-1,-1), 4),
        ('BOTTOMPADDING',(0,0), (-1,-1), 4),
        ('LEFTPADDING',  (0,0), (-1,-1), 4),
        ('RIGHTPADDING', (0,0), (-1,-1), 4),
    ])
    # 추석(9월) 행 강조
    for i, row in df.iterrows():
        if '-09' in row['연월']:
            t_row = i + 1
            style.add('BACKGROUND', (1, t_row), (3, t_row), BLUE_LIGHT)
        if '-11' in row['연월']:
            t_row = i + 1
            style.add('BACKGROUND', (4, t_row), (6, t_row), RED_LIGHT)
    t.setStyle(style)
    return t

# ── 인사이트 카드 ────────────────────────────────────────────
def insight_card(title, body, bg, border):
    data = [[
        Paragraph(f'<b>{title}</b>', sty('it', fontSize=9, fontName='MalgunBd', textColor=border)),
        Paragraph(body, sty('ib', fontSize=8.5, textColor=colors.HexColor('#212121')))
    ]]
    t = Table(data, colWidths=[30*mm, W - 2*MARGIN - 30*mm])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), bg),
        ('BOX',        (0,0), (-1,-1), 1, border),
        ('VALIGN',     (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 7),
        ('BOTTOMPADDING', (0,0), (-1,-1), 7),
        ('LEFTPADDING',   (0,0), (-1,-1), 8),
        ('RIGHTPADDING',  (0,0), (-1,-1), 8),
    ]))
    return t

# ── 요약 통계 카드 ────────────────────────────────────────────
def summary_stats(df):
    exp_vals = df['수출액_예측'].values
    imp_vals = df['수입액_예측'].values
    rows = [
        ['항목', '수출액', '수입액'],
        ['예측 기간 합계', f"{exp_vals.sum()/1e9:.2f} 십억 달러", f"{imp_vals.sum()/1e9:.2f} 십억 달러"],
        ['월평균 예측값', f"{exp_vals.mean()/1e6:.1f} 백만 달러", f"{imp_vals.mean()/1e6:.1f} 백만 달러"],
        ['최고치 월', df.loc[exp_vals.argmax(), '연월'], df.loc[imp_vals.argmax(), '연월']],
        ['최고치 금액', f"{exp_vals.max()/1e6:.1f} 백만 달러", f"{imp_vals.max()/1e6:.1f} 백만 달러"],
        ['최저치 월', df.loc[exp_vals.argmin(), '연월'], df.loc[imp_vals.argmin(), '연월']],
    ]
    cw = [45*mm, 65*mm, 65*mm]
    t = Table(rows, colWidths=cw)
    t.setStyle(TableStyle([
        ('BACKGROUND',   (0,0), (-1,0), HEADER_BG),
        ('TEXTCOLOR',    (0,0), (-1,0), colors.white),
        ('FONTNAME',     (0,0), (-1,0), 'MalgunBd'),
        ('ALIGN',        (1,0), (-1,-1), 'CENTER'),
        ('FONTNAME',     (0,1), (0,-1), 'MalgunBd'),
        ('FONTSIZE',     (0,0), (-1,-1), 9),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, GREY_BG]),
        ('GRID',         (0,0), (-1,-1), 0.4, colors.HexColor('#BDBDBD')),
        ('TOPPADDING',   (0,0), (-1,-1), 5),
        ('BOTTOMPADDING',(0,0), (-1,-1), 5),
        ('LEFTPADDING',  (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
    ]))
    return t

# ── 페이지 번호 ───────────────────────────────────────────────
def on_page(canvas, doc):
    canvas.saveState()
    canvas.setFont('Malgun', 8)
    canvas.setFillColor(colors.grey)
    canvas.drawRightString(W - MARGIN, 12*mm, f'- {doc.page} -')
    canvas.drawString(MARGIN, 12*mm, '한국 전자상거래 수출입 수요예측 보고서 | 2026')
    canvas.setStrokeColor(colors.HexColor('#BDBDBD'))
    canvas.line(MARGIN, 14*mm, W - MARGIN, 14*mm)
    canvas.restoreState()

# ═══════════════════════════════════════════════════════════════
# 보고서 조립
# ═══════════════════════════════════════════════════════════════
OUTPUT_PDF = '2026_ecommerce_forecast_report.pdf'
doc = SimpleDocTemplate(
    OUTPUT_PDF, pagesize=A4,
    leftMargin=MARGIN, rightMargin=MARGIN,
    topMargin=MARGIN, bottomMargin=25*mm,
)

story = []
sp = lambda n=6: Spacer(1, n*mm)

# ── 표지 ─────────────────────────────────────────────────────
story += [
    sp(20),
    cover_box(),
    sp(6),
    Paragraph('Korea E-Commerce Trade Trend & Demand Forecast Report', S_SUB),
    sp(4),
    Paragraph(f'작성일: {datetime.now().strftime("%Y년 %m월 %d일")} &nbsp;|&nbsp; 예측 대상: 2026년 3월 ~ 12월', S_META),
    Paragraph('데이터 출처: 관세청 전자상거래 무역통계 (bandtrass.or.kr) &nbsp;|&nbsp; 모델: SARIMAX + 이벤트 회귀', S_META),
    sp(10),
    hr(),
    sp(6),
]

# 표지 EDA 이미지
story += img('eda_trend.png', 165, '[그림 1] 한국 전자상거래 수출입 월별 추세 (2014~2026.02)')
story += [PageBreak()]

# ── 1장: 개요 ───────────────────────────────────────────────
story += [
    Paragraph('1. 분석 개요', S_H1), hr(),
    sp(2),
    Paragraph('1.1 분석 목적', S_H2),
    Paragraph(
        '본 보고서는 관세청 전자상거래 무역통계 데이터(2014년 1월~2026년 2월, 총 146개월)를 활용하여 '
        '2026년 3월부터 12월까지의 월별 전자상거래 수출액 및 수입액을 예측합니다. '
        '한국 이커머스 시장의 계절적 특성(설날, 추석, 연말 쇼핑 시즌, 광군제 등)을 모델에 반영하여 '
        '정책 입안 및 비즈니스 의사결정에 활용할 수 있는 근거 자료를 제공합니다.',
        S_BODY),
    sp(4),
    Paragraph('1.2 데이터 현황', S_H2),
]

meta_rows = [
    ['구분', '내용'],
    ['데이터 출처', '관세청 전자상거래 무역통계 (bandtrass.or.kr)'],
    ['데이터 기간', '2014년 1월 ~ 2026년 2월 (146개월)'],
    ['주요 변수', '전자상거래 수출·수입 금액(달러), 건수'],
    ['예측 대상', '2026년 3월 ~ 12월 (10개월)'],
    ['사용 모델', 'SARIMAX(1,1,1)(1,1,1,12) + 한국 이커머스 이벤트 외생 회귀 변수'],
    ['신뢰 구간', '90% 예측 구간 (yhat_lower ~ yhat_upper)'],
]
mt = Table(meta_rows, colWidths=[45*mm, W - 2*MARGIN - 45*mm])
mt.setStyle(TableStyle([
    ('BACKGROUND',   (0,0), (-1,0), HEADER_BG),
    ('TEXTCOLOR',    (0,0), (-1,0), colors.white),
    ('FONTNAME',     (0,0), (-1,0), 'MalgunBd'),
    ('FONTNAME',     (0,1), (0,-1), 'MalgunBd'),
    ('FONTSIZE',     (0,0), (-1,-1), 9),
    ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, GREY_BG]),
    ('GRID',         (0,0), (-1,-1), 0.4, colors.HexColor('#BDBDBD')),
    ('TOPPADDING',   (0,0), (-1,-1), 5),
    ('BOTTOMPADDING',(0,0), (-1,-1), 5),
    ('LEFTPADDING',  (0,0), (-1,-1), 6),
]))
story += [mt, sp(4)]

story += [
    Paragraph('1.3 모델링 방법론', S_H2),
    Paragraph(
        'SARIMAX(계절성 자기회귀 누적이동평균 + 외생변수) 모델을 적용하였으며, '
        '아래 한국 이커머스 특성 변수를 외생 회귀 항으로 포함하였습니다.',
        S_BODY),
    sp(2),
]

method_rows = [
    ['이벤트 변수', '적용 기간', '효과'],
    ['설날 시즌', '음력 설날 해당 월', '선물·소비재 구매 증가'],
    ['추석 시즌', '음력 추석 해당 월', '선물 수요 + 귀성 관련 소비'],
    ['연말 쇼핑 시즌', '11~12월', '블랙프라이데이·크리스마스·연말 소비 급증'],
    ['어린이날 시즌', '5월', '아동 관련 상품 구매 증가'],
    ['COVID-19 충격', '2020년 1~3월', '초기 물류 차질 및 소비 위축 구조 변화'],
]
mt2 = Table(method_rows, colWidths=[45*mm, 45*mm, W - 2*MARGIN - 90*mm])
mt2.setStyle(TableStyle([
    ('BACKGROUND',   (0,0), (-1,0), colors.HexColor('#283593')),
    ('TEXTCOLOR',    (0,0), (-1,0), colors.white),
    ('FONTNAME',     (0,0), (-1,0), 'MalgunBd'),
    ('FONTSIZE',     (0,0), (-1,-1), 8.5),
    ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, GREY_BG]),
    ('GRID',         (0,0), (-1,-1), 0.4, colors.HexColor('#BDBDBD')),
    ('TOPPADDING',   (0,0), (-1,-1), 4),
    ('BOTTOMPADDING',(0,0), (-1,-1), 4),
    ('LEFTPADDING',  (0,0), (-1,-1), 5),
]))
story += [mt2, sp(4), PageBreak()]

# ── 2장: EDA ─────────────────────────────────────────────────
story += [
    Paragraph('2. 탐색적 데이터 분석 (EDA)', S_H1), hr(), sp(2),
    Paragraph('2.1 수출입 추세', S_H2),
    Paragraph(
        '2014년~2026년 초까지의 전자상거래 수출입 추세를 분석한 결과, '
        '수출은 2020년 이후 COVID-19 반사 효과와 온라인 소비 확대로 뚜렷한 성장세를 보였으며, '
        '수입은 해외직구 활성화로 더욱 가파른 성장을 기록하였습니다.', S_BODY),
    sp(3),
]
story += img('eda_trend.png', 165, '[그림 2] 전자상거래 수출입 추세 및 YoY 성장률')
story += [sp(5)]

story += [
    Paragraph('2.2 시계열 분해 – 수출 (추세·계절성·잔차)', S_H2), sp(2),
]
story += img('components_export.png', 165, '[그림 3] 수출 시계열 분해 (Trend / Seasonality / Residual)')
story += [sp(3), PageBreak()]

story += [
    Paragraph('2.3 시계열 분해 – 수입 (추세·계절성·잔차)', S_H2), sp(2),
]
story += img('components_import.png', 165, '[그림 4] 수입 시계열 분해 (Trend / Seasonality / Residual)')
story += [sp(4), PageBreak()]

# ── 3장: 예측 결과 ───────────────────────────────────────────
story += [
    Paragraph('3. 예측 결과', S_H1), hr(), sp(2),
    Paragraph('3.1 예측 차트 (2026년 3~12월)', S_H2), sp(2),
]
story += img('forecast_result.png', 165, '[그림 5] 전자상거래 수출·수입 예측 결과 및 90% 신뢰구간')
story += [sp(3), PageBreak()]

story += [
    Paragraph('3.2 월별 예측 요약 차트', S_H2), sp(2),
]
story += img('forecast_summary_bar.png', 165, '[그림 6] 2026년 월별 수출·수입 예측 바 차트 (오차 막대: 90% CI)')
story += [sp(4)]

story += [
    Paragraph('3.3 예측 요약 통계', S_H2), sp(2),
    summary_stats(df),
    sp(5), PageBreak(),
]

# ── 4장: 상세 예측 테이블 ────────────────────────────────────
story += [
    Paragraph('4. 월별 상세 예측 결과 테이블', S_H1), hr(), sp(2),
    Paragraph(
        '아래 표는 2026년 3월부터 12월까지 월별 전자상거래 수출액과 수입액의 예측값과 '
        '90% 신뢰구간 하한·상한을 나타냅니다. 단위는 달러(USD)입니다.', S_BODY),
    sp(3),
    forecast_table(df),
    sp(4),
    Paragraph(
        '<b>* 강조 표시:</b> 파란색 배경 = 수출 성수기(추석 9월), 빨간색 배경 = 수입 성수기(광군제·블프 11월)',
        S_NOTE),
    sp(8),
]

# ── 5장: 시사점 ──────────────────────────────────────────────
story += [
    Paragraph('5. 주요 시사점 및 결론', S_H1), hr(), sp(3),
    insight_card(
        '수출 성수기',
        '2026년 9월(추석)에 수출액이 최고점(약 1억 7,347만 달러)을 기록할 것으로 예측됩니다. '
        '이 시기를 겨냥한 수출 물류 및 재고 확보 전략이 필요합니다.',
        BLUE_LIGHT, BLUE_DARK,
    ),
    sp(3),
    insight_card(
        '수입 성수기',
        '2026년 11월(광군제·블랙프라이데이)에 수입액 최고점(약 2억 7,937만 달러)이 예상됩니다. '
        '해외직구 수요 대응을 위한 통관 인력 및 배송 인프라 준비가 권고됩니다.',
        RED_LIGHT, RED_DARK,
    ),
    sp(3),
    insight_card(
        '수출입 구조',
        '예측 기간 전반에 걸쳐 수입액이 수출액의 약 1.7배 수준을 유지하며, '
        '무역 적자 구조가 지속될 것으로 전망됩니다. '
        '수출 경쟁력 강화를 위한 플랫폼·브랜드 전략이 중요합니다.',
        colors.HexColor('#F3E5F5'), colors.HexColor('#6A1B9A'),
    ),
    sp(3),
    insight_card(
        '모델 신뢰도',
        '수출 모델 AIC 기준 안정적인 적합 수준을 달성하였으나, '
        '환율 변동·글로벌 공급망 이슈 등 외부 변수에 의한 예측 오차 가능성이 존재합니다. '
        '분기별 재학습을 통한 예측 업데이트를 권장합니다.',
        colors.HexColor('#E8F5E9'), colors.HexColor('#2E7D32'),
    ),
    sp(8),
    hr(colors.HexColor('#9E9E9E'), 0.5),
    sp(2),
    Paragraph(
        f'본 보고서는 자동 생성되었습니다 | 생성일시: {datetime.now().strftime("%Y-%m-%d %H:%M")} | '
        '모델: SARIMAX + 한국 이커머스 이벤트 회귀 | 데이터: bandtrass.or.kr',
        sty('footer', fontSize=7.5, alignment=TA_CENTER, textColor=colors.grey)
    ),
]

# ── 빌드 ─────────────────────────────────────────────────────
doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
print(f"PDF 보고서 저장 완료: {OUTPUT_PDF}")
