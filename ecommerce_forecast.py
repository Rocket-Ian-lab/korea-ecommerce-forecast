# -*- coding: utf-8 -*-
"""
한국 전자상거래 수출입 통계 시계열 예측
- Facebook Prophet 스타일 분석 (추세 + 계절성 + 한국 이커머스 이벤트)
- 백엔드: statsmodels SARIMA + 이벤트 회귀 (Python 3.14 호환)
데이터 출처: bandtrass.or.kr
"""

import sys, io, warnings
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
warnings.filterwarnings('ignore')

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import matplotlib.patches as mpatches
from matplotlib.gridspec import GridSpec
import platform
from statsmodels.tsa.statespace.sarimax import SARIMAX
from statsmodels.tsa.seasonal import seasonal_decompose

# ─────────────────────────────────────────────
# 0. 한글 폰트 설정
# ─────────────────────────────────────────────
if platform.system() == 'Windows':
    plt.rcParams['font.family'] = 'Malgun Gothic'
else:
    plt.rcParams['font.family'] = 'AppleGothic'
plt.rcParams['axes.unicode_minus'] = False

FILE_PATH = r'C:\Users\arron\Downloads\전자상거래무역.xlsx'

# ═══════════════════════════════════════════════════════════════
# 1. 데이터 로드 및 전처리
# ═══════════════════════════════════════════════════════════════
def load_sheet(sheet_name: str) -> pd.DataFrame:
    """
    시트를 로드하고 불필요한 상단 행을 자동 감지·제거 후 전처리합니다.
    - '연월' 컬럼: datetime 형식
    - 금액 컬럼: 콤마 제거 후 숫자형
    """
    raw = pd.read_excel(FILE_PATH, sheet_name=sheet_name, header=None)

    # 헤더 자동 감지: 첫 컬럼이 2000~2030 범위 연도인 첫 행
    header_row = None
    for i, val in enumerate(raw.iloc[:, 0]):
        try:
            y = int(str(val).strip())
            if 2000 <= y <= 2030:
                header_row = i
                break
        except (ValueError, TypeError):
            continue

    if header_row is None:
        raise ValueError(f"'{sheet_name}' 시트에서 데이터 시작 행을 찾을 수 없습니다.")

    data = raw.iloc[header_row:].copy().reset_index(drop=True)
    data.columns = ['연', '월', '전자상거래_금액', '전자상거래_건수', '전체_금액', '전체_건수']

    for col in data.columns:
        data[col] = pd.to_numeric(
            data[col].astype(str).str.replace(',', '', regex=False).str.strip(),
            errors='coerce'
        )

    data = data.dropna(subset=['연', '월', '전자상거래_금액'])
    data['연월'] = pd.to_datetime(
        data['연'].astype(int).astype(str) + '-' +
        data['월'].astype(int).astype(str).str.zfill(2),
        format='%Y-%m'
    )
    return data.sort_values('연월').reset_index(drop=True)


print("[1/5] 데이터 로딩 및 전처리...")
xl = pd.ExcelFile(FILE_PATH)
sheets = xl.sheet_names
print(f"  시트: {sheets}")

# 시트명 기반으로 수출/수입 구분
export_sheet = next((s for s in sheets if '수출' in s), sheets[1])
import_sheet  = next((s for s in sheets if '수입' in s), sheets[2])

export_df = load_sheet(export_sheet)
import_df  = load_sheet(import_sheet)

print(f"  수출: {export_df['연월'].min().strftime('%Y-%m')} ~ {export_df['연월'].max().strftime('%Y-%m')} ({len(export_df)}개월)")
print(f"  수입: {import_df['연월'].min().strftime('%Y-%m')} ~ {import_df['연월'].max().strftime('%Y-%m')} ({len(import_df)}개월)")

# ═══════════════════════════════════════════════════════════════
# 2. EDA – 수출입 추세 시각화
# ═══════════════════════════════════════════════════════════════
print("\n[2/5] EDA 차트 생성...")

def add_yoy_growth(ax, df, col, color, label):
    """전년 동월 대비 증감률(%) 표시"""
    yoy = df.set_index('연월')[col].pct_change(12) * 100
    ax2 = ax.twinx()
    ax2.bar(df['연월'], yoy, color=color, alpha=0.18, label='YoY 성장률(%)')
    ax2.set_ylabel('YoY 성장률 (%)', fontsize=9, color=color)
    ax2.tick_params(axis='y', labelcolor=color, labelsize=8)
    ax2.axhline(0, color=color, linewidth=0.5, linestyle='--', alpha=0.4)
    return ax2

fig = plt.figure(figsize=(18, 12))
gs = GridSpec(3, 2, figure=fig, hspace=0.45, wspace=0.35)

ax_exp = fig.add_subplot(gs[0, :])
ax_imp = fig.add_subplot(gs[1, :])
ax_e_cnt = fig.add_subplot(gs[2, 0])
ax_i_cnt = fig.add_subplot(gs[2, 1])

fig.suptitle('한국 전자상거래 수출입 통계 EDA\n(Korea E-Commerce Trade Statistics)',
             fontsize=15, fontweight='bold')

# 수출 금액
ax_exp.plot(export_df['연월'], export_df['전자상거래_금액'] / 1e6,
            color='#1565C0', linewidth=2, label='전자상거래 수출액')
ax_exp.fill_between(export_df['연월'], export_df['전자상거래_금액'] / 1e6, alpha=0.12, color='#1565C0')
ax_exp.set_title('전자상거래 수출액 추세 (백만 달러)', fontsize=12)
ax_exp.set_ylabel('백만 달러')
ax_exp.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'{x:,.0f}'))
ax_exp.legend(loc='upper left', fontsize=9)
ax_exp.grid(axis='y', alpha=0.3, linestyle='--')
ax2_exp = add_yoy_growth(ax_exp, export_df, '전자상거래_금액', '#90CAF9', '수출 YoY')

# 이벤트 마커
events = {'2020-01': 'COVID-19 시작', '2022-03': '러-우 전쟁'}
for ds, ev in events.items():
    dt = pd.Timestamp(ds)
    if export_df['연월'].min() <= dt <= export_df['연월'].max():
        ax_exp.axvline(dt, color='red', linestyle=':', alpha=0.7, linewidth=1.3)
        ax_exp.text(dt, ax_exp.get_ylim()[0], f' {ev}', fontsize=7.5,
                    color='red', rotation=90, va='bottom')

# 수입 금액
ax_imp.plot(import_df['연월'], import_df['전자상거래_금액'] / 1e6,
            color='#BF360C', linewidth=2, label='전자상거래 수입액')
ax_imp.fill_between(import_df['연월'], import_df['전자상거래_금액'] / 1e6, alpha=0.12, color='#BF360C')
ax_imp.set_title('전자상거래 수입액 추세 (백만 달러)', fontsize=12)
ax_imp.set_ylabel('백만 달러')
ax_imp.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'{x:,.0f}'))
ax_imp.legend(loc='upper left', fontsize=9)
ax_imp.grid(axis='y', alpha=0.3, linestyle='--')
add_yoy_growth(ax_imp, import_df, '전자상거래_금액', '#FFAB91', '수입 YoY')

# 수출 건수
ax_e_cnt.bar(export_df['연월'], export_df['전자상거래_건수'] / 1e3,
             color='#1565C0', alpha=0.7, width=20)
ax_e_cnt.set_title('수출 건수 (천 건)', fontsize=11)
ax_e_cnt.set_ylabel('천 건')
ax_e_cnt.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'{x:,.0f}'))
ax_e_cnt.grid(axis='y', alpha=0.3, linestyle='--')

# 수입 건수
ax_i_cnt.bar(import_df['연월'], import_df['전자상거래_건수'] / 1e3,
             color='#BF360C', alpha=0.7, width=20)
ax_i_cnt.set_title('수입 건수 (천 건)', fontsize=11)
ax_i_cnt.set_ylabel('천 건')
ax_i_cnt.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'{x:,.0f}'))
ax_i_cnt.grid(axis='y', alpha=0.3, linestyle='--')

plt.savefig('eda_trend.png', dpi=150, bbox_inches='tight')
plt.close()
print("  저장: eda_trend.png")

# ═══════════════════════════════════════════════════════════════
# 3. 계절 분해 (Prophet 스타일 컴포넌트 시각화)
# ═══════════════════════════════════════════════════════════════
def plot_decomposition(df, trade_label, color, fname):
    ts = df.set_index('연월')['전자상거래_금액'].asfreq('MS')
    result = seasonal_decompose(ts, model='multiplicative', period=12, extrapolate_trend='freq')

    fig, axes = plt.subplots(4, 1, figsize=(15, 11))
    fig.suptitle(f'전자상거래 {trade_label} – Prophet 스타일 시계열 분해\n(추세 + 계절성 + 잔차)',
                 fontsize=13, fontweight='bold')

    items = [
        (ts / 1e6, '원본 시계열 (백만 달러)', color),
        (result.trend / 1e6, '추세 (Trend)', '#34A853'),
        (result.seasonal, '계절성 지수 (Seasonality)', '#FBBC04'),
        (result.resid, '잔차 (Residual)', '#EA4335'),
    ]
    for ax, (series, title, c) in zip(axes, items):
        ax.plot(series.index, series.values, color=c, linewidth=1.5)
        ax.set_title(title, fontsize=11)
        ax.grid(axis='y', alpha=0.3, linestyle='--')
        if '원본' in title or '추세' in title:
            ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'{x:,.0f}'))
        if '계절' in title:
            ax.axhline(1, color='gray', linewidth=0.8, linestyle='--')

    plt.tight_layout()
    plt.savefig(fname, dpi=130, bbox_inches='tight')
    plt.close()
    print(f"  저장: {fname}")

print("\n  계절 분해 차트 생성...")
plot_decomposition(export_df, '수출', '#1565C0', 'components_export.png')
plot_decomposition(import_df, '수입', '#BF360C', 'components_import.png')

# ═══════════════════════════════════════════════════════════════
# 4. SARIMAX 모델 – 한국 이커머스 이벤트 회귀 변수 포함
#    (Facebook Prophet의 holidays 효과를 외생 변수로 구현)
# ═══════════════════════════════════════════════════════════════

def build_event_regressors(date_index: pd.DatetimeIndex) -> pd.DataFrame:
    """
    한국 이커머스 주요 이벤트를 월별 더미 변수로 생성합니다.
    Prophet의 holidays 파라미터와 동일한 역할을 합니다.
    """
    df_ev = pd.DataFrame(index=date_index)

    # 설날 영향 월 (설날 포함 월)
    seollal_months = {
        2014: 1, 2015: 2, 2016: 2, 2017: 1,
        2018: 2, 2019: 2, 2020: 1, 2021: 2,
        2022: 2, 2023: 1, 2024: 2, 2025: 1, 2026: 2,
    }
    # 추석 영향 월
    chuseok_months = {
        2014: 9, 2015: 9, 2016: 9, 2017: 10,
        2018: 9, 2019: 9, 2020: 10, 2021: 9,
        2022: 9, 2023: 9, 2024: 9, 2025: 10, 2026: 9,
    }

    df_ev['설날'] = [
        1.0 if (d.year in seollal_months and d.month == seollal_months[d.year]) else 0.0
        for d in date_index
    ]
    df_ev['추석'] = [
        1.0 if (d.year in chuseok_months and d.month == chuseok_months[d.year]) else 0.0
        for d in date_index
    ]
    # 연말 쇼핑 시즌 (11~12월)
    df_ev['연말쇼핑'] = [1.0 if d.month in (11, 12) else 0.0 for d in date_index]
    # 어린이날 시즌 (5월)
    df_ev['어린이날'] = [1.0 if d.month == 5 else 0.0 for d in date_index]
    # COVID 충격 (2020년 1분기)
    df_ev['COVID충격'] = [
        1.0 if (d.year == 2020 and d.month <= 3) else 0.0
        for d in date_index
    ]
    return df_ev


FORECAST_START = pd.Timestamp('2026-03-01')
FORECAST_END   = pd.Timestamp('2026-12-01')


def train_and_forecast(df: pd.DataFrame, trade_label: str):
    """
    SARIMAX(1,1,1)(1,1,1,12) + 이벤트 외생 변수로 모델 학습 후
    2026-03 ~ 2026-12 예측값과 신뢰구간을 반환합니다.
    """
    ts = df.set_index('연월')['전자상거래_금액'].asfreq('MS').copy()

    # 전체 기간 (학습 + 예측) 날짜 인덱스
    future_idx = pd.date_range(ts.index[0], FORECAST_END, freq='MS')
    exog_all = build_event_regressors(future_idx)

    # 학습 기간 외생 변수
    exog_train = exog_all.loc[ts.index]

    print(f"  [{trade_label}] SARIMAX 모델 학습 중 (약 10~30초)...")
    model = SARIMAX(
        endog=ts,
        exog=exog_train,
        order=(1, 1, 1),
        seasonal_order=(1, 1, 1, 12),
        enforce_stationarity=False,
        enforce_invertibility=False,
    )
    result = model.fit(disp=False, maxiter=200)
    print(f"  [{trade_label}] 학습 완료 (AIC={result.aic:.1f})")

    # 예측 기간 외생 변수 (학습 마지막 이후 ~ 2026-12)
    pred_idx = pd.date_range(ts.index[-1] + pd.offsets.MonthBegin(1), FORECAST_END, freq='MS')
    exog_pred = exog_all.loc[pred_idx]

    n_steps = len(pred_idx)
    forecast_obj = result.get_forecast(steps=n_steps, exog=exog_pred)
    fc_mean = forecast_obj.predicted_mean
    fc_ci   = forecast_obj.conf_int(alpha=0.10)  # 90% CI

    # 학습 기간 적합값
    fitted = result.fittedvalues

    return ts, fitted, fc_mean, fc_ci, result


print("\n[3/5] SARIMAX 모델 학습 (Prophet 스타일)...")
exp_ts, exp_fitted, exp_fc, exp_ci, exp_result = train_and_forecast(export_df, '수출')
imp_ts, imp_fitted, imp_fc, imp_ci, imp_result = train_and_forecast(import_df, '수입')

# ═══════════════════════════════════════════════════════════════
# 5. 예측 결과 테이블
# ═══════════════════════════════════════════════════════════════
print("\n[4/5] 예측 결과 집계...")

def make_result_df(fc_mean, fc_ci, prefix):
    result = pd.DataFrame({
        '연월': fc_mean.index.strftime('%Y-%m'),
        f'{prefix}_예측': fc_mean.values.round(0).astype(int),
        f'{prefix}_하한(90%)': fc_ci.iloc[:, 0].values.round(0).astype(int),
        f'{prefix}_상한(90%)': fc_ci.iloc[:, 1].values.round(0).astype(int),
    })
    return result

export_result = make_result_df(exp_fc, exp_ci, '수출액')
import_result = make_result_df(imp_fc, imp_ci, '수입액')
combined = pd.merge(export_result, import_result, on='연월')

print("\n  ┌─────────────────────────────────────────────────────────────────────────────────────────┐")
print("  │  2026년 3월~12월 전자상거래 수출입 예측 결과 (단위: 달러)                                │")
print("  ├──────────┬────────────────┬────────────────┬────────────────┬────────────────┤")
print("  │  연월    │  수출액_예측   │  수출_하한     │  수입액_예측   │  수입_하한     │")
print("  ├──────────┼────────────────┼────────────────┼────────────────┼────────────────┤")
for _, row in combined.iterrows():
    print(f"  │ {row['연월']} │ {row['수출액_예측']:>14,} │ {row['수출액_하한(90%)']:>14,} │ {row['수입액_예측']:>14,} │ {row['수입액_하한(90%)']:>14,} │")
print("  └──────────┴────────────────┴────────────────┴────────────────┴────────────────┘")

# ═══════════════════════════════════════════════════════════════
# 6. 예측 결과 차트
# ═══════════════════════════════════════════════════════════════
print("\n[5/5] 예측 차트 생성...")

def plot_forecast_panel(ax, ts, fitted, fc_mean, fc_ci, panel_label, color):
    """실적 + 적합값 + 예측 + 신뢰구간을 그립니다."""
    # 실적
    ax.plot(ts.index, ts.values / 1e6, color=color, linewidth=1.8,
            label='실적 데이터', zorder=3)
    # 적합값
    ax.plot(fitted.index, fitted.values / 1e6, color='gray', linewidth=1.0,
            linestyle='--', alpha=0.65, label='모델 적합값')
    # 예측
    ax.plot(fc_mean.index, fc_mean.values / 1e6, color=color, linewidth=2.5,
            linestyle='-', marker='o', markersize=6, label='예측값', zorder=4)
    # 신뢰구간
    ax.fill_between(fc_mean.index,
                    fc_ci.iloc[:, 0].values / 1e6,
                    fc_ci.iloc[:, 1].values / 1e6,
                    color=color, alpha=0.18, label='90% 예측 구간')
    # 예측값 라벨
    for dt, val in zip(fc_mean.index, fc_mean.values):
        ax.annotate(f'{val/1e6:,.1f}',
                    xy=(dt, val / 1e6),
                    xytext=(0, 10), textcoords='offset points',
                    fontsize=7.5, ha='center', color=color, fontweight='bold')
    # 예측 시작 구분선
    ax.axvline(FORECAST_START, color='black', linestyle=':', linewidth=1.3, alpha=0.6)
    ylim = ax.get_ylim()
    ax.text(FORECAST_START, ylim[0] + (ylim[1]-ylim[0]) * 0.03,
            '  예측 시작 ->', fontsize=8.5, color='black', fontstyle='italic')

    ax.set_title(f'전자상거래 {panel_label} (백만 달러)', fontsize=13, pad=8)
    ax.set_ylabel('금액 (백만 달러)')
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f'{x:,.0f}'))
    ax.legend(loc='upper left', fontsize=9, framealpha=0.8)
    ax.grid(axis='y', linestyle='--', alpha=0.35)

fig, axes = plt.subplots(2, 1, figsize=(19, 13))
fig.suptitle('한국 전자상거래 수출입 예측\n(SARIMAX + 한국 이커머스 계절 이벤트 | 2026년 3~12월)',
             fontsize=15, fontweight='bold')

plot_forecast_panel(axes[0], exp_ts, exp_fitted, exp_fc, exp_ci, '수출액 예측', '#1565C0')
plot_forecast_panel(axes[1], imp_ts, imp_fitted, imp_fc, imp_ci, '수입액 예측', '#BF360C')
axes[1].set_xlabel('연월')

plt.tight_layout()
plt.savefig('forecast_result.png', dpi=150, bbox_inches='tight')
plt.close()
print("  저장: forecast_result.png")

# ─── 예측 요약 바 차트 ───────────────────────────────────────
fig, axes = plt.subplots(1, 2, figsize=(16, 6))
fig.suptitle('2026년 전자상거래 수출입 월별 예측 요약', fontsize=14, fontweight='bold')

months = combined['연월']
x = np.arange(len(months))
width = 0.35

for ax, col_p, col_lo, col_hi, label, color in [
    (axes[0], '수출액_예측', '수출액_하한(90%)', '수출액_상한(90%)', '수출액', '#1565C0'),
    (axes[1], '수입액_예측', '수입액_하한(90%)', '수입액_상한(90%)', '수입액', '#BF360C'),
]:
    pred = combined[col_p].values / 1e6
    lo   = combined[col_lo].values / 1e6
    hi   = combined[col_hi].values / 1e6
    err  = np.array([pred - lo, hi - pred])

    bars = ax.bar(x, pred, color=color, alpha=0.8, width=0.6,
                  yerr=err, capsize=4, ecolor='#555', label='예측값 (±90% CI)')
    ax.set_xticks(x)
    ax.set_xticklabels([m[5:] + '월' for m in months], fontsize=9)
    ax.set_title(f'전자상거래 {label} 월별 예측 (2026)', fontsize=12)
    ax.set_ylabel('백만 달러')
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f'{v:,.0f}'))
    ax.grid(axis='y', alpha=0.3, linestyle='--')
    for xi, v in zip(x, pred):
        ax.text(xi, v + (hi - pred).mean() * 1.1, f'{v:,.1f}',
                ha='center', fontsize=7.5, fontweight='bold', color=color)

plt.tight_layout()
plt.savefig('forecast_summary_bar.png', dpi=150, bbox_inches='tight')
plt.close()
print("  저장: forecast_summary_bar.png")

# ═══════════════════════════════════════════════════════════════
# 7. CSV 저장
# ═══════════════════════════════════════════════════════════════
OUTPUT_FILE = '2026_ecommerce_forecast.csv'
combined.to_csv(OUTPUT_FILE, index=False, encoding='utf-8-sig')
print(f"\n저장 완료: {OUTPUT_FILE}")

print("\n========================================")
print("  모든 작업 완료!")
print("========================================")
print("생성된 파일:")
print("  eda_trend.png            - EDA 추세 차트 (수출입 + 건수 + YoY)")
print("  components_export.png    - 수출 시계열 분해 (추세/계절성/잔차)")
print("  components_import.png    - 수입 시계열 분해 (추세/계절성/잔차)")
print("  forecast_result.png      - 예측 결과 차트 (실적+예측+신뢰구간)")
print("  forecast_summary_bar.png - 2026년 월별 예측 바 차트")
print(f"  {OUTPUT_FILE}    - 예측 결과 CSV")
