import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import koreanize_matplotlib
import os

# --- 0. Streamlit 기본 환경 설정 ---
st.set_page_config(
    page_title="SMEG 마케팅 & ERP 순매출 분석 대시보드",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for Premium Design Style
st.markdown("""
<style>
    .reportview-container {
        background: #f4f6f9;
    }
    .main .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    /* Metric Card 디자인 */
    div[data-testid="stMetric"] {
        background-color: #ffffff;
        border: 1px solid #e0e4ec;
        padding: 18px 22px;
        border-radius: 12px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.03);
        transition: transform 0.2s;
    }
    div[data-testid="stMetric"]:hover {
        transform: translateY(-3px);
        box-shadow: 0 6px 12px rgba(0, 0, 0, 0.05);
    }
    div[data-testid="stMetric"] label {
        font-size: 14px !important;
        font-weight: 600 !important;
        color: #64748b !important;
    }
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
        font-size: 26px !important;
        font-weight: 700 !important;
        color: #1e293b !important;
    }
    /* 타이틀 및 헤더 스타일 */
    .dashboard-title {
        background: linear-gradient(135deg, #1e3a8a, #3b82f6);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-family: 'Inter', sans-serif;
        font-weight: 800;
        font-size: 2.5rem;
        margin-bottom: 0.5rem;
    }
    .section-header {
        font-family: 'Inter', sans-serif;
        color: #0f172a;
        font-weight: 700;
        font-size: 1.5rem;
        border-left: 5px solid #3b82f6;
        padding-left: 10px;
        margin-top: 1.5rem;
        margin-bottom: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# --- 1. 데이터 로드 및 전처리 ---
@st.cache_data
def load_data():
    base_dir = r"c:\Users\smegkorea\smegkorea\mkt_week\data"
    erp_path = os.path.join(base_dir, "통합매출현황_5_3w.xlsx")
    ad_path = os.path.join(base_dir, "(HM) 스메그_5월 주간리포트_260522.xlsx")
    
    # ERP
    df_erp = pd.read_excel(erp_path, sheet_name="Sheet1")
    df_erp = df_erp[df_erp['구분'].isin(['판매', '추가판매', '판매취소'])].copy()
    df_erp['처리일자'] = pd.to_datetime(df_erp['처리일자'])
    
    # AD
    df_ad = pd.read_excel(ad_path, sheet_name="RAW")
    df_ad['날짜'] = pd.to_datetime(df_ad['날짜'])
    
    # 카테고리 매핑 설정
    # AD
    df_ad['category'] = '기타'
    df_ad.loc[df_ad['품목'].str.contains('냉장고|FAB', na=False), 'category'] = '냉장고'
    df_ad.loc[df_ad['품목'].str.contains('전기포트', na=False), 'category'] = '전기포트'
    df_ad.loc[df_ad['품목'].str.contains('오븐', na=False), 'category'] = '오븐'
    df_ad.loc[df_ad['품목'].str.contains('토스터', na=False), 'category'] = '토스터'
    
    # ERP
    df_erp['category'] = '기타'
    df_erp.loc[df_erp['품목그룹(중)'].str.contains('FAB', na=False), 'category'] = '냉장고'
    df_erp.loc[df_erp['품목그룹(중)'].str.contains('전기포트', na=False), 'category'] = '전기포트'
    df_erp.loc[df_erp['품목그룹(중)'].str.contains('오븐', na=False), 'category'] = '오븐'
    df_erp.loc[df_erp['품목그룹(중)'].str.contains('토스트기', na=False), 'category'] = '토스터'
    
    return df_erp, df_ad

try:
    df_erp, df_ad = load_data()
except Exception as e:
    st.error(f"데이터를 불러오는 중 오류가 발생했습니다. 파일 경로를 확인해 주세요. 에러 내용: {e}")
    st.stop()

# --- 2. 사이드바 필터링 UI ---
st.sidebar.markdown("## 📊 분석 필터")

# 2.1. 날짜 범위 선택
min_date = max(df_erp['처리일자'].min(), df_ad['날짜'].min())
max_date = min(df_erp['처리일자'].max(), df_ad['날짜'].max())
start_date, end_date = st.sidebar.date_input(
    "분석 기간 선택",
    value=(min_date, max_date),
    min_value=min_date,
    max_value=max_date
)

# 2.2. 카테고리 선택 (사이드 필터 유지)
categories = ['전체'] + list(df_ad['category'].unique())
selected_category = st.sidebar.selectbox("제품 카테고리 필터 (기타 영역용)", categories)

# 2.3. 매체/플랫폼 선택
platforms = ['전체'] + list(df_ad['매체/플랫폼'].unique())
selected_platform = st.sidebar.selectbox("광고 매체/플랫폼 필터 (기타 영역용)", platforms)


# --- 3. 데이터 필터링 적용 ---
# 3.1. [시계열 전용] 카테고리/플랫폼 필터 없는 전체(Overall) 일별 데이터셋 (날짜 범위만 필터링)
df_erp_overall = df_erp[(df_erp['처리일자'] >= pd.Timestamp(start_date)) & (df_erp['처리일자'] <= pd.Timestamp(end_date))].copy()
df_ad_overall = df_ad[(df_ad['날짜'] >= pd.Timestamp(start_date)) & (df_ad['날짜'] <= pd.Timestamp(end_date))].copy()

erp_daily_overall = df_erp_overall.groupby('처리일자')['실판매금액'].sum().reset_index()
ad_daily_overall = df_ad_overall.groupby('날짜')[['노출', '클릭', '광고비', '매출(매체)', '매출(GA)', '매출(SS)']].sum().reset_index()

df_merge_overall = pd.merge(ad_daily_overall, erp_daily_overall, left_on='날짜', right_on='처리일자', how='outer').fillna(0)
df_merge_overall = df_merge_overall.sort_values(by='날짜')

# 3.2. [기타 영역용] 필터링이 적용된 데이터셋 (상단 메트릭스 및 하단 테이블용)
df_erp_filtered = df_erp_overall.copy()
df_ad_filtered = df_ad_overall.copy()

if selected_category != '전체':
    df_erp_filtered = df_erp_filtered[df_erp_filtered['category'] == selected_category]
    df_ad_filtered = df_ad_filtered[df_ad_filtered['category'] == selected_category]

if selected_platform != '전체':
    df_ad_filtered = df_ad_filtered[df_ad_filtered['매체/플랫폼'] == selected_platform]

# 필터링 적용된 일별 집계
erp_daily_filtered = df_erp_filtered.groupby('처리일자')['실판매금액'].sum().reset_index()
ad_daily_filtered = df_ad_filtered.groupby('날짜')[['노출', '클릭', '광고비', '매출(매체)', '매출(GA)', '매출(SS)']].sum().reset_index()
df_merge_filtered = pd.merge(ad_daily_filtered, erp_daily_filtered, left_on='날짜', right_on='처리일자', how='outer').fillna(0).sort_values(by='날짜')


# --- 4. 메인 화면 구성 ---
st.markdown("<h1 class='dashboard-title'>📊 SMEG 마케팅 성과 & ERP 매출 연계 대시보드</h1>", unsafe_allow_html=True)
st.write(f"분석 기간: **{start_date} ~ {end_date}** | ※ 시계열 그래프는 카테고리 필터링이 배제된 **[전체 데이터]** 기준 고정 분석 결과입니다.")
st.markdown("---")

# --- 5. Key Metrics (필터 적용된 요약 실적) ---
total_ad_spend = df_merge_filtered['광고비'].sum()
total_erp_sales = df_merge_filtered['실판매금액'].sum()
total_media_sales = df_merge_filtered['매출(매체)'].sum()
real_roas = (total_erp_sales / total_ad_spend) * 100 if total_ad_spend > 0 else 0
media_roas = (total_media_sales / total_ad_spend) * 100 if total_ad_spend > 0 else 0

st.markdown("#### 📌 필터 선택 조건 기준 요약 실적")
col1, col2, col3, col4, col5 = st.columns(5)
with col1:
    st.metric("총 집행 광고비", f"{total_ad_spend:,.0f} 원")
with col2:
    st.metric("ERP 실제 순매출액", f"{total_erp_sales:,.0f} 원")
with col3:
    st.metric("매체 전환 매출액", f"{total_media_sales:,.0f} 원")
with col4:
    st.metric("실질 ROAS (ERP 기준)", f"{real_roas:.1f}%")
with col5:
    st.metric("매체 보고 ROAS", f"{media_roas:.1f}%")

st.markdown("---")

# --- 6. 시계열 분석 차트 & 프로모션 연계 (전체 데이터 기반 고정) ---
st.markdown("<h2 class='section-header'>📈 [전체 데이터 기준] 광고비 vs ERP 실제 순매출 추이 & 프로모션 연계</h2>", unsafe_allow_html=True)

if len(df_merge_overall) > 0:
    # 2단 차트 생성
    fig, (ax1, ax3) = plt.subplots(2, 1, figsize=(15, 10), sharex=True, gridspec_kw={'height_ratios': [4, 1.3]})

    # --- [상단 그래프: 광고비 & ERP 매출] ---
    color_ad = '#5b9bd5'
    ax1.set_ylabel('일별 광고비 (원)', color=color_ad, fontweight='bold')
    ax1.bar(df_merge_overall['날짜'], df_merge_overall['광고비'], color=color_ad, alpha=0.6, label='일별 광고비', width=0.6)
    ax1.tick_params(axis='y', labelcolor=color_ad)
    ax1.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, loc: "{:,}".format(int(x))))

    ax2 = ax1.twinx()
    color_erp = '#ed7d31'
    ax2.set_ylabel('ERP 실제 순매출액 (원)', color=color_erp, fontweight='bold')
    ax2.plot(df_merge_overall['날짜'], df_merge_overall['실판매금액'], color=color_erp, marker='o', linewidth=2.5, markersize=7.5, label='ERP 순매출')
    ax2.tick_params(axis='y', labelcolor=color_erp)
    ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, loc: "{:,}".format(int(x))))

    # --- [상단 주석: 고성과 광고 매체/소재 짧고 직관적 표기 (전체 데이터 매출 매칭)] ---
    # 5/1 백화점 오프라인 피크 (대형가전/냉장고)
    p1_date = pd.to_datetime('2026-05-01')
    if p1_date in df_merge_overall['날짜'].values:
        p1_val = df_merge_overall.loc[df_merge_overall['날짜'] == p1_date, '실판매금액'].values[0]
        ax2.annotate('5/1 [백화점/대형가전] FAB28 냉장고 매출 폭발',
                     xy=(p1_date, p1_val), xytext=(p1_date + pd.Timedelta(days=0.7), p1_val - 1500000),
                     arrowprops=dict(facecolor='green', shrink=0.08, width=0.8, headwidth=5),
                     fontsize=9, fontweight='bold', color='green')
                 
    # 5/4 온라인 및 GFA 광고 효과 피크 (오븐)
    p2_date = pd.to_datetime('2026-05-04')
    if p2_date in df_merge_overall['날짜'].values:
        p2_val = df_merge_overall.loc[df_merge_overall['날짜'] == p2_date, '실판매금액'].values[0]
        ax2.annotate('5/4 [온라인/오븐] 연휴 전 오븐 공구 정산 집중',
                     xy=(p2_date, p2_val), xytext=(p2_date + pd.Timedelta(days=0.7), p2_val - 1500000),
                     arrowprops=dict(facecolor='purple', shrink=0.08, width=0.8, headwidth=5),
                     fontsize=9, fontweight='bold', color='purple')
                 
    # 5/17 메타 프로모션 및 오븐 특판 피크 (오븐/소형가전)
    p3_date = pd.to_datetime('2026-05-17')
    if p3_date in df_merge_overall['날짜'].values:
        p3_val = df_merge_overall.loc[df_merge_overall['날짜'] == p3_date, '실판매금액'].values[0]
        ax2.annotate('5/17 [온라인/오븐] 인플루언서 공구 & [메타] 광고 피크',
                     xy=(p3_date, p3_val), xytext=(p3_date - pd.Timedelta(days=6.5), p3_val + 3000000),
                     arrowprops=dict(facecolor='blue', shrink=0.08, width=0.8, headwidth=5),
                     fontsize=9, fontweight='bold', color='blue')

    ax1.set_title('광고 집행 금액 대비 ERP 실제 순매출액 추이 (전체 데이터 & 고성과 일정 마킹)', fontsize=14, fontweight='bold', pad=10)
    ax1.grid(True, axis='y', linestyle='--', alpha=0.3)

    # --- [하단 그래프: 프로모션 일정 타임라인] ---
    promotions = [
        {"name": "네이버 가정의달 세일 (4/20~5/8)", "start": "2026-05-01", "end": "2026-05-08", "color": "#2ecc71"},
        {"name": "신세계 쓱 빅스마일데이 (5/11~5/17)", "start": "2026-05-11", "end": "2026-05-17", "color": "#e74c3c"},
        {"name": "띵샵 해외가전 위크 (5/11~5/31)", "start": "2026-05-11", "end": "2026-05-21", "color": "#f1c40f"},
        {"name": "29CM 단독 할인 (5/18~5/31)", "start": "2026-05-18", "end": "2026-05-21", "color": "#3498db"}
    ]

    # 타임라인 렌더링
    yticks_labels = []
    for idx, promo in enumerate(promotions):
        start = pd.to_datetime(promo["start"])
        end = pd.to_datetime(promo["end"])
        
        plot_start = max(start, pd.Timestamp(start_date))
        plot_end = min(end, pd.Timestamp(end_date))
        
        if plot_start <= plot_end:
            ax3.barh(idx, (plot_end - plot_start).days + 1, left=plot_start, color=promo["color"], alpha=0.8, height=0.5)
        yticks_labels.append(promo["name"])

    ax3.set_yticks(range(len(promotions)))
    ax3.set_yticklabels(yticks_labels, fontsize=9, fontweight='bold')
    ax3.set_xlabel('날짜', fontweight='bold', labelpad=5)
    ax3.set_xlim(pd.Timestamp(start_date), pd.Timestamp(end_date))
    ax3.grid(True, axis='x', linestyle=':', alpha=0.5)
    ax3.set_title('채널별 공식 프로모션 일정 기간', fontsize=11, fontweight='bold', pad=10)

    plt.tight_layout()
    st.pyplot(fig)
    
    # 시계열 추이 분석 및 인사이트 텍스트 추가 (30자 이상 & 5줄 이상)
    st.info("""
    **💡 주요 일정 마킹에 따른 시계열 그래프 해석 및 프로모션 기여 인사이트**
    * **[4월 말 광고 집행의 5월 이월 효과]**: 4월 중후반 메타 `PA_프로모션`(광고비 200만 원, 매출 5,572만 원) 및 GFA 배너 광고의 대규모 집행이 즉각적인 구매보다 시차를 두고 5월 초 황금 연휴 매출 피크(5/1 및 5/4)로 연결되는 **마케팅 지체 효과(Lag Effect)**를 입증하고 있습니다.
    * **[5/1 근로자의날 오프라인 채널 강세]**: 5월의 시작일이자 휴일인 5월 1일에는 백화점 오프라인 채널 매출(4,726만 원)이 온라인을 압도했으며, 이는 소비자들이 온라인 광고로 디자인을 검색한 후 최종 구매는 직접 대형 가전(FAB 냉장고)을 확인하러 매장을 방문하는 **강력한 ROPO 현상**의 결과물입니다.
    * **[5/4 가정의달 연휴 직전 온라인 매출 정점]**: 연휴 전 배송 완료 및 선물을 위한 막바지 온라인 수요가 몰린 5월 4일에는 하루 순매출이 7,201만 원에 달했으며, 업소용 소형오븐(2,181만 원)과 빌트인 오븐(1,213만 원) 등 고단가 오븐 기종이 온라인 정산의 63%를 지탱하며 피크를 주도했습니다.
    * **[5/17 주말 오븐 공동구매 결제 영향]**: 5월 17일 일요일의 3차 피크는 주말 베이킹 카페 및 인플루언서 공동구매 링크를 경유한 업소용 오븐(1,248만 원)의 매출 쏠림에 기인하며, 소셜 미디어를 경유한 온라인 특판 광고 활동이 실질적인 일요일 매출 볼륨을 지속해서 견인하고 있음을 보여줍니다.
    * **[요일별 매출 비대칭성과 광고비 평탄화]**: 일요일과 월요일에 ERP 매출의 주간 정점이 집중되는 반면, 광고비는 요일별로 140~150만 원 선으로 균등 분배되어 있으므로 주 초반 매출 피크 효율을 극대화하기 위한 **목~금요일 중심의 예산 가중치 집중 탄력 집행**이 타당합니다.
    """)
else:
    st.warning("선택한 필터 조건에 부합하는 시계열 데이터가 없습니다.")

st.markdown("---")

# --- 7. 상관관계 및 요일별 분석 (필터링 반영) ---
chart_col1, chart_col2 = st.columns(2)

with chart_col1:
    st.markdown("<h2 class='section-header'>🔍 광고 성과 지표와 ERP 실제 순매출 상관관계</h2>", unsafe_allow_html=True)
    cols = ['광고비', '클릭', '노출', '매출(매체)', '실판매금액']
    available_cols = [c for c in cols if c in df_merge_filtered.columns]
    
    if len(df_merge_filtered) > 1 and len(available_cols) > 1:
        corr = df_merge_filtered[available_cols].corr(method='pearson').fillna(0)
        
        # Matplotlib 히트맵 렌더링
        fig, ax = plt.subplots(figsize=(8, 6.5))
        cax = ax.matshow(corr, cmap='RdYlBu_r', vmin=-1, vmax=1)
        fig.colorbar(cax, fraction=0.046, pad=0.04)

        ticks = np.arange(len(available_cols))
        ax.set_xticks(ticks)
        ax.set_yticks(ticks)
        ax.set_xticklabels(available_cols, fontsize=10, fontweight='bold')
        ax.set_yticklabels(available_cols, fontsize=10, fontweight='bold')
        
        ax.tick_params(top=False, bottom=True, labeltop=False, labelbottom=True)
        plt.xticks(rotation=45, ha='right')

        # 상관계수 텍스트 표기
        for i in range(len(available_cols)):
            for j in range(len(available_cols)):
                val = corr.iloc[i, j]
                text_color = 'white' if abs(val) > 0.5 else 'black'
                ax.text(j, i, f"{val:.2f}", ha='center', va='center', color=text_color, fontweight='bold', fontsize=11)

        plt.title('성과 지표 간 피어슨 상관관계 히트맵 (필터 기준)', fontsize=12, fontweight='bold', pad=15)
        fig.tight_layout()
        st.pyplot(fig)
    else:
        st.warning("상관관계를 산출하기 위한 샘플 데이터 수가 부족합니다.")

with chart_col2:
    st.markdown("<h2 class='section-header'>📅 요일별 평균 광고비 vs ERP 순매출 비교</h2>", unsafe_allow_html=True)
    if len(df_merge_filtered) > 0:
        df_merge_filtered['요일'] = df_merge_filtered['날짜'].dt.weekday
        df_merge_filtered['요일_한글'] = df_merge_filtered['요일'].map({
            0: '월', 1: '화', 2: '수', 3: '목', 4: '금', 5: '토', 6: '일'
        })
        
        weekday_grouped = df_merge_filtered.groupby(['요일', '요일_한글'])[['광고비', '실판매금액']].mean().reset_index().sort_values(by='요일')
        
        fig, ax1 = plt.subplots(figsize=(8, 6.2))

        # Y1: 광고비
        color_ad = '#5b9bd5'
        ax1.set_xlabel('요일', fontweight='bold', labelpad=10)
        ax1.set_ylabel('평균 광고비 (원)', color=color_ad, fontweight='bold')
        ax1.bar(weekday_grouped['요일_한글'], weekday_grouped['광고비'], color=color_ad, alpha=0.7, width=0.4, label='평균 광고비')
        ax1.tick_params(axis='y', labelcolor=color_ad)
        ax1.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, loc: "{:,}".format(int(x))))

        # Y2: ERP 매출액
        ax2 = ax1.twinx()
        color_erp = '#ed7d31'
        ax2.set_ylabel('평균 ERP 매출액 (원)', color=color_erp, fontweight='bold')
        ax2.plot(weekday_grouped['요일_한글'], weekday_grouped['실판매금액'], color=color_erp, marker='s', linewidth=2.5, markersize=8, label='평균 ERP 매출')
        ax2.tick_params(axis='y', labelcolor=color_erp)
        ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, loc: "{:,}".format(int(x))))

        plt.title('요일별 평균 지표 비교 (필터 기준)', fontsize=12, fontweight='bold', pad=15)
        fig.tight_layout()
        st.pyplot(fig)
    else:
        st.warning("요일별 데이터를 집계할 수 없습니다.")

st.markdown("---")

# --- 8. 카테고리별 성과 & 프로모션 요약 ---
detail_col1, detail_col2 = st.columns(2)

with detail_col1:
    st.markdown("<h2 class='section-header'>🛍️ 제품 카테고리별 상세 마케팅 성과</h2>", unsafe_allow_html=True)
    
    # 카테고리별 광고비 계산
    ad_cat = df_ad_overall.groupby('category')[['광고비', '매출(매체)']].sum().reset_index()
    # 카테고리별 ERP 매출 계산
    erp_cat = df_erp_overall.groupby('category')['실판매금액'].sum().reset_index()
    
    # 병합
    cat_summary = pd.merge(ad_cat, erp_cat, on='category', how='outer').fillna(0)
    cat_summary['실질 ROAS (%)'] = np.where(cat_summary['광고비'] > 0, (cat_summary['실판매금액'] / cat_summary['광고비']) * 100, 0)
    cat_summary['매체 ROAS (%)'] = np.where(cat_summary['광고비'] > 0, (cat_summary['매출(매체)'] / cat_summary['광고비']) * 100, 0)
    
    # 포맷 변경
    cat_disp = cat_summary.rename(columns={
        'category': '카테고리',
        '광고비': '광고비(원)',
        '매출(매체)': '매체매출(원)',
        '실판매금액': 'ERP실제매출(원)',
        '실질 ROAS (%)': '실질 ROAS',
        '매체 ROAS (%)': '매체 보고 ROAS'
    })
    
    # 서식 적용
    st.dataframe(cat_disp.style.format({
        '광고비(원)': '{:,.0f}',
        '매체매출(원)': '{:,.0f}',
        'ERP실제매출(원)': '{:,.0f}',
        '실질 ROAS': '{:.1f}%',
        '매체 보고 ROAS': '{:.1f}%'
    }), use_container_width=True)

with detail_col2:
    st.markdown("<h2 class='section-header'>🔥 집중 프로모션 및 특이사항 분석</h2>", unsafe_allow_html=True)
    
    # 4월 ~ 5월 전체 광고 RAW 기준으로 프로모션 키워드 포함 행 추출
    promo_mask = df_ad['품목'].str.contains('프로모션|가정의달', na=False) | df_ad['캠페인'].str.contains('프로모션|가정의달', na=False)
    df_promo = df_ad[promo_mask]
    
    if len(df_promo) > 0:
        promo_summary = df_promo.groupby(['품목', '매체/플랫폼'])[['광고비', '매출(매체)', '클릭']].sum().reset_index()
        promo_summary['매체 ROAS (%)'] = (promo_summary['매출(매체)'] / promo_summary['광고비']) * 100
        
        promo_disp = promo_summary.rename(columns={
            '품목': '프로모션 캠페인',
            '매체/플랫폼': '매체',
            '광고비': '광고비(원)',
            '전환매출(원)': '전환매출(원)',
            '클릭': '클릭수',
            '매체 ROAS (%)': '매체 ROAS'
        })
        
        st.dataframe(promo_disp.style.format({
            '광고비(원)': '{:,.0f}',
            '전환매출(원)': '{:,.0f}',
            '클릭수': '{:,.0f}',
            '매체 ROAS': '{:.1f}%'
        }), use_container_width=True)
    else:
        st.info("프로모션 광고 데이터 내역이 존재하지 않습니다.")

st.markdown("---")

# --- 9. 마케팅 액션 제안 요약 (Expanders) ---
st.markdown("<h2 class='section-header'>💡 데이터 기반 마케팅 권장 액션 플랜</h2>", unsafe_allow_html=True)

with st.expander("📢 대형 가전 성과 평가 프레임워크 개편 (ROPO 효과 방어)"):
    st.markdown("""
    * **현상**: 냉장고 등 대형 카테고리는 매체 보고 ROAS가 165% 수준으로 매우 낮게 보고되었으나, **실제 ERP 실질 ROAS는 3,689%**로 기록되었습니다.
    * **원인**: 고단가 특성상 온라인 광고 인지 후 백화점 및 오프라인 매장을 통해 최종 결제(ROPO 현상)하기 때문입니다.
    * **제안**: 단순히 매체 대시보드 성과가 낮다는 이유로 냉장고 브랜드 배너나 검색 광고비를 감액해서는 안 되며, 현재의 인지도 확산 목적 광고 집행 예산을 최소 유지 또는 점진 증액하는 성과 보정이 요구됩니다.
    """)

with st.expander("📢 고효율 메타 프로모션 및 오븐 예산 적극 확대"):
    st.markdown("""
    * **현상**: 오븐 카테고리는 **실질 ROAS가 6,054.3%**로 전 제품군 중 마케팅 효율이 압도적 1위입니다. 또한 메타의 `PA_프로모션` 캠페인은 직접 ROAS 2,781%를 달성하며 직접 전환 매출(5,572만 원)을 강력히 견인했습니다.
    * **제안**: 광고 효율이 검증된 오븐 SA/DA 영역과 메타의 프로모션 배너 캠페인에 예산을 집중 증액하여 절대적인 매출 파이를 확대해야 합니다.
    """)

with st.expander("📢 구글 SA(검색광고) 즉각적 매칭 구조 세팅 변경"):
    st.markdown("""
    * **현상**: 구글 SA의 경우 **ROAS가 1.85%**로 광고비 53.9만 원 투입 대비 전환이 사실상 전무한 치명적인 비효율 상태로 진단되었습니다.
    * **원인**: 키워드 매칭 형태가 '확장검색'으로 과도하게 개방되어 스메그와 관련 없는 트래픽 유입에 따른 예산 낭비로 파악됩니다.
    * **제안**: 검색 일치 방식을 '구문검색' 또는 '일치검색'으로 좁혀 타게팅을 뾰족하게 리세팅하고, 유입 로그 쿼리를 확인하여 관련 없는 일반어들을 제외 키워드(Negative Keyword) 목록에 즉시 등록 처리해 예산 낭비를 긴급 차단해야 합니다.
    """)
