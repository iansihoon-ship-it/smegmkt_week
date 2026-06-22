import pandas as pd
import numpy as np
import os
import json
import re

# 1. 파일 경로 설정
base_dir = r"c:\Users\smegkorea\smegkorea\mkt_week\data"
prev_erp_path = os.path.join(base_dir, "통합매출현황_6_1w.xlsx")
curr_erp_path = os.path.join(base_dir, "통합매출현황_6_3w.xlsx")
curr_ad_path = os.path.join(base_dir, "(HM) 스메그_6월 주간리포트_260619.xlsx")

html_template_path = r"c:\Users\smegkorea\smegkorea\mkt_week\smeg_5_3w_dashboard.html"
analysis_out = r"c:\Users\smegkorea\smegkorea\mkt_week\analysis_dashboard.html"
smeg_out = r"c:\Users\smegkorea\smegkorea\mkt_week\smeg_5_3w_dashboard.html"

print("1. 데이터 파일 로딩 시작...")
df_erp_1w = pd.read_excel(prev_erp_path, sheet_name="Sheet1")
df_erp_3w = pd.read_excel(curr_erp_path, sheet_name="Sheet1")
df_ad_raw = pd.read_excel(curr_ad_path, sheet_name="RAW")
print("데이터 파일 로딩 완료.")

# 2. 데이터 전처리 및 주차별 분류
# ERP 병합 및 중복제거
df_erp_raw = pd.concat([df_erp_1w, df_erp_3w], ignore_index=True)
df_erp_raw = df_erp_raw.drop_duplicates(subset=['처리번호', '품목코드', '수량', '실판매금액']).copy()
df_erp_raw['처리일자'] = pd.to_datetime(df_erp_raw['처리일자'])
df_erp_clean = df_erp_raw[df_erp_raw['구분'].isin(['판매', '추가판매', '판매취소'])].copy()

# AD 날짜 포맷
df_ad_raw['날짜'] = pd.to_datetime(df_ad_raw['날짜'])

# 주차 정의 (W1: 5/29~6/4, W2: 6/5~6/11, W3: 6/12~6/18)
W1_start, W1_end = '2026-05-29', '2026-06-04'
W2_start, W2_end = '2026-06-05', '2026-06-11'
W3_start, W3_end = '2026-06-12', '2026-06-18'

df_erp_clean.loc[(df_erp_clean['처리일자'] >= W1_start) & (df_erp_clean['처리일자'] <= W1_end), 'wk'] = 'W1'
df_erp_clean.loc[(df_erp_clean['처리일자'] >= W2_start) & (df_erp_clean['처리일자'] <= W2_end), 'wk'] = 'W2'
df_erp_clean.loc[(df_erp_clean['처리일자'] >= W3_start) & (df_erp_clean['처리일자'] <= W3_end), 'wk'] = 'W3'
df_erp = df_erp_clean[df_erp_clean['wk'].notna()].copy()

df_ad_raw.loc[(df_ad_raw['날짜'] >= W1_start) & (df_ad_raw['날짜'] <= W1_end), 'wk'] = 'W1'
df_ad_raw.loc[(df_ad_raw['날짜'] >= W2_start) & (df_ad_raw['날짜'] <= W2_end), 'wk'] = 'W2'
df_ad_raw.loc[(df_ad_raw['날짜'] >= W3_start) & (df_ad_raw['날짜'] <= W3_end), 'wk'] = 'W3'
df_ad = df_ad_raw[df_ad_raw['wk'].notna()].copy()

# 카테고리 매핑 설정
def get_cat(item, grp):
    item = str(item) if pd.notna(item) else ""
    grp = str(grp) if pd.notna(grp) else ""
    if '냉장고' in item or 'FAB' in item or '냉장고' in grp or 'FAB' in grp:
        return '냉장고'
    elif '전기포트' in item or '전기포트' in grp:
        return '전기포트'
    elif '오븐' in item or '오븐' in grp:
        return '오븐'
    elif '토스터' in item or '토스트기' in item or '토스트기' in grp:
        return '토스터'
    elif '커피' in item or '머신' in item or '커피' in grp or '머신' in grp:
        return '커피머신'
    else:
        return '기타'

df_ad['category'] = df_ad.apply(lambda r: get_cat(r['품목'], r['대분류']), axis=1)
df_erp['category'] = df_erp.apply(lambda r: get_cat(r['품목명'], r['품목그룹(중)']), axis=1)

WEEKS = ['W1', 'W2', 'W3']

# --- D 객체 빌드 ---
D = {}

# 1. D.ad 매체별 주차 실적
media_list = ["네이버 SA(검색)", "네이버 SS(쇼핑)", "메타 DA(디스플레이)"]
media_colors = {
    "네이버 SA(검색)": "#36c6b0",
    "네이버 SS(쇼핑)": "#5a9bf0",
    "메타 DA(디스플레이)": "#e0a93b"
}

ad_obj = {}
for m_name in media_list:
    ad_obj[m_name] = {"color": media_colors[m_name]}
    for wk in WEEKS:
        df_wk = df_ad[df_ad['wk'] == wk]
        if m_name == "네이버 SA(검색)":
            df_m = df_wk[(df_wk['매체/플랫폼'].str.contains('네이버|구글', na=False)) & (df_wk['광고유형'].str.contains('SA', na=False))]
        elif m_name == "네이버 SS(쇼핑)":
            df_m = df_wk[(df_wk['매체/플랫폼'].str.contains('네이버', na=False)) & (df_wk['광고유형'].str.contains('SS', na=False))]
        elif m_name == "메타 DA(디스플레이)":
            df_m = df_wk[df_wk['매체/플랫폼'].str.contains('메타|Meta|GFA', na=False) | ((df_wk['매체/플랫폼'].str.contains('구글', na=False)) & (df_wk['광고유형'].str.contains('DA', na=False)))]
        else:
            df_m = pd.DataFrame()
            
        if len(df_m) > 0 and df_m['광고비'].sum() > 0:
            cost = int(df_m['광고비'].sum())
            rev = int(df_m['매출(매체)'].sum())
            conv = int(df_m['클릭'].count())
            clk = int(df_m['클릭'].sum())
            imp = int(df_m['노출'].sum())
            roas = round(rev / cost, 1) if cost > 0 else 0
            ad_obj[m_name][wk] = {
                "cost": cost,
                "rev": rev,
                "roas": roas,
                "conv": conv,
                "clk": clk,
                "imp": imp
            }
        else:
            ad_obj[m_name][wk] = None

D["ad"] = ad_obj

# 2. D.weekly_sales & weekly_qty
weekly_sales = {}
weekly_qty = {}
for wk in WEEKS:
    df_wk = df_erp[df_erp['wk'] == wk]
    weekly_sales[wk] = int(df_wk['실판매금액'].sum())
    weekly_qty[wk] = int(df_wk['수량'].sum())

D["weekly_sales"] = weekly_sales
D["weekly_qty"] = weekly_qty

# 3. D.daily (5/29 ~ 6/18 일일 데이터 매핑)
df_erp_daily = df_erp.groupby('처리일자')['실판매금액'].sum().reset_index()
df_ad_daily = df_ad.groupby('날짜')['광고비'].sum().reset_index()
df_daily = pd.merge(df_ad_daily, df_erp_daily, left_on='날짜', right_on='처리일자', how='outer')
df_daily['날짜'] = df_daily['날짜'].fillna(df_daily['처리일자'])
df_daily = df_daily.drop(columns=['처리일자'])
df_daily = df_daily.fillna(0)
df_daily = df_daily.sort_values(by='날짜')

daily_list = []
for idx, row in df_daily.iterrows():
    d_dt = row['날짜']
    if d_dt >= pd.to_datetime(W1_start) and d_dt <= pd.to_datetime(W1_end):
        wk_val = "W1"
    elif d_dt >= pd.to_datetime(W2_start) and d_dt <= pd.to_datetime(W2_end):
        wk_val = "W2"
    else:
        wk_val = "W3"
        
    daily_list.append({
        "date": d_dt.strftime("%m/%d"),
        "rev": int(row['실판매금액']),
        "cost": int(row['광고비']),
        "wk": wk_val
    })

D["daily"] = daily_list

# 4. D.channel
channels = ["온라인", "백화점", "홀세일", "특판"]
chan_colors = {"온라인": "#c8102e", "백화점": "#36c6b0", "홀세일": "#5a9bf0", "특판": "#e0a93b"}
channel_list = []
for chan in channels:
    chan_data = {"name": chan, "color": chan_colors[chan]}
    total_val = 0
    for wk in WEEKS:
        if chan == "온라인":
            df_c = df_erp[(df_erp['wk'] == wk) & df_erp['거래처대분류'].isin(['온라인', '쇼핑몰', '자사몰'])]
        elif chan == "백화점":
            df_c = df_erp[(df_erp['wk'] == wk) & df_erp['거래처대분류'].isin(['백화점'])]
        elif chan == "홀세일":
            df_c = df_erp[(df_erp['wk'] == wk) & df_erp['거래처대분류'].isin(['할인점', '도매', '대리점', '홀세일'])]
        elif chan == "특판":
            df_c = df_erp[(df_erp['wk'] == wk) & df_erp['거래처대분류'].isin(['특판', '업소용'])]
        
        val = int(df_c['실판매금액'].sum())
        chan_data[wk] = val
        total_val += val
    chan_data["total"] = total_val
    channel_list.append(chan_data)

D["channel"] = channel_list

# 5. D.product (품목 대분류)
prod_sales = {}
for wk in WEEKS:
    df_wk = df_erp[df_erp['wk'] == wk]
    for cat in ['오븐', '냉장고', '전기포트', '토스터', '커피머신', '기타']:
        if cat not in prod_sales:
            prod_sales[cat] = {"name": cat}
        val = int(df_wk[df_wk['category'] == cat]['실판매금액'].sum())
        prod_sales[cat][wk] = val

product_list = []
for cat, data in prod_sales.items():
    tot = sum(data[wk] for wk in WEEKS if wk in data)
    data["total"] = tot
    product_list.append(data)

product_list = sorted(product_list, key=lambda x: x["total"], reverse=True)
D["product"] = product_list

# 6. D.topcodes (6월 전체 TOP 10)
df_codes = df_erp.groupby('품목명')['실판매금액'].sum().reset_index()
df_codes = df_codes.sort_values(by='실판매금액', ascending=False).head(10)
topcodes_list = []
for idx, row in df_codes.iterrows():
    topcodes_list.append({"name": row['품목명'], "rev": int(row['실판매금액'])})

D["topcodes"] = topcodes_list

# 7. D.budget (6월 전체 광고 예산 대비 3주차까지 집행률)
total_spent = int(df_ad['광고비'].sum())
budget_obj = {
    "total": 48100000,
    "spent": total_spent,
    "rate": round(total_spent / 48100000, 3)
}
D["budget"] = budget_obj

# 8. D.corr
weekly_metrics = {
    "cost": [],
    "adrev": [],
    "erp": [],
    "roas": []
}
for wk in WEEKS:
    df_ad_wk = df_ad[df_ad['wk'] == wk]
    df_erp_wk = df_erp[df_erp['wk'] == wk]
    cost = df_ad_wk['광고비'].sum()
    adrev = df_ad_wk['매출(매체)'].sum()
    erp = df_erp_wk['실판매금액'].sum()
    roas = adrev / cost if cost > 0 else 0
    weekly_metrics["cost"].append(round(cost / 1e6, 2))
    weekly_metrics["adrev"].append(round(adrev / 1e6, 2))
    weekly_metrics["erp"].append(round(erp / 1e8, 2))
    weekly_metrics["roas"].append(round(roas, 1))

# 일별 데이터를 기반으로 21일 피어슨 상관계수 산출
daily_dfs = []
for d in pd.date_range(start=W1_start, end=W3_end):
    df_a_d = df_ad[df_ad['날짜'] == d]
    df_e_d = df_erp[df_erp['처리일자'] == d]
    df_e_d_online = df_e_d[df_e_d['거래처대분류'].isin(['온라인', '자사몰', '쇼핑몰'])]
    
    cost = df_a_d['광고비'].sum()
    adrev = df_a_d['매출(매체)'].sum()
    roas = adrev / cost if cost > 0 else 0
    
    daily_dfs.append({
        "광고비": cost,
        "노출": df_a_d['노출'].sum(),
        "클릭": df_a_d['클릭'].sum(),
        "전환": df_a_d['클릭'].count(),
        "광고매출": adrev,
        "ROAS": roas,
        "전체실매출": df_e_d['실판매금액'].sum(),
        "온라인실매출": df_e_d_online['실판매금액'].sum()
    })
df_daily_corr = pd.DataFrame(daily_dfs)
corr_matrix = df_daily_corr.corr(method='pearson').fillna(0).round(2).values.tolist()

online_erp = []
ad_rev = []
rate_online = []
for wk in WEEKS:
    df_e_online = df_erp[(df_erp['wk'] == wk) & df_erp['거래처대분류'].isin(['온라인', '쇼핑몰', '자사몰'])]
    df_a = df_ad[df_ad['wk'] == wk]
    o_erp = df_e_online['실판매금액'].sum()
    a_rev = df_a['매출(매체)'].sum()
    online_erp.append(round(o_erp / 1e6, 1))
    ad_rev.append(round(a_rev / 1e6, 1))
    rate_online.append(round((a_rev / o_erp) * 100, 1) if o_erp > 0 else 0)

D["corr"] = {
    "heatmap": {
        "labels": ['광고비','노출','클릭','전환','광고매출','ROAS','전체실매출','온라인실매출'],
        "M": corr_matrix,
        "hi": [
            [5,6,'a'],[6,5,'a'],
            [0,6,'b'],[1,6,'b'],[2,6,'b'],[3,6,'b'],
            [0,7,'c'],[1,7,'c'],[2,7,'c'],[3,7,'c']
        ]
    },
    "weekly": weekly_metrics,
    "contrib": {
        "online_erp": online_erp,
        "ad_rev": ad_rev,
        "rate_online": rate_online
    }
}

# 9. 프로모션 일정 Gantt 매핑
D["promos"] = [
    {"name": "롯데 띵삼 위크", "ch": "롯데몰", "start": -18, "end": 0, "note": "롯데 띵삼 위크 종료", "color": "#5a9bf0"},
    {"name": "29CM 밀크프로머", "ch": "29CM", "start": -11, "end": 0, "note": "MFF02 단독전 종료", "color": "#e0a93b"},
    {"name": "신세계V 커피머신", "ch": "신세계몰", "start": -4, "end": 2, "note": "커피머신 연합전 종료", "color": "#8b5cf6"},
    {"name": "GDN SSG 프로모션", "ch": "SSG닷컴", "start": 10, "end": 16, "note": "SSG 프로모션 라이브 진행", "color": "#10b981"},
    {"name": "29CM 협력광고", "ch": "29CM", "start": 17, "end": 23, "note": "29CM 협력광고 기획전 진행", "color": "#f59e0b"},
    {"name": "썸머페스타", "ch": "자사몰/스토어", "start": 3, "end": 32, "note": "썸머페스타 프로모션 진행 中", "color": "#e8546b"}
]

D["events"] = [
    {"di": 3, "mark": "A", "color": "#e8546b", "title": "썸머페스타 개시 · 오븐 반등", "rev": 57896000, "detail": "썸머페스타 및 식기세척기 카탈로그 GFA 개시로 일매출 57.9M 반등.", "date": "6/01"},
    {"di": 9, "mark": "B", "color": "#5a9bf0", "title": "빌트인 가전 결제 집중", "rev": 61429000, "detail": "인덕션 및 빌트인 오븐 패키지 등 온라인 대량 결제 유입으로 일매출 61.4M 피크.", "date": "6/07"},
    {"di": 16, "mark": "C", "color": "#8b5cf6", "title": "주말 백화점/온라인 수요 집중", "rev": 50684200, "detail": "전자동 커피머신 및 FAB30 냉장고 등 고관여 품목 결제로 일매출 50.7M 달성.", "date": "6/14"}
]

# 10. 프로모션 성과 집계 (썸머페스타 동적 집계)
df_ad_summer = df_ad[df_ad['캠페인'].astype(str).str.contains('썸머|페스타|PA프로모션', na=False)]
adcost_summer = round(df_ad_summer['광고비'].sum() / 1e6, 1)
adrev_summer = round(df_ad_summer['매출(매체)'].sum() / 1e6, 1)
roas_summer = round(adrev_summer / adcost_summer, 1) if adcost_summer > 0 else 0.0

df_erp_summer = df_erp[(df_erp['처리일자'] >= '2026-06-01') & (df_erp['거래처대분류'].isin(['온라인', '본사몰', '자사몰', '스마트스토어', '쇼핑몰', '본사']))]
erp_summer = round(df_erp_summer['실판매금액'].sum() / 1e6, 1)
cnt_summer = int(df_erp_summer['수량'].sum())
daily_summer = round(erp_summer / 18, 1) # 6/1 ~ 6/18 총 18일

D["promoPerf"] = [
    {
        "name": "롯데 띵삼 위크",
        "period": "5/11~5/29",
        "color": "#5a9bf0",
        "erp": 118.0,
        "daily": 6.2,
        "cnt": 107,
        "top": "업소용 오븐 93M · 빌트인 11M · 냉동냉장고 9M",
        "adch": "메타 DA+GFA",
        "adcost": 10.9,
        "adrev": 97.6,
        "roas": 9.0,
        "note": "롯데 띵삼 위크 종료. 누적 매출 1억 1800만 원 달성"
    },
    {
        "name": "29CM 밀크프로머",
        "period": "5/18~5/29",
        "color": "#e0a93b",
        "erp": 47.7,
        "daily": 4.0,
        "cnt": 149,
        "top": "커피머신 9M · 소형1 7M · 반죽기 4M",
        "adch": "GFA 카탈로그",
        "adcost": 1.8,
        "adrev": 7.7,
        "roas": 4.3,
        "note": "MFF02 출시기념 단독전 종료. 커피·소형가전 위주 볼륨 형성"
    },
    {
        "name": "신세계V 커피머신",
        "period": "5/25~5/31",
        "color": "#8b5cf6",
        "erp": 4.1,
        "daily": 0.6,
        "cnt": 20,
        "top": "전기포트 2.5M · 토스터 1.1M · 커피머신 0.5M",
        "adch": "-",
        "adcost": 0.0,
        "adrev": 0.0,
        "roas": 0.0,
        "note": "신세계몰 커피머신 연합전 종료. 최종 누적 매출 413만 원 달성"
    },
    {
        "name": "썸머페스타",
        "period": "6/01~6/18*",
        "color": "#e8546b",
        "erp": erp_summer,
        "daily": daily_summer,
        "cnt": cnt_summer,
        "top": "오븐 60.1M · FAB 냉장고 13.9M · 전기포트 10.4M",
        "adch": "네이버/메타/구글",
        "adcost": adcost_summer,
        "adrev": adrev_summer,
        "roas": roas_summer,
        "note": "6월 썸머페스타 프로모션 진행 중. 자사몰/스토어 온라인 매출 볼륨 리드"
    }
]

print("2. JSON 데이터 객체 D 생성 완료.")

# 3. HTML 파일 로드 및 문자열 갈아끼우기
with open(html_template_path, "r", encoding="utf-8") as f:
    html_content = f.read()

# 분석기간 텍스트 헤더 업데이트 (6W3용)
html_content = html_content.replace(
    '<div class="dt mono">2026.05.01 – 05.22</div>',
    '<div class="dt mono">2026.05.29 – 06.18</div>'
)

# 대시보드 메인 타이틀 업데이트 (5월 -> 6월)
html_content = html_content.replace(
    '<title>스메그 5월 주간 통합 대시보드</title>',
    '<title>스메그 6월 주간 통합 대시보드</title>'
)
html_content = html_content.replace(
    '스메그 코리아 · 5월 주간 통합 대시보드',
    '스메그 코리아 · 6월 주간 통합 대시보드'
)

# a. const D = { ... } 치환
d_pattern = r"const D = \{.*?const won="
new_d_block = "const D = " + json.dumps(D, ensure_ascii=False, indent=2) + ";\n\nconst won="
html_content = re.sub(d_pattern, new_d_block, html_content, flags=re.DOTALL)

# b. 주차 칩 필터 버튼 업데이트
chip_area_pattern = r'<div class="filter-group" id="wkFilter">.*?</div>'
new_chip_area = """<div class="filter-group" id="wkFilter">
  <button class="chip on" data-wk="all">전체</button>
  <button class="chip" data-wk="W1">W1 · 5/29~06/04</button>
  <button class="chip" data-wk="W2">W2 · 6/05~11</button>
  <button class="chip" data-wk="W3">W3 · 6/12~18</button>
</div>"""
html_content = re.sub(chip_area_pattern, new_chip_area, html_content, flags=re.DOTALL)

# c. Gantt 일수 days=22 -> days=21 축소 (wkLines=[7,14]는 템플릿과 동일하게 유지)
html_content = html_content.replace("const days=22", "const days=21")

# d. renderCorr 차트 레이블 6월 주차로 변경
old_labels = "labels:['W1 (5/01~07)','W2 (5/08~14)','W3 (5/15~22)'],"
new_labels = "labels:['W1 (5/29~6/04)','W2 (6/05~11)','W3 (6/12~18)'],"
html_content = html_content.replace(old_labels, new_labels)

# e. 6월 3주차 전체 기간 주차 정의 푸터 메시지 업데이트
html_content = html_content.replace(
    "W1 5/01~07 · W2 5/08~14 · W3 5/15~22 (광고리포트 주차 구분에 정렬)",
    "W1 5/29~06/04 · W2 6/05~11 · W3 6/12~18 (광고리포트 주차 구분에 정렬)"
)

# f. 핵심 발견(인사이트) 코멘트 업데이트 (6W3용)
corr_find_pattern = r"document\.getElementById\('corrFind'\)\.innerHTML=`.*?`;"
new_corr_find = """document.getElementById('corrFind').innerHTML=`
    <li class="pos"><b>[상업 오븐 매출 방어] W2 전환 효율 극대화 및 W3 직접 광고 성과 수성</b> — 오븐 카테고리는 W3 전체 ERP 매출 하락(-20.8%)에도 불구하고, 검색광고 중심의 매체 광고매출(41.9M)과 ROAS(21.6x)를 견고하게 지탱하며 견인차 역할을 유지함.</li>
    <li class="neu"><b>[냉장고 실구매 지연 극복] W3 전환 매출 급감 대응 혜택형 소재 교체</b> — 냉장고 품목 W3 광고 클릭 유입은 7% 증가했으나 전환 매출이 거의 발생하지 않아 ERP 매출이 14.5M으로 급감함. 단순 노출형 배너 예산을 20% 절감하고 기획전 상세 혜택 소구형 소재로 즉시 전환함.</li>
    <li class="pos"><b>[커피머신 매체 트래킹 유실 증명] W3 매출 24.8M 달성 및 ROPO 기여</b> — 커피머신 광고비(+21.7%) 및 클릭(+20.6%) 증액이 실제 ERP 매출 상승(W3 2,484만 원, 51건)을 성공적으로 이끌었으나, 제휴몰 결제 이탈로 매체 전환액은 0원으로 잡힘. 강력한 ROPO 기여를 확인하여 예산 수성 확정.</li>
    <li class="neg"><b>[DA 광고 피로도 대응] W3 메타/GFA 성과 폭락에 따른 타겟 모수 리프레쉬</b> — W2에 메타 ROAS 30.99x로 정점을 찍은 후 W3에 3.86x로 급락하며 기타 가전 매출이 60% 급감함. 타겟 모수를 유사 타겟(3~5%)으로 넓히고, 썸머 테마 제품 이미지 소재 교체를 완료함.</li>
    <li class="neu"><b>[구글 SA 0원 지출 방어] 키워드 매칭 유형 축소 및 비효율 검색어 제외</b> — 구글 SA(검색광고)의 누적 무전환 효율에 대응하기 위해 광고 그룹 검색어 매칭 유형을 '확장'에서 '구문/일치'로 제한하고, 부정 쿼리(중고, 렌탈 등) 제외어 필터링 처리를 진행함.</li>`;"""
html_content = re.sub(corr_find_pattern, new_corr_find, html_content, flags=re.DOTALL)

# 4. 파일 쓰기
with open(analysis_out, "w", encoding="utf-8") as f:
    f.write(html_content)
print(f"3. 성공적으로 최종 분석 대시보드 저장 완료: {analysis_out}")

with open(smeg_out, "w", encoding="utf-8") as f:
    f.write(html_content)
print(f"4. 성공적으로 기준 템플릿 대시보드 업데이트 완료: {smeg_out}")

print("=== 대시보드 자동 빌드 및 갱신 성공 ===")
