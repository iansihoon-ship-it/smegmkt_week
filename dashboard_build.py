import pandas as pd
import numpy as np
import os
import json
import re

# 1. 파일 경로 설정
base_dir = r"c:\Users\smegkorea\smegkorea\mkt_week\data"
erp_path = os.path.join(base_dir, "통합매출현황_5_4w.xlsx")
ad_path = os.path.join(base_dir, "(HM) 스메그_5월 주간리포트_260529.xlsx")

html_template_path = r"c:\Users\smegkorea\smegkorea\mkt_week\smeg_5_3w_dashboard.html"
analysis_out = r"c:\Users\smegkorea\smegkorea\mkt_week\analysis_dashboard.html"
smeg_out = r"c:\Users\smegkorea\smegkorea\mkt_week\smeg_5_3w_dashboard.html"

print("1. 데이터 파일 로딩 시작 (대용량 광고 리포트 로딩으로 수십 초 소요될 수 있음)...")
df_erp_raw = pd.read_excel(erp_path, sheet_name="Sheet1")
# openpyxl 락 현상을 차단하기 위해 calamine 대신 일반 pandas 로딩 사용
df_ad_raw = pd.read_excel(ad_path, sheet_name="RAW")
print("데이터 파일 로딩 완료.")

# 2. 데이터 전처리
df_erp = df_erp_raw[df_erp_raw['구분'].isin(['판매', '추가판매', '판매취소'])].copy()
df_erp['처리일자'] = pd.to_datetime(df_erp['처리일자'])

df_ad = df_ad_raw.copy()
df_ad['날짜'] = pd.to_datetime(df_ad['날짜'])

# 카테고리 매핑 설정
df_ad['category'] = '기타'
df_ad.loc[df_ad['품목'].str.contains('냉장고|FAB', na=False), 'category'] = '냉장고'
df_ad.loc[df_ad['품목'].str.contains('전기포트', na=False), 'category'] = '전기포트'
df_ad.loc[df_ad['품목'].str.contains('오븐', na=False), 'category'] = '오븐'
df_ad.loc[df_ad['품목'].str.contains('토스터|토스트기', na=False), 'category'] = '토스터'

df_erp['category'] = '기타'
df_erp.loc[df_erp['품목그룹(중)'].str.contains('FAB', na=False), 'category'] = '냉장고'
df_erp.loc[df_erp['품목그룹(중)'].str.contains('전기포트', na=False), 'category'] = '전기포트'
df_erp.loc[df_erp['품목그룹(중)'].str.contains('오븐', na=False), 'category'] = '오븐'
df_erp.loc[df_erp['품목그룹(중)'].str.contains('토스트기', na=False), 'category'] = '토스터'

# 주차 정의 (W1: 5/1~7, W2: 5/8~14, W3: 5/15~22, W4: 5/23~29)
weeks_def = {
    'W1': ('2026-05-01', '2026-05-07'),
    'W2': ('2026-05-08', '2026-05-14'),
    'W3': ('2026-05-15', '2026-05-22'),
    'W4': ('2026-05-23', '2026-05-29')
}
WEEKS = ['W1', 'W2', 'W3', 'W4']

def get_week(date):
    for wk, (start, end) in weeks_def.items():
        if pd.Timestamp(start) <= date <= pd.Timestamp(end):
            return wk
    return None

df_erp['wk'] = df_erp['처리일자'].apply(get_week)
df_ad['wk'] = df_ad['날짜'].apply(get_week)

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
            df_m = df_wk[df_wk['매체/플랫폼'].str.contains('메타|Meta|GFA', na=False)]
        else:
            df_m = pd.DataFrame()
            
        if len(df_m) > 0 and df_m['광고비'].sum() > 0:
            cost = int(df_m['광고비'].sum())
            rev = int(df_m['매출(매체)'].sum())
            conv = int(df_m['클릭'].count()) # 임의값 대신 클릭 건수를 카운트한 전환 대리값
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

# 3. D.daily (5/01 ~ 5/29 필터링 적용 및 wk 속성 추가)
df_erp_daily = df_erp.groupby('처리일자')['실판매금액'].sum().reset_index()
df_ad_daily = df_ad.groupby('날짜')['광고비'].sum().reset_index()
df_daily = pd.merge(df_ad_daily, df_erp_daily, left_on='날짜', right_on='처리일자', how='outer')
df_daily['날짜'] = df_daily['날짜'].fillna(df_daily['처리일자'])
df_daily = df_daily.drop(columns=['처리일자'])
df_daily = df_daily.fillna(0)
df_daily = df_daily.sort_values(by='날짜')

# 5월 1일 ~ 5월 29일 범위로 엄격하게 필터링
df_daily = df_daily[(df_daily['날짜'] >= '2026-05-01') & (df_daily['날짜'] <= '2026-05-29')]

daily_list = []
for idx, row in df_daily.iterrows():
    wk_val = get_week(row['날짜'])
    daily_list.append({
        "date": row['날짜'].strftime("%m/%d"),
        "rev": int(row['실판매금액']),
        "cost": int(row['광고비']),
        "wk": wk_val if wk_val else "W4"  # 5/29 등 경계 보완용
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
    for cat in ['오븐', '냉장고', '전기포트', '토스터', '기타']:
        if cat not in prod_sales:
            prod_sales[cat] = {"name": cat}
        val = int(df_wk[df_wk['category'] == cat]['실판매금액'].sum())
        prod_sales[cat][wk] = val

product_list = []
for cat, data in prod_sales.items():
    tot = sum(data[wk] for wk in WEEKS if wk in data)
    data["total"] = tot
    product_list.append(data)

# 정밀 정렬 (매출 높은 카테고리 순 정렬)
product_list = sorted(product_list, key=lambda x: x["total"], reverse=True)
D["product"] = product_list

# 6. D.topcodes (실판매금액 TOP 10)
df_codes = df_erp.groupby('품목명')['실판매금액'].sum().reset_index()
df_codes = df_codes.sort_values(by='실판매금액', ascending=False).head(10)
topcodes_list = []
for idx, row in df_codes.iterrows():
    topcodes_list.append({"name": row['품목명'], "rev": int(row['실판매금액'])})

D["topcodes"] = topcodes_list

# 7. D.budget (예산 집행율)
total_spent = int(df_ad['광고비'].sum())
budget_obj = {
    "total": 48100000,
    "spent": total_spent,
    "rate": round(total_spent / 48100000, 3)
}
D["budget"] = budget_obj

# 8. D.corr (상관 및 기여 지표 계산)
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

weekly_df_list = []
for wk in WEEKS:
    df_a = df_ad[df_ad['wk'] == wk]
    df_e = df_erp[df_erp['wk'] == wk]
    df_e_online = df_e[df_e['거래처대분류'].isin(['온라인', '자사몰', '쇼핑몰'])]
    
    weekly_df_list.append({
        "광고비": df_a['광고비'].sum(),
        "노출": df_a['노출'].sum(),
        "클릭": df_a['클릭'].sum(),
        "전환": df_a['클릭'].count(),
        "광고매출": df_a['매출(매체)'].sum(),
        "ROAS": df_a['매출(매체)'].sum() / df_a['광고비'].sum() if df_a['광고비'].sum() > 0 else 0,
        "전체실매출": df_e['실판매금액'].sum(),
        "온라인실매출": df_e_online['실판매금액'].sum()
    })
df_wk_corr = pd.DataFrame(weekly_df_list)
corr_matrix = df_wk_corr.corr(method='pearson').fillna(0).round(2).values.tolist()

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

# 9. promos & events & promoPerf 데이터 정적 복사 및 4주차 인덱스(30) 패치
D["promos"] = [
    {"name": "네이버 가정의달", "ch": "네이버", "start": -11, "end": 7, "note": "가정의 달+세일, SDA 13~30%", "color": "#36c6b0"},
    {"name": "신세계 빅스마일", "ch": "신세계몰", "start": 10, "end": 16, "note": "SSG 빅스마일데이, SDA 8~50%", "color": "#c8102e"},
    {"name": "롯데 띵삼 위크", "ch": "롯데몰", "start": 10, "end": 28, "note": "롯데 띵삼 위크, LDA 10~48%", "color": "#5a9bf0"},
    {"name": "29CM 밀크프로머", "ch": "29CM", "start": 17, "end": 28, "note": "MFF02 출시기념 단독 할인", "color": "#e0a93b"},
    {"name": "신세계V 커피머신", "ch": "신세계몰", "start": 24, "end": 30, "note": "커피머신 연합전 할인", "color": "#8b5cf6"}
]

D["events"] = [
    {"di": 3, "mark": "A", "color": "#36c6b0", "title": "가정의 달 피크 · SA 입찰 20%↑", "rev": 72016800, "detail": "네이버 SA 핵심 키워드 입찰 상향 + 자사명 스토어 랜딩 추가 운영 효과.", "date": "5/04"},
    {"di": 10, "mark": "B", "color": "#c8102e", "title": "빅스마일데이 개시 · 최고 매출", "rev": 53637000, "detail": "신세계 SSG 빅스마일데이 시작일. 메타 DA 집행 본격화로 월간 최고 일매출.", "date": "5/11"},
    {"di": 7, "mark": "C", "color": "#5a9bf0", "title": "메타 DA 집행 시작 · 롯데 띵삼", "rev": 47431900, "detail": "가정의 달 종료 직후 메타 DA 신규 집행 및 띵삼 위크 시너지.", "date": "5/08"},
    {"di": 17, "mark": "D", "color": "#e0a93b", "title": "29CM 단독전 · GFA 리타게팅", "rev": 48814800, "detail": "29CM 밀크프로머 출시기념 단독 할인 + GFA 예산 증액.", "date": "5/18"},
    {"di": 14, "mark": "E", "color": "#e8546b", "title": "SS 효율 저점 · 오븐 구매 감소", "rev": 17245500, "detail": "네이버 쇼핑검색 효율 저점 및 오븐 비수기 영향.", "date": "5/15"},
    {"di": 25, "mark": "F", "color": "#8b5cf6", "title": "본사몰 오븐 매출 집중 · 29CM 연장", "rev": 50218700, "detail": "본사몰 오븐(ALFA43K 등 고단가) 매출 집중 및 29CM 기획전 연장 시너지로 4주차 최고 일매출 달성.", "date": "5/26"}
]

D["promoPerf"] = [
    {
        "name": "네이버 가정의 달",
        "period": "5/01~5/08*",
        "color": "#36c6b0",
        "erp": 189.0,
        "daily": 23.7,
        "cnt": 531,
        "top": "업소용 오븐 72M · 소형3 33M · 전기포트 28M",
        "adch": "네이버 SA+SS",
        "adcost": 4.9,
        "adrev": 80.4,
        "roas": 16.4,
        "note": "오븐·소형가전 폭넓게 견인. 광고효율 최고 구간"
    },
    {
        "name": "신세계 빅스마일",
        "period": "5/11~5/17",
        "color": "#c8102e",
        "erp": 132.0,
        "daily": 18.8,
        "cnt": 240,
        "top": "빌트인 55M · 소형1 29M · 냉동냉장고 19M",
        "adch": "메타 DA",
        "adcost": 5.6,
        "adrev": 78.1,
        "roas": 14.0,
        "note": "빌트인 고단가 품목 집중. DA 디스플레이가 트래픽 견인"
    },
    {
        "name": "롯데 띵삼 위크",
        "period": "5/11~5/29*",
        "color": "#5a9bf0",
        "erp": 118.0,
        "daily": 6.2,
        "cnt": 107,
        "top": "업소용 오븐 93M · 빌트인 11M · 냉동냉장고 9M",
        "adch": "메타 DA+GFA",
        "adcost": 10.9,
        "adrev": 97.6,
        "roas": 9.0,
        "note": "업소용 오븐(ALFA43)에 매출 집중. 객단가 높음"
    },
    {
        "name": "29CM 밀크프로머",
        "period": "5/18~5/29*",
        "color": "#e0a93b",
        "erp": 47.7,
        "daily": 4.0,
        "cnt": 149,
        "top": "커피머신 9M · 소형1 7M · 반죽기 4M",
        "adch": "GFA 카탈로그",
        "adcost": 1.8,
        "adrev": 7.7,
        "roas": 4.3,
        "note": "MFF02 출시기념 단독전. 소형·커피 라인 중심, 볼륨 소규모"
    },
    {
        "name": "신세계V 커피머신",
        "period": "5/25~5/31*",
        "color": "#8b5cf6",
        "erp": 3.2,
        "daily": 0.6,
        "cnt": 16,
        "top": "전기포트 2.1M · 토스터 1.1M · 커피머신 0.5M",
        "adch": "-",
        "adcost": 0.0,
        "adrev": 0.0,
        "roas": 0.0,
        "note": "신세계몰 커피머신 연합전. 행사 초기 온라인 매출 발생 (가용데이터 5/29까지)"
    }
]

# 4주차 데이터 반영 완료
print("2. JSON 데이터 객체 D 생성 완료.")
# 3. HTML 파일 로드 및 문자열 갈아끼우기
with open(html_template_path, "r", encoding="utf-8") as f:
    html_content = f.read()

# 분석기간 텍스트 헤더 업데이트
html_content = html_content.replace(
    '<div class="dt mono">2026.05.01 – 05.22</div>',
    '<div class="dt mono">2026.05.01 – 05.29</div>'
)

# a. const D = { ... } 치환
# 3주차 대시보드의 D 선언 위치부터 won 헬퍼 함수 직전까지 교체
d_pattern = r"const D = \{.*?const won="
new_d_block = "const D = " + json.dumps(D, ensure_ascii=False, indent=2) + ";\n\nconst won="
html_content = re.sub(d_pattern, new_d_block, html_content, flags=re.DOTALL)

# b. W4 칩 필터 버튼 추가 (중복 방지를 위한 안전 체크)
if 'data-wk="W4"' not in html_content:
    old_chips = '<button class="chip" data-wk="W3">W3 · 5/15~22</button>'
    new_chips = '<button class="chip" data-wk="W3">W3 · 5/15~22</button>\n  <button class="chip" data-wk="W4">W4 · 5/23~29</button>'
    html_content = html_content.replace(old_chips, new_chips)

# c. Gantt 일수 days=22 -> days=31 확장 및 wkLines=[7,14] -> wkLines=[7,14,22,29] 갱신
html_content = html_content.replace("const days=22", "const days=31")
html_content = html_content.replace("const wkLines=[7,14]", "const wkLines=[7,14,22,29]")

# d. WEEKS 리스트 W4 추가
html_content = html_content.replace("const WEEKS=['W1','W2','W3']", "const WEEKS=['W1','W2','W3','W4']")

# e. renderCorr 차트 레이블 W4 주차 추가
old_labels = "labels:['W1 (5/01~07)','W2 (5/08~14)','W3 (5/15~22)'],"
new_labels = "labels:['W1 (5/01~07)','W2 (5/08~14)','W3 (5/15~22)','W4 (5/23~29)'],"
html_content = html_content.replace(old_labels, new_labels)

# f. getWeekdayData 함수의 주차 슬라이싱 인덱스를 wk 매칭 필터링으로 변경
old_get_weekday = """function getWeekdayData(wk) {
  let filtered = [];
  if (wk === 'all') {
    filtered = D.daily;
  } else {
    let startIdx = wk === 'W1' ? 0 : wk === 'W2' ? 7 : 14;
    let endIdx = wk === 'W1' ? 6 : wk === 'W2' ? 13 : 21;
    filtered = D.daily.slice(startIdx, endIdx + 1);
  }"""
new_get_weekday = """function getWeekdayData(wk) {
  let filtered = wk === 'all' ? D.daily : D.daily.filter(d => d.wk === wk);"""

html_content = html_content.replace(old_get_weekday, new_get_weekday)

# g. 5월 4주차 전체 기간 주차 정의 푸터 메시지 업데이트
html_content = html_content.replace(
    "W1 5/01~07 · W2 5/08~14 · W3 5/15~22 (광고리포트 주차 구분에 정렬)",
    "W1 5/01~07 · W2 5/08~14 · W3 5/15~22 · W4 5/23~29 (광고리포트 주차 구분에 정렬)"
)

# h. 기여율 차트(contribChart) 레이블 W4 주차 추가
html_content = html_content.replace(
    "data:{labels:['W1','W2','W3'],datasets:[",
    "data:{labels:['W1','W2','W3','W4'],datasets:["
)

# 4. 파일 쓰기
with open(analysis_out, "w", encoding="utf-8") as f:
    f.write(html_content)
print(f"3. 성공적으로 최종 분석 대시보드 저장 완료: {analysis_out}")

with open(smeg_out, "w", encoding="utf-8") as f:
    f.write(html_content)
print(f"4. 성공적으로 기준 템플릿 대시보드 업데이트 완료: {smeg_out}")

print("=== 대시보드 자동 빌드 및 갱신 성공 ===")

