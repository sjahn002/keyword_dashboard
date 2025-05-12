import streamlit as st
import pandas as pd
import plotly.express as px
import re
import os
import plotly.graph_objects as go

# 1. 파일 업로드
st.set_page_config(layout="wide")
st.title("SEO 키워드 2x2 매트릭스 대시보드")

# 기본 샘플 데이터 경로
SAMPLE_DATA_PATH = os.path.join(os.path.dirname(__file__), "sample_data", "sample.xlsx")

# 2. 데이터 통합 및 전처리
dfs = []

# 샘플 데이터 로드
if os.path.exists(SAMPLE_DATA_PATH):
    df = pd.read_excel(SAMPLE_DATA_PATH)
    dfs.append(df)
    st.info("기본 샘플 데이터를 사용합니다.")
else:
    st.error(f"샘플 데이터 파일을 찾을 수 없습니다: {SAMPLE_DATA_PATH}")
    st.stop()

st.markdown("""
#### 1. 연관 키워드 엑셀 파일 여러 개 업로드
- 네이버 키워드 도구에서 추출한 여러 엑셀 파일을 업로드하세요.
- 파일 내 '연관키워드' 컬럼이 반드시 존재해야 합니다.
- 파일을 업로드하면 현재 표시된 샘플 데이터가 대체됩니다.
""")

uploaded_files = st.file_uploader(
    "엑셀 파일 업로드 (여러 개 가능)",
    type=["xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    dfs = []  # 샘플 데이터 초기화
    for file in uploaded_files:
        df = pd.read_excel(file)
        dfs.append(df)
    st.info("업로드된 파일로 데이터가 업데이트되었습니다.")

combined_df = pd.concat(dfs, ignore_index=True)

# 숫자 컬럼 전처리
def clean_numeric(col):
    col = col.astype(str).str.replace(',', '', regex=False)
    col = col.replace('< 10', '5')
    return pd.to_numeric(col, errors='coerce').fillna(0).astype(int)

numeric_columns = ['월간검색수(PC)', '월간검색수(모바일)', '월평균클릭수(PC)', '월평균클릭수(모바일)', '월평균노출 광고수']
for col in numeric_columns:
    if col in combined_df.columns:
        combined_df[col] = clean_numeric(combined_df[col])

for col in ['월평균클릭률(PC)', '월평균클릭률(모바일)']:
    if col in combined_df.columns:
        combined_df[col] = combined_df[col].astype(str).str.replace('%', '', regex=False)
        combined_df[col] = pd.to_numeric(combined_df[col], errors='coerce').fillna(0)

combined_df['총 검색수'] = combined_df['월간검색수(PC)'] + combined_df['월간검색수(모바일)']
combined_df['총 클릭수'] = combined_df['월평균클릭수(PC)'] + combined_df['월평균클릭수(모바일)']

# 키워드 정규화 함수
def normalize_keyword(keyword):
    keyword = str(keyword).lower()
    keyword = re.sub(r'[^\w\s]', '', keyword)
    keyword = re.sub(r'\s+', ' ', keyword)
    return keyword.strip()

combined_df['연관키워드'] = combined_df['연관키워드'].apply(normalize_keyword)
combined_df = combined_df.drop_duplicates(subset=['연관키워드'])

# 3. 분류 필터 정의
suitable_filters = {
    '유아/초등 타겟 영어 교육': r'(?=.*?(유아|아기|어린이|아동|초등|유치원|영유아|초1|초2|초3|초4|초5|초6|키즈|1세|2세|3세|4세|5세|6세|7세|8세|9세|10세|11세|12세|1살|2살|3살|4살|5살|6살|7살|8살|9살|10살|11살|12살|개월|예비초|영어유치원|학년|방과후|엄마|아이).*?영어)|(?=.*?영어.*?(유아|아기|어린이|아동|초등|유치원|영유아|초1|초2|초3|초4|초5|초6|키즈|1세|2세|3세|4세|5세|6세|7세|8세|9세|10세|11세|12세|1살|2살|3살|4살|5살|6살|7살|8살|9살|10살|11살|12살|개월|예비초|영어유치원|학년|방과후|엄마|아이))',
    '미국 교육커리큘럼': r'(?=.*?(미국|공교육|교과서|커리큘럼|IXL|북미|아메리칸|미국식|교육과정|학제|영어권|미국교과|미국식교육|미교|미국학교|미국초등|미국유치원|미교리딩|미국교과서리딩|미국교과서읽는리딩단계))',
    '유아/초등 영어 콘텐츠 (G1 이상)': r'(?=.*?(영어문법|영문법|영어단어|영단어|영어교구|영어학습지|영어교재|영어프로그램|영어앱|영어책|원서|영어독서|영어발음|영어학습|영어공부).*?(유아|초등|어린이|아동|키즈|아이))',
    '유아/초등 영어 콘텐츠 (Pre-K, K)': r'(?=.*?(영어놀이|영어동요|영어동화|알파벳|사이트워드|파닉스|영어게임|영어애니메이션|영어학습게임|영어만화|애니메이션영어))',
    '국제학교/글로벌 교육': r'(?=.*?(국제학교|인터내셔널스쿨|글로벌학교|국제교육|글로벌교육|외국인학교|온라인국제학교|채드윅|스쿨링|해외학교|글로벌스쿨|국제초등학교|국제유치원|외국교육|외국학교|국제교과|IB|국제학생|글로벌인재|국제교육과정|글로벌교육과정|인터내셔널교육|인터내셔널스쿨|온라인스쿨|캐나다온라인고등학교|로렐스프링스스쿨|로렐스프링스|LAURELSPRINGSSCHOOL|ICNA))',
    '조기 영어 교육': r'(?=.*?(조기영어|조기교육|영어조기교육|영어조기|영어조기학습|영어조기학습법|영어조기학습방법|영어조기학습프로그램|영어조기학습교재|영어조기학습교구|영어조기학습앱|영어조기학습어플|영어조기학습방문|영어조기학습방문교사|영어조기학습방문학습|영어조기학습방문교육|영어조기학습방문프로그램|영어조기학습방문교재|영어조기학습방문교구|영어조기학습방문앱|영어조기학습방문어플))',
    '프리미엄 지역 유아/초등 영어': r'(?=.*?(강남|대치|목동|청담|삼성동|도곡|양재|개포|송파|잠실|분당|판교|동탄|광교|송도|위례|일산|하남).*?(초등|유아|어린이|아동|키즈|영어|영어학원|영어교육|영어학습|영어공부))'
}

unsuitable_filters = {
    '타겟 연령 외': r'.*(중학|고등|대학|성인|직장인|노인|50대|40대|30대|20대|청소년|중등|고1|고2|고3|중1|중2|중3).*(?!.*(초등|유아|어린이|아동|키즈)).*(?!.*(국제학교|온라인국제학교))',
    '교육 분야 외': r'.*(일본어|중국어|프랑스어|스페인어|독일어|베트남어|태국어|러시아어|아랍어).*(?!.*영어|.*글로벌)|.*(수학|과학|사회|국어|한국어|한국사|물리|화학|생물|지구과학|역사|문학|한문|컴퓨터|코딩|프로그래밍|경제|미술|체육|음악|무용|태권도|발레).*(?!.*영어|.*국제학교|.*온라인국제학교|.*글로벌)',
    '시험/자격증 관련': r'.*(토익|토플|아이엘츠|오픽|텝스|HSK|JLPT|DELE|DELF|TSC|JPT|TOPIK|EJU|AP|수능|내신|모의고사|TOEIC|TOEFL|IELTS|OPIC|TEPS).*(?!.*초등|.*유아|.*어린이|.*아동|.*키즈|.*국제학교|.*온라인국제학교)|.*(SAT|SSAT).*(?!.*초등|.*유아|.*어린이|.*아동|.*키즈|.*국제학교)',
    '오프라인 중심': r'.*(방문학습|방문교사|대면|현장체험학습|체험학습|체험활동|캠프|기숙).*(?!.*온라인|.*화상|.*인터넷|.*국제학교|.*온라인국제학교|.*초등영어)',
    '특수 교육 관련': r'.*(검정고시|재수|편입|입시|윈터스쿨|서머스쿨|논술|특목고|영재|올림피아드|경시대회|대회|마이스터|특성화).*(?!.*초등|.*유아|.*어린이|.*아동|.*키즈|.*국제학교|.*온라인국제학교|.*영어)',
    '브랜드명': r'.*(눈높이|구몬|웅진|대교|YBM|YBM토익|튼튼영어|윤선생|EBSe|와이즈만|라이즈|하바|크라운|뽀로로|핑크퐁|몬테소리|발도르프|키즈랜드|숲유치원|이투스|메가|대성|스카이에듀|강남구청|시원스쿨).*(?!.*국제학교|.*온라인국제학교|.*초등영어|.*유아영어)',
    '비프리미엄 지역 및 업무/대학 지역': r'.*(노원구|도봉구|강동구|은평구|중랑구|광화문|여의도|종로|홍대|신촌|용산|광진구|구로구|금천구|서대문구|성동구|성북구|영등포구|동작구|관악구|양천구|강서구|마포구).*(?!.*(초등|유아|어린이|아동|키즈|국제학교|인터내셔널))',
    '기타 상품': r'(?=.*(육아|여행|장난감|놀이공원|인형|블럭|퍼즐|레고|책장|가구|영양제|건강|운동|다이어트))(?!.*(?:영어|영어교육|영어학습|영어공부|영어학원|영어교재|영어교구|영어프로그램|영어앱|영어학습지|영어동화|영어동요|영어책|영어독서|영어발음|영어문법|영어단어|국제학교|온라인국제학교))',
    '직장인/성인/비즈니스 타겟 키워드': r'.*(비즈니스영어|비지니스영어|강남역|역삼역|직장인영어|성인영어|영어과외알바|영어회화알바|영어학원창업|영어공부방창업|영어학원매매|영어PT|왕초보영어|기초영어|주말영어|토요일영어|종로영어|한달영어|평생영어|6개월영어|영어회화주말반|영어회화단기|비즈니스영어학원|비즈니스영어과외|비즈니스영어회화|비즈니스영어인강|직장인화상영어|영어가맹|영어학원가맹|영어학원체인점|영어프랜차이즈|이력서영어|면접영어|인터뷰영어|취업영어|스피킹|토킹|프리토킹|회사|직장|취업|면접|이력서|토요일|평일|평생|한달|6개월|알바|창업|매매|가맹|PT).*(?!(?:.*(?:초등|유아|어린이|아동|키즈|국제학교|인터내셔널학교)|^(?:초등|유아|어린이|아동|키즈|국제학교|인터내셔널학교).*))',
    '사전/번역 관련': r'.*(사전|번역|번역기|번역사|통역|통역사)'
}

additional_filters = {
    '온라인 영어 교육': r'(온라인|화상|인터넷|비대면|원격|디지털|스마트|태블릿|패드|앱|어플|홈스쿨|홈스쿨링|홈러닝|자기주도|자기주도학습|엄마표|e러닝|이러닝|인터넷강의|온라인강의|온라인수업|온라인학습|온라인교육|스마트러닝)(?=.*영어)',
    '영어 콘텐츠': r'(영어책|영어독서|영어발음|영어문법|영어단어|영어학습지|영어교재|영어교구|영어프로그램|영어앱|영어학습|영어공부)',
    '일반 영어 교육': r'(영어|원어민|영어학원|영어공부|영어학습|영어교육|영어수업|영어강의|영어과외|영어회화|영어인강|영어학습지|영어교재|영어교구|영어프로그램|영어앱|영어학습|영어공부)'
}

# 4. 키워드 분류
final_df = combined_df.copy()
final_df['키워드_분류'] = '미분류'
final_df['키워드_상세분류'] = '미분류'

# 부적합 키워드 필터 적용
for category, pattern in unsuitable_filters.items():
    mask = final_df['연관키워드'].str.contains(pattern, regex=True, na=False)
    final_df.loc[mask, '키워드_분류'] = '부적합'
    final_df.loc[mask, '키워드_상세분류'] = category

# 적합 키워드 필터 적용 (부적합이 아닌 것만)
for category, pattern in suitable_filters.items():
    mask = (final_df['연관키워드'].str.contains(pattern, regex=True, na=False)) & (final_df['키워드_분류'] == '미분류')
    final_df.loc[mask, '키워드_분류'] = '적합'
    final_df.loc[mask, '키워드_상세분류'] = category

# 확장 가능 키워드 필터 적용 (미분류만)
for category, pattern in additional_filters.items():
    mask = (final_df['연관키워드'].str.contains(pattern, regex=True, na=False)) & (final_df['키워드_분류'] == '미분류')
    final_df.loc[mask, '키워드_분류'] = '확장 가능 키워드'
    final_df.loc[mask, '키워드_상세분류'] = category

# 부적합 필터 재적용 (확장 가능 키워드 포함)
for category, pattern in unsuitable_filters.items():
    mask = (final_df['연관키워드'].str.contains(pattern, regex=True, na=False)) & (final_df['키워드_분류'] != '부적합')
    final_df.loc[mask, '키워드_분류'] = '부적합'
    final_df.loc[mask, '키워드_상세분류'] = category

# 5. 질적 분류
final_df['키워드_분류_질적'] = '미분류'
mask_purple = final_df['키워드_상세분류'].isin([
    '국제학교/글로벌 교육',
    '프리미엄 지역 유아/초등 영어'
])
final_df.loc[mask_purple, '키워드_분류_질적'] = '전략적 Sweet Spot'
mask_blue = final_df['키워드_상세분류'].isin([
    '미국 교육커리큘럼'
])
final_df.loc[mask_blue, '키워드_분류_질적'] = '특화 영역'
mask_red = final_df['키워드_상세분류'].isin([
    '유아/초등 타겟 영어 교육',
    '유아/초등 영어 콘텐츠 (Pre-K, K)',    
    '브랜드명'
])
final_df.loc[mask_red, '키워드_분류_질적'] = '타겟 경쟁 영역'
mask_expandable = final_df['키워드_상세분류'].isin([
    '영어 콘텐츠',
    '일반 영어 교육',
    '온라인 영어 교육'
])
final_df.loc[mask_expandable, '키워드_분류_질적'] = '확장 가능 키워드'
mask_junk = final_df['키워드_상세분류'].isin([
    '기타 상품',
    '비프리미엄 지역 및 업무/대학 지역'
])
final_df.loc[mask_junk, '키워드_분류_질적'] = '정크 키워드'
mask_off_target = final_df['키워드_상세분류'].isin([
    '교육 분야 외',
    '시험/자격증 관련', 
    '오프라인 중심',
    '직장인/성인/비즈니스 타겟 키워드',
    '타겟 연령 외',
    '특수 교육 관련',
    '사전/번역 관련'
])
final_df.loc[mask_off_target, '키워드_분류_질적'] = '타겟 외 경쟁 영역'

# 6. 통계 집계 및 트렙맵 시각화
gb = final_df.groupby('키워드_분류_질적')
classification_stats = gb.agg({
    '총 검색수': ['mean', 'count'],
    '총 클릭수': ['mean'],
    '월평균클릭률(PC)': 'mean',
    '월평균클릭률(모바일)': 'mean',
    '월평균노출 광고수': 'mean'
}).round(2)
classification_stats.columns = ['평균_검색수', '키워드_개수', '평균_클릭수', '평균_클릭률_PC', '평균_클릭률_모바일', '평균_노출광고수']
classification_stats = classification_stats.reset_index()

labels_kr = {
    '전략적 Sweet Spot': '전략적 Sweet Spot\n(Purple Ocean)',
    '특화 영역': '특화 키워드\n(Blue Ocean)',
    '타겟 경쟁 영역': '경쟁 키워드\n(Red Ocean)',
    '확장 가능 키워드': '확장 가능 키워드',
    '정크 키워드': '정크 키워드',
    '타겟 외 경쟁 영역': '타겟 외 경쟁 영역',
    '미분류': '미분류'
}
classification_stats['label'] = classification_stats['키워드_분류_질적'].map(labels_kr)
color_map = {
    '전략적 Sweet Spot': '#b39ddb',
    '특화 영역': '#90caf9',
    '타겟 경쟁 영역': '#ef9a9a',
    '확장 가능 키워드': '#fff59d',
    '정크 키워드': '#bdbdbd',
    '타겟 외 경쟁 영역': '#ffe082',
    '미분류': '#eeeeee'
}
classification_stats['color'] = classification_stats['키워드_분류_질적'].map(color_map)

# 예시용 분류명 매핑 (실제 분류명/통계로 대체)
area_defs = [
    # x0, y0, x1, y1, 분류명, 색상
    [0.5, 0.75, 1, 1, '타겟 경쟁 영역', '#ffebee'],
    [0.5, 0.5, 1, 0.75, '전략적 Sweet Spot', '#ede7f6'],
    [0.5, 0, 1, 0.5, '특화 영역', '#e3f2fd'],
    [0.25, 0.5, 0.5, 1, '확장 가능 키워드', '#fffde7'],
    [0, 0.5, 0.25, 1, '타겟 외 경쟁 영역', '#ffe082'],
    [0, 0, 0.5, 0.5, '정크 키워드', '#e0e0e0'],
]

# 각 영역별 중앙 좌표, 텍스트, hovertext 준비
data = []
for x0, y0, x1, y1, area_name, color in area_defs:
    stat = classification_stats[classification_stats['키워드_분류_질적'] == area_name].iloc[0] if not classification_stats[classification_stats['키워드_분류_질적'] == area_name].empty else {}
    
    # 해당 영역의 세부 카테고리와 키워드 개수 계산
    area_categories = final_df[final_df['키워드_분류_질적'] == area_name].groupby('키워드_상세분류').size()
    category_text = "<br>".join([f"{cat}: {count:,}개" for cat, count in area_categories.items()])
    
    # 영역 내 텍스트 구성
    area_text = f"<b><span style='font-size: 24px; color: black;'>{area_name}</span></b><br>"
    area_text += f"키워드: {stat.get('키워드_개수', '-'):,}개<br>"
    area_text += f"검색수: {stat.get('평균_검색수', '-'):,}<br>"
    area_text += f"클릭수: {stat.get('평균_클릭수', '-'):,}<br>"
    area_text += f"클릭률: {stat.get('평균_클릭률_PC', '-'):.1f}%<br>"
    area_text += f"광고수: {stat.get('평균_노출광고수', '-'):,}<br><br>"
    area_text += f"<b>세부 카테고리:</b><br>{category_text}"
    
    data.append(dict(
        x=(x0+x1)/2, y=(y0+y1)/2, x0=x0, y0=y0, x1=x1, y1=y1,
        area_name=area_name, color=color, area_text=area_text
    ))

fig = go.Figure()

# 사각형 영역 그리기
for d in data:
    fig.add_shape(
        type="rect",
        x0=d['x0'], y0=d['y0'], x1=d['x1'], y1=d['y1'],
        line=dict(color="black", width=2),
        fillcolor=d['color'],
        layer="below"
    )

# 영역 내 텍스트 배치
fig.add_trace(go.Scatter(
    x=[d['x'] for d in data],
    y=[d['y'] for d in data],
    text=[d['area_text'] for d in data],
    mode="text",
    textposition="middle center",
    hoverinfo="none",
    marker=dict(opacity=0),
    showlegend=False,
    textfont=dict(size=14)  # 텍스트 크기 조정
))

fig.update_xaxes(
    showticklabels=False, showgrid=False, zeroline=False,
    range=[0, 1], title_text="타겟 관련성(아이덴티티) →",
    title_font=dict(size=20)
)
fig.update_yaxes(
    showticklabels=False, showgrid=False, zeroline=False,
    range=[0, 1], title_text="↑ 확장성(트래픽)",
    title_font=dict(size=20)
)
fig.update_layout(
    width=1200, height=1200,  # 그래프 크기 증가
    margin=dict(l=40, r=40, t=40, b=40),
    plot_bgcolor="white"
)

st.plotly_chart(fig, use_container_width=True)

# 7. 분류별 샘플 키워드 표
st.subheader("분류별 샘플 키워드")

# 표시할 컬럼 정의
display_columns = [
    '연관키워드', 
    '키워드_상세분류',
    '총 검색수',
    '총 클릭수',
    '월평균클릭률(PC)',
    '월평균노출 광고수'
]

# 정렬 기준 선택
sort_by = st.selectbox(
    "정렬 기준 선택",
    ["총 검색수 (기본)", "월평균 클릭률", "총 클릭수", "월평균노출 광고수"],
    index=0
)

# 정렬 기준에 따른 컬럼 매핑
sort_column_map = {
    "총 검색수 (기본)": "총 검색수",
    "월평균 클릭률": "월평균클릭률(PC)",
    "총 클릭수": "총 클릭수",
    "월평균노출 광고수": "월평균노출 광고수"
}

# 중요도 순서 정의
importance_order = [
    '전략적 Sweet Spot',
    '특화 영역',
    '타겟 경쟁 영역',
    '확장 가능 키워드',
    '정크 키워드',
    '타겟 외 경쟁 영역',
    '미분류'
]

# 각 분류별로 데이터 표시 (중요도 순서대로)
for category in importance_order:
    category_df = final_df[final_df['키워드_분류_질적'] == category]
    
    if not category_df.empty:
        # 통계 계산
        avg_search = category_df['총 검색수'].mean()
        avg_clicks = category_df['총 클릭수'].mean()
        avg_ctr = category_df['월평균클릭률(PC)'].mean()
        avg_ads = category_df['월평균노출 광고수'].mean()
        
        # 통계 표시
        st.markdown(f"### {labels_kr.get(category, category)}")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("평균 검색수", f"{avg_search:,.0f}")
        with col2:
            st.metric("평균 클릭수", f"{avg_clicks:,.0f}")
        with col3:
            st.metric("평균 클릭률", f"{avg_ctr:.2f}%")
        with col4:
            st.metric("평균 노출광고수", f"{avg_ads:,.0f}")
        
        # 경쟁 키워드인 경우 적합/부적합으로 나누어 표시
        if category == '타겟 경쟁 영역':
            # 적합 키워드
            suitable_df = category_df[category_df['키워드_분류'] == '적합']
            if not suitable_df.empty:
                st.markdown("#### 적합 키워드")
                sorted_suitable = suitable_df.sort_values(
                    by=sort_column_map[sort_by],
                    ascending=False
                ).head(10)
                
                st.dataframe(
                    sorted_suitable[display_columns].style.format({
                        '총 검색수': '{:,.0f}',
                        '총 클릭수': '{:,.0f}',
                        '월평균클릭률(PC)': '{:.2f}%',
                        '월평균노출 광고수': '{:,.0f}'
                    }),
                    use_container_width=True
                )
            
            # 부적합 키워드
            unsuitable_df = category_df[category_df['키워드_분류'] == '부적합']
            if not unsuitable_df.empty:
                st.markdown("#### 부적합 키워드")
                sorted_unsuitable = unsuitable_df.sort_values(
                    by=sort_column_map[sort_by],
                    ascending=False
                ).head(10)
                
                st.dataframe(
                    sorted_unsuitable[display_columns].style.format({
                        '총 검색수': '{:,.0f}',
                        '총 클릭수': '{:,.0f}',
                        '월평균클릭률(PC)': '{:.2f}%',
                        '월평균노출 광고수': '{:,.0f}'
                    }),
                    use_container_width=True
                )
        else:
            # 다른 카테고리는 기존대로 표시
            sorted_df = category_df.sort_values(
                by=sort_column_map[sort_by],
                ascending=False
            ).head(10)
            
            st.dataframe(
                sorted_df[display_columns].style.format({
                    '총 검색수': '{:,.0f}',
                    '총 클릭수': '{:,.0f}',
                    '월평균클릭률(PC)': '{:.2f}%',
                    '월평균노출 광고수': '{:,.0f}'
                }),
                use_container_width=True
            )
        
        st.markdown("---")

# 8. 전체 데이터 다운로드
st.download_button(
    label="전체 분류 데이터 다운로드 (CSV)",
    data=final_df.to_csv(index=False, encoding='utf-8-sig'),
    file_name="키워드_질적분류_결과.csv",
    mime="text/csv"
) 