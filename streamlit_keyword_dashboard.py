import streamlit as st
import pandas as pd
import plotly.express as px
import re
import os
import plotly.graph_objects as go
import io
import streamlit.components.v1 as components

# 1. 파일 업로드
st.set_page_config(layout="wide", initial_sidebar_state="collapsed")

# CSS 스타일 추가
st.markdown("""
<style>
    /* 전체적인 폰트 및 여백 설정 */
    .stApp {
        font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, system-ui, Roboto, sans-serif;
        line-height: 1.6;
    }
    
    /* 제목 스타일 */
    h1 {
        font-size: clamp(1.5rem, 4vw, 2.5rem) !important;
        font-weight: 700 !important;
        margin-bottom: 1.5rem !important;
        color: #1E1E1E !important;
    }
    
    h2 {
        font-size: clamp(1.2rem, 3vw, 2rem) !important;
        font-weight: 600 !important;
        margin: 1.5rem 0 1rem 0 !important;
        color: #2C3E50 !important;
    }
    
    h3 {
        font-size: clamp(1rem, 2.5vw, 1.5rem) !important;
        font-weight: 600 !important;
        margin: 1rem 0 !important;
        color: #34495E !important;
    }
    
    /* 메트릭 카드 스타일 */
    .stMetric {
        background: #FFFFFF;
        border-radius: 12px;
        padding: 1rem;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        transition: all 0.3s ease;
    }
    
    .stMetric:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0,0,0,0.1);
    }
    
    /* 데이터프레임 스타일 */
    .dataframe {
        border-radius: 8px !important;
        overflow: hidden !important;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05) !important;
    }
    
    .dataframe th {
        background-color: #F8F9FA !important;
        color: #2C3E50 !important;
        font-weight: 600 !important;
        padding: 12px !important;
    }
    
    .dataframe td {
        padding: 10px !important;
        color: #4A4A4A !important;
    }
    
    /* 버튼 스타일 */
    .stButton>button {
        background-color: #4A90E2 !important;
        color: white !important;
        border-radius: 8px !important;
        padding: 0.5rem 1rem !important;
        border: none !important;
        font-weight: 500 !important;
        transition: all 0.3s ease !important;
    }
    
    .stButton>button:hover {
        background-color: #357ABD !important;
        transform: translateY(-1px) !important;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1) !important;
    }
    
    /* 파일 업로더 스타일 */
    .stFileUploader {
        border: 2px dashed #E0E0E0 !important;
        border-radius: 12px !important;
        padding: 1rem !important;
        background-color: #F8F9FA !important;
    }
    
    /* 알림 메시지 스타일 */
    .stAlert {
        border-radius: 8px !important;
        padding: 1rem !important;
    }
    
    /* 구분선 스타일 */
    hr {
        margin: 2rem 0 !important;
        border: none !important;
        border-top: 1px solid #E0E0E0 !important;
    }
    
    /* 반응형 컨테이너 */
    .stContainer {
        padding: 1rem !important;
    }
    
    @media (max-width: 768px) {
        .stContainer {
            padding: 0.5rem !important;
        }
        
        .stMetric {
            padding: 0.75rem !important;
        }
        
        /* 모바일에서 텍스트 크기 조정 */
        h1 {
            font-size: 1.5rem !important;
        }
        
        h2 {
            font-size: 1.2rem !important;
        }
        
        h3 {
            font-size: 1rem !important;
        }
        
        /* 모바일에서 데이터프레임 스타일 조정 */
        .dataframe {
            font-size: 0.9rem !important;
        }
        
        .dataframe th, .dataframe td {
            padding: 8px !important;
        }
    }
</style>
""", unsafe_allow_html=True)

st.title("SEO 키워드 분석기")

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
#### 대시보드 소개
본 대시보드는 네이버 검색 데이터를 활용한 키워드 분석 솔루션입니다. 키워드의 검색량, 클릭률, 경쟁도 등 핵심 지표를 분석하여 전략적 의사결정을 지원하며, 자동화된 키워드 분류와 시각화 기능을 통해 효율적인 SEO 전략을 수립할 수 있습니다.

[📚 상세 사용 가이드 보러가기](https://docs.google.com/document/d/1AVVoIKKelMUVIJydk6Xzmw1e2ibxBjDi4Le-VVoDch8/edit?usp=sharing)

#### 네이버 키워드 도구 안내
네이버 키워드 도구는 네이버 검색에서 자주 사용되는 키워드와 관련 검색량, 경쟁도 등의 정보를 제공하는 분석 도구입니다. 네이버 키워드 도구에서서 수집한 데이터를 대시보드에 업로드하시면 자동으로 분석이 진행됩니다.

#### 데이터 수집 방법
1. 성과형 디스플레이 광고 계정으로 로그인합니다.  
2. 네이버 광고 센터(https://manage.searchad.naver.com) 에 접속합니다.
3. '도구' 메뉴에서 '키워드 도구'를 선택합니다.
4. 분석을 원하는 키워드를 입력합니다.
5. '다운로드' 버튼을 클릭하여 데이터를 저장합니다.

#### 대시보드 활용 가이드
1. 수집한 엑셀 파일을 하단 업로드 영역에 드래그 앤 드롭하거나 'Browse files' 버튼을 클릭하여 직접 선택합니다.
2. 여러 파일을 동시에 업로드하실 수 있으며, 중복 키워드는 자동으로 제거됩니다.
3. 업로드된 데이터는 다음과 같이 자동 분석됩니다.
   - 키워드별 검색량, 클릭률, 경쟁도 분석
   - 카테고리별 키워드 분류 및 통계 산출
   - 카테고리별 주요 키워드 추출
4. 분석 결과는 카테고리별로 엑셀 파일로 다운로드하실 수 있습니다.

#### 핵심 기능
- 키워드 분석: 검색량, 클릭률, 경쟁도 등 주요 지표의 분석 및 시각화
- 자동 분류: 사전 정의된 규칙 기반의 키워드 자동 분류
- 데이터 시각화: 직관적인 차트와 그래프를 통한 데이터 표현
- 엑셀 내보내기: 카테고리별 분석 결과의 엑셀 파일 다운로드

#### 이용 시 주의사항
- 업로드하시는 엑셀 파일에는 반드시 '연관키워드' 컬럼이 포함되어야 합니다.
- 새로운 파일을 업로드하시면 기존 샘플 데이터는 자동으로 대체됩니다.
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
    '유아/초등 타겟 영어교육': r'(?=.*?(유아|아기|어린이|아동|초등|유치원|영유아|초1|초2|초3|초4|초5|초6|키즈|1세|2세|3세|4세|5세|6세|7세|8세|9세|10세|11세|12세|1살|2살|3살|4살|5살|6살|7살|8살|9살|10살|11살|12살|개월|예비초|영어유치원|학년|방과후|엄마|아이).*?영어)|(?=.*?영어.*?(유아|아기|어린이|아동|초등|유치원|영유아|초1|초2|초3|초4|초5|초6|키즈|1세|2세|3세|4세|5세|6세|7세|8세|9세|10세|11세|12세|1살|2살|3살|4살|5살|6살|7살|8살|9살|10살|11살|12살|개월|예비초|영어유치원|학년|방과후|엄마|아이))',
    '미국 교육 커리큘럼': r'(?=.*?(미국|공교육|교과서|커리큘럼|IXL|북미|아메리칸|미국식|교육과정|학제|영어권|미국교과|미국식교육|미교|미국학교|미국초등|미국유치원|미교리딩|미국교과서리딩|미국교과서읽는리딩단계))',
    'Pre-K, K 유아/초등 영어 콘텐츠': r'(?=.*?(영어놀이|영어동요|영어동화|알파벳|사이트워드|파닉스|영어게임|영어애니메이션|영어학습게임|영어만화|애니메이션영어))',
    '국제학교/글로벌 교육': r'(?=.*?(국제학교|인터내셔널스쿨|글로벌학교|국제교육|글로벌교육|외국인학교|온라인국제학교|채드윅|스쿨링|해외학교|글로벌스쿨|국제초등학교|국제유치원|외국교육|외국학교|국제교과|IB|국제학생|글로벌인재|국제교육과정|글로벌교육과정|인터내셔널교육|인터내셔널스쿨|온라인스쿨|캐나다온라인고등학교|로렐스프링스스쿨|로렐스프링스|LAURELSPRINGSSCHOOL|ICNA))',
    '프리미엄 학군 유아/초등 영어': r'(?=.*?(강남|대치|목동|청담|삼성동|도곡|양재|개포|송파|잠실|분당|판교|동탄|광교|송도|위례|일산|하남).*?(초등|유아|어린이|아동|키즈|영어|영어학원|영어교육|영어학습|영어공부))',
    '유아/초등 영어교육': r'(?=.*?(영어문법|영문법|영어단어|영단어|영어교구|영어학습지|영어교재|영어프로그램|영어앱|영어책|원서|영어독서|영어발음|영어학습|영어공부).*?(유아|초등|어린이|아동|키즈|아이))'
}

additional_filters = {
    '타겟 없는 온라인 영어 교육': r'(?=.*(온라인|화상|인터넷|비대면|원격|디지털|스마트|태블릿|패드|앱|어플|홈스쿨|홈스쿨링|홈러닝|자기주도|자기주도학습|엄마표|e러닝|이러닝|인터넷강의|온라인강의|온라인수업|온라인학습|온라인교육|스마트러닝))(?=.*영어)',
    '타겟 없는 영어 콘텐츠 (교재 등)': r'(?=.*(영어책|영어독서|영어발음|영어문법|영어단어|영어학습지|영어교재|영어교구|영어프로그램|영어앱|영어학습|영어공부))',
    '타겟없는 일반 영어 교육': r'(?=.*(영어|원어민|영어학원|영어공부|영어학습|영어교육|영어수업|영어강의|영어과외|영어회화|영어인강|영어학습지|영어교재|영어교구|영어프로그램|영어앱|영어학습|영어공부))'
}

unsuitable_filters = {
    '중등/고등/대학': r'.*(중학|고등|대학|성인|직장인|노인|50대|40대|30대|20대|청소년|중등|고1|고2|고3|중1|중2|중3).*(?!.*(초등|유아|어린이|아동|키즈)).*(?!.*(국제학교|온라인국제학교))',
    '제2외국어/수학/한국사 등 교육 분야 외': r'.*(일본어|중국어|프랑스어|스페인어|독일어|베트남어|태국어|러시아어|아랍어).*(?!.*영어|.*글로벌)|.*(수학|과학|사회|국어|한국어|한국사|물리|화학|생물|지구과학|역사|문학|한문|컴퓨터|코딩|프로그래밍|경제|미술|체육|음악|무용|태권도|발레).*(?!.*영어|.*국제학교|.*온라인국제학교|.*글로벌)',
    '시험/자격증 관련': r'.*(토익|토플|아이엘츠|오픽|텝스|HSK|JLPT|DELE|DELF|TSC|JPT|TOPIK|EJU|AP|수능|내신|모의고사|TOEIC|TOEFL|IELTS|OPIC|TEPS).*(?!.*초등|.*유아|.*어린이|.*아동|.*키즈|.*국제학교|.*온라인국제학교)|.*(SAT|SSAT).*(?!.*초등|.*유아|.*어린이|.*아동|.*키즈|.*국제학교)',
    '캠프/기숙학원 등 오프라인 중심': r'.*(방문학습|방문교사|대면|현장체험학습|체험학습|체험활동|캠프|기숙).*(?!.*온라인|.*화상|.*인터넷|.*국제학교|.*온라인국제학교|.*초등영어)',
    '대안학교/경시대회 등': r'.*(검정고시|재수|편입|입시|윈터스쿨|서머스쿨|논술|특목고|영재|올림피아드|경시대회|대회|마이스터|특성화).*(?!.*초등|.*유아|.*어린이|.*아동|.*키즈|.*국제학교|.*온라인국제학교|.*영어)',
    '경쟁 브랜드명': r'.*(눈높이|구몬|웅진|대교|YBM|YBM토익|튼튼영어|윤선생|EBSe|와이즈만|라이즈|하바|크라운|뽀로로|핑크퐁|몬테소리|발도르프|키즈랜드|숲유치원|이투스|메가|대성|스카이에듀|강남구청|시원스쿨).*(?!.*국제학교|.*온라인국제학교|.*초등영어|.*유아영어)',
    '비프리미엄 지역 및 업무/대학 지역': r'.*(노원구|도봉구|강동구|은평구|중랑구|광화문|여의도|종로|홍대|신촌|용산|광진구|구로구|금천구|서대문구|성동구|성북구|영등포구|동작구|관악구|양천구|강서구|마포구).*(?!.*(초등|유아|어린이|아동|키즈|국제학교|인터내셔널))',
    '육아/여행 등 기타상품': r'(?=.*(육아|여행|장난감|놀이공원|인형|블럭|퍼즐|레고|책장|가구|영양제|건강|운동|다이어트))(?!.*(?:영어|영어교육|영어학습|영어공부|영어학원|영어교재|영어교구|영어프로그램|영어앱|영어학습지|영어동화|영어동요|영어책|영어독서|영어발음|영어문법|영어단어|국제학교|온라인국제학교))',
    '직장인/성인/비즈니스 타겟 키워드': r'.*(비즈니스영어|비지니스영어|강남역|역삼역|직장인영어|성인영어|영어과외알바|영어회화알바|영어학원창업|영어공부방창업|영어학원매매|영어PT|왕초보영어|기초영어|주말영어|토요일영어|종로영어|한달영어|평생영어|6개월영어|영어회화주말반|영어회화단기|비즈니스영어학원|비즈니스영어과외|비즈니스영어회화|비즈니스영어인강|직장인화상영어|영어가맹|영어학원가맹|영어학원체인점|영어프랜차이즈|이력서영어|면접영어|인터뷰영어|취업영어|스피킹|토킹|프리토킹|회사|직장|취업|면접|이력서|토요일|평일|평생|한달|6개월|알바|창업|매매|가맹|PT).*(?!(?:.*(?:초등|유아|어린이|아동|키즈|국제학교|인터내셔널학교)|^(?:초등|유아|어린이|아동|키즈|국제학교|인터내셔널학교).*))',
    '사전/번역 관련': r'.*(사전|번역|번역기|번역사|통역|통역사)'
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
    '프리미엄 학군 유아/초등 영어'
])
final_df.loc[mask_purple, '키워드_분류_질적'] = '전략적 Sweet Spot'
mask_blue = final_df['키워드_상세분류'].isin([
    '미국 교육 커리큘럼'
])
final_df.loc[mask_blue, '키워드_분류_질적'] = '특화 영역'
mask_red = final_df['키워드_상세분류'].isin([
    '유아/초등 타겟 영어교육',
    'Pre-K, K 유아/초등 영어 콘텐츠',    
    '경쟁 브랜드명'
])
final_df.loc[mask_red, '키워드_분류_질적'] = '타겟 경쟁 영역'
mask_expandable = final_df['키워드_상세분류'].isin([
    '타겟없는 일반 영어 교육',
    '타겟 없는 영어 콘텐츠 (교재 등)',
    '타겟 없는 온라인 영어 교육'
])
final_df.loc[mask_expandable, '키워드_분류_질적'] = '확장 가능 키워드'
mask_junk = final_df['키워드_상세분류'].isin([
    '육아/여행 등 기타상품',
    '비프리미엄 지역 및 업무/대학 지역'
])
final_df.loc[mask_junk, '키워드_분류_질적'] = '정크 키워드'
mask_off_target = final_df['키워드_상세분류'].isin([
    '제2외국어/수학/한국사 등 교육 분야 외',
    '시험/자격증 관련', 
    '캠프/기숙학원 등 오프라인 중심',
    '직장인/성인/비즈니스 타겟 키워드',
    '중등/고등/대학',
    '대안학교/경시대회 등',
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

# importance_order와 labels_kr 키 일치 보장
importance_order = [
    '전략적 Sweet Spot',
    '특화 영역',
    '타겟 경쟁 영역',
    '확장 가능 키워드',
    '정크 키워드',
    '타겟 외 경쟁 영역',
    '미분류'
]
labels_kr = {
    '전략적 Sweet Spot': '전략적 Sweet Spot\n(Purple Ocean)',
    '특화 영역': '특화 키워드\n(Blue Ocean)',
    '타겟 경쟁 영역': '경쟁 키워드\n(Red Ocean)',
    '확장 가능 키워드': '확장 가능 키워드',
    '정크 키워드': '정크 키워드',
    '타겟 외 경쟁 영역': '타겟 외 경쟁 영역',
    '미분류': '미분류'
}

# 색상 팔레트 업데이트 - 더 밝고 선명한 색상으로 변경
color_map = {
    '전략적 Sweet Spot': '#7B68EE',  # 밝은 보라색
    '특화 영역': '#00CED1',          # 밝은 청록색
    '타겟 경쟁 영역': '#FF6B6B',     # 밝은 산호색
    '확장 가능 키워드': '#FFD700',    # 골드
    '정크 키워드': '#B0C4DE',        # 밝은 회색
    '타겟 외 경쟁 영역': '#FFA07A',   # 밝은 연어색
    '미분류': '#F0F8FF'              # 매우 밝은 하늘색
}

# 예시용 분류명 매핑 (실제 분류명/통계로 대체)
area_defs = [
    # x0, y0, x1, y1, 분류명, 색상
    [0.5, 0.75, 1, 1, '타겟 경쟁 영역', '#FFF0F0'],
    [0.5, 0.5, 1, 0.75, '전략적 Sweet Spot', '#F0F0FF'],
    [0.5, 0, 1, 0.5, '특화 영역', '#F0FFFF'],
    [0.25, 0.5, 0.5, 1, '확장 가능 키워드', '#FFFFF0'],
    [0, 0.5, 0.25, 1, '타겟 외 경쟁 영역', '#FFF8F0'],
    [0, 0, 0.5, 0.5, '정크 키워드', '#F8F8F8'],
]

# JS로 화면 폭 감지 및 session_state에 저장
if 'is_narrow' not in st.session_state:
    st.session_state['is_narrow'] = True  # 기본값을 좁은 모드로 설정

# JavaScript 디버깅을 위한 컴포넌트
components.html(
    """
    <script>
    function updateStreamlitState(isNarrow) {
        const message = {
            type: 'streamlit:setComponentValue',
            value: isNarrow
        };
        window.parent.postMessage(message, '*');
    }

    function checkWidth() {
        const width = window.innerWidth;
        const isNarrow = width < 900;
        updateStreamlitState(isNarrow);
    }

    // 초기 실행
    checkWidth();

    // 리사이즈 이벤트 리스너
    let resizeTimeout;
    window.addEventListener('resize', function() {
        clearTimeout(resizeTimeout);
        resizeTimeout = setTimeout(checkWidth, 100);
    });
    </script>
    """,
    height=0,
)

# 상세 정보 보기 버튼
if st.button("상세 정보 보기" if st.session_state['is_narrow'] else "간단 정보 보기"):
    st.session_state['is_narrow'] = not st.session_state['is_narrow']
    st.rerun()

# 레이아웃 업데이트
data = []
for x0, y0, x1, y1, area_name, color in area_defs:
    # 해당 영역의 통계 데이터 가져오기
    area_stats = classification_stats[classification_stats['키워드_분류_질적'] == area_name]
    stat = area_stats.iloc[0].to_dict() if not area_stats.empty else {
        '키워드_개수': 0,
        '평균_검색수': 0,
        '평균_클릭수': 0,
        '평균_클릭률_PC': 0,
        '평균_노출광고수': 0
    }
    
    # 해당 영역의 세부 카테고리와 키워드 개수 계산
    area_categories = final_df[final_df['키워드_분류_질적'] == area_name].groupby('키워드_상세분류').size()
    category_text = "<br>".join([f"{cat}: {count:,}개" for cat, count in area_categories.items()])

    # 분기: 넓은 화면(폭 충분) vs 좁은 화면(폭 부족)
    if st.session_state.get('is_narrow', False):
        # 좁은 화면: 간단한 정보만 표시
        area_text = f"<b><span style='font-size: 16px; color: #2C3E50;'>{area_name}</span></b><br>"
        area_text += f"키워드: {stat['키워드_개수']:,}개<br>"
        area_text += f"검색수: {stat['평균_검색수']:,}회"
        hover_text = f"<b>{area_name}</b><br>키워드: {stat['키워드_개수']:,}개<br>검색수: {stat['평균_검색수']:,}회<br>클릭수: {stat['평균_클릭수']:,}회<br>클릭률: {stat['평균_클릭률_PC']:.1f}%<br>광고수: {stat['평균_노출광고수']:,}개<br><b>세부 카테고리:</b><br>{category_text}"
    else:
        # 넓은 화면: 모든 정보 표시
        area_text = f"<b><span style='font-size: 20px; color: #2C3E50;'>{area_name}</span></b><br>"
        area_text += f"키워드: {stat['키워드_개수']:,}개<br>"
        area_text += f"검색수: {stat['평균_검색수']:,}회<br>"
        area_text += f"클릭수: {stat['평균_클릭수']:,}회<br>"
        area_text += f"클릭률: {stat['평균_클릭률_PC']:.1f}%<br>"
        area_text += f"광고수: {stat['평균_노출광고수']:,}개<br><br>"
        area_text += f"<b>세부 카테고리:</b><br>{category_text}"
        hover_text = ""

    data.append(dict(
        x=(x0+x1)/2, y=(y0+y1)/2, x0=x0, y0=y0, x1=x1, y1=y1,
        area_name=area_name, color=color, area_text=area_text, hover_text=hover_text
    ))

# 레이아웃 업데이트
fig = go.Figure()

# 사각형 영역 그리기
for d in data:
    fig.add_shape(
        type="rect",
        x0=d['x0'], y0=d['y0'], x1=d['x1'], y1=d['y1'],
        line=dict(color="#E0E0E0", width=1),
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
    hoverinfo="text",
    hovertext=[d['hover_text'] for d in data],
    marker=dict(opacity=0),
    showlegend=False,
    textfont=dict(size=14, color="#2C3E50")
))

# 레이아웃 업데이트
fig.update_layout(
    width=1200 if not st.session_state.get('is_narrow', False) else 600,
    height=1200 if not st.session_state.get('is_narrow', False) else 600,
    margin=dict(l=40, r=40, t=40, b=40),
    plot_bgcolor="white",
    paper_bgcolor="white",
    font=dict(
        family="Pretendard, -apple-system, BlinkMacSystemFont, system-ui, Roboto, sans-serif",
        size=14,
        color="#2C3E50"
    ),
    title=dict(
        text="2x2 매트릭스 보기",
        font=dict(
            size=24 if not st.session_state.get('is_narrow', False) else 18,
            color="#1E1E1E"
        )
    ),
    hovermode="closest",
    hoverlabel=dict(
        bgcolor="white",
        font_size=14,
        font_family="Pretendard, -apple-system, BlinkMacSystemFont, system-ui, Roboto, sans-serif"
    )
)

# 축 업데이트
fig.update_xaxes(
    showticklabels=False,
    showgrid=False,
    zeroline=False,
    range=[0, 1],
    title_text="타겟 관련성(아이덴티티) →",
    title_font=dict(size=20, color="#2C3E50"),
    linecolor="#E0E0E0"  # 축 선 색상을 연한 회색으로 변경
)

fig.update_yaxes(
    showticklabels=False,
    showgrid=False,
    zeroline=False,
    range=[0, 1],
    title_text="↑ 확장성(트래픽)",
    title_font=dict(size=20, color="#2C3E50"),
    linecolor="#E0E0E0"  # 축 선 색상을 연한 회색으로 변경
)

# 반응형 레이아웃을 위한 컨테이너 설정
st.plotly_chart(fig, use_container_width=True, config={'responsive': True})

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
        st.markdown(f"### {labels_kr.get(category, '')}")
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
                        '월평균클릭률(PC)': '{:.2f}%','월평균노출 광고수': '{:,.0f}'
                    }),
                    use_container_width=True
                )
                # 적합 전체 다운로드 버튼
                excel_data = io.BytesIO()
                with pd.ExcelWriter(excel_data, engine='xlsxwriter') as writer:
                    suitable_df.to_excel(writer, index=False, sheet_name='적합 키워드')
                excel_data.seek(0)
                st.download_button(
                    label="적합 키워드 전체 다운로드 (Excel)",
                    data=excel_data,
                    file_name=f"{labels_kr.get(category, category)}_적합_전체.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
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
                        '월평균클릭률(PC)': '{:.2f}%','월평균노출 광고수': '{:,.0f}'
                    }),
                    use_container_width=True
                )
                # 부적합 전체 다운로드 버튼
                excel_data = io.BytesIO()
                with pd.ExcelWriter(excel_data, engine='xlsxwriter') as writer:
                    unsuitable_df.to_excel(writer, index=False, sheet_name='부적합 키워드')
                excel_data.seek(0)
                st.download_button(
                    label="부적합 키워드 전체 다운로드 (Excel)",
                    data=excel_data,
                    file_name=f"{labels_kr.get(category, category)}_부적합_전체.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
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
                    '월평균클릭률(PC)': '{:.2f}%','월평균노출 광고수': '{:,.0f}'
                }),
                use_container_width=True
            )
            # 카테고리 전체 다운로드 버튼
            excel_data = io.BytesIO()
            with pd.ExcelWriter(excel_data, engine='xlsxwriter') as writer:
                category_df.to_excel(writer, index=False, sheet_name=labels_kr.get(category, category))
            excel_data.seek(0)
            st.download_button(
                label=f"{labels_kr.get(category, category)} 전체 다운로드 (Excel)",
                data=excel_data,
                file_name=f"{labels_kr.get(category, category)}_전체.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        st.markdown("---")

# 8. 전체 데이터 다운로드
def get_excel_download_link(df, filename):
    # 컬럼 순서 재배치 및 정렬
    df = df[['키워드_분류', '키워드_상세분류', '연관키워드'] + [col for col in df.columns if col not in ['키워드_분류', '키워드_상세분류', '연관키워드']]]

    # 키워드 분류 순서 정의
    classification_order = ['적합', '확장 가능 키워드', '부적합', '미분류']
    df['키워드_분류'] = pd.Categorical(df['키워드_분류'], categories=classification_order, ordered=True)

    # 정렬
    df = df.sort_values(['키워드_분류', '키워드_상세분류', '연관키워드'], ascending=[True, True, True])

    # 키워드 분류별 상세 통계 생성
    stats_columns = ['월간검색수(PC)', '월간검색수(모바일)', '월평균클릭수(PC)', '월평균클릭수(모바일)', 
                    '월평균클릭률(PC)', '월평균클릭률(모바일)', '경쟁정도', '월평균노출 광고수', 
                    '총 검색수', '총 클릭수']

    # 빈 DataFrame 생성
    classification_stats = pd.DataFrame()

    # 각 분류 조합에 대해 통계 계산
    for 분류 in classification_order:
        for 상세분류 in df[df['키워드_분류'] == 분류]['키워드_상세분류'].unique():
            mask = (df['키워드_분류'] == 분류) & (df['키워드_상세분류'] == 상세분류)
            subset = df[mask]
            
            if not subset.empty:
                stats = {
                    '키워드_건수': len(subset),
                    '총 검색수': subset['총 검색수'].sum(),
                    '총 클릭수': subset['총 클릭수'].sum(),           
                    '경쟁정도': subset['경쟁정도'].mode().iloc[0] if '경쟁정도' in subset.columns else '-',
                    '월평균노출 광고수': subset['월평균노출 광고수'].mean(),         
                    '월간검색수(PC)': subset['월간검색수(PC)'].sum(),
                    '월간검색수(모바일)': subset['월간검색수(모바일)'].sum(),
                    '월평균클릭수(PC)': subset['월평균클릭수(PC)'].sum(),
                    '월평균클릭수(모바일)': subset['월평균클릭수(모바일)'].sum(),
                    '월평균클릭률(PC)': subset['월평균클릭률(PC)'].mean(),
                    '월평균클릭률(모바일)': subset['월평균클릭률(모바일)'].mean()
                }
                
                # MultiIndex 생성
                idx = pd.MultiIndex.from_tuples([(분류, 상세분류)], names=['키워드_분류', '키워드_상세분류'])
                temp_df = pd.DataFrame([stats], index=idx)
                classification_stats = pd.concat([classification_stats, temp_df])

    # 소수점 둘째자리까지 반올림
    classification_stats = classification_stats.round(2)

    # 엑셀 파일로 저장
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # 통계 시트 저장
        classification_stats.to_excel(writer, sheet_name='통계')
        # 원본 데이터 시트 저장
        df.to_excel(writer, sheet_name='원본데이터', index=False)
        
        # 워크시트 가져오기
        stats_worksheet = writer.sheets['통계']
        data_worksheet = writer.sheets['원본데이터']
        
        # 통계 시트 열 너비 조정
        for idx, col in enumerate(classification_stats.columns):
            max_length = max(
                classification_stats[col].astype(str).apply(len).max(),
                len(col)
            ) + 2
            stats_worksheet.set_column(idx, idx, max_length)
        
        # 원본 데이터 시트 열 너비 조정
        for idx, col in enumerate(df.columns):
            max_length = max(
                df[col].astype(str).apply(len).max(),
                len(col)
            ) + 2
            data_worksheet.set_column(idx, idx, max_length)
    
    output.seek(0)
    return output

# 다운로드 버튼 생성
excel_data = get_excel_download_link(final_df, "키워드_질적분류_결과.xlsx")
st.download_button(
    label="전체 분류 데이터 다운로드 (Excel)",
    data=excel_data,
    file_name="키워드_질적분류_결과.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
) 