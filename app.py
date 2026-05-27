import streamlit as st
import pandas as pd
import plotly.express as px
import re

# ─────────────────────────────────────────────────────────
# 🟢 1. 기본 설정
# ─────────────────────────────────────────────────────────
st.set_page_config(page_title="공급량 실적 및 계획 분석", layout="wide")

def set_korean_font():
    try:
        import matplotlib as mpl
        mpl.rcParams['axes.unicode_minus'] = False
        mpl.rc('font', family='Malgun Gothic') 
    except: pass

set_korean_font()

# ─────────────────────────────────────────────────────────
# 🟢 2. 용도 매핑
# ─────────────────────────────────────────────────────────
MAPPING_SUPPLY = {
    "취사용": "가정용", "개별난방용": "가정용", "중앙난방용": "가정용", 
    "개별난방": "가정용", "중앙난방": "가정용",
    "영업용": "영업용",
    "일반용(1)": "업무용", "일반용1": "업무용", "일반용1(영업)": "업무용", "일반용1(업무)": "업무용",
    "일반용(2)": "업무용", "일반용2": "업무용", 
    "업무난방용": "업무난방용", "냉난방용": "업무용", "냉방용": "업무용", 
    "주한미군": "업무용", 
    "산업용": "산업용",
    "수송용(CNG)": "수송용", "CNG": "수송용",
    "수송용(BIO)": "수송용", "BIO": "수송용",
    "열병합용": "열병합용", "열병합용1": "열병합용",
    "연료전지": "연료전지", "연료전지용": "연료전지",
    "자가열전용": "자가열전용",
    "열전용설비용": "열전용설비용(주택외)", "열전용설비용(주택외)": "열전용설비용(주택외)"
}

ORDER_LIST = [
    "가정용", "영업용", "업무용", "업무난방용", "산업용", 
    "열병합용", "연료전지", "자가열전용", "열전용설비용(주택외)", "수송용"
]

# ─────────────────────────────────────────────────────────
# 🟢 3. 전처리 함수
# ─────────────────────────────────────────────────────────
@st.cache_data(ttl=600)
def load_all_sheets(uploaded_file):
    if uploaded_file is None: return {}
    data_dict = {}
    try:
        excel = pd.ExcelFile(uploaded_file, engine='openpyxl')
        for sheet in excel.sheet_names:
            data_dict[sheet] = excel.parse(sheet)
    except:
        uploaded_file.seek(0)
        df = pd.read_csv(uploaded_file, encoding='utf-8-sig')
        data_dict["default"] = df
    return data_dict

def clean_df(df):
    if df is None: return pd.DataFrame()
    df = df.copy()
    if len(df.columns) > 0 and isinstance(df.columns[0], str) and "데이터 학습기간" in df.columns[0]:
        new_header = df.iloc[0] 
        df = df[1:] 
        df.columns = new_header
    df.columns = df.columns.astype(str).str.strip()
    cols = [c for c in df.columns if not ("Unnamed" in c or re.search(r'^열\s*\d+', c) or c == '0')]
    df = df[cols]
    if '날짜' in df.columns:
        df['날짜'] = pd.to_datetime(df['날짜'], errors='coerce')
        if '연' not in df.columns: df['연'] = df['날짜'].dt.year
        if '월' not in df.columns: df['월'] = df['날짜'].dt.month
    return df

def make_long_data(df, label):
    df = clean_df(df)
    if df.empty or '연' not in df.columns: return pd.DataFrame()
    if '월' not in df.columns: df['월'] = 1 

    records = []
    df['연'] = pd.to_numeric(df['연'], errors='coerce')
    df['월'] = pd.to_numeric(df['월'], errors='coerce')
    df = df.dropna(subset=['연'])
    
    exclude_cols = ['연', '월', '날짜', '평균기온', '총공급량', '총합계', '비교(V-W)', '소 계', '소계']

    for col in df.columns:
        if col in exclude_cols: continue
        val_series = pd.to_numeric(df[col], errors='coerce').fillna(0)
        if val_series.sum() == 0: continue

        group = MAPPING_SUPPLY.get(col, col)

        sub = df[['연', '월']].copy()
        sub['그룹'] = group
        sub['구분'] = label
        sub['값'] = val_series
        sub = sub[sub['값'] != 0]
        records.append(sub)
        
    if not records: return pd.DataFrame()
    return pd.concat(records, ignore_index=True)

# ─────────────────────────────────────────────────────────
# 🟢 4. 심플 대시보드 렌더링
# ─────────────────────────────────────────────────────────
def render_simple_dashboard(df, unit):
    st.subheader(f"📊 공급량 실적 및 계획 통합 분석 ({unit})")
    
    # 정렬을 위한 헬퍼 컬럼 생성 (과거 실적 -> 기존계획 -> new 계획 순서 보장)
    order_dict = {"과거 실적": 1, "기존계획량": 2, "new 계획량": 3}
    df['sort_key'] = df['연'].astype(int) * 10 + df['구분'].map(order_dict)
    df = df.sort_values(['sort_key', '월'])
    
    # X축 및 범례에 사용될 통합 라벨 (예: 2026 (new 계획량))
    df['연_구분'] = df['연'].astype(str) + " (" + df['구분'] + ")"
    unique_x_labels = df['연_구분'].unique().tolist()
    
    # 용도 그룹 정렬
    current_groups = df['그룹'].unique()
    valid_order = [g for g in ORDER_LIST if g in current_groups]
    rest_groups = [g for g in current_groups if g not in valid_order]
    final_group_order = valid_order + sorted(rest_groups)

    # 연도 필터링
    years = sorted(df['연'].unique().tolist())
    selected_years = st.multiselect("📅 조회할 연도 선택", options=years, default=years)
    if not selected_years: return
    
    df_filt = df[df['연'].isin(selected_years)]
    filtered_x_labels = [x for x in unique_x_labels if int(x[:4]) in selected_years]

    st.markdown("---")
    
    # ==========================================
    # 1️⃣ 상단: 전체량 분석
    # ==========================================
    st.markdown("### 1️⃣ 전체량 분석")
    c1, c2 = st.columns(2)
    
    with c1:
        st.markdown("#### 📈 전체 월별 공급량 추이")
        mon_grp = df_filt.groupby(['연_구분', '월'])['값'].sum().reset_index()
        fig1 = px.line(mon_grp, x='월', y='값', color='연_구분', markers=True, 
                       category_orders={"연_구분": filtered_x_labels})
        fig1.update_xaxes(dtick=1)
        st.plotly_chart(fig1, use_container_width=True)
        
    with c2:
        st.markdown("#### 🧱 연도/구분별 용도 구성비")
        yr_grp = df_filt.groupby(['연_구분', '그룹'])['값'].sum().reset_index()
        fig2 = px.bar(yr_grp, x='연_구분', y='값', color='그룹', text_auto='.2s',
                      category_orders={"연_구분": filtered_x_labels, "그룹": final_group_order})
        st.plotly_chart(fig2, use_container_width=True)
        
    st.markdown("##### 📋 전체량 상세 수치")
    piv1 = df_filt.pivot_table(index='연_구분', columns='그룹', values='값', aggfunc='sum').fillna(0)
    # 정렬 적용
    piv1 = piv1.reindex(index=filtered_x_labels, columns=[c for c in final_group_order if c in piv1.columns])
    piv1['총계'] = piv1.sum(axis=1)
    st.dataframe(piv1.style.format("{:,.0f}"), use_container_width=True)

    st.markdown("---")

    # ==========================================
    # 2️⃣ 하단: 용도별 분석
    # ==========================================
    st.markdown("### 2️⃣ 용도별 구성 분석")
    
    st.markdown("#### 📈 연도/구분별 용도 꺾은선 추이")
    fig3 = px.line(yr_grp, x='연_구분', y='값', color='그룹', markers=True,
                   category_orders={"연_구분": filtered_x_labels, "그룹": final_group_order})
    st.plotly_chart(fig3, use_container_width=True)
    
    st.markdown("##### 📋 용도별 상세 수치 (비교 테이블)")
    # 비교하기 쉽게 용도를 인덱스(세로)로, 연도/구분을 컬럼(가로)으로 배치
    piv2 = df_filt.pivot_table(index='그룹', columns='연_구분', values='값', aggfunc='sum').fillna(0)
    piv2 = piv2.reindex(index=[c for c in final_group_order if c in piv2.index], columns=filtered_x_labels)
    st.dataframe(piv2.style.format("{:,.0f}"), use_container_width=True)

# ─────────────────────────────────────────────────────────
# 🟢 5. 메인 실행
# ─────────────────────────────────────────────────────────
def main():
    st.title("🔥 도시가스 공급량 심플 분석")
    
    with st.sidebar:
        st.header("⚙️ 기본 설정")
        unit = st.radio("단위 선택", ["열량 (GJ)", "부피 (천m³)"], index=0)
        
        st.markdown("---")
        st.subheader("📂 데이터 업로드")
        up_supply = st.file_uploader("공급량 데이터 (실적, 기존계획, new계획 포함 파일)", type=["xlsx", "csv"])
    
    if up_supply:
        data_dict = load_all_sheets(up_supply)
        
        # 시트 이름에 따라 자동 분류
        df_act = next((df for name, df in data_dict.items() if "실적" in name and "계획" not in name), None)
        df_plan = next((df for name, df in data_dict.items() if "사업계획" in name and "실천" not in name), None)
        df_action = next((df for name, df in data_dict.items() if "실천" in name), None)
        
        # 이름 매칭 실패 시 시트 순서대로 할당 (1번째: 실적, 2번째: 기존계획, 3번째: 실천계획)
        if df_act is None and len(data_dict) >= 1: df_act = list(data_dict.values())[0]
        if df_plan is None and len(data_dict) >= 2: df_plan = list(data_dict.values())[1]
        if df_action is None and len(data_dict) >= 3: df_action = list(data_dict.values())[2]
        
        # 데이터 Long form 변환 및 태깅
        long_act = make_long_data(df_act, "과거 실적") if df_act is not None else pd.DataFrame()
        long_plan = make_long_data(df_plan, "기존계획량") if df_plan is not None else pd.DataFrame()
        long_action = make_long_data(df_action, "new 계획량") if df_action is not None else pd.DataFrame()
        
        df_final = pd.concat([long_act, long_plan, long_action], ignore_index=True)
        
        if not df_final.empty:
            # GJ 단위 변환 적용
            if "GJ" in unit:
                df_final['값'] = df_final['값'] / 1000
                
            render_simple_dashboard(df_final, unit)
        else:
            st.warning("⚠️ 유효한 데이터를 추출하지 못했습니다. 파일 양식을 확인해주세요.")
    else:
        st.info("👈 좌측 사이드바에서 공급량 파일을 업로드해주세요.")

if __name__ == "__main__":
    main()
