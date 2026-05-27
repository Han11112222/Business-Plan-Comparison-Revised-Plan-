import streamlit as st
import pandas as pd
import plotly.express as px
import re
import os

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
        if hasattr(uploaded_file, 'seek'):
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
    
    df['연'] = df['연'].astype(int)
    
    order_dict = {"실적": 1, "기존계획량": 2, "실천계획량": 3}
    df['sort_key'] = df['연'] * 10 + df['구분'].map(order_dict)
    df = df.sort_values(['sort_key', '월'])
    
    df['연_구분'] = df['연'].astype(str) + " (" + df['구분'] + ")"
    unique_x_labels = df['연_구분'].unique().tolist()
    
    current_groups = df['그룹'].unique()
    valid_order = [g for g in ORDER_LIST if g in current_groups]
    rest_groups = [g for g in current_groups if g not in valid_order]
    final_group_order = valid_order + sorted(rest_groups)

    years = sorted(df['연'].unique().tolist())

    st.markdown("---")
    
    # ==========================================
    # 1️⃣ 상단: 전체량 분석
    # ==========================================
    st.markdown("### 1️⃣ 전체량 분석")
    
    default_years_1 = years[-5:] if len(years) >= 5 else years
    
    col_y1, col_t1 = st.columns(2)
    with col_y1:
        selected_years_1 = st.multiselect("📅 [전체량] 조회할 연도 선택", options=years, default=default_years_1, key="sec1")
    with col_t1:
        selected_types_1 = st.multiselect("📊 [전체량] 조회할 구분 선택", options=["실적", "기존계획량", "실천계획량"], default=["실적", "기존계획량", "실천계획량"], key="type_sec1")
    
    if selected_years_1 and selected_types_1:
        df_filt1 = df[df['연'].isin(selected_years_1) & df['구분'].isin(selected_types_1)]
        filtered_x_labels_1 = [x for x in unique_x_labels if int(x[:4]) in selected_years_1 and any(t in x for t in selected_types_1)]

        color_map = {}
        palette = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2']
        for i, lbl in enumerate(unique_x_labels):
            color_map[lbl] = palette[i % len(palette)]

        st.markdown("#### 📈 전체 월별 공급량 추이")
        mon_grp = df_filt1.groupby(['연_구분', '월'])['값'].sum().reset_index()
        fig1 = px.line(mon_grp, x='월', y='값', color='연_구분', markers=True, 
                       category_orders={"연_구분": filtered_x_labels_1},
                       color_discrete_map=color_map)
        fig1.update_xaxes(tickvals=list(range(1, 13)), ticktext=[f"{i}월" for i in range(1, 13)])
        # 🟢 [요청 반영] 우측 상단 단위 추가
        fig1.add_annotation(x=1, y=1.05, xref="paper", yref="paper", text=f"단위: {unit}", showarrow=False, font=dict(size=12, color="gray"))
        st.plotly_chart(fig1, use_container_width=True)
            
        st.markdown("#### 🧱 연도/구분별 용도 구성비")
        yr_grp1 = df_filt1.groupby(['연_구분', '그룹'])['값'].sum().reset_index()
        
        yr_grp1_bar = yr_grp1[yr_grp1['연_구분'] != '2026 (실적)']
        # 🟢 [요청 반영] 2026(실적) 라벨 공간 자체를 X축에서 완전 소거하기 위한 필터링 정의
        filtered_x_labels_1_bar = [x for x in filtered_x_labels_1 if x != '2026 (실적)']
        
        fig2 = px.bar(yr_grp1_bar, x='연_구분', y='값', color='그룹', text_auto='.2s',
                      category_orders={"연_구분": filtered_x_labels_1_bar, "그룹": final_group_order})
        # 🟢 [요청 반영] 우측 상단 단위 추가
        fig2.add_annotation(x=1, y=1.05, xref="paper", yref="paper", text=f"단위: {unit}", showarrow=False, font=dict(size=12, color="gray"))
        st.plotly_chart(fig2, use_container_width=True)
            
        st.markdown("##### 📋 전체량 상세 수치")
        piv1 = df_filt1.pivot_table(index='연_구분', columns='그룹', values='값', aggfunc='sum').fillna(0)
        piv1 = piv1.reindex(index=filtered_x_labels_1, columns=[c for c in final_group_order if c in piv1.columns])
        piv1['총계'] = piv1.sum(axis=1)
        st.dataframe(piv1.style.format("{:,.0f}"), use_container_width=True)
        
        # 🟢 [요청 반영] 전체량 표 하단 [세부 분석] 버튼 및 표 신설
        if st.checkbox("🔍 세부 분석 (월별 용도별 전체량)", key="show_detail_1"):
            st.markdown("##### 📅 월별 용도별 세부량")
            piv1_detail = df_filt1.pivot_table(index=['연_구분', '월'], columns='그룹', values='값', aggfunc='sum').fillna(0)
            piv1_detail = piv1_detail.reindex(columns=[c for c in final_group_order if c in piv1_detail.columns])
            piv1_detail['총계'] = piv1_detail.sum(axis=1)
            st.dataframe(piv1_detail.style.format("{:,.0f}"), use_container_width=True)

    st.markdown("---")

    # ==========================================
    # 2️⃣ 하단: 용도별 분석
    # ==========================================
    st.markdown("### 2️⃣ 용도별 구성 분석")
    
    default_years_2 = [y for y in years if y in [2025, 2026]]
    if not default_years_2: default_years_2 = years[-2:] if len(years) >= 2 else years
        
    col_y2, col_t2 = st.columns(2)
    with col_y2:
        selected_years_2 = st.multiselect("📅 [용도별] 조회할 연도 선택", options=years, default=default_years_2, key="sec2")
    with col_t2:
        selected_types_2 = st.multiselect("📊 [용도별] 조회할 구분 선택", options=["실적", "기존계획량", "실천계획량"], default=["실적", "기존계획량", "실천계획량"], key="type_sec2")
    
    if selected_years_2 and selected_types_2:
        df_filt2 = df[df['연'].isin(selected_years_2) & df['구분'].isin(selected_types_2)]
        
        df_filt2 = df_filt2[df_filt2['연_구분'] != '2026 (실적)']
        filtered_x_labels_2 = [x for x in unique_x_labels if int(x[:4]) in selected_years_2 and any(t in x for t in selected_types_2) and x != '2026 (실적)']

        st.markdown("#### 📈 연도/구분별 용도 꺾은선 추이 (월별 비교)")
        
        sec2_mon_grp = df_filt2.groupby(['연_구분', '그룹', '월'])['값'].sum().reset_index()
        
        fig3 = px.line(sec2_mon_grp, x='월', y='값', color='그룹', line_dash='연_구분', markers=True,
                       category_orders={"연_구분": filtered_x_labels_2, "그룹": final_group_order})
        
        fig3.update_xaxes(tickvals=list(range(1, 13)), ticktext=[f"{i}월" for i in range(1, 13)])
        fig3.for_each_trace(lambda t: t.update(visible=True if '가정용' in t.name else 'legendonly'))
        # 🟢 [요청 반영] 우측 상단 단위 추가
        fig3.add_annotation(x=1, y=1.05, xref="paper", yref="paper", text=f"단위: {unit}", showarrow=False, font=dict(size=12, color="gray"))
        st.plotly_chart(fig3, use_container_width=True)
        
        st.markdown("##### 📋 용도별 상세 수치 (비교 테이블)")
        piv2 = df_filt2.pivot_table(index='연_구분', columns='그룹', values='값', aggfunc='sum').fillna(0)
        piv2 = piv2.reindex(index=filtered_x_labels_2, columns=[c for c in final_group_order if c in piv2.columns])
        piv2['총계'] = piv2.sum(axis=1)
        st.dataframe(piv2.style.format("{:,.0f}"), use_container_width=True)
        
        # 🟢 [요청 반영] 용도별 표 하단 [세부 분석] 버튼 및 표 신설
        if st.checkbox("🔍 세부 분석 (월별 용도별 상세량)", key="show_detail_2"):
            st.markdown("##### 📅 월별 용도별 세부량")
            piv2_detail = df_filt2.pivot_table(index=['연_구분', '월'], columns='그룹', values='값', aggfunc='sum').fillna(0)
            piv2_detail = piv2_detail.reindex(columns=[c for c in final_group_order if c in piv2_detail.columns])
            piv2_detail['총계'] = piv2_detail.sum(axis=1)
            st.dataframe(piv2_detail.style.format("{:,.0f}"), use_container_width=True)

# ─────────────────────────────────────────────────────────
# 🟢 5. 메인 실행
# ─────────────────────────────────────────────────────────
def main():
    st.title("🔥 도시가스 공급량 심플 분석")
    
    with st.sidebar:
        st.header("⚙️ 기본 설정")
        unit = st.radio("단위 선택", ["열량 (GJ)", "부피 (천m³)"], index=0)
        
        heating_value = 42.563
        if "부피" in unit:
            heating_value = st.number_input("(기준열량 MJ/Nm3 : 42.563 )", value=42.563, format="%.3f")
        
        st.markdown("---")
        st.subheader("📂 데이터 업로드")
        up_supply = st.file_uploader("공급량 데이터 업로드 (새 파일이 있으면 우선 반영됩니다)", type=["xlsx", "csv"])
    
    default_file = "공급량실적_계획_실적_MJ.xlsx"
    target_file = None
    
    if up_supply is not None:
        target_file = up_supply
    elif os.path.exists(default_file):
        target_file = default_file

    if target_file:
        data_dict = load_all_sheets(target_file)
        
        df_act = data_dict.get("공급량_실적")
        df_plan = data_dict.get("공급량_사업계획")
        df_action = data_dict.get("공급량_실천사업계획")
        
        if df_act is None: df_act = next((df for name, df in data_dict.items() if "실적" in name and "계획" not in name), None)
        if df_plan is None: df_plan = next((df for name, df in data_dict.items() if "사업계획" in name and "실천" not in name), None)
        if df_action is None: df_action = next((df for name, df in data_dict.items() if "실천" in name), None)
        
        if df_act is None and len(data_dict) >= 1: df_act = list(data_dict.values())[0]
        if df_plan is None and len(data_dict) >= 2: df_plan = list(data_dict.values())[1]
        if df_action is None and len(data_dict) >= 3: df_action = list(data_dict.values())[2]
        
        long_act = make_long_data(df_act, "실적") if df_act is not None else pd.DataFrame()
        long_plan = make_long_data(df_plan, "기존계획량") if df_plan is not None else pd.DataFrame()
        long_action = make_long_data(df_action, "실천계획량") if df_action is not None else pd.DataFrame()
        
        if not long_act.empty:
            df_hist = long_act[long_act['연'] < 2026]
            df_act_2026 = long_act[(long_act['연'] == 2026) & (long_act['월'] <= 3)]
            
            df_plan_2026_apr_dec = long_plan[(long_plan['연'] == 2026) & (long_plan['월'] >= 4)]
            df_plan_2026_jan_mar = df_act_2026.copy()
            df_plan_2026_jan_mar['구분'] = "기존계획량"
            df_plan_2026_apr_dec = df_plan_2026_apr_dec.copy()
            df_plan_2026_apr_dec['구분'] = "기존계획량"
            df_plan_2026 = pd.concat([df_plan_2026_jan_mar, df_plan_2026_apr_dec], ignore_index=True)
            
            df_action_2026_apr_dec = long_action[(long_action['연'] == 2026) & (long_action['월'] >= 4)]
            df_action_2026_jan_mar = df_act_2026.copy()
            df_action_2026_jan_mar['구분'] = "실천계획량"
            df_action_2026_apr_dec = df_action_2026_apr_dec.copy()
            df_action_2026_apr_dec['구분'] = "실천계획량"
            df_action_2026 = pd.concat([df_action_2026_jan_mar, df_action_2026_apr_dec], ignore_index=True)
            
            df_final = pd.concat([df_hist, df_act_2026, df_plan_2026, df_action_2026], ignore_index=True)
            
            df_final['연'] = pd.to_numeric(df_final['연'], errors='coerce')
            df_final = df_final[df_final['연'] <= 2026]
            
            if "GJ" in unit:
                df_final['값'] = df_final['값'] / 1000
            elif "부피" in unit:
                df_final['값'] = df_final['값'] / heating_value / 1000
                
            render_simple_dashboard(df_final, unit)
        else:
            st.warning("⚠️ 유효한 실적 데이터를 추출하지 못했습니다. 파일 구조를 확인해주세요.")
    else:
        st.info("👈 좌측 사이드바에서 공급량 파일을 업로드하거나, 프로젝트 폴더 내에 공급량실적_계획_실적_MJ.xlsx 파일을 배치해 주세요.")

if __name__ == "__main__":
    main()
