import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
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
# 🟢 4-1. 심플 대시보드 렌더링 (Basic report)
# ─────────────────────────────────────────────────────────
def render_simple_dashboard(df, unit, long_plan=None, long_action=None, heating_value=42.563):
    st.subheader(f"📊 공급량 실적 및 계획 통합 분석 ({unit})")
    
    # ==========================================
    # 1️⃣ 2026년 계획 vs 예상실적 비교
    # ==========================================
    st.markdown("### 1️⃣ 2026년 계획 vs 예상실적 비교")
    
    if long_plan is not None and long_action is not None and not long_plan.empty and not long_action.empty:
        df_p2026 = long_plan[long_plan['연'] == 2026].copy()
        df_a2026 = long_action[long_action['연'] == 2026].copy()
        
        df_p2026['구분_비교'] = "계획"
        df_a2026['구분_비교'] = "예상실적"
        
        df_comp = pd.concat([df_p2026, df_a2026], ignore_index=True)
        
        if "GJ" in unit:
            df_comp['값'] = df_comp['값'] / 1000
        elif "부피" in unit:
            df_comp['값'] = df_comp['값'] / heating_value / 1000
            
        if not df_comp.empty:
            options = ["전체"] + ORDER_LIST
            selected_option = st.selectbox("📂 조회할 항목 선택", options=options, index=0, key="sb_2026_comp")
            
            if selected_option == "전체":
                df_filtered = df_comp
                title_suffix = "전체량"
            else:
                df_filtered = df_comp[df_comp['그룹'] == selected_option]
                title_suffix = selected_option
                
            col_bar, col_line = st.columns([3, 7])
            
            with col_bar:
                df_tot = df_filtered.groupby('구분_비교')['값'].sum().reset_index()
                
                plan_val = df_tot.loc[df_tot['구분_비교'] == '계획', '값'].sum()
                texts = []
                for _, row in df_tot.iterrows():
                    val_str = f"{row['값']:,.0f}"
                    if row['구분_비교'] == '예상실적' and plan_val > 0:
                        ratio = (row['값'] / plan_val) * 100
                        texts.append(f"{val_str}<br>({ratio:.1f}%)")
                    else:
                        texts.append(val_str)
                df_tot['텍스트'] = texts

                fig_bar = px.bar(df_tot, x='구분_비교', y='값', text='텍스트', color='구분_비교',
                                 color_discrete_map={"계획": "#1f77b4", "예상실적": "#ff7f0e"})
                fig_bar.update_traces(textfont_size=18)
                fig_bar.update_layout(title=f"2026년 {title_suffix} 연간 총합 비교", showlegend=False)
                fig_bar.add_annotation(x=1, y=1.05, xref="paper", yref="paper", text=f"단위: {unit}", showarrow=False, font=dict(size=12, color="gray"))
                st.plotly_chart(fig_bar, use_container_width=True)
                
            with col_line:
                df_mon = df_filtered.groupby(['구분_비교', '월'])['값'].sum().reset_index().sort_values('월')
                fig_line = px.line(df_mon, x='월', y='값', color='구분_비교', markers=True,
                                   color_discrete_map={"계획": "#1f77b4", "예상실적": "#ff7f0e"})
                fig_line.update_xaxes(tickvals=list(range(1, 13)), ticktext=[f"{i}월" for i in range(1, 13)])
                fig_line.update_layout(title=f"2026년 {title_suffix} 월별 추이")
                fig_line.add_annotation(x=1, y=1.05, xref="paper", yref="paper", text=f"단위: {unit}", showarrow=False, font=dict(size=12, color="gray"))
                st.plotly_chart(fig_line, use_container_width=True)
            
            st.markdown("##### 📋 세부 수치")
            
            df_table = df_filtered.pivot_table(index='구분_비교', columns='월', values='값', aggfunc='sum').fillna(0)
            
            month_cols = [m for m in range(1, 13) if m in df_table.columns]
            df_table = df_table[month_cols]
            df_table.rename(columns={m: f"{m}월" for m in month_cols}, inplace=True)
            
            df_table['연간 총합'] = df_table.sum(axis=1)
            
            if '계획' in df_table.index and '예상실적' in df_table.index:
                df_table = df_table.reindex(['계획', '예상실적'])
                df_table.loc['증감'] = df_table.loc['예상실적'] - df_table.loc['계획']
                df_table.loc['대비'] = 0.0
                mask = df_table.loc['계획'] != 0
                df_table.loc['대비', mask] = (df_table.loc['예상실적', mask] / df_table.loc['계획', mask]) * 100
                
            df_table.index.name = "구분"
            
            df_display = df_table.copy().astype(object)
            for col in df_display.columns:
                for idx in df_display.index:
                    val = df_table.loc[idx, col]
                    if idx == '대비':
                        df_display.loc[idx, col] = f"{val:,.1f}%"
                    else:
                        df_display.loc[idx, col] = f"{val:,.0f}"
            
            df_display = df_display.reset_index()
            
            styled_df = df_display.style.set_properties(**{'text-align': 'right'})
            
            try:
                html_str = styled_df.hide(axis='index').to_html()
            except:
                html_str = styled_df.hide_index().to_html()
                
            custom_css = """
            <style>
                .custom-table table { width: 100%; border-collapse: collapse; font-family: sans-serif; font-size: 14px; margin-bottom: 1rem; }
                .custom-table th, .custom-table td { border: 1px solid #e2e6ea; padding: 8px; color: #31333F; }
                .custom-table th { background-color: #ffffff; }
                .custom-table th:first-child, .custom-table td:first-child { background-color: #f8f9fa !important; text-align: center !important; font-weight: bold !important; }
                .custom-table th:last-child, .custom-table td:last-child { background-color: #e2e6ea !important; text-align: right !important; font-weight: bold !important; }
            </style>
            """
            st.markdown(f'<div class="custom-table">{html_str}</div>', unsafe_allow_html=True)
            st.markdown(custom_css, unsafe_allow_html=True)

    else:
        st.info("💡 2026년 계획 및 예상실적 비교를 위해 '공급량_사업계획' 및 '공급량_실천사업계획' 데이터를 분석하고 있습니다.")

    st.markdown("---")
    
    df['연'] = df['연'].astype(int)
    
    order_dict = {"실적": 1, "계획": 2, "예상실적": 3}
    df['sort_key'] = df['연'] * 10 + df['구분'].map(order_dict)
    df = df.sort_values(['sort_key', '월'])
    
    df['연_구분'] = df['연'].astype(str) + " (" + df['구분'] + ")"
    unique_x_labels = df['연_구분'].unique().tolist()
    
    current_groups = df['그룹'].unique()
    valid_order = [g for g in ORDER_LIST if g in current_groups]
    rest_groups = [g for g in current_groups if g not in valid_order]
    final_group_order = valid_order + sorted(rest_groups)

    years = sorted(df['연'].unique().tolist())
    
    # ==========================================
    # 2️⃣ 중간: 전체량 분석
    # ==========================================
    st.markdown("### 2️⃣ 전체량 분석")
    
    default_years_1 = years[-5:] if len(years) >= 5 else years
    
    col_y1, col_t1 = st.columns(2)
    with col_y1:
        selected_years_1 = st.multiselect("📅 [전체량] 조회할 연도 선택", options=years, default=default_years_1, key="sec1")
    with col_t1:
        selected_types_1 = st.multiselect("📊 [전체량] 조회할 구분 선택", options=["실적", "계획", "예상실적"], default=["실적", "계획", "예상실적"], key="type_sec1")
    
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
        fig1.add_annotation(x=1, y=1.05, xref="paper", yref="paper", text=f"단위: {unit}", showarrow=False, font=dict(size=12, color="gray"))
        st.plotly_chart(fig1, use_container_width=True)
            
        st.markdown("#### 🧱 연도/구분별 용도 구성비")
        yr_grp1 = df_filt1.groupby(['연_구분', '그룹'])['값'].sum().reset_index()
        
        yr_grp1_bar = yr_grp1[~yr_grp1['연_구분'].isin(['2026 (실적)', '2026 (계획)'])]
        filtered_x_labels_1_bar = [x for x in filtered_x_labels_1 if x not in ['2026 (실적)', '2026 (계획)']]
        
        fig2 = px.bar(yr_grp1_bar, x='연_구분', y='값', color='그룹', text_auto=',.0f',
                      category_orders={"연_구분": filtered_x_labels_1_bar, "그룹": final_group_order})
        
        yr_grp1_tot = yr_grp1_bar.groupby('연_구분')['값'].sum().reset_index()
        for _, row in yr_grp1_tot.iterrows():
            fig2.add_annotation(
                x=row['연_구분'],
                y=row['값'],
                text=f"{row['값']:,.0f}",
                showarrow=False,
                yanchor="bottom",
                yshift=15,
                font=dict(size=14, color="blue")
            )
            
        fig2.add_annotation(x=1, y=1.05, xref="paper", yref="paper", text=f"단위: {unit}", showarrow=False, font=dict(size=12, color="gray"))
        st.plotly_chart(fig2, use_container_width=True)
            
        st.markdown("##### 📋 전체량 상세 수치")
        piv1 = df_filt1.pivot_table(index='연_구분', columns='그룹', values='값', aggfunc='sum').fillna(0)
        piv1 = piv1.reindex(index=filtered_x_labels_1, columns=[c for c in final_group_order if c in piv1.columns])
        piv1['총계'] = piv1.sum(axis=1)
        st.dataframe(piv1.style.format("{:,.0f}"), use_container_width=True)

    st.markdown("---")

    # ==========================================
    # 3️⃣ 하단: 용도별 분석
    # ==========================================
    st.markdown("### 3️⃣ 용도별 구성 분석")
    
    default_years_2 = [y for y in years if y in [2025, 2026]]
    if not default_years_2: default_years_2 = years[-2:] if len(years) >= 2 else years
        
    col_y2, col_t2 = st.columns(2)
    with col_y2:
        selected_years_2 = st.multiselect("📅 [용도별] 조회할 연도 선택", options=years, default=default_years_2, key="sec2")
    with col_t2:
        selected_types_2 = st.multiselect("📊 [용도별] 조회할 구분 선택", options=["실적", "계획", "예상실적"], default=["실적", "계획", "예상실적"], key="type_sec2")
    
    if selected_years_2 and selected_types_2:
        df_filt2 = df[df['연'].isin(selected_years_2) & df['구분'].isin(selected_types_2)]
        
        filtered_x_labels_2 = [x for x in unique_x_labels if int(x[:4]) in selected_years_2 and any(t in x for t in selected_types_2)]

        selected_group = st.radio("📂 조회할 용도 선택", options=final_group_order, index=0, horizontal=True, key="rb_group_sec3")

        st.markdown("#### 📈 연도/구분별 용도 꺾은선 추이 (월별 비교)")
        
        df_fig3 = df_filt2[df_filt2['그룹'] == selected_group]
        sec2_mon_grp = df_fig3.groupby(['연_구분', '월'])['값'].sum().reset_index()
        
        color_map = {}
        palette = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2']
        for i, lbl in enumerate(unique_x_labels):
            color_map[lbl] = palette[i % len(palette)]
        
        fig3 = px.line(sec2_mon_grp, x='월', y='값', color='연_구분', markers=True,
                       category_orders={"연_구분": filtered_x_labels_2},
                       color_discrete_map=color_map)
        
        fig3.update_xaxes(tickvals=list(range(1, 13)), ticktext=[f"{i}월" for i in range(1, 13)])
        fig3.add_annotation(x=1, y=1.05, xref="paper", yref="paper", text=f"단위: {unit}", showarrow=False, font=dict(size=12, color="gray"))
        st.plotly_chart(fig3, use_container_width=True)
        
        st.markdown("##### 📋 용도별 상세 수치 (비교 테이블)")
        piv2 = df_filt2.pivot_table(index='연_구분', columns='그룹', values='값', aggfunc='sum').fillna(0)
        piv2 = piv2.reindex(index=filtered_x_labels_2, columns=[c for c in final_group_order if c in piv2.columns])
        piv2['총계'] = piv2.sum(axis=1)
        st.dataframe(piv2.style.format("{:,.0f}"), use_container_width=True)


# ─────────────────────────────────────────────────────────
# 🟢 4-2. One Page Review 렌더링 (2번 탭)
# ─────────────────────────────────────────────────────────
def render_one_page_review(long_plan, long_action, unit, heating_value):
    st.subheader(f"📑 2026년 DSE 예상실적(실천사업계획) 요약 보고서 ({unit})")
    
    if long_plan.empty or long_action.empty:
        st.warning("데이터가 부족합니다. 공급량_사업계획 및 공급량_실천사업계획 데이터를 확인해주세요.")
        return
        
    # 1. 2026년 데이터만 필터링 및 단위 변환
    p2026 = long_plan[long_plan['연'] == 2026].copy()
    a2026 = long_action[long_action['연'] == 2026].copy()

    def apply_unit(df):
        if "GJ" in unit: df['값'] = df['값'] / 1000
        elif "부피" in unit: df['값'] = df['값'] / heating_value / 1000
        return df

    p2026 = apply_unit(p2026)
    a2026 = apply_unit(a2026)

    # 2. 그룹별/기간별 집계 함수 (1~3월, 4~12월)
    def agg_data(df, prefix):
        df['기간'] = df['월'].apply(lambda x: '1~3월' if x <= 3 else '4~12월')
        pivot = df.pivot_table(index='그룹', columns='기간', values='값', aggfunc='sum').fillna(0)
        
        for col in ['1~3월', '4~12월']:
            if col not in pivot.columns: pivot[col] = 0
            
        pivot['합계'] = pivot['1~3월'] + pivot['4~12월']
        pivot.columns = [f'{prefix}_{c}' for c in pivot.columns]
        return pivot

    df_p = agg_data(p2026, '당초')
    df_a = agg_data(a2026, '변경')

    # 데이터 병합
    df_summary = pd.concat([df_p, df_a], axis=1).fillna(0)

    # ORDER_LIST에 맞게 행 정렬
    valid_order = [g for g in ORDER_LIST if g in df_summary.index]
    rest = [g for g in df_summary.index if g not in valid_order]
    df_summary = df_summary.reindex(valid_order + rest)

    # 총계 행 추가
    df_summary.loc['총계'] = df_summary.sum()

    # 기간별 증감 및 달성률 계산
    df_summary['증감_1~3월'] = df_summary['변경_1~3월'] - df_summary['당초_1~3월']
    df_summary['증감_4~12월'] = df_summary['변경_4~12월'] - df_summary['당초_4~12월']
    df_summary['증감_합계'] = df_summary['변경_합계'] - df_summary['당초_합계']
    df_summary['달성률(%)'] = (df_summary['변경_합계'] / df_summary['당초_합계'] * 100).fillna(0)

    # ==========================================
    # 💡 1. 핵심 KPI 요약 카드
    # ==========================================
    st.markdown("#### 🎯 2026년 핵심 요약 (Total KPI)")
    plan_tot = df_summary.loc['총계', '당초_합계']
    act_tot = df_summary.loc['총계', '변경_합계']
    diff_tot = df_summary.loc['총계', '증감_합계']
    rate_tot = df_summary.loc['총계', '달성률(%)']

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("① 당초 계획 총계", f"{plan_tot:,.0f}")
    col2.metric("② 변경(예상) 총계", f"{act_tot:,.0f}")
    col3.metric("③ 총 증감량 (②-①)", f"{diff_tot:,.0f}", delta=f"{diff_tot:,.0f}", delta_color="normal")
    col4.metric("④ 총 달성률", f"{rate_tot:,.1f}%")

    st.markdown("---")

    # ==========================================
    # 💡 2. 증감 요인 폭포수 차트 (Waterfall)
    # ==========================================
    st.markdown("#### 🌊 증감 요인 폭포수 차트 (Waterfall)")
    
    # 👉 스위치 버튼 추가: 용도별 vs 기간별
    wf_view = st.radio("🔘 차트 구분 선택", ["용도별 구분", "기간별 구분 (1~3월 실적, 4~12월 계획)"], horizontal=True, key="wf_view_radio")
    
    if wf_view == "용도별 구분":
        df_wf = df_summary.drop('총계')
        wf_labels = ["당초 계획 합계"] + df_wf.index.tolist() + ["변경 예상 합계"]
        wf_measures = ["absolute"] + ["relative"] * len(df_wf) + ["total"]
        wf_values = [plan_tot] + df_wf['증감_합계'].tolist() + [act_tot]
        text_labels = [f"{plan_tot:,.0f}"] + [f"{v:,.0f}" if v != 0 else "" for v in df_wf['증감_합계']] + [f"{act_tot:,.0f}"]
        chart_title = "당초 계획 대비 용도별 증감 브릿지"
    else:
        diff_1_3_tot = df_summary.loc['총계', '증감_1~3월']
        diff_4_12_tot = df_summary.loc['총계', '증감_4~12월']
        wf_labels = ["당초 계획 합계", "1~3월 실적반영 (증감)", "4~12월 계획수정 (증감)", "변경 예상 합계"]
        wf_measures = ["absolute", "relative", "relative", "total"]
        wf_values = [plan_tot, diff_1_3_tot, diff_4_12_tot, act_tot]
        text_labels = [f"{v:,.0f}" for v in wf_values]
        chart_title = "당초 계획 대비 기간별 1~3월 및 4~12월 증감 브릿지"

    fig_wf = go.Figure(go.Waterfall(
        orientation="v",
        measure=wf_measures,
        x=wf_labels,
        y=wf_values,
        text=text_labels,
        textposition="outside",
        decreasing={"marker":{"color":"#ff7f0e"}}, # 감소는 주황색
        increasing={"marker":{"color":"#1f77b4"}}, # 증가는 파란색
        totals={"marker":{"color":"#2ca02c"}}      # 총합은 초록색
    ))
    fig_wf.update_layout(title=chart_title, margin=dict(t=40, b=40))
    st.plotly_chart(fig_wf, use_container_width=True)

    st.markdown("---")

    # ==========================================
    # 💡 3. 스마트 요약 테이블 (에러 수정)
    # ==========================================
    st.markdown("#### 🚥 실천사업계획 요약 테이블")
    
    table_cols = [
        '당초_1~3월', '당초_4~12월', '당초_합계',
        '변경_1~3월', '변경_4~12월', '변경_합계',
        '증감_1~3월', '증감_4~12월', '증감_합계', '달성률(%)'
    ]
    df_display = df_summary[table_cols].copy()
    
    format_dict = {c: "{:,.0f}" for c in table_cols if c != '달성률(%)'}
    format_dict['달성률(%)'] = "{:,.1f}%"
    
    # 💥 에러 방지: matplotlib 의존성(background_gradient)을 제거하고, 파이썬 기본 스타일링 함수를 적용
    def custom_heatmap(val):
        if pd.isna(val) or val == 0:
            return ''
        elif val > 0:
            return 'background-color: #e6f2ff; color: #0055a4;' # 연한 파랑 배경, 파란 글씨
        else:
            return 'background-color: #ffe6e6; color: #cc0000;' # 연한 빨강 배경, 붉은 글씨

    # pandas 버전에 맞춰 호환되도록 applymap/map 사용
    try:
        styled_df = df_display.style.map(custom_heatmap, subset=['증감_1~3월', '증감_4~12월', '증감_합계']).format(format_dict)
    except:
        styled_df = df_display.style.applymap(custom_heatmap, subset=['증감_1~3월', '증감_4~12월', '증감_합계']).format(format_dict)
    
    st.dataframe(styled_df, use_container_width=True)

    st.markdown("---")

    # ==========================================
    # 💡 4. 기간별 물량 구성비 차트 (가로 누적 막대)
    # ==========================================
    st.markdown("#### 📊 변경(예상실적) 기준 1~3월 vs 4~12월 비중")
    
    a2026['기간'] = a2026['월'].apply(lambda x: '1~3월 실적' if x <= 3 else '4~12월 예상')
    bar_grp = a2026.groupby(['그룹', '기간'])['값'].sum().reset_index()
    
    current_g = a2026['그룹'].unique()
    v_ord = [g for g in ORDER_LIST if g in current_g]
    r_ord = [g for g in current_g if g not in v_ord]
    f_ord = v_ord + r_ord
    
    fig_bar_h = px.bar(bar_grp, x='값', y='그룹', color='기간', orientation='h', text_auto=',.0f',
                       category_orders={"그룹": f_ord[::-1]},
                       color_discrete_map={"1~3월 실적": "#1f77b4", "4~12월 예상": "#aec7e8"})
    fig_bar_h.update_layout(title=f"각 용도별 1~3월과 4~12월 물량 구성비 (단위: {unit})", barmode='stack')
    st.plotly_chart(fig_bar_h, use_container_width=True)


# ─────────────────────────────────────────────────────────
# 🟢 5. 메인 실행
# ─────────────────────────────────────────────────────────
def main():
    st.title("🔥 2026년 DSE 예상실적(실천사업계획)비교분석")
    
    with st.sidebar:
        st.header("⚙️ 메뉴 및 기본 설정")
        
        menu = st.radio("📋 보고서 탭 선택", ["1. Basic report", "2. One page review"])
        st.markdown("---")
        
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
        long_plan = make_long_data(df_plan, "계획") if df_plan is not None else pd.DataFrame()
        long_action = make_long_data(df_action, "예상실적") if df_action is not None else pd.DataFrame()
        
        if not long_act.empty:
            df_final = pd.concat([long_act, long_plan, long_action], ignore_index=True)
            
            df_final['연'] = pd.to_numeric(df_final['연'], errors='coerce')
            df_final = df_final[df_final['연'] <= 2026]
            
            if menu == "1. Basic report":
                if "GJ" in unit:
                    df_final['값'] = df_final['값'] / 1000
                elif "부피" in unit:
                    df_final['값'] = df_final['값'] / heating_value / 1000
                render_simple_dashboard(df_final, unit, long_plan=long_plan, long_action=long_action, heating_value=heating_value)
            
            elif menu == "2. One page review":
                render_one_page_review(long_plan, long_action, unit, heating_value)
                
        else:
            st.warning("⚠️ 유효한 실적 데이터를 추출하지 못했습니다. 파일 구조를 확인해주세요.")
    else:
        st.info("👈 좌측 사이드바에서 공급량 파일을 업로드하거나, 프로젝트 폴더 내에 공급량실적_계획_실적_MJ.xlsx 파일을 배치해 주세요.")

if __name__ == "__main__":
    main()
