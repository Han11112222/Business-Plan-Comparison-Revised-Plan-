import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import io
import re
from sklearn.linear_model import LinearRegression
from sklearn.preprocessing import PolynomialFeatures
from sklearn.pipeline import make_pipeline

# ─────────────────────────────────────────────────────────
# 🟢 1. 기본 설정
# ─────────────────────────────────────────────────────────
st.set_page_config(page_title="도시가스 통합 분석", layout="wide")

def set_korean_font():
    try:
        import matplotlib as mpl
        mpl.rcParams['axes.unicode_minus'] = False
        mpl.rc('font', family='Malgun Gothic') 
    except: pass

set_korean_font()

# ─────────────────────────────────────────────────────────
# 🟢 2. 용도별 매핑 및 정렬 순서
# ─────────────────────────────────────────────────────────

MAPPING_SALES = {
    "취사용": "가정용", "개별난방용": "가정용", "중앙난방용": "가정용", "자가열전용": "가정용",
    "개별난방": "가정용", "중앙난방": "가정용", "가정용소계": "가정용",
    "일반용": "영업용", "업무난방용": "업무용", "냉방용": "업무용", "주한미군": "업무용",
    "산업용": "산업용", "수송용(CNG)": "수송용", "수송용(BIO)": "수송용",
    "열병합용": "열병합", "열병합용1": "열병합", "열병합용2": "열병합",
    "연료전지용": "연료전지", "연료전지": "연료전지",
    "열전용설비용": "열전용설비용"
}

MAPPING_SUPPLY_SPECIFIC = {
    "취사용": "가정용", "개별난방용": "가정용", "중앙난방용": "가정용", 
    "개별난방": "가정용", "중앙난방": "가정용",
    "영업용": "영업용",
    "일반용(1)": "업무용", "일반용1": "업무용", "일반용1(영업)": "업무용", "일반용1(업무)": "업무용",
    "일반용(2)": "업무용", "일반용2": "업무용", 
    "업무난방용": "업무난방용", "냉난방용": "업무용", "냉방용": "업무용", 
    "주한미군": "업무용", 
    "산업용": "산업용",
    "수송용(CNG)": "수송용", "CNG": "수송용",
    "수송용(BIO)": "수송용", "BIO": "수송용"
}

MAPPING_DETAIL = {
    "취사용": "취사용", 
    "개별난방용": "개별난방용", "개별난방": "개별난방용",
    "중앙난방용": "중앙난방용", "중앙난방": "중앙난방용",
    "영업용": "영업용",
    "일반용(1)": "일반용(1)", "일반용1": "일반용(1)", "일반용1(영업)": "일반용(1)",
    "일반용(2)": "일반용(2)", "일반용2": "일반용(2)",
    "업무난방용": "업무난방용", "업무난방": "업무난방용",
    "냉난방용": "냉난방용", "냉방용": "냉난방용",
    "주한미군": "주한미군",
    "산업용": "산업용",
    "열병합용": "열병합용", "열병합용1": "열병합용",
    "연료전지": "연료전지", "연료전지용": "연료전지",
    "자가열전용": "자가열전용",
    "열전용설비용": "열전용설비용(주택외)", "열전용설비용(주택외)": "열전용설비용(주택외)",
    "수송용(CNG)": "수송용(CNG)", "CNG": "수송용(CNG)",
    "수송용(BIO)": "수송용(BIO)", "BIO": "수송용(BIO)"
}

ORDER_LIST_DETAIL = [
    "취사용", "개별난방용", "중앙난방용",        
    "영업용",                                 
    "일반용(1)", "일반용(2)", "업무난방용", "냉난방용", "주한미군", 
    "산업용",                                 
    "열병합용",                               
    "연료전지",                               
    "자가열전용",                             
    "열전용설비용(주택외)",                    
    "수송용(CNG)", "수송용(BIO)"              
]

MAPPING_FINAL_GROUP = {
    "취사용": "가정용", "개별난방용": "가정용", "중앙난방용": "가정용",
    "영업용": "영업용",
    "일반용(1)": "업무용", "일반용(2)": "업무용", "업무난방용": "업무용", "냉난방용": "업무용", "주한미군": "업무용",
    "산업용": "산업용",
    "열병합용": "열병합용",
    "연료전지": "연료전지",
    "자가열전용": "자가열전용",
    "열전용설비용(주택외)": "열전용설비용(주택외)",
    "수송용(CNG)": "수송용", "수송용(BIO)": "수송용"
}

ORDER_LIST_FINAL_GROUP = [
    "가정용", "영업용", "업무용", "산업용", "열병합용", 
    "연료전지", "자가열전용", "열전용설비용(주택외)", "수송용"
]

# ─────────────────────────────────────────────────────────
# 🟢 3. 파일 로딩 및 전처리
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
        try:
            df = pd.read_csv(uploaded_file, encoding='utf-8-sig')
            data_dict["default"] = df
        except:
            uploaded_file.seek(0)
            try:
                df = pd.read_csv(uploaded_file, encoding='cp949')
                data_dict["default"] = df
            except: pass
    return data_dict

def clean_df(df):
    if df is None: return pd.DataFrame()
    df = df.copy()
    
    if len(df.columns) > 0 and isinstance(df.columns[0], str) and "데이터 학습기간" in df.columns[0]:
        new_header = df.iloc[0] 
        df = df[1:] 
        df.columns = new_header

    df.columns = df.columns.astype(str).str.strip()
    
    cols = []
    for c in df.columns:
        if "Unnamed" in c: continue
        if re.search(r'^열\s*\d+', c): continue 
        if c == '0': continue
        cols.append(c)
    df = df[cols]
    
    if '날짜' in df.columns:
        df['날짜'] = pd.to_datetime(df['날짜'], errors='coerce')
        if '연' not in df.columns: df['연'] = df['날짜'].dt.year
        if '월' not in df.columns: df['월'] = df['날짜'].dt.month
        
    return df

def make_long_data(df, label, mode='sales'):
    df = clean_df(df)
    if df.empty or '연' not in df.columns: return pd.DataFrame()
    
    if '월' not in df.columns:
         df['월'] = 1 

    records = []
    df['연'] = pd.to_numeric(df['연'], errors='coerce')
    df['월'] = pd.to_numeric(df['월'], errors='coerce')
    df = df.dropna(subset=['연'])
    
    exclude_cols = ['연', '월', '날짜', '평균기온', '총공급량', '총합계', '비교(V-W)', '소 계', '소계']

    for col in df.columns:
        if col in exclude_cols: continue
        
        val_series = pd.to_numeric(df[col], errors='coerce').fillna(0)
        if val_series.sum() == 0: continue

        if mode == 'detail':
            group = MAPPING_DETAIL.get(col)
            if not group: group = col 
        elif mode == 'sales':
            group = MAPPING_SALES.get(col)
            if not group: continue 
        else: # supply
            if df[col].dtype == object: continue
            group = MAPPING_SUPPLY_SPECIFIC.get(col, col)

        sub = df[['연', '월']].copy()
        sub['그룹'] = group
        sub['용도'] = col
        sub['구분'] = label
        sub['값'] = val_series
        
        sub = sub[sub['값'] != 0]
        records.append(sub)
        
    if not records: return pd.DataFrame()
    return pd.concat(records, ignore_index=True)

# ─────────────────────────────────────────────────────────
# 🟢 4. 분석 화면 (기존)
# ─────────────────────────────────────────────────────────
def render_analysis_dashboard(long_df, unit_label):
    st.subheader(f"📊 실적 분석 ({unit_label})")
    
    df_act = long_df[long_df['구분'].str.contains('실적')].copy()
    if df_act.empty: st.error("실적 데이터 없음"); return
    
    all_years = sorted([int(y) for y in df_act['연'].unique()])
    if len(all_years) >= 10: default_years = all_years[-10:]
    else: default_years = all_years
        
    selected_years = st.multiselect("연도 선택", options=all_years, default=default_years)
    if not selected_years: return
    
    df_filtered = df_act[df_act['연'].isin(selected_years)]
    st.markdown("---")
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f"#### 📈 월별 추이")
        mon_grp = df_filtered.groupby(['연', '월'])['값'].sum().reset_index()
        fig1 = px.line(mon_grp, x='월', y='값', color='연', markers=True)
        fig1.update_xaxes(dtick=1, tickformat="d")
        st.plotly_chart(fig1, use_container_width=True)
    with col2:
        st.markdown(f"#### 🧱 용도별 구성비")
        yr_grp = df_filtered.groupby(['연', '그룹'])['값'].sum().reset_index()
        fig2 = px.bar(yr_grp, x='연', y='값', color='그룹', text_auto='.2s')
        fig2.update_xaxes(dtick=1, tickformat="d")
        st.plotly_chart(fig2, use_container_width=True)
    
    st.markdown("##### 📋 상세 수치")
    piv = df_filtered.pivot_table(index='연', columns='그룹', values='값', aggfunc='sum').fillna(0)
    piv['소계'] = piv.sum(axis=1)
    st.dataframe(piv.style.format("{:,.0f}"), use_container_width=True)

# ─────────────────────────────────────────────────────────
# 🟢 5. 예측 화면 및 기타 기능 (생략 없이 원본 유지)
# ─────────────────────────────────────────────────────────
def generate_trend_insight(hist_df, pred_df):
    if hist_df.empty or pred_df.empty: return ""
    hist_yearly = hist_df.groupby('연')['값'].sum().sort_index()
    pred_yearly = pred_df.groupby('연')['값'].sum().sort_index()
    
    diffs = hist_yearly.diff()
    max_up_year = diffs.idxmax() if not diffs.dropna().empty else None
    max_down_year = diffs.idxmin() if not diffs.dropna().empty else None
    
    start_val = pred_yearly.iloc[0]
    end_val = pred_yearly.iloc[-1]
    years = len(pred_yearly)
    if start_val > 0:
        cagr = (end_val / start_val) ** (1 / years) - 1
        trend_str = "지속적인 증가세" if cagr > 0.01 else "감소세" if cagr < -0.01 else "보합세"
    else: trend_str = "변동"

    insight = f"💡 **[AI 분석]** 과거 데이터를 분석한 결과, **{int(max_up_year) if max_up_year else '-'}년의 상승**과 **{int(max_down_year) if max_down_year else '-'}년의 하락/조정**을 종합하여 볼 때, 향후 2035년까지는 **{trend_str}**가 유지될 것으로 전망됩니다."
    return insight

def render_prediction_2035(long_df, unit_label, start_pred_year, train_years_selected, is_supply_mode, custom_sort_list=None):
    st.subheader(f"🔮 2035 장기 예측 ({unit_label})")
    filter_cond = long_df['연'].isin(train_years_selected)
    if is_supply_mode:
        filter_cond = filter_cond | (long_df['구분'] == '확정계획')
        
    df_train = long_df[filter_cond].copy()
    if df_train.empty: st.warning("학습 데이터가 부족합니다."); return
    
    st.markdown("##### 📊 추세 분석 모델 선택")
    pred_method = st.radio("방법", ["선형 회귀", "2차 곡선", "3차 곡선", "로그 추세", "지수 평활", "CAGR"], horizontal=True)
    
    df_grp = long_df.groupby(['연', '그룹', '구분'])['값'].sum().reset_index()
    df_train_grp = df_train.groupby(['연', '그룹'])['값'].sum().reset_index()
    groups = df_grp['그룹'].unique()
    future_years = np.arange(start_pred_year, 2036).reshape(-1, 1)
    results = []
    total_hist_vals = []
    total_pred_vals = []

    for grp in groups:
        sub_train = df_train_grp[df_train_grp['그룹'] == grp]
        sub_full = df_grp[df_grp['그룹'] == grp]
        if len(sub_train) < 2: continue
        
        X = sub_train['연'].values.reshape(-1, 1)
        y = sub_train['값'].values
        pred = []
        try:
            if "선형" in pred_method: model = LinearRegression(); model.fit(X, y); pred = model.predict(future_years)
            elif "2차" in pred_method: model = make_pipeline(PolynomialFeatures(2), LinearRegression()); model.fit(X, y); pred = model.predict(future_years)
            elif "3차" in pred_method: model = make_pipeline(PolynomialFeatures(3), LinearRegression()); model.fit(X, y); pred = model.predict(future_years)
            elif "로그" in pred_method: 
                model = LinearRegression(); model.fit(np.log(X - X.min() + 1), y); pred = model.predict(np.log(future_years - X.min() + 1))
            elif "지수" in pred_method:
                fit = np.polyfit(X.flatten(), np.log(y + 1), 1)
                pred = np.exp(fit[1] + fit[0] * future_years.flatten())
            else: 
                cagr = (y[-1]/y[0])**(1/(len(y)-1)) - 1
                pred = [y[-1] * ((1+cagr)**(i+1)) for i in range(len(future_years))]
        except:
            model = LinearRegression(); model.fit(X, y); pred = model.predict(future_years)
            
        pred = [max(0, p) for p in pred]
        added_years = set()
        
        hist_mask = sub_full['연'].isin(train_years_selected)
        if is_supply_mode and start_pred_year == 2029:
             hist_mask = hist_mask & (sub_full['연'] < 2026)
        elif not is_supply_mode:
             hist_mask = hist_mask & (sub_full['연'] < start_pred_year)
        
        hist_data = sub_full[hist_mask]
        for _, row in hist_data.iterrows():
            if row['연'] not in added_years:
                results.append({'연': row['연'], '그룹': grp, '값': row['값'], '구분': '실적'})
                total_hist_vals.append({'연': row['연'], '값': row['값']})
                added_years.add(row['연'])
            
        if is_supply_mode and start_pred_year == 2029:
            plan_data = sub_full[sub_full['연'].between(2026, 2028)]
            for _, row in plan_data.iterrows():
                if row['연'] not in added_years:
                    results.append({'연': row['연'], '그룹': grp, '값': row['값'], '구분': '확정계획'})
                    added_years.add(row['연'])
                
        for yr, v in zip(future_years.flatten(), pred): 
            if yr not in added_years: 
                results.append({'연': yr, '그룹': grp, '값': v, '구분': '예측(AI)'})
                total_pred_vals.append({'연': yr, '값': v})
                added_years.add(yr)
        
    df_res = pd.DataFrame(results)
    display_order = {} 
    
    if custom_sort_list:
        display_order = {'그룹': custom_sort_list}
        current_groups = df_res['그룹'].unique()
        valid_order = [g for g in custom_sort_list if g in current_groups]
        rest_groups = [g for g in current_groups if g not in valid_order]
        final_sort_order = valid_order + sorted(rest_groups)
        df_res['그룹'] = pd.Categorical(df_res['그룹'], categories=final_sort_order, ordered=True)
        df_res = df_res.sort_values(['연', '그룹'])
    else:
        df_res = df_res.sort_values(['연', '그룹'])
    
    insight_text = generate_trend_insight(pd.DataFrame(total_hist_vals), pd.DataFrame(total_pred_vals))
    if insight_text: st.success(insight_text)
    
    st.markdown("---")
    st.markdown("#### 📈 전체 장기 전망 (추세선)")
    fig = px.line(df_res, x='연', y='값', color='그룹', line_dash='구분', markers=True, category_orders=display_order)
    fig.add_vline(x=start_pred_year-0.5, line_dash="dash", line_color="green")
    fig.add_vrect(x0=start_pred_year-0.5, x1=2035.5, fillcolor="green", opacity=0.05, annotation_text="예측 값", annotation_position="inside top")
    
    if is_supply_mode and start_pred_year == 2029:
        fig.add_vrect(x0=2025.5, x1=2028.5, fillcolor="yellow", opacity=0.1, annotation_text="확정계획", annotation_position="inside top")
    
    fig.update_xaxes(dtick=1, tickformat="d")
    st.plotly_chart(fig, use_container_width=True)
    
    st.markdown("---")
    st.markdown("#### 🧱 연도별 공급량 구성 (누적 스택)")
    fig_stack = px.bar(df_res, x='연', y='값', color='그룹', title="연도별 용도 구성비", text_auto='.2s', category_orders=display_order)
    fig_stack.update_xaxes(dtick=1, tickformat="d")
    st.plotly_chart(fig_stack, use_container_width=True)

def render_final_check(long_df, unit_label):
    st.subheader(f"🏁 최종 확정 데이터 시각화 ({unit_label})")
    col_t1, col_t2 = st.columns([4, 1])
    with col_t2: apply_usage_group = st.checkbox("☑️ 용도별 적용")
    
    df_res = long_df.copy()
    if apply_usage_group:
        df_res['New_Group'] = df_res['그룹'].map(MAPPING_FINAL_GROUP).fillna(df_res['그룹'])
        df_res = df_res.groupby(['연', 'New_Group'])['값'].sum().reset_index()
        df_res.rename(columns={'New_Group': '그룹'}, inplace=True)
        target_order = ORDER_LIST_FINAL_GROUP
    else:
        target_order = ORDER_LIST_DETAIL

    current_groups = df_res['그룹'].unique()
    valid_order = [g for g in target_order if g in current_groups]
    rest_groups = [g for g in current_groups if g not in valid_order]
    final_sort_order = valid_order + sorted(rest_groups)
    
    df_res['그룹'] = pd.Categorical(df_res['그룹'], categories=final_sort_order, ordered=True)
    df_res = df_res.sort_values(['연', '그룹'])
    display_order = {'그룹': final_sort_order}
    
    st.markdown("#### 📈 연도별 추세 (Line Chart)")
    fig = px.line(df_res, x='연', y='값', color='그룹', markers=True, category_orders=display_order)
    fig.update_xaxes(dtick=1, tickformat="d")
    st.plotly_chart(fig, use_container_width=True)
    
    st.markdown("#### 🧱 연도별 공급량 구성 (Stacked Bar)")
    fig_stack = px.bar(df_res, x='연', y='값', color='그룹', text_auto='.2s', category_orders=display_order)
    fig_stack.update_xaxes(dtick=1, tickformat="d")
    st.plotly_chart(fig_stack, use_container_width=True)
    
    with st.expander("📋 최종 데이터 상세 (Click)"):
        piv = df_res.pivot_table(index='연', columns='그룹', values='값', aggfunc='sum').fillna(0)
        cols_in_piv = piv.columns.tolist()
        sorted_cols = [c for c in final_sort_order if c in cols_in_piv]
        remaining = [c for c in cols_in_piv if c not in sorted_cols]
        piv = piv[sorted_cols + remaining]
        piv['소계'] = piv.sum(axis=1)
        st.dataframe(piv.style.format("{:,.0f}"), use_container_width=True)

# ─────────────────────────────────────────────────────────
# 🟢 6. 사업계획 비교 (신규 기능 추가)
# ─────────────────────────────────────────────────────────
def render_plan_comparison(df_act, df_plan, df_action, unit_label):
    st.subheader(f"📊 2026 사업계획 비교 및 10개년 추이 ({unit_label})")
    
    # 데이터를 각각 Long form으로 변환
    long_act = make_long_data(df_act, "실적", 'supply')
    long_plan = make_long_data(df_plan, "기존 사업계획", 'supply')
    long_action = make_long_data(df_action, "실천사업계획", 'supply')
    
    # 단위 변환 (GJ)
    if "GJ" in unit_label:
        if not long_act.empty: long_act['값'] /= 1000
        if not long_plan.empty: long_plan['값'] /= 1000
        if not long_action.empty: long_action['값'] /= 1000

    st.markdown("---")
    st.markdown("#### 1️⃣ 2026년 기존 사업계획 vs 실천사업계획 비교")
    
    comp_data = []
    
    if not long_plan.empty:
        plan_2026 = long_plan[long_plan['연'] == 2026].groupby('그룹')['값'].sum().reset_index()
        plan_2026['구분'] = '기존 사업계획'
        comp_data.append(plan_2026)
        
    if not long_action.empty:
        action_2026 = long_action[long_action['연'] == 2026].groupby('그룹')['값'].sum().reset_index()
        action_2026['구분'] = '실천사업계획'
        comp_data.append(action_2026)
        
    if not comp_data:
        st.warning("비교할 2026년 데이터가 부족합니다.")
    else:
        comp_df = pd.concat(comp_data, ignore_index=True)
        
        # 묶은 막대 그래프
        fig_comp = px.bar(comp_df, x='그룹', y='값', color='구분', barmode='group', text_auto='.2s',
                          title="2026년 용도별 사업계획 비교")
        st.plotly_chart(fig_comp, use_container_width=True)
        
        # 피벗 테이블 생성
        piv_2026 = comp_df.pivot_table(index='그룹', columns='구분', values='값', aggfunc='sum').fillna(0)
        
        if '기존 사업계획' in piv_2026.columns and '실천사업계획' in piv_2026.columns:
            piv_2026['차이(실천-기존)'] = piv_2026['실천사업계획'] - piv_2026['기존 사업계획']
            
        # 열 순서 정렬
        ordered_cols = [c for c in ['기존 사업계획', '실천사업계획', '차이(실천-기존)'] if c in piv_2026.columns]
        piv_2026 = piv_2026[ordered_cols]
        
        st.dataframe(piv_2026.style.format("{:,.0f}"), use_container_width=True)

    st.markdown("---")
    st.markdown("#### 2️⃣ 2017년 ~ 2026년 용도별 공급량 추이")
    
    trend_data = []
    if not long_act.empty:
        act_hist = long_act[(long_act['연'] >= 2017) & (long_act['연'] <= 2025)]
        trend_data.append(act_hist)
        
    if not long_action.empty:
        action_2026_trend = long_action[long_action['연'] == 2026].copy()
        action_2026_trend['구분'] = '2026실천계획'
        trend_data.append(action_2026_trend)
        
    if not trend_data:
        st.warning("추이를 표시할 데이터가 없습니다.")
    else:
        trend_df = pd.concat(trend_data, ignore_index=True)
        trend_grp = trend_df.groupby(['연', '그룹'])['값'].sum().reset_index()
        
        # 누적 막대 그래프
        fig_trend = px.bar(trend_grp, x='연', y='값', color='그룹', text_auto='.2s',
                           title="용도별 공급량 10개년 추이 (2017~2026)")
        fig_trend.update_xaxes(dtick=1, tickformat="d")
        st.plotly_chart(fig_trend, use_container_width=True)
        
        # 데이터 표
        piv_trend = trend_grp.pivot_table(index='연', columns='그룹', values='값', aggfunc='sum').fillna(0)
        piv_trend['총합계'] = piv_trend.sum(axis=1)
        st.dataframe(piv_trend.style.format("{:,.0f}"), use_container_width=True)

# ─────────────────────────────────────────────────────────
# 🟢 7. 메인 실행
# ─────────────────────────────────────────────────────────
def main():
    st.title("🔥 도시가스 판매/공급 통합 분석")
    
    with st.sidebar:
        st.header("설정")
        mode = st.radio("분석 모드", ["1. 판매량", "2. 공급량"], index=1)
        
        sub_mode = ""
        if mode.startswith("2"):
            # 🟢 [수정됨] 5) 사업계획 비교 옵션 추가
            sub_mode = st.radio("기능 선택", ["1) 실적분석", "2) 2035 예측", "3) 상품별 예측", "4) 최종값 확인", "5) 사업계획 비교"])
        elif mode.startswith("1"):
            sub_mode = st.radio("기능 선택", ["1) 실적분석"])
        
        idx = 0 
        if mode.startswith("1"): idx = 0 
        
        unit_opts = ["열량 (GJ)", "부피 (천m³)"]
        unit = st.radio("단위 선택", unit_opts, index=idx)
        unit_key = "열량" if "열량" in unit else "부피"
        
        st.markdown("---")
        st.subheader("파일 업로드")
        
        up_sales = st.file_uploader("1. 판매량(계획_실적).xlsx", type=["xlsx", "csv"], key="s", accept_multiple_files=True)
        up_supply = st.file_uploader("2. 공급량실적_계획_실적_MJ.xlsx", type=["xlsx", "csv"], key="p")
        up_final = st.file_uploader("3. 최종값.xlsx (결과파일)", type=["xlsx", "csv"], key="f")
        st.markdown("---")
    
    df_final = pd.DataFrame()
    start_year = 2026
    is_supply = False
    
    # 🟢 [모드 1] 판매량
    if mode.startswith("1"):
        start_year = 2026
        if up_sales:
            data_dict = load_all_sheets(up_sales[0] if isinstance(up_sales, list) else up_sales)
            target_df = None
            for sheet_name, df in data_dict.items():
                if "실적" in sheet_name and unit_key in sheet_name:
                    target_df = df; break
            
            if target_df is None:
                for sheet_name, df in data_dict.items():
                    if "실적" in sheet_name: target_df = df; break
            
            if target_df is None and len(data_dict) > 0:
                target_df = list(data_dict.values())[0]

            if target_df is not None:
                long_a = make_long_data(target_df, "실적", 'sales')
                long_a = long_a[long_a['연'] <= 2025] 
                df_final = pd.concat([long_a], ignore_index=True)
        else: st.info("👈 [판매량 파일]을 업로드하세요."); return

    # 🟢 [모드 2] 공급량
    elif mode.startswith("2"):
        start_year = 2029 
        is_supply = True
        
        if "최종값" in sub_mode:
            if up_final:
                data_dict = load_all_sheets(up_final)
                if len(data_dict) > 0:
                    df_raw = list(data_dict.values())[0]
                    df_final = make_long_data(df_raw, "최종값", mode='detail')
            else:
                st.info("👈 [최종값 파일]을 업로드하세요."); return
        
        # 🟢 [수정됨] 사업계획 비교 로직 분기 처리
        elif "사업계획 비교" in sub_mode:
            if up_supply:
                data_dict = load_all_sheets(up_supply)
                
                # 시트명 기반으로 데이터프레임 추출
                df_act = next((df for name, df in data_dict.items() if "실적" in name and "계획" not in name), None)
                df_plan = next((df for name, df in data_dict.items() if "사업계획" in name and "실천" not in name), None)
                df_action = next((df for name, df in data_dict.items() if "실천" in name), None)
                
                # 시트 이름 매칭 실패 시 fallback (딕셔너리 순서)
                if df_act is None and len(data_dict) >= 1: df_act = list(data_dict.values())[0]
                if df_plan is None and len(data_dict) >= 2: df_plan = list(data_dict.values())[1]
                if df_action is None and len(data_dict) >= 3: df_action = list(data_dict.values())[2]
                
                if df_act is not None and df_plan is not None and df_action is not None:
                    render_plan_comparison(df_act, df_plan, df_action, unit)
                else:
                    st.warning("⚠️ 업로드된 파일 내에 '실적', '사업계획', '실천사업계획' 관련 시트나 데이터가 충분하지 않습니다.")
                return # 렌더링 후 조기 종료 (아래 공통 실행 생략)
            else:
                st.info("👈 좌측에서 [공급량 파일]을 업로드하세요."); return

        else: # 실적분석, 2035 예측, 상품별 예측
            if up_supply:
                data_dict = load_all_sheets(up_supply)
                df_hist, df_plan = None, None
                for name, df in data_dict.items():
                    if "실적" in name: df_hist = df; break
                for name, df in data_dict.items():
                    if "계획" in name: df_plan = df; break
                
                if df_hist is None and len(data_dict) > 0: df_hist = list(data_dict.values())[0]
                
                if df_hist is not None:
                    long_h = make_long_data(df_hist, "실적", 'supply')
                    df_final = long_h
                    if df_plan is not None:
                        long_p = make_long_data(df_plan, "확정계획", 'supply')
                        df_final = pd.concat([long_h, long_p], ignore_index=True)
            else: st.info("👈 [공급량 파일]을 업로드하세요."); return

    # ── 공통 실행 ──
    if not df_final.empty:
        if (mode.startswith("2") or mode.startswith("3")) and "GJ" in unit:
            if "최종값" not in sub_mode: 
                df_final['값'] = df_final['값'] / 1000

        if "최종값" not in sub_mode:
            with st.sidebar:
                st.markdown("### 📅 데이터 학습 기간 설정")
                all_years = sorted([int(y) for y in df_final['연'].unique()])
                default_yrs = all_years 
                train_years = st.multiselect("학습 연도 (2025년 포함됨)", options=all_years, default=default_yrs)

        if "최종값" in sub_mode:
            render_final_check(df_final, unit)
        
        elif "실적분석" in sub_mode:
            render_analysis_dashboard(df_final, unit)
        
        elif "2035" in sub_mode:
            render_prediction_2035(df_final, unit, start_year, train_years, is_supply, custom_sort_list=None)
        
        elif "상품별" in sub_mode:
            df_detail = pd.DataFrame()
            if mode.startswith("1") and up_sales:
                dd = load_all_sheets(up_sales[0] if isinstance(up_sales, list) else up_sales)
                tgt = None
                for sn, d in dd.items():
                    if "실적" in sn and unit_key in sn: tgt = d; break
                if tgt is None:
                    for sn, d in dd.items():
                        if "실적" in sn: tgt = d; break
                if tgt is not None:
                    df_detail = make_long_data(tgt, "실적", mode='detail')
                    df_detail = df_detail[df_detail['연'] <= 2025]

            elif mode.startswith("2") and up_supply:
                dd = load_all_sheets(up_supply)
                dh, dp = None, None
                for n, d in dd.items():
                    if "실적" in n: dh = d; break
                for n, d in dd.items():
                    if "계획" in n: dp = d; break
                if dh is None and len(dd)>0: dh = list(dd.values())[0]

                if dh is not None:
                    ld_h = make_long_data(dh, "실적", mode='detail')
                    df_detail = ld_h
                    if dp is not None:
                        ld_p = make_long_data(dp, "확정계획", mode='detail')
                        df_detail = pd.concat([ld_h, ld_p], ignore_index=True)
            
            if not df_detail.empty:
                if (mode.startswith("2")) and "GJ" in unit:
                    df_detail['값'] = df_detail['값'] / 1000
                
                render_prediction_2035(df_detail, unit, start_year, train_years, is_supply, custom_sort_list=ORDER_LIST_DETAIL)

if __name__ == "__main__":
    main()
