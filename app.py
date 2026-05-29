import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np

# 페이지 기본 설정
st.set_page_config(page_title="기온 분석 및 이상기온 대시보드", layout="wide")

st.title("🌡️ 기온 분석 및 이상기온 모니터링 대시보드")

# 1. 데이터 로드 (구글 시트 연동 유지)
@st.cache_data(ttl=600)
def load_data():
    sheet_id = "13HrIz6OytYDykXeXzXJ02I6XbaKin1YaKBoO2kBd6Bs"
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid=0"
    df = pd.read_csv(url)
    
    if '공급량(M3)' in df.columns:
        df['공급량(M3)'] = pd.to_numeric(df['공급량(M3)'].astype(str).str.replace(',', ''), errors='coerce')
        
    df['일자'] = pd.to_datetime(df['일자'])
    df['연도'] = df['일자'].dt.year
    df['월'] = df['일자'].dt.month
    df['일'] = df['일자'].dt.day
    return df

df = load_data()

max_year = int(df['연도'].max())
target_year = max_year

# 2. 통합 컨트롤 UI
st.subheader("🗓️ 전체 분석 월 선택")
selected_month = st.selectbox("", list(range(1, 13)), index=0, format_func=lambda x: f"{x}월")

st.markdown("---")

monthly_avg = df[df['월'] == selected_month].groupby('연도')['평균기온(℃)'].mean().reset_index()

if '공급량(M3)' in df.columns:
    supply_data = df[df['월'] == selected_month].groupby('연도')['공급량(M3)'].sum().reset_index()
else:
    supply_data = pd.DataFrame({'연도': monthly_avg['연도'], '공급량(M3)': 0})

def get_supply(yr):
    val = supply_data[supply_data['연도'] == yr]['공급량(M3)'].values
    return val[0] if len(val) > 0 and not pd.isna(val[0]) else 0

valid_supplies = [get_supply(y) for y in monthly_avg['연도'] if get_supply(y) > 0]
max_s, min_s = (max(valid_supplies), min(valid_supplies)) if valid_supplies else (1, 0)

def scale_size(v):
    if v == 0 or max_s == min_s: return 15
    return 15 + ((v - min_s) / (max_s - min_s)) * 35 

# ---------------------------------------------------------
# 3. 최상단: 월 평균기온 현황 요약 표 (★ 배경 하이라이트 및 레이아웃 수정)
# ---------------------------------------------------------
st.header(f"📋 {selected_month}월 평균기온 현황 요약")

if target_year in monthly_avg['연도'].values and len(monthly_avg[monthly_avg['연도'] < target_year]) >= 7:
    past_7_years = monthly_avg[(monthly_avg['연도'] >= target_year - 7) & (monthly_avg['연도'] < target_year)].copy()
    past_7_years_list = past_7_years['연도'].tolist()
    
    curr_val = round(monthly_avg[monthly_avg['연도'] == target_year]['평균기온(℃)'].values[0], 1)
    mean_7yr = past_7_years['평균기온(℃)'].mean()
    
    max_idx = past_7_years['평균기온(℃)'].idxmax()
    min_idx = past_7_years['평균기온(℃)'].idxmin()
    past_5_years = past_7_years.drop(index=[max_idx, min_idx])
    
    mean_5yr = round(past_5_years['평균기온(℃)'].mean(), 1)
    std_5yr = np.sqrt(np.sum((past_5_years['평균기온(℃)'] - mean_5yr)**2) / 5)
    is_abnormal = abs(curr_val - mean_5yr) > std_5yr

    past_7_temps = [round(monthly_avg[monthly_avg['연도'] == yr]['평균기온(℃)'].values[0], 1) for yr in past_7_years_list]
    max_t, min_t = max(past_7_temps), min(past_7_temps)

    # 표 구조 변경: 1. 월 평균, 2. 판정 결과, 3. 5년 평균, 4. 표준편차
    table_data = {'구분': ['월 평균', '판정 결과', '5년 평균 :', '표준 편차 :']}
    
    # 당해연도 이상기온 문구 설정
    if is_abnormal:
        abnormal_type = "(이상고온)" if curr_val > mean_5yr else "(이상저온)"
        target_str = f"{curr_val} {abnormal_type}"
        judgment_str = f"🚨 이상 {abnormal_type}"
    else:
        target_str = f"{curr_val}"
        judgment_str = "✅ 정상"
    
    for i, yr in enumerate(past_7_years_list):
        t = past_7_temps[i]
        # 첫 번째 연도(맨 앞 칸)에만 5년 평균과 표준편차를 넣고 나머지는 비움
        if i == 0:
            table_data[str(yr)] = [f"{t}", '', f"{mean_5yr:.1f}℃", f"{std_5yr:.4f}"]
        else:
            table_data[str(yr)] = [f"{t}", '', '', '']

    table_data['7년 평균'] = [f"{mean_7yr:.1f}", '', '', '']
    table_data[str(target_year)] = [target_str, judgment_str, '', '']

    df_table = pd.DataFrame(table_data)

    # ★ 데이터프레임 스타일링 함수
    def apply_highlight(x):
        df_style = pd.DataFrame('', index=x.index, columns=x.columns)
        
        # 1. 월 평균(첫 번째 줄)의 전체 배경색 및 7년 최고/최저 글자색 변경
        for col in x.columns:
            if col in ['구분', '7년 평균']:
                continue
                
            val_str = x.loc[0, col]
            try:
                # "0.3 (이상저온)" 등에서 숫자만 추출하여 비교
                val = float(str(val_str).split()[0])
                
                # 5년 평균 기준 배경색 설정
                bg_color = '#ffebee' if val > mean_5yr else ('#e3f2fd' if val < mean_5yr else '')
                
                text_color = 'black'
                font_weight = 'normal'
                
                # 과거 7년에 대해서만 최고/최저 글자색 변경
                if col in [str(y) for y in past_7_years_list]:
                    if val == max_t:
                        text_color = '#d32f2f' # 최고기온 붉은색 글자
                        font_weight = 'bold'
                    elif val == min_t:
                        text_color = '#1976d2' # 최저기온 푸른색 글자
                        font_weight = 'bold'
                elif col == str(target_year):
                    if is_abnormal:
                        font_weight = 'bold'

                style_str = f'color: {text_color}; font-weight: {font_weight};'
                if bg_color:
                    style_str += f' background-color: {bg_color};'
                
                df_style.loc[0, col] = style_str
            except Exception:
                pass

        # 2. 당해연도 '판정 결과'(두 번째 줄) 텍스트 색상 강조
        if is_abnormal:
            col_idx = x.columns.get_loc(str(target_year))
            text_color = '#d32f2f' if curr_val > mean_5yr else '#1976d2'
            df_style.iloc[1, col_idx] = f'color: {text_color}; font-weight: bold;'
            
        return df_style

    styled_table = df_table.style.apply(apply_highlight, axis=None)
    st.dataframe(styled_table, hide_index=True)

    # 하단에 이상기온 판단 기준 간단 명시
    st.caption("※ **이상기온 판단 기준**: 과거 7년 중 최고·최저 기온을 제외한 5년 평균 기온과의 차이 절대값이 5년 표준편차보다 큰 경우")

else:
    st.info("데이터가 충분하지 않습니다.")

st.markdown("---")

# ---------------------------------------------------------
# 4. 이상기온 판별 상세 (공급량 비례 버블 차트 유지)
# ---------------------------------------------------------
st.header(f"🚨 {selected_month}월 이상기온 판별 상세 (한국가스공사 기준)")

if is_abnormal:
    st.error(f"🚨 **주의:** {target_year}년 {selected_month}월은 가스공사 기준 **'이상기온'**으로 판별되었습니다! (도시가스 실적 변동에 유의하세요)", icon="🚨")
else:
    st.success(f"✅ {target_year}년 {selected_month}월은 가스공사 기준 **'정상 기온'** 범위 내에 있습니다.", icon="✅")

st.write("")

m_col1, m_col2, m_col3, m_col4 = st.columns(4)
m_col1.metric("최근 7년 평균", f"{mean_7yr:.1f}℃")
m_col2.metric("5년 평균 (최대/최소 제외)", f"{mean_5yr:.1f}℃")
m_col3.metric("5년 표준편차", f"{std_5yr:.2f}")

status_text = "이상기온 발생!" if is_abnormal else "정상 범위"
status_color = "inverse" if is_abnormal else "normal"
m_col4.metric(f"기준 연도({target_year}) 평균", f"{curr_val:.1f}℃", delta=status_text, delta_color=status_color)

fig_abnormal = go.Figure()
lower_bound = mean_5yr - std_5yr
upper_bound = mean_5yr + std_5yr

fig_abnormal.add_vrect(
    x0=lower_bound, x1=upper_bound,
    fillcolor="#00BFFF", opacity=0.1, layer="below", 
    line_width=1, line_dash="solid", line_color="#00BFFF",
)

fig_abnormal.add_annotation(x=lower_bound, y=0.9, text=f"<b>하한: {lower_bound:.2f}℃</b>", showarrow=False, font=dict(color="#00BFFF", size=13), xanchor="right", yanchor="top")
fig_abnormal.add_annotation(x=upper_bound, y=0.9, text=f"<b>상한: {upper_bound:.2f}℃</b>", showarrow=False, font=dict(color="#00BFFF", size=13), xanchor="left", yanchor="top")

fig_abnormal.add_vline(x=mean_5yr, line=dict(color="#2ca02c", width=2, dash="dash"), annotation_text=f"5년 평균 ({mean_5yr:.1f}℃)", annotation_position="bottom right", annotation_font=dict(color="#2ca02c"))

sizes_7yr, texts_7yr, x_vals = [], [], []
for _, row in past_7_years.iterrows():
    y = int(row['연도'])
    if y == target_year - 1: continue 
    s = get_supply(y)
    sizes_7yr.append(scale_size(s))
    texts_7yr.append(f"<b>{y}년</b><br>평균기온: {row['평균기온(℃)']:.1f}℃<br>공급량: {s:,.0f} M3")
    x_vals.append(row['평균기온(℃)'])

fig_abnormal.add_trace(go.Scatter(
    x=x_vals, y=[0]*len(x_vals), mode="markers", name="과거 7년",
    marker=dict(size=sizes_7yr, color="rgba(128, 128, 128, 0.5)", line=dict(color="gray", width=1)),
    hoverinfo="text", text=texts_7yr
))

y1_year = target_year - 1
if y1_year in past_7_years['연도'].values:
    y1_val = past_7_years[past_7_years['연도'] == y1_year]['평균기온(℃)'].values[0]
    y1_s = get_supply(y1_year)
    fig_abnormal.add_trace(go.Scatter(
        x=[y1_val], y=[0], mode="markers+text", name=f"작년({y1_year}년)",
        marker=dict(size=scale_size(y1_s), color="#9467bd", symbol="diamond", line=dict(color="white", width=2)),
        text=[f"<b>{y1_year}년 ({y1_val:.1f}℃)</b>"], textposition="top center", textfont=dict(size=13, color="#9467bd"),
        hoverinfo="text", hovertext=f"<b>{y1_year}년</b><br>평균기온: {y1_val:.1f}℃<br>공급량: {y1_s:,.0f} M3"
    ))

curr_s = get_supply(target_year)
fig_abnormal.add_trace(go.Scatter(
    x=[curr_val], y=[0], mode="markers+text", name=f"{target_year}년",
    marker=dict(size=scale_size(curr_s), color="#FF8C00", symbol="circle", line=dict(color="white", width=2)),
    text=[f"<b>{target_year}년 ({curr_val}℃)</b>"], textposition="bottom center", textfont=dict(size=15, color="#FF8C00"),
    hoverinfo="text", hovertext=f"<b>{target_year}년</b><br>평균기온: {curr_val:.1f}℃<br>공급량: {curr_s:,.0f} M3"
))

fig_abnormal.update_layout(
    height=320, 
    yaxis=dict(showticklabels=False, range=[-1.3, 1.3], showgrid=False, zeroline=False), 
    xaxis=dict(title="평균기온(℃)", gridcolor="#f0f0f0", showline=True, linecolor='lightgray'),
    showlegend=False, margin=dict(l=40, r=40, t=20, b=40), plot_bgcolor="white"
)
st.plotly_chart(fig_abnormal, use_container_width=True)

st.markdown("---")

# ---------------------------------------------------------
# 5. 일별 평균기온 매트릭스 (31일 및 평균 누락 강제 방지 유지)
# ---------------------------------------------------------
st.header("📊 상세 일별 기온 매트릭스")

min_year_val = int(df['연도'].min())
selected_years = st.slider("기온 매트릭스 연도 범위 설정", min_year_val, max_year, (min_year_val, max_year))

filtered_df = df[(df['월'] == selected_month) & 
                 (df['연도'] >= selected_years[0]) & 
                 (df['연도'] <= selected_years[1])]

pivot_df = filtered_df.pivot(index='일', columns='연도', values='평균기온(℃)')
pivot_df = pivot_df.reindex(list(range(1, 32)))

avg_series = filtered_df.groupby('연도')['평균기온(℃)'].mean()
pivot_df.loc['평균'] = avg_series 

pivot_df.index = pivot_df.index.astype(str)
pivot_df.columns = pivot_df.columns.astype(str)

fig_heatmap = px.imshow(
    pivot_df,
    labels=dict(x="연도", y="일", color="평균기온(℃)"),
    x=pivot_df.columns,
    y=pivot_df.index,
    color_continuous_scale='RdBu_r',
    aspect="auto",
    text_auto='.1f'  
)

fig_heatmap.update_layout(
    height=1000, 
    margin=dict(l=40, r=40, t=80, b=40),
    yaxis=dict(
        type='category',
        tickmode='linear',
        dtick=1,
        autorange='reversed'
    ),
    xaxis=dict(
        type='category', 
        tickmode='linear', 
        dtick=1, 
        side='top'
    )
)
fig_heatmap.update_traces(textfont=dict(size=14))

st.plotly_chart(fig_heatmap, use_container_width=True)

# =====================================================================
# 6. 추가된 기능: 10년치 평균기온 매트릭스 및 미실적 월 예측
# =====================================================================
st.markdown("---")

# 하단 '평균기온' 활성화 토글 버튼
if st.toggle("📈 평균기온 10년 분석 및 미실적 월 예측 활성화"):
    st.header("🔮 연도별 월별 평균기온 및 예측 시뮬레이션")
    
    # 1. 10년치 연도별 월별 평균기온 박스 생성
    # 전체 월별 평균기온 계산 및 피벗 테이블 생성
    monthly_all = df.groupby(['연도', '월'])['평균기온(℃)'].mean().reset_index()
    pivot_all = monthly_all.pivot(index='연도', columns='월', values='평균기온(℃)')
    
    # 1~12월 컬럼이 모두 존재하도록 보장
    for m in range(1, 13):
        if m not in pivot_all.columns:
            pivot_all[m] = np.nan
    pivot_all = pivot_all[list(range(1, 13))]
    
    # 최근 10년 범위 설정
    start_year = target_year - 9
    recent_10_years = list(range(start_year, target_year + 1))
    
    # 데이터가 없는 연도에 빈 행 추가
    for y in recent_10_years:
        if y not in pivot_all.index:
            pivot_all.loc[y] = [np.nan] * 12
            
    # 최근 10년 데이터만 추출 후 정렬
    pivot_10yr = pivot_all.loc[recent_10_years].copy().sort_index()
    
    st.subheader("📊 최근 10년 연도별/월별 평균기온 매트릭스")
    # 사진과 유사하게 배경에 색상이 들어간 형태로 스타일링 적용 (결측치는 빈칸 표시)
    styled_pivot = pivot_10yr.style.format("{:.1f}", na_rep="") \
                                   .background_gradient(cmap='RdYlBu_r', axis=None) \
                                   .set_properties(**{'text-align': 'center'})
    
    st.dataframe(styled_pivot, use_container_width=True)
    
    # 2. 당해연도(target_year) 기온(실적이 없는 월) 예측
    st.subheader(f"💡 {target_year}년 미실적 월 평균기온 예측 (4가지 시나리오)")
    
    # target_year 기준 실적이 없는(결측치인) 월 찾기
    target_year_data = pivot_10yr.loc[target_year]
    missing_months = target_year_data[target_year_data.isna()].index.tolist()
    
    if missing_months:
        for m in missing_months:
            with st.expander(f"📌 {target_year}년 {m}월 기온 예측 상세", expanded=True):
                # 과거 10년 중 당해 연도를 제외한 해당 월의 실적 데이터 추출
                hist_data = pivot_all.loc[start_year:target_year-1, m].dropna()
                
                if len(hist_data) >= 3:
                    # 1) 3y평균 (최근 3년 단순 산술 평균)
                    mean_3y = hist_data.iloc[-3:].mean()
                    
                    # 2 & 3) 이상기온(최고/최저) 제외 산술 평균, Max, Min
                    max_val = hist_data.max()
                    min_val = hist_data.min()
                    
                    # 최고치와 최저치를 제외한 나머지 데이터
                    hist_ex_abnormal = hist_data[~hist_data.index.isin([hist_data.idxmax(), hist_data.idxmin()])]
                    mean_ex_abnormal = hist_ex_abnormal.mean() if len(hist_ex_abnormal) > 0 else hist_data.mean()
                    
                    # 4) 선형추세를 이용한 평균기온 (1차 다항식 회귀)
                    x = hist_data.index.values
                    y = hist_data.values
                    z = np.polyfit(x, y, 1)
                    p = np.poly1d(z)
                    trend_val = p(target_year)
                    
                    # 4분할 화면으로 깔끔하게 지표 표시
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("① 3년 평균 (산술)", f"{mean_3y:.1f}℃")
                    c2.metric("② 10년 이상기온 제외 평균", f"{mean_ex_abnormal:.1f}℃", help="과거 데이터 중 Max, Min 1개씩 제외")
                    c3.metric("③ 10년 Max / Min", f"{max_val:.1f}℃ / {min_val:.1f}℃")
                    c4.metric("④ 선형추세 예측기온", f"{trend_val:.1f}℃", help=f"과거 {len(hist_data)}년 추세 반영")
                else:
                    st.warning(f"{m}월의 과거 데이터가 부족하여 예측 모델을 돌릴 수 없습니다.")
    else:
        st.info(f"✅ {target_year}년의 모든 월 데이터가 이미 업데이트되어 있어 미실적 월이 없습니다.")
