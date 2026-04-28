"""
员工离职数据分析系统 - Streamlit版本
"""
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import re
import warnings
warnings.filterwarnings('ignore')

st.set_page_config(page_title="员工离职数据分析", layout="wide")

st.markdown("""
<div style="background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%); padding: 20px; border-radius: 10px; margin-bottom: 20px;">
    <h2 style="color: white; margin: 0;">员工离职数据分析系统</h2>
</div>
""", unsafe_allow_html=True)


def extract_month(sheet_name):
    patterns = [r'(\d{4})[年-](\d{1,2})[月]', r'(\d{4})(\d{2})']
    for pattern in patterns:
        match = re.search(pattern, str(sheet_name))
        if match:
            return f"{match.group(1)}年{int(match.group(2)):02d}月"
    return None


def categorize_tenure(df):
    tenure_col = '累计司龄（年）'
    if tenure_col not in df.columns:
        return pd.DataFrame()
    def tenure_category(years):
        if pd.isna(years): return '未知'
        elif years <= 0.5: return '0.5年及以下(新人)'
        elif years <= 1: return '0.5-1年'
        elif years <= 3: return '1-3年'
        elif years <= 5: return '3-5年'
        else: return '5年以上'
    df_copy = df.copy()
    df_copy['司龄段'] = df_copy[tenure_col].apply(tenure_category)
    return df_copy.groupby('司龄段').size().reset_index(name='人数')


def get_monthly_analysis(month, turnover_data, months, monthly_data):
    if turnover_data is None: return None
    month_turnover = turnover_data[turnover_data['离职月份'] == month].copy()
    month_idx = months.index(month) if month in months else -1
    start_count = len(monthly_data.get(month, pd.DataFrame()))
    if month_idx < len(months) - 1:
        next_month = months[month_idx + 1]
        end_count = len(monthly_data.get(next_month, pd.DataFrame()))
    else:
        month_turnover_count = len(turnover_data[turnover_data['离职月份'] == month]) if turnover_data is not None else 0
        end_count = start_count - month_turnover_count
    avg_count = (start_count + end_count) / 2
    turnover_count = len(month_turnover)
    turnover_rate = (turnover_count / avg_count * 100) if avg_count > 0 else 0
    
    dept_turnover = pd.DataFrame()
    if len(month_turnover) > 0 and '一级组织' in month_turnover.columns:
        dept_stats = month_turnover.groupby('一级组织').size().reset_index(name='离职人数')
        dept_rates = []
        for dept in dept_stats['一级组织']:
            dt_count = len(month_turnover[month_turnover['一级组织'] == dept])
            d_count = 0
            if month in monthly_data and '一级组织' in monthly_data[month].columns:
                d_count = len(monthly_data[month][monthly_data[month]['一级组织'] == dept])
            dept_rates.append(round((dt_count / d_count * 100) if d_count > 0 else 0, 2))
        dept_stats['离职率(%)'] = dept_rates
        dept_turnover = dept_stats.sort_values('离职率(%)', ascending=False)
    
    turnover_type = pd.DataFrame()
    if len(month_turnover) > 0 and '离职类型' in month_turnover.columns:
        turnover_type = month_turnover.groupby('离职类型').size().reset_index(name='人数')
        turnover_type['占比(%)'] = (turnover_type['人数'] / turnover_count * 100).round(2)
        turnover_type = turnover_type.sort_values('人数', ascending=False)
    
    turnover_reason = pd.DataFrame()
    if len(month_turnover) > 0 and '离职原因' in month_turnover.columns:
        turnover_reason = month_turnover.groupby('离职原因').size().reset_index(name='人数')
        turnover_reason['占比(%)'] = (turnover_reason['人数'] / turnover_count * 100).round(2)
        turnover_reason = turnover_reason.sort_values('人数', ascending=False)
    
    turnover_tenure = categorize_tenure(month_turnover)
    if len(turnover_tenure) > 0:
        turnover_tenure['占比(%)'] = (turnover_tenure['人数'] / turnover_count * 100).round(2)
        turnover_tenure = turnover_tenure.sort_values('人数', ascending=False)
    
    turnover_level = pd.DataFrame()
    if len(month_turnover) > 0 and '职级' in month_turnover.columns:
        turnover_level = month_turnover.groupby('职级').size().reset_index(name='人数')
        turnover_level['占比(%)'] = (turnover_level['人数'] / turnover_count * 100).round(2)
        turnover_level = turnover_level.sort_values('人数', ascending=False)
    
    return {'month': month, 'avg_count': round(avg_count, 2), 'turnover_count': turnover_count,
            'turnover_rate': round(turnover_rate, 2), 'dept_turnover': dept_turnover,
            'turnover_type': turnover_type, 'turnover_reason': turnover_reason,
            'turnover_tenure': turnover_tenure, 'turnover_level': turnover_level}


uploaded_file = st.file_uploader("上传Excel文件", type=['xlsx', 'xls'])
if uploaded_file:
    xlsx = pd.ExcelFile(uploaded_file)
    monthly_data = {}
    turnover_data = None
    for sheet in xlsx.sheet_names:
        try:
            df = pd.read_excel(xlsx, sheet_name=sheet)
            if '姓名' in df.columns and '身份类别' in df.columns:
                month = extract_month(sheet)
                if month: monthly_data[month] = df[df['身份类别'] != '实习生'].copy()
            elif '离职类型' in df.columns:
                turnover_data = df
                if '最后工作日' in turnover_data.columns:
                    turnover_data['最后工作日'] = pd.to_datetime(turnover_data['最后工作日'], errors='coerce')
                if '离职月份' not in turnover_data.columns:
                    turnover_data['离职月份'] = turnover_data['最后工作日'].dt.strftime('%Y年%m月')
        except: continue
    
    months = sorted(monthly_data.keys(), reverse=True)
    if months:
        selected_month = st.selectbox("选择分析月份", months)
        if st.button("开始分析"):
            analysis = get_monthly_analysis(selected_month, turnover_data, months, monthly_data)
            if analysis:
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("分析月份", analysis['month'])
                c2.metric("平均人数", analysis['avg_count'])
                c3.metric("离职人数", analysis['turnover_count'])
                c4.metric("离职率", str(analysis['turnover_rate']) + '%')
                st.divider()
                st.subheader("各部门离职分析")
                if len(analysis['dept_turnover']) > 0:
                    cc1, cc2 = st.columns(2)
                    with cc1:
                        st.markdown("**离职人数**")
                        st.bar_chart(analysis['dept_turnover'].set_index('一级组织')['离职人数'])
                    with cc2:
                        st.markdown("**离职率(%)**")
                        st.bar_chart(analysis['dept_turnover'].set_index('一级组织')['离职率(%)'])
                    st.dataframe(analysis['dept_turnover'], use_container_width=True, hide_index=True)
                st.divider()
                cl, cr = st.columns(2)
                with cl:
                    st.subheader("离职类型分布")
                    if len(analysis['turnover_type']) > 0:
                        st.dataframe(analysis['turnover_type'], use_container_width=True, hide_index=True)
                with cr:
                    st.subheader("离职类型占比")
                    if len(analysis['turnover_type']) > 0:
                        st.bar_chart(analysis['turnover_type'].set_index('离职类型')['人数'])
                st.divider()
                tl, tr = st.columns(2)
                with tl:
                    st.subheader("离职司龄分布")
                    if len(analysis['turnover_tenure']) > 0:
                        st.dataframe(analysis['turnover_tenure'], use_container_width=True, hide_index=True)
                with tr:
                    st.subheader("司龄占比")
                    if len(analysis['turnover_tenure']) > 0:
                        st.bar_chart(analysis['turnover_tenure'].set_index('司龄段')['人数'])
                st.divider()
                rl, rr = st.columns(2)
                with rl:
                    st.subheader("离职原因分析")
                    if len(analysis['turnover_reason']) > 0:
                        st.dataframe(analysis['turnover_reason'], use_container_width=True, hide_index=True)
                with rr:
                    st.subheader("离职原因占比")
                    if len(analysis['turnover_reason']) > 0:
                        st.bar_chart(analysis['turnover_reason'].set_index('离职原因')['人数'])
                st.divider()
                ll, lr = st.columns(2)
                with ll:
                    st.subheader("离职职级分布")
                    if len(analysis['turnover_level']) > 0:
                        st.dataframe(analysis['turnover_level'], use_container_width=True, hide_index=True)
                with lr:
                    st.subheader("职级占比")
                    if len(analysis['turnover_level']) > 0:
                        st.bar_chart(analysis['turnover_level'].set_index('职级')['人数'])
                st.divider()
                if st.button("导出Excel"):
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        pd.DataFrame([{'月份': analysis['month'], '平均人数': analysis['avg_count'],
                                       '离职人数': analysis['turnover_count'], '离职率': str(analysis['turnover_rate']) + '%'}]).to_excel(writer, sheet_name='概览', index=False)
                        analysis['dept_turnover'].to_excel(writer, sheet_name='部门离职', index=False)
                        analysis['turnover_type'].to_excel(writer, sheet_name='离职类型', index=False)
                        analysis['turnover_reason'].to_excel(writer, sheet_name='离职原因', index=False)
                        analysis['turnover_tenure'].to_excel(writer, sheet_name='司龄分布', index=False)
                        analysis['turnover_level'].to_excel(writer, sheet_name='职级分布', index=False)
                    st.download_button(label="下载Excel文件", data=output.getvalue(),
                                       file_name=analysis['month'] + '离职分析.xlsx',
                                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')