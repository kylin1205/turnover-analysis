"""
员工离职数据分析系统
离职率 = 离职人数 ÷ 平均人数
平均人数 = (期初人数 + 期末人数) / 2
"""
import streamlit as st
import pandas as pd
from io import BytesIO
import re
import warnings
warnings.filterwarnings('ignore')

st.set_page_config(page_title="员工离职数据分析", layout="wide")
st.markdown("<h2>员工离职数据分析系统</h2>", unsafe_allow_html=True)

def get_month(month_num):
    return "2026年" + str(month_num).zfill(2) + "月"

def extract_month(sheet_name):
    m = re.search(r'(\d{1,2})月', str(sheet_name))
    return int(m.group(1)) if m else None

def analyze(month_num, period_data, turnover_data):
    ms = get_month(month_num)
    start_df = period_data.get(month_num, {}).get('start')
    end_df = period_data.get(month_num, {}).get('end')
    
    if start_df is None or len(start_df) == 0:
        st.error("未找到该月期初数据"); return None
    
    start_cnt = len(start_df)
    end_cnt = len(end_df) if end_df is not None and len(end_df) > 0 else start_cnt
    avg_cnt = (start_cnt + end_cnt) / 2
    
    month_turn = pd.DataFrame()
    if turnover_data is not None and '离职月份' in turnover_data.columns:
        month_turn = turnover_data[turnover_data['离职月份'] == ms].copy()
    turn_cnt = len(month_turn)
    turn_rate = round((turn_cnt / avg_cnt * 100), 2) if avg_cnt > 0 else 0
    
    dept_df = pd.DataFrame()
    if len(month_turn) > 0 and '一级组织' in month_turn.columns:
        d1 = month_turn.groupby('一级组织').size().reset_index(name='离职人数')
        d2 = pd.DataFrame()
        if start_df is not None and '一级组织' in start_df.columns:
            d2 = start_df.groupby('一级组织').size().reset_index(name='期初人数')
        d3 = pd.DataFrame()
        if end_df is not None and len(end_df) > 0 and '一级组织' in end_df.columns:
            d3 = end_df.groupby('一级组织').size().reset_index(name='期末人数')
        dept_df = d1.copy()
        if len(d2) > 0:
            dept_df = dept_df.merge(d2, on='一级组织', how='left')
        else:
            dept_df['期初人数'] = 0
        if len(d3) > 0:
            dept_df = dept_df.merge(d3, on='一级组织', how='left')
        else:
            dept_df['期末人数'] = dept_df['期初人数']
        dept_df['期初人数'] = dept_df['期初人数'].fillna(0).astype(int)
        dept_df['期末人数'] = dept_df['期末人数'].fillna(0).astype(int)
        dept_df['平均人数'] = (dept_df['期初人数'] + dept_df['期末人数']) / 2
        dept_df['离职率(%)'] = dept_df.apply(lambda x: round((x['离职人数']/x['平均人数']*100) if x['平均人数']>0 else 0, 2), axis=1)
        dept_df = dept_df.sort_values('离职率(%)', ascending=False)
    
    type_df = pd.DataFrame()
    if len(month_turn) > 0 and '离职类型' in month_turn.columns:
        type_df = month_turn.groupby('离职类型').size().reset_index(name='人数')
        type_df['占比(%)'] = round(type_df['人数']/turn_cnt*100, 2) if turn_cnt > 0 else 0
        type_df = type_df.sort_values('人数', ascending=False)
    
    reason_df = pd.DataFrame()
    if len(month_turn) > 0 and '离职原因' in month_turn.columns:
        reason_df = month_turn.groupby('离职原因').size().reset_index(name='人数')
        reason_df['占比(%)'] = round(reason_df['人数']/turn_cnt*100, 2) if turn_cnt > 0 else 0
        reason_df = reason_df.sort_values('人数', ascending=False)
    
    level_df = pd.DataFrame()
    if len(month_turn) > 0 and '职级' in month_turn.columns:
        level_df = month_turn.groupby('职级').size().reset_index(name='人数')
        level_df['占比(%)'] = round(level_df['人数']/turn_cnt*100, 2) if turn_cnt > 0 else 0
        level_df = level_df.sort_values('人数', ascending=False)
    
    return {'month': ms, 'start': start_cnt, 'end': end_cnt, 'avg': round(avg_cnt, 2), 'turn_cnt': turn_cnt, 'turn_rate': turn_rate,
            'dept': dept_df, 'type': type_df, 'reason': reason_df, 'level': level_df}

f = st.file_uploader("上传Excel", type=['xlsx', 'xls'])
if f:
    st.info("已上传: " + f.name)
    xlsx = pd.ExcelFile(f)
    period = {}; turn = None
    for sh in xlsx.sheet_names:
        try:
            df = pd.read_excel(xlsx, sheet_name=sh)
            mn = extract_month(sh)
            if sh == "离职数据":
                turn = df.copy()
                if '最后工作日' in df.columns:
                    turn['最后工作日'] = pd.to_datetime(df['最后工作日'], errors='coerce')
                    turn['离职月份'] = turn['最后工作日'].dt.strftime('%Y年%m月')
                st.success("离职数据: " + str(len(df)))
            elif mn is not None:
                if mn not in period:
                    period[mn] = {"start": None, "end": None}
                if "期末" in sh:
                    period[mn]["end"] = df
                else:
                    period[mn]["start"] = df
                st.success(str(mn) + "月数据: " + str(len(df)))
        except Exception as e:
            st.error("错误: " + str(e))
    
    months = sorted([m for m in period.keys() if period[m]["start"] is not None], reverse=True)
    if months:
        sel = st.selectbox("月份", months)
        if st.button("分析"):
            r = analyze(sel, period, turn)
            if r:
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("月份", r['month'])
                c2.metric("离职人数", r['turn_cnt'])
                c3.metric("平均人数", r['avg'])
                c4.metric("离职率", str(r['turn_rate']) + '%')
                
                st.subheader("各部门离职分析")
                if len(r['dept']) > 0:
                    st.dataframe(r['dept'], use_container_width=True, hide_index=True)
                    st.bar_chart(r['dept'].set_index('一级组织')['离职率(%)'])
                
                st.subheader("离职类型")
                if len(r['type']) > 0:
                    st.dataframe(r['type'], use_container_width=True, hide_index=True)
                
                st.subheader("离职原因")
                if len(r['reason']) > 0:
                    st.dataframe(r['reason'], use_container_width=True, hide_index=True)
                
                st.subheader("职级分布")
                if len(r['level']) > 0:
                    st.dataframe(r['level'], use_container_width=True, hide_index=True)
                
                if st.button("导出Excel"):
                    out = BytesIO()
                    with pd.ExcelWriter(out, engine='openpyxl') as w:
                        pd.DataFrame([{'月份': r['month'], '期初': r['start'], '期末': r['end'], '平均人数': r['avg'], '离职人数': r['turn_cnt'], '离职率': str(r['turn_rate'])+'%'}]).to_excel(w, sheet_name='概览', index=False)
                        if len(r['dept']) > 0: r['dept'].to_excel(w, sheet_name='部门离职', index=False)
                        if len(r['type']) > 0: r['type'].to_excel(w, sheet_name='离职类型', index=False)
                        if len(r['reason']) > 0: r['reason'].to_excel(w, sheet_name='离职原因', index=False)
                        if len(r['level']) > 0: r['level'].to_excel(w, sheet_name='职级', index=False)
                    st.download_button("下载Excel", out.getvalue(), r["month"]+"离职分析.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
