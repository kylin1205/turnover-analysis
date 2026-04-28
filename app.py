"""
员工离职数据分析系统 - Flask版本
离职率 = 离职人数 ÷ 平均人数
平均人数 = (期初人数 + 期末人数) / 2
"""
from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
import pandas as pd
import numpy as np
import os
import re
from datetime import datetime
from io import BytesIO
import warnings
warnings.filterwarnings('ignore')

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = '/tmp'
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

processor = None

def extract_month_num(sheet_name):
    match = re.search(r'(\d{1,2})月', str(sheet_name))
    return int(match.group(1)) if match else None

def get_month_str(month_num):
    return "2026年" + str(month_num).zfill(2) + "月"

class DataProcessor:
    def __init__(self):
        self.period_data = {}
        self.turnover_data = None
        
    def load_file(self, filepath):
        xlsx = pd.ExcelFile(filepath)
        for sheet in xlsx.sheet_names:
            try:
                df = pd.read_excel(xlsx, sheet_name=sheet)
                month_num = extract_month_num(sheet)
                
                if sheet == "离职数据":
                    self.turnover_data = df.copy()
                    if '最后工作日' in df.columns:
                        self.turnover_data['最后工作日'] = pd.to_datetime(df['最后工作日'], errors='coerce')
                        self.turnover_data['离职月份'] = self.turnover_data['最后工作日'].dt.strftime('%Y年%m月')
                
                elif month_num is not None:
                    if month_num not in self.period_data:
                        self.period_data[month_num] = {'start': None, 'end': None}
                    if "期末" in sheet:
                        self.period_data[month_num]['end'] = df
                    else:
                        self.period_data[month_num]['start'] = df
            except Exception as e:
                continue
    
    def get_available_months(self):
        return sorted([m for m in self.period_data.keys() if self.period_data[m]['start'] is not None], reverse=True)
    
    def get_summary(self):
        months = self.get_available_months()
        total_turnover = len(self.turnover_data) if self.turnover_data is not None else 0
        latest_month = months[0] if months else None
        latest_count = len(self.period_data.get(latest_month, {}).get('start', pd.DataFrame())) if latest_month else 0
        return {
            'available_months': [get_month_str(m) for m in months],
            'total_employees': latest_count,
            'total_turnover': total_turnover
        }
    
    def get_monthly_analysis(self, month_num):
        start_df = self.period_data.get(month_num, {}).get('start')
        end_df = self.period_data.get(month_num, {}).get('end')
        month_str = get_month_str(month_num)
        
        if start_df is None or len(start_df) == 0:
            return None
        
        start_cnt = len(start_df)
        end_cnt = len(end_df) if end_df is not None and len(end_df) > 0 else start_cnt
        avg_cnt = (start_cnt + end_cnt) / 2
        
        month_turn = pd.DataFrame()
        if self.turnover_data is not None and '离职月份' in self.turnover_data.columns:
            month_turn = self.turnover_data[self.turnover_data['离职月份'] == month_str].copy()
        turn_cnt = len(month_turn)
        turn_rate = round((turn_cnt / avg_cnt * 100), 2) if avg_cnt > 0 else 0
        
        # 部门离职
        dept_turnover = []
        if len(month_turn) > 0 and '一级组织' in month_turn.columns:
            d1 = month_turn.groupby('一级组织').size().reset_index(name='离职人数')
            d2 = pd.DataFrame()
            d3 = pd.DataFrame()
            if '一级组织' in start_df.columns:
                d2 = start_df.groupby('一级组织').size().reset_index(name='期初人数')
            if end_df is not None and len(end_df) > 0 and '一级组织' in end_df.columns:
                d3 = end_df.groupby('一级组织').size().reset_index(name='期末人数')
            
            dept = d1.copy()
            if len(d2) > 0:
                dept = dept.merge(d2, on='一级组织', how='left')
            else:
                dept['期初人数'] = 0
            if len(d3) > 0:
                dept = dept.merge(d3, on='一级组织', how='left')
            else:
                dept['期末人数'] = dept['期初人数']
            dept['期初人数'] = dept['期初人数'].fillna(0).astype(int)
            dept['期末人数'] = dept['期末人数'].fillna(0).astype(int)
            dept['平均人数'] = (dept['期初人数'] + dept['期末人数']) / 2
            dept['离职率(%)'] = dept.apply(lambda x: round((x['离职人数']/x['平均人数']*100) if x['平均人数']>0 else 0, 2), axis=1)
            dept = dept.sort_values('离职率(%)', ascending=False)
            dept_turnover = dept.to_dict('records')
        
        # 离职类型
        type_data = []
        if len(month_turn) > 0 and '离职类型' in month_turn.columns:
            t = month_turn.groupby('离职类型').size().reset_index(name='人数')
            t['占比'] = round(t['人数']/turn_cnt*100, 2) if turn_cnt > 0 else 0
            t = t.sort_values('人数', ascending=False)
            type_data = t.to_dict('records')
        
        # 离职原因
        reason_data = []
        if len(month_turn) > 0 and '离职原因' in month_turn.columns:
            r = month_turn.groupby('离职原因').size().reset_index(name='人数')
            r['占比'] = round(r['人数']/turn_cnt*100, 2) if turn_cnt > 0 else 0
            r = r.sort_values('人数', ascending=False)
            reason_data = r.to_dict('records')
        
        # 司龄
        tenure_data = []
        tenure_col = '累计司龄（年）'
        if len(month_turn) > 0 and tenure_col in month_turn.columns:
            def cat(y):
                if pd.isna(y): return '未知'
                elif y <= 0.5: return '0.5年及以下'
                elif y <= 1: return '0.5-1年'
                elif y <= 3: return '1-3年'
                elif y <= 5: return '3-5年'
                else: return '5年以上'
            month_turn['司龄段'] = month_turn[tenure_col].apply(cat)
            t = month_turn.groupby('司龄段').size().reset_index(name='人数')
            t['占比'] = round(t['人数']/turn_cnt*100, 2) if turn_cnt > 0 else 0
            t = t.sort_values('人数', ascending=False)
            tenure_data = t.to_dict('records')
        
        # 职级
        level_data = []
        if len(month_turn) > 0 and '职级' in month_turn.columns:
            l = month_turn.groupby('职级').size().reset_index(name='人数')
            l['占比'] = round(l['人数']/turn_cnt*100, 2) if turn_cnt > 0 else 0
            l = l.sort_values('人数', ascending=False)
            level_data = l.to_dict('records')
        
        return {
            'month': month_str,
            'avg_count': round(avg_cnt, 2),
            'start_count': start_cnt,
            'end_count': end_cnt,
            'turnover_count': turn_cnt,
            'turnover_rate': turn_rate,
            'dept_turnover': dept_turnover,
            'turnover_type': type_data,
            'turnover_reason': reason_data,
            'turnover_tenure': tenure_data,
            'turnover_level': level_data
        }

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/upload', methods=['POST'])
def upload():
    global processor
    if 'file' not in request.files:
        return jsonify({'success': False, 'message': '没有文件'})
    f = request.files['file']
    if f.filename == '':
        return jsonify({'success': False, 'message': '文件名为空'})
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(f.filename))
    f.save(filepath)
    try:
        processor = DataProcessor()
        processor.load_file(filepath)
        months = processor.get_available_months()
        if not months:
            return jsonify({'success': False, 'message': '未找到有效月份数据'})
        return jsonify({'success': True, 'message': '上传成功', 'summary': processor.get_summary()})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)})

@app.route('/api/analyze', methods=['POST'])
def analyze():
    global processor
    data = request.json
    month_str = data.get('month', '')
    m = re.search(r'(\d+)年(\d+)月', month_str)
    if not m or not processor:
        return jsonify({'success': False, 'message': '参数错误'})
    month_num = int(m.group(2))
    result = processor.get_monthly_analysis(month_num)
    if not result:
        return jsonify({'success': False, 'message': '未找到数据'})
    return jsonify({'success': True, 'data': result})

@app.route('/api/export', methods=['POST'])
def export():
    global processor
    data = request.json
    month_str = data.get('month', '')
    m = re.search(r'(\d+)年(\d+)月', month_str)
    if not m or not processor:
        return jsonify({'success': False, 'message': '参数错误'})
    month_num = int(m.group(2))
    result = processor.get_monthly_analysis(month_num)
    if not result:
        return jsonify({'success': False, 'message': '未找到数据'})
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as w:
        pd.DataFrame([{
            '月份': result['month'],
            '期初人数': result['start_count'],
            '期末人数': result['end_count'],
            '平均人数': result['avg_count'],
            '离职人数': result['turnover_count'],
            '离职率(%)': result['turnover_rate']
        }]).to_excel(w, sheet_name='概览', index=False)
        if result['dept_turnover']:
            pd.DataFrame(result['dept_turnover']).to_excel(w, sheet_name='部门离职', index=False)
        if result['turnover_type']:
            pd.DataFrame(result['turnover_type']).to_excel(w, sheet_name='离职类型', index=False)
        if result['turnover_reason']:
            pd.DataFrame(result['turnover_reason']).to_excel(w, sheet_name='离职原因', index=False)
        if result['turnover_tenure']:
            pd.DataFrame(result['turnover_tenure']).to_excel(w, sheet_name='司龄分布', index=False)
        if result['turnover_level']:
            pd.DataFrame(result['turnover_level']).to_excel(w, sheet_name='职级分布', index=False)
    output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name=result['month']+'离职分析.xlsx')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)