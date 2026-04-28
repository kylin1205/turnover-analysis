"""
员工离职数据分析系统 - Flask版本
离职率 = 离职人数 ÷ 平均人数
支持月份单选或多月汇总分析
"""
from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
import pandas as pd
import os
import re
from io import BytesIO
import warnings
warnings.filterwarnings("ignore")

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = "/tmp"
processor = None

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
                if sheet == "离职数据":
                    self.turnover_data = df.copy()
                    if "最后工作日" in df.columns:
                        self.turnover_data["最后工作日"] = pd.to_datetime(df["最后工作日"], errors="coerce")
                        self.turnover_data["离职月份"] = self.turnover_data["最后工作日"].dt.strftime("%Y年%m月")
                    continue
                m = re.search(r"^(\d{1,2})月", sheet)
                if not m: continue
                outer_month = int(m.group(1))
                inner = re.search(r"（(\d{1,2})月[^）]*期末", sheet)
                if inner:
                    pm = int(inner.group(1))
                    if pm not in self.period_data: self.period_data[pm] = {"start": None, "end": None}
                    self.period_data[pm]["end"] = df
                elif "期末" in sheet:
                    if outer_month not in self.period_data: self.period_data[outer_month] = {"start": None, "end": None}
                    self.period_data[outer_month]["end"] = df
                if "期初" in sheet:
                    if outer_month not in self.period_data: self.period_data[outer_month] = {"start": None, "end": None}
                    self.period_data[outer_month]["start"] = df
            except: pass
    
    def get_available_months(self):
        return sorted([m for m in self.period_data.keys() if self.period_data[m]["start"] is not None])
    
    def get_summary(self):
        months = self.get_available_months()
        total = len(self.turnover_data) if self.turnover_data is not None else 0
        latest = months[-1] if months else None
        cnt = len(self.period_data.get(latest, {}).get("start", pd.DataFrame())) if latest else 0
        options = []
        for m in months: options.append({"value": get_month_str(m), "label": get_month_str(m), "months": [m]})
        for i in range(len(months)):
            for j in range(i + 1, len(months)):
                if months[j] - months[i] == j - i:
                    rms = list(range(months[i], months[j] + 1))
                    label = get_month_str(months[i]) + "至" + get_month_str(months[j])
                    options.append({"value": label, "label": label, "months": rms})
        return {"available_months": [get_month_str(m) for m in months], "month_options": options, "total_employees": cnt, "total_turnover": total}
    
    def get_monthly_analysis(self, month_nums):
        if not isinstance(month_nums, list): month_nums = [month_nums]
        ts, te, tt = 0, 0, 0
        all_turn = pd.DataFrame()
        ds, de, dt = {}, {}, {}
        for mn in month_nums:
            sdf = self.period_data.get(mn, {}).get("start")
            edf = self.period_data.get(mn, {}).get("end")
            ms = get_month_str(mn)
            if sdf is not None:
                ts += len(sdf)
                if "一级组织" in sdf.columns:
                    for d, c in sdf.groupby("一级组织").size().items(): ds[d] = ds.get(d, 0) + c
            if edf is not None and len(edf) > 0:
                te += len(edf)
                if "一级组织" in edf.columns:
                    for d, c in edf.groupby("一级组织").size().items(): de[d] = de.get(d, 0) + c
            if self.turnover_data is not None and "离职月份" in self.turnover_data.columns:
                mt = self.turnover_data[self.turnover_data["离职月份"] == ms].copy()
                tt += len(mt)
                all_turn = pd.concat([all_turn, mt], ignore_index=True)
                if "一级组织" in mt.columns:
                    for d, c in mt.groupby("一级组织").size().items(): dt[d] = dt.get(d, 0) + c
        avg = (ts + te) / 2 if (ts + te) > 0 else 0
        rate = round((tt / avg * 100), 2) if avg > 0 else 0
        lbl = get_month_str(month_nums[0]) if len(month_nums) == 1 else get_month_str(month_nums[0]) + "至" + get_month_str(month_nums[-1])
        da = []
        for dept in set(list(ds.keys()) + list(de.keys()) + list(dt.keys())):
            d, e, t = ds.get(dept, 0), de.get(dept, 0), dt.get(dept, 0)
            a = (d + e) / 2
            r = round((t / a * 100), 2) if a > 0 else 0
            da.append({"一级组织": dept, "期初人数": d, "期末人数": e, "平均人数": round(a, 1), "离职人数": t, "离职率": r})
        da = sorted(da, key=lambda x: x["离职率"], reverse=True)
        td = []
        if len(all_turn) > 0 and "离职类型" in all_turn.columns:
            t = all_turn.groupby("离职类型").size().reset_index(name="人数")
            t["占比"] = round(t["人数"] / tt * 100, 2) if tt > 0 else 0
            td = t.sort_values("人数", ascending=False).to_dict("records")
        rd = []
        if len(all_turn) > 0 and "离职原因" in all_turn.columns:
            r = all_turn.groupby("离职原因").size().reset_index(name="人数")
            r["占比"] = round(r["人数"] / tt * 100, 2) if tt > 0 else 0
            rd = r.sort_values("人数", ascending=False).to_dict("records")
        ted = []
        if len(all_turn) > 0 and "累计司龄（年）" in all_turn.columns:
            def cat(y):
                if pd.isna(y): return "未知"
                elif y <= 0.5: return "0.5年及以下"
                elif y <= 1: return "0.5-1年"
                elif y <= 3: return "1-3年"
                elif y <= 5: return "3-5年"
                else: return "5年以上"
            tmp = all_turn.copy()
            tmp["司龄段"] = tmp["累计司龄（年）"].apply(cat)
            t = tmp.groupby("司龄段").size().reset_index(name="人数")
            t["占比"] = round(t["人数"] / tt * 100, 2) if tt > 0 else 0
            ted = t.sort_values("人数", ascending=False).to_dict("records")
        ld = []
        if len(all_turn) > 0 and "职级" in all_turn.columns:
            l = all_turn.groupby("职级").size().reset_index(name="人数")
            l["占比"] = round(l["人数"] / tt * 100, 2) if tt > 0 else 0
            ld = l.sort_values("人数", ascending=False).to_dict("records")
        return {"month": lbl, "months": month_nums, "avg_count": round(avg, 1), "start_count": ts, "end_count": te, "turnover_count": tt, "turnover_rate": rate, "dept_turnover": da, "turnover_type": td, "turnover_reason": rd, "turnover_tenure": ted, "turnover_level": ld}



@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/upload", methods=["POST"])
def upload():
    global processor
    if "file" not in request.files:
        return jsonify({"success": False, "message": "没有文件"})
    f = request.files["file"]
    filepath = os.path.join(app.config["UPLOAD_FOLDER"], secure_filename(f.filename))
    f.save(filepath)
    try:
        processor = DataProcessor()
        processor.load_file(filepath)
        months = processor.get_available_months()
        if not months:
            return jsonify({"success": False, "message": "未找到有效月份数据"})
        return jsonify({"success": True, "message": "上传成功", "summary": processor.get_summary()})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)})

@app.route("/api/analyze", methods=["POST"])
def analyze():
    global processor
    data = request.json
    month_value = data.get("month", "")
    if not month_value or not processor:
        return jsonify({"success": False, "message": "参数错误"})
    rm = re.search(r"(\d+)年(\d+)月至(\d+)年(\d+)月", month_value)
    if rm:
        sm, em = int(rm.group(2)), int(rm.group(4))
        mns = list(range(sm, em + 1))
    else:
        m = re.search(r"(\d+)年(\d+)月", month_value)
        mns = [int(m.group(2))] if m else []
    result = processor.get_monthly_analysis(mns)
    if not result:
        return jsonify({"success": False, "message": "未找到数据"})
    return jsonify({"success": True, "data": result})

@app.route("/api/export", methods=["POST"])
def export():
    global processor
    data = request.json
    month_value = data.get("month", "")
    if not month_value or not processor:
        return jsonify({"success": False, "message": "参数错误"})
    rm = re.search(r"(\d+)年(\d+)月至(\d+)年(\d+)月", month_value)
    if rm:
        sm, em = int(rm.group(2)), int(rm.group(4))
        mns = list(range(sm, em + 1))
    else:
        m = re.search(r"(\d+)年(\d+)月", month_value)
        mns = [int(m.group(2))] if m else []
    result = processor.get_monthly_analysis(mns)
    if not result:
        return jsonify({"success": False, "message": "未找到数据"})
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as w:
        pd.DataFrame([{"月份": result["month"], "期初人数": result["start_count"], "期末人数": result["end_count"], "平均人数": result["avg_count"], "离职人数": result["turnover_count"], "离职率": str(result["turnover_rate"]) + "%"}]).to_excel(w, sheet_name="概览", index=False)
        if result["dept_turnover"]: pd.DataFrame(result["dept_turnover"]).to_excel(w, sheet_name="部门离职", index=False)
        if result["turnover_type"]: pd.DataFrame(result["turnover_type"]).to_excel(w, sheet_name="离职类型", index=False)
        if result["turnover_reason"]: pd.DataFrame(result["turnover_reason"]).to_excel(w, sheet_name="离职原因", index=False)
        if result["turnover_tenure"]: pd.DataFrame(result["turnover_tenure"]).to_excel(w, sheet_name="司龄分布", index=False)
        if result["turnover_level"]: pd.DataFrame(result["turnover_level"]).to_excel(w, sheet_name="职级分布", index=False)
    output.seek(0)
    return send_file(output, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", as_attachment=True, download_name=result["month"] + "离职分析.xlsx")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)

