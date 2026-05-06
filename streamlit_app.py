"""
员工离职数据分析系统 - Streamlit Cloud版本
"""
import streamlit as st
import pandas as pd
import re
from io import BytesIO
import warnings
warnings.filterwarnings("ignore")

st.set_page_config(page_title="员工离职数据分析", layout="wide")

def get_month_str(n):
    return "2026年" + str(n).zfill(2) + "月"

def cat_tenure(y):
    if pd.isna(y): return "未知"
    elif y <= 0.5: return "0.5年及以下"
    elif y <= 1: return "0.5-1年"
    elif y <= 3: return "1-3年"
    elif y <= 5: return "3-5年"
    else: return "5年以上"

class Proc:
    def __init__(self, xlsx):
        self.period = {}
        self.turn = None
        for sh in xlsx.sheet_names:
            try:
                df = pd.read_excel(xlsx, sheet_name=sh)
                # Check if this is a turnover sheet (contains 离职 but not 期初/期末)
                if "离职" in sh and "期初" not in sh and "期末" not in sh:
                    self.turn = df.copy()
                    if "最后工作日" in df.columns:
                        self.turn["最后工作日"] = pd.to_datetime(self.turn["最后工作日"], errors="coerce")
                        self.turn["离职月份"] = self.turn["最后工作日"].dt.strftime("%Y年%m月")
                    continue
                m = re.search(r"^(\d{1,2})月", sh)
                if not m: continue
                om = int(m.group(1))
                inner = re.search(r"（(\d{1,2})月[^）]*期末", sh)
                if inner:
                    pm = int(inner.group(1))
                    if pm not in self.period: self.period[pm] = {"s": None, "e": None}
                    self.period[pm]["e"] = df
                elif "期末" in sh:
                    if om not in self.period: self.period[om] = {"s": None, "e": None}
                    self.period[om]["e"] = df
                if "期初" in sh:
                    if om not in self.period: self.period[om] = {"s": None, "e": None}
                    self.period[om]["s"] = df
            except: pass
    
    def months(self):
        return sorted([m for m in self.period.keys() if self.period[m]["s"] is not None])
    
    def analyze(self, mns):
        if not isinstance(mns, list): mns = [mns]
        ts, te, tt = 0, 0, 0
        all_t = pd.DataFrame()
        ds, de, dt = {}, {}, {}
        for mn in mns:
            s, e = self.period.get(mn, {}).get("s"), self.period.get(mn, {}).get("e")
            ms = get_month_str(mn)
            if s is not None:
                ts += len(s)
                if "一级组织" in s.columns:
                    for d, c in s.groupby("一级组织").size().items(): ds[d] = ds.get(d, 0) + c
            if e is not None and len(e) > 0:
                te += len(e)
                if "一级组织" in e.columns:
                    for d, c in e.groupby("一级组织").size().items(): de[d] = de.get(d, 0) + c
            if self.turn is not None and "离职月份" in self.turn.columns:
                mt = self.turn[self.turn["离职月份"] == ms].copy()
                tt += len(mt)
                all_t = pd.concat([all_t, mt], ignore_index=True)
                if "一级组织" in mt.columns:
                    for d, c in mt.groupby("一级组织").size().items(): dt[d] = dt.get(d, 0) + c
        avg = (ts + te) / 2 if (ts + te) > 0 else 0
        rate = round((tt / avg * 100), 2) if avg > 0 else 0
        lbl = get_month_str(mns[0]) if len(mns) == 1 else get_month_str(mns[0]) + "至" + get_month_str(mns[-1])
        da = []
        for dept in set(list(ds.keys()) + list(de.keys()) + list(dt.keys())):
            d, e, t = ds.get(dept, 0), de.get(dept, 0), dt.get(dept, 0)
            a = (d + e) / 2
            r = round((t / a * 100), 2) if a > 0 else 0
            da.append({"一级组织": dept, "期初人数": d, "期末人数": e, "平均人数": round(a, 1), "离职人数": t, "离职率": r})
        da = sorted(da, key=lambda x: x["离职率"], reverse=True)
        td = []
        if len(all_t) > 0 and "离职类型" in all_t.columns:
            t = all_t.groupby("离职类型").size().reset_index(name="人数")
            t["占比"] = round(t["人数"] / tt * 100, 2) if tt > 0 else 0
            td = t.sort_values("人数", ascending=False).to_dict("records")
        rd = []
        if len(all_t) > 0 and "离职原因" in all_t.columns:
            r = all_t.groupby("离职原因").size().reset_index(name="人数")
            r["占比"] = round(r["人数"] / tt * 100, 2) if tt > 0 else 0
            rd = r.sort_values("人数", ascending=False).to_dict("records")
        ted = []
        if len(all_t) > 0 and "累计司龄（年）" in all_t.columns:
            tmp = all_t.copy()
            tmp["司龄段"] = tmp["累计司龄（年）"].apply(cat_tenure)
            t = tmp.groupby("司龄段").size().reset_index(name="人数")
            t["占比"] = round(t["人数"] / tt * 100, 2) if tt > 0 else 0
            ted = t.sort_values("人数", ascending=False).to_dict("records")
        ld = []
        if len(all_t) > 0 and "职级" in all_t.columns:
            l = all_t.groupby("职级").size().reset_index(name="人数")
            l["占比"] = round(l["人数"] / tt * 100, 2) if tt > 0 else 0
            ld = l.sort_values("人数", ascending=False).to_dict("records")
        return {"month": lbl, "avg": round(avg, 1), "start": ts, "end": te, "turn": tt, "rate": rate, "dept": da, "type": td, "reason": rd, "tenure": ted, "level": ld}

st.markdown('<div style="background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%); padding: 20px; border-radius: 10px; margin-bottom: 20px;"><h2 style="color: white; margin: 0;">员工离职数据分析系统</h2></div>', unsafe_allow_html=True)

st.header("数据导入")
f = st.file_uploader("上传Excel文件", type=["xlsx", "xls"])

if f:
    st.success("已上传: " + f.name)
    xlsx = pd.ExcelFile(f)
    p = Proc(xlsx)
    ms = p.months()
    
    options = []
    for m in ms: options.append((get_month_str(m), [m]))
    for i in range(len(ms)):
        for j in range(i + 1, len(ms)):
            if ms[j] - ms[i] == j - i:
                options.append((get_month_str(ms[i]) + "至" + get_month_str(ms[j]), list(range(ms[i], ms[j] + 1))))
    
    sel = st.selectbox("选择月份", [o[0] for o in options])
    sel_mns = dict(options).get(sel, [1])
    
    if st.button("分析", type="primary"):
        r = p.analyze(sel_mns)
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("月份", r["month"])
        c2.metric("平均人数", r["avg"])
        c3.metric("离职人数", r["turn"])
        c4.metric("离职率", str(r["rate"]) + "%")
        
        if r["dept"]:
            st.subheader("各部门离职分析")
            st.dataframe(pd.DataFrame(r["dept"]), use_container_width=True, hide_index=True)
        
        if r["type"]:
            st.subheader("离职类型分布")
            st.dataframe(pd.DataFrame(r["type"]), use_container_width=True, hide_index=True)
        
        if r["reason"]:
            st.subheader("离职原因分布")
            st.dataframe(pd.DataFrame(r["reason"]), use_container_width=True, hide_index=True)
        
        if r["tenure"]:
            st.subheader("离职司龄分布")
            st.dataframe(pd.DataFrame(r["tenure"]), use_container_width=True, hide_index=True)
        
        if r["level"]:
            st.subheader("职级分布")
            st.dataframe(pd.DataFrame(r["level"]), use_container_width=True, hide_index=True)
        
        # Export
        st.subheader("导出数据")
        out = BytesIO()
        with pd.ExcelWriter(out, engine="openpyxl") as w:
            pd.DataFrame([{"月份": r["month"], "期初人数": r["start"], "期末人数": r["end"], "平均人数": r["avg"], "离职人数": r["turn"], "离职率": str(r["rate"]) + "%"}]).to_excel(w, sheet_name="概览", index=False)
            if r["dept"]: pd.DataFrame(r["dept"]).to_excel(w, sheet_name="部门离职", index=False)
            if r["type"]: pd.DataFrame(r["type"]).to_excel(w, sheet_name="离职类型", index=False)
            if r["reason"]: pd.DataFrame(r["reason"]).to_excel(w, sheet_name="离职原因", index=False)
            if r["tenure"]: pd.DataFrame(r["tenure"]).to_excel(w, sheet_name="司龄分布", index=False)
            if r["level"]: pd.DataFrame(r["level"]).to_excel(w, sheet_name="职级分布", index=False)
        
        st.download_button(label="下载Excel报告", data=out.getvalue(), file_name=r["month"] + "离职分析.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("请上传Excel文件开始分析")
