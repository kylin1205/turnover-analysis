"""
员工离职数据分析系统 - Streamlit Cloud版本 V3.1
"""
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import re
from io import BytesIO
import json

import warnings
warnings.filterwarnings("ignore")

def get_download_config(name):
    return {
        "toImageButtonOptions": {
            "format": "png",
            "filename": name,
            "height": 600,
            "width": 1200,
            "scale": 2
        },
        "displaylogo": False,
        "modeBarButtonsToRemove": ["select2d", "lasso2d"]
    }

st.set_page_config(page_title="员工离职数据分析", layout="wide")

COLORS = {"primary": "#4F46E5", "danger": "#EF4444", "warning": "#F59E0B", "purple": "#8B5CF6", "success": "#10B981"}

def get_month_str(n):
    return "2026年" + str(n).zfill(2) + "月"

def cat_tenure(y):
    if pd.isna(y): return "未知"
    elif y <= 0.5: return "0.5年及以下"
    elif y <= 1: return "0.5-1年"
    elif y <= 3: return "1-3年"
    elif y <= 5: return "3-5年"
    else: return "5年以上"

def cat_level(level):
    if pd.isna(level): return "未知"
    level = str(level).strip()
    if level.startswith("中"): return "中级"
    elif level.startswith("高"): return "高级"
    elif level.startswith("初"): return "初级"
    elif level.startswith("资深"): return "资深"
    elif level.startswith("总监"): return "总监"
    else: return level

def generate_ai_analysis(r):
    """生成AI智能分析报告"""
    lines = []
    
    # 1. 整体概况
    lines.append("## 📊 整体概况")
    lines.append("")
    主动离职 = 0
    被动离职 = 0
    if r["type"]:
        for t in r["type"]:
            if "主动" in str(t.get("离职类型", "")):
                主动离职 = t.get("人数", 0)
            elif "被动" in str(t.get("离职类型", "")):
                被动离职 = t.get("人数", 0)
    
    rate_str = f"**{r['rate']}%**"
    lines.append(f"- 分析周期：{r['month']}")
    lines.append(f"- 离职率：{rate_str}（期间平均人数 {r['avg']} 人）")
    lines.append(f"- 离职总人数：{r['turn']} 人（主动离职 {主动离职} 人，被动离职 {被动离职} 人）")
    lines.append("")
    
    # 2. 主动离职分析
    if 主动离职 > 0:
        lines.append("## 💡 主动离职分析")
        lines.append("")
        # 找出主动离职的主要原因
        main_reasons = []
        if r["reason"]:
            for reason in r["reason"][:5]:
                main_reasons.append(f"{reason['离职原因']}（{reason['人数']}人，占比{reason['占比']}%）")
        
        lines.append(f"主动离职主要原因：")
        for reason in main_reasons:
            lines.append(f"- {reason}")
        lines.append("")
        lines.append("**💼 建议措施：**")
        lines.append("- 招聘环节强化对工作负荷及岗位挑战的识别")
        lines.append("- 对在岗员工定期复盘职业路径匹配度")
        lines.append("- 提前干预因预期落差或生活变动导致的主动性流失")
        lines.append("")
    
    # 3. 被动离职分析
    if 被动离职 > 0:
        lines.append("## ⚠️ 被动离职分析")
        lines.append("")
        lines.append(f"- 被动离职 {被动离职} 人，占比 {round(被动离职/r['turn']*100, 1)}%")
        lines.append("")
        lines.append("**🔧 建议措施：**")
        lines.append("- 加强候选人能力匹配度评估")
        lines.append("- 强化试用期过程预警与低绩效员工改进辅导")
        lines.append("- 因架构调整终止合约时落实妥善交接与职业缓冲支持")
        lines.append("")
    
    # 4. 部门风险分析
    if r["dept"]:
        lines.append("## 🏢 部门离职风险分析")
        lines.append("")
        high_risk = [d for d in r["dept"] if d["离职率"] >= 5]
        medium_risk = [d for d in r["dept"] if 2 <= d["离职率"] < 5]
        
        if high_risk:
            lines.append("**🔴 高风险部门（离职率≥5%）：**")
            for d in high_risk:
                lines.append(f"- **{d['一级组织']}**：离职 {d['离职人数']} 人，离职率 {d['离职率']}%，需重点关注")
            lines.append("")
        
        if medium_risk:
            lines.append("**🟡 中风险部门（离职率2%-5%）：**")
            for d in medium_risk:
                lines.append(f"- {d['一级组织']}：离职 {d['离职人数']} 人，离职率 {d['离职率']}%")
            lines.append("")
    
    # 5. 司龄分析
    if r["tenure"]:
        lines.append("## 📅 司龄分布分析")
        lines.append("")
        new_employee = 0
        for t in r["tenure"]:
            if "0.5" in t["司龄段"]:
                new_employee = t["人数"]
                break
        
        if new_employee > 0:
            lines.append(f"- 新人（司龄≤0.5年）离职 {new_employee} 人，需关注新员工适应性")
        
        # 找出离职最多的司龄段
        if r["tenure"]:
            top_tenure = max(r["tenure"], key=lambda x: x["人数"])
            lines.append(f"- 离职最多司龄段：{top_tenure['司龄段']}（{top_tenure['人数']}人，占比{top_tenure['占比']}%）")
        lines.append("")
        lines.append("**🎯 建议措施：**")
        lines.append("- 新员工入职引导与岗位匹配度评估")
        lines.append("- 关注1-3年员工的职业发展通道")
        lines.append("- 建立人才梯队储备机制")
        lines.append("")
    
    # 6. 职级分析
    if r["level"]:
        lines.append("## 👥 职级分布分析")
        lines.append("")
        for level in r["level"][:5]:
            lines.append(f"- **{level['职级合并']}**：{level['人数']}人（占比{level['占比']}%）")
        lines.append("")
        
        # 中坚力量分析
        middle_level = [l for l in r["level"] if l["职级合并"] in ["中级", "高级"]]
        if middle_level:
            total_middle = sum(l["人数"] for l in middle_level)
            lines.append(f"**⚡ 中坚力量（中级+高级）流失：**{total_middle}人，占比{round(total_middle/r['turn']*100, 1)}%")
            if total_middle / r["turn"] > 0.5:
                lines.append("> ⚠️ 中坚力量流失严重，需关注人才梯队断层风险")
        lines.append("")
    
    # 7. 风险预警
    lines.append("## 🚨 风险预警")
    lines.append("")
    risks = []
    
    if r["rate"] > 5:
        risks.append(f"- 离职率{r['rate']}%偏高，超过安全阈值5%")
    
    if 被动离职 / r["turn"] > 0.3 if r["turn"] > 0 else False:
        risks.append(f"- 被动离职占比{round(被动离职/r['turn']*100, 1)}%较高，可能存在招聘评估或绩效管理问题")
    
    if r["dept"] and any(d["离职率"] > 8 for d in r["dept"]):
        risks.append("- 存在离职率过高的部门，建议深入调研原因")
    
    if r["tenure"]:
        new_rate = sum(t["人数"] for t in r["tenure"] if "0.5" in t["司龄段"]) / r["turn"] if r["turn"] > 0 else 0
        if new_rate > 0.3:
            risks.append(f"- 新人离职率{round(new_rate*100, 1)}%偏高，需改善新员工体验")
    
    if risks:
        for risk in risks:
            lines.append(risk)
    else:
        lines.append("- 当前离职率在可控范围内")
    lines.append("")
    
    # 8. 改善建议
    lines.append("## ✅ 综合改善建议")
    lines.append("")
    lines.append("| 维度 | 建议措施 | 优先级 |")
    lines.append("|------|----------|--------|")
    lines.append("| 招聘 | 强化岗位匹配度评估，增加工作内容真实性预览 | 高 |")
    lines.append("| 入职 | 完善新员工引导机制，关注前6个月留存 | 高 |")
    lines.append("| 发展 | 建立职业发展双通道，定期职业路径回顾 | 中 |")
    lines.append("| 绩效 | 强化绩效预警机制，及早干预低绩效员工 | 中 |")
    lines.append("| 沟通 | 定期员工满意度调研，建立反馈渠道 | 中 |")
    lines.append("| 梯队 | 建立关键岗位继任计划，防止人才断层 | 高 |")
    lines.append("")
    
    return "\n".join(lines)

class Proc:
    def __init__(self, xlsx):
        self.period = {}
        self.turn = None
        for sh in xlsx.sheet_names:
            try:
                df = pd.read_excel(xlsx, sheet_name=sh)
                if "离职" in sh and "期初" not in sh and "期末" not in sh:
                    self.turn = df.copy()
                    m = re.search(r"^(\d{1,2})月", sh)
                    if m:
                        self.turn["离职月份"] = get_month_str(int(m.group(1)))
                    elif "最后工作日" in df.columns:
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
            if t > 0:
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
            rd = r.sort_values("人数", ascending=False).head(10).to_dict("records")
        ted = []
        if len(all_t) > 0 and "累计司龄（年）" in all_t.columns:
            tmp = all_t.copy()
            tmp["司龄段"] = tmp["累计司龄（年）"].apply(cat_tenure)
            t = tmp.groupby("司龄段").size().reset_index(name="人数")
            t["占比"] = round(t["人数"] / tt * 100, 2) if tt > 0 else 0
            ted = t.sort_values("人数", ascending=False).to_dict("records")
        ld = []
        if len(all_t) > 0 and "职级" in all_t.columns:
            tmp = all_t.copy()
            tmp["职级合并"] = tmp["职级"].apply(cat_level)
            l = tmp.groupby("职级合并").size().reset_index(name="人数")
            l["占比"] = round(l["人数"] / tt * 100, 2) if tt > 0 else 0
            ld = l.sort_values("人数", ascending=False).to_dict("records")
        return {"month": lbl, "avg": round(avg, 1), "start": ts, "end": te, "turn": tt, "rate": rate, "dept": da, "type": td, "reason": rd, "tenure": ted, "level": ld}

st.markdown('<div style="background: linear-gradient(135deg, #4F46E5 0%, #7C3AED 100%); padding: 20px; border-radius: 10px; margin-bottom: 20px;"><h2 style="color: white; margin: 0;">员工离职数据分析系统 V3.1</h2></div>', unsafe_allow_html=True)

st.header("数据导入")
f = st.file_uploader("上传Excel文件", type=["xlsx", "xls"])

if f:
    st.success("已上传: " + f.name)
    xlsx = pd.ExcelFile(f)
    p = Proc(xlsx)
    ms = p.months()
    
    if not ms:
        st.error("未找到期初数据，请检查文件格式！")
    else:
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
            
            st.subheader("关键指标")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("分析月份", r["month"])
            c2.metric("平均人数", str(r["avg"]) + "人")
            c3.metric("离职人数", str(r["turn"]) + "人")
            c4.metric("离职率", str(r["rate"]) + "%")
            
            st.markdown("---")
            
            if r["dept"]:
                st.subheader("各部门离职分析")
                df_dept = pd.DataFrame(r["dept"])
                
                fig = go.Figure()
                fig.add_trace(go.Bar(x=df_dept["一级组织"], y=df_dept["离职人数"], name="离职人数", marker_color=COLORS["primary"], text=df_dept["离职人数"], textposition="outside"))
                fig.add_trace(go.Scatter(x=df_dept["一级组织"], y=df_dept["离职率"], name="离职率(%)", marker_color=COLORS["danger"], text=[f"{x:.2f}%" for x in df_dept["离职率"]], textposition="top center", yaxis="y2", mode="lines+markers+text", line=dict(width=3), marker=dict(size=10)))
                fig.update_layout(title=dict(text="各部门离职人数与离职率", font=dict(size=20)), xaxis=dict(title="部门", tickangle=-45, tickfont=dict(size=11), gridcolor="lightgray"), yaxis=dict(title="离职人数", title_font=dict(color=COLORS["primary"], size=14), tickfont=dict(size=12), gridcolor="lightgray"), yaxis2=dict(title="离职率(%)", title_font=dict(color=COLORS["danger"], size=14), overlaying="y", side="right", tickfont=dict(size=12)), legend=dict(x=0.5, y=1.12, xanchor="center", orientation="h", font=dict(size=13)), height=550, margin=dict(b=120), plot_bgcolor="white", paper_bgcolor="white")
                st.plotly_chart(fig, use_container_width=True, config=get_download_config(r["month"] + "_部门离职分析"))
                st.dataframe(df_dept, use_container_width=True, hide_index=True)
            
            st.markdown("---")
            
            if r["type"]:
                st.subheader("离职类型分布")
                df_type = pd.DataFrame(r["type"]).copy()
                for idx, row in df_type.iterrows():
                    if "主动" in str(row.get("离职类型", "")):
                        df_type.loc[idx, "离职类型"] = "主动离职"
                    elif "被动" in str(row.get("离职类型", "")):
                        df_type.loc[idx, "离职类型"] = "被动离职"
                merged = df_type.groupby("离职类型")["人数"].sum().reset_index()
                merged["占比"] = round(merged["人数"] / r["turn"] * 100, 2)
                
                fig_pie = px.pie(merged, values="人数", names="离职类型", title="离职类型占比", color_discrete_sequence=[COLORS["success"], COLORS["danger"]], hole=0.5)
                fig_pie.update_traces(textposition="outside", textinfo="label+percent+value", textfont=dict(size=14), marker=dict(line=dict(color="white", width=3)), hovertemplate="<b>%{label}</b><br>人数: %{value}人<br>占比: %{percent}<extra></extra>")
                fig_pie.update_layout(title=dict(text="离职类型分布", font=dict(size=20)), height=500, legend=dict(font=dict(size=14)))
                st.plotly_chart(fig_pie, use_container_width=True, config=get_download_config(r["month"] + "_离职类型分布"))
                st.dataframe(df_type, use_container_width=True, hide_index=True)
            
            st.markdown("---")
            
            if r["tenure"]:
                st.subheader("离职司龄分布")
                df_tenure = pd.DataFrame(r["tenure"])
                
                fig_tenure = px.pie(df_tenure, values="人数", names="司龄段", title="司龄分布", color_discrete_sequence=px.colors.qualitative.Set3, hole=0.5)
                fig_tenure.update_traces(textposition="outside", textinfo="label+percent+value", textfont=dict(size=13), marker=dict(line=dict(color="white", width=3)), hovertemplate="<b>%{label}</b><br>人数: %{value}人<br>占比: %{percent}<extra></extra>")
                fig_tenure.update_layout(title=dict(text="离职司龄分布", font=dict(size=20)), height=500, legend=dict(font=dict(size=14)))
                st.plotly_chart(fig_tenure, use_container_width=True, config=get_download_config(r["month"] + "_司龄分布"))
                st.dataframe(df_tenure, use_container_width=True, hide_index=True)
            
            st.markdown("---")
            
            if r["reason"]:
                st.subheader("离职原因分布")
                df_reason = pd.DataFrame(r["reason"])
                
                fig_reason = go.Figure()
                fig_reason.add_trace(go.Bar(x=df_reason["离职原因"], y=df_reason["人数"], marker_color=COLORS["warning"], text=[f"{d}人({p}%)" for d, p in zip(df_reason["人数"], df_reason["占比"])], textposition="outside", width=0.7))
                fig_reason.update_layout(title=dict(text="离职原因分布（纵向）", font=dict(size=20)), xaxis=dict(title="离职原因", tickangle=-30, tickfont=dict(size=11), gridcolor="lightgray"), yaxis=dict(title="人数", title_font=dict(size=14), tickfont=dict(size=12), gridcolor="lightgray"), height=max(450, len(df_reason) * 60), margin=dict(b=120, l=80), plot_bgcolor="white", paper_bgcolor="white")
                st.plotly_chart(fig_reason, use_container_width=True, config=get_download_config(r["month"] + "_离职原因分布"))
                st.dataframe(df_reason, use_container_width=True, hide_index=True)
            
            st.markdown("---")
            
            if r["level"]:
                st.subheader("职级分布（合并后）")
                df_level = pd.DataFrame(r["level"])
                
                fig_level = go.Figure()
                fig_level.add_trace(go.Bar(x=df_level["职级合并"], y=df_level["人数"], marker_color=COLORS["purple"], text=[f"{d}人({p}%)" for d, p in zip(df_level["人数"], df_level["占比"])], textposition="outside", width=0.7))
                fig_level.update_layout(title=dict(text="职级分布（人数+占比）", font=dict(size=20)), xaxis=dict(title="职级", tickfont=dict(size=14), gridcolor="lightgray"), yaxis=dict(title="人数", title_font=dict(size=14), tickfont=dict(size=12), gridcolor="lightgray"), height=450, plot_bgcolor="white", paper_bgcolor="white")
                st.plotly_chart(fig_level, use_container_width=True, config=get_download_config(r["month"] + "_职级分布"))
                st.dataframe(df_level.rename(columns={"职级合并": "职级"}), use_container_width=True, hide_index=True)
            
            st.markdown("---")
            
            # AI智能分析
            st.subheader("AI智能分析")
            with st.spinner("正在生成分析报告..."):
                analysis = generate_ai_analysis(r)
            st.markdown(analysis)
            st.markdown("---")
            
            # 导出Excel
            st.subheader("导出数据")
            out = BytesIO()
            with pd.ExcelWriter(out, engine="openpyxl") as w:
                pd.DataFrame([{"月份": r["month"], "期初人数": r["start"], "期末人数": r["end"], "平均人数": r["avg"], "离职人数": r["turn"], "离职率": str(r["rate"]) + "%"}]).to_excel(w, sheet_name="概览", index=False)
                if r["dept"]: pd.DataFrame(r["dept"]).to_excel(w, sheet_name="部门离职", index=False)
                if r["type"]: pd.DataFrame(r["type"]).to_excel(w, sheet_name="离职类型", index=False)
                if r["reason"]: pd.DataFrame(r["reason"]).to_excel(w, sheet_name="离职原因", index=False)
                if r["tenure"]: pd.DataFrame(r["tenure"]).to_excel(w, sheet_name="司龄分布", index=False)
                if r["level"]: pd.DataFrame(r["level"]).to_excel(w, sheet_name="职级分布", index=False)
            
            st.download_button(
                label="下载Excel报告",
                data=out.getvalue(),
                file_name=r["month"] + "离职分析.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.info("请上传Excel文件开始分析")
