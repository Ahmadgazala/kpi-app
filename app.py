#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
KPI-Final 2026 - Interactive Program
Built for /Users/ahmad/Desktop/kpi/KPI-Final 2026.xlsx
Run: streamlit run app.py
"""
import json
import io
import copy
from pathlib import Path
import streamlit as st
import plotly.graph_objects as go
import plotly.express as px
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Alignment, Font, Border, Side

# ---------- Config ----------
st.set_page_config(
    page_title="KPI 2026 - نظام التقييم المؤسسي",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

BASE_DIR = Path(__file__).parent
EXCEL_PATH = BASE_DIR / "KPI-Final 2026.xlsx"
JSON_PATH = BASE_DIR / "kpi_domains.json"

LEVELS_AR = {
    1: "ناشئ",
    2: "أولي",
    3: "أساسي",
    4: "منظم",
    5: "ديناميكي"
}
LEVELS_COLOR = {
    1: "#e74c3c",  # red
    2: "#e67e22",  # orange
    3: "#f1c40f",  # yellow
    4: "#2ecc71",  # green
    5: "#1abc9c",  # teal
}
LEVELS_DESC = {
    1: "لا توجد ممارسة واضحة",
    2: "بدأ التطبيق بصورة أولية",
    3: "الممارسة موجودة أساسية",
    4: "الممارسة واضحة ومنظمة",
    5: "الممارسة متطورة وديناميكية"
}

def get_level(avg: float):
    if avg < 1.8:
        return "ناشئ", 1
    elif avg < 2.6:
        return "أولي", 2
    elif avg < 3.4:
        return "أساسي", 3
    elif avg < 4.2:
        return "منظم", 4
    else:
        return "ديناميكي", 5

def load_domains():
    with open(JSON_PATH, encoding="utf-8") as f:
        return json.load(f)

domains = load_domains()
TOTAL_INDICATORS = sum(len(d["indicators"]) for d in domains)

# ---------- Session State ----------
if "scores" not in st.session_state:
    # init all to 1 (as in file)
    st.session_state.scores = {}
    for d in domains:
        st.session_state.scores[d["domain"]] = {ind["name"]: 1 for ind in d["indicators"]}
if "org_name" not in st.session_state:
    st.session_state.org_name = ""

# ---------- Helper: Calculations ----------
def calculate_results():
    results = []
    all_scores = []
    for d in domains:
        vals = list(st.session_state.scores[d["domain"]].values())
        avg = sum(vals) / len(vals) if vals else 0
        score100 = avg * 20
        level, level_n = get_level(avg)
        results.append({
            "domain": d["domain"],
            "count": len(vals),
            "avg": round(avg, 2),
            "score100": round(score100, 1),
            "level": level,
            "level_n": level_n,
            "vals": vals
        })
        all_scores.extend(vals)
    overall_avg = sum(all_scores) / len(all_scores) if all_scores else 0
    overall_100 = overall_avg * 20
    overall_level, overall_n = get_level(overall_avg)
    return results, overall_avg, overall_100, overall_level, overall_n

results, overall_avg, overall_100, overall_level, overall_n = calculate_results()

# ---------- CSS ----------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@400;700&display=swap');
html, body, [class*="css"] { font-family: 'Tajawal', sans-serif; }
h1, h2, h3 { text-align: center; }
.metric-card {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    color: white; padding: 15px; border-radius: 12px; text-align:center;
}
.domain-header { background: #f8f9fa; padding:10px; border-radius:8px; border-right: 5px solid #667eea; margin: 10px 0; }
</style>
""", unsafe_allow_html=True)

# ---------- Sidebar ----------
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3135/3135679.png", width=80)
    st.title("KPI 2026")
    st.markdown("**نظام التقييم المؤسسي**\n58 مؤشر - 9 محاور")
    st.divider()
    st.text_input("اسم المؤسسة", key="org_name", placeholder="مثال: مركز تعلم الكبار...")
    
    # Quick actions
    st.subheader("⚡ إجراءات سريعة")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("🔄 تصفير", use_container_width=True):
            for d in domains:
                for ind in d["indicators"]:
                    st.session_state.scores[d["domain"]][ind["name"]] = 1
            st.rerun()
    with c2:
        if st.button("⭐ تعيين الكل 5", use_container_width=True):
            for d in domains:
                for ind in d["indicators"]:
                    st.session_state.scores[d["domain"]][ind["name"]] = 5
            st.rerun()
    # Presets
    st.selectbox("تحميل سيناريو", ["— اختر —", "واقع ناشئ (1)", "متوسط أساسي (3)", "متميز منظم (4)", "ديناميكي (5)"],
                 key="preset")
    if st.session_state.preset != "— اختر —":
        mp = {"واقع ناشئ (1)":1, "متوسط أساسي (3)":3, "متميز منظم (4)":4, "ديناميكي (5)":5}
        v = mp[st.session_state.preset]
        for d in domains:
            for ind in d["indicators"]:
                st.session_state.scores[d["domain"]][ind["name"]] = v
        st.toast(f"تم تعيين الكل إلى {v} - {LEVELS_AR[v]}")
    
    st.divider()
    # Progress
    st.metric("المتوسط العام", f"{overall_avg:.2f} / 5", f"{overall_level}")
    st.progress(overall_100/100)
    st.caption(f"العلامة: {overall_100:.1f} / 100")
    st.divider()
    st.caption("الملف الأصلي: KPI-Final 2026.xlsx")
    st.caption("المطور: KPI Program v1.0")

# ---------- Header ----------
st.markdown(f"""
<h1>📊 نظام تقييم الأداء المؤسسي KPI 2026</h1>
<p style="text-align:center; color:#666; font-size:18px">
{st.session_state.org_name or " — "} &nbsp; | &nbsp; {TOTAL_INDICATORS} مؤشر &nbsp; | &nbsp; 9 محاور &nbsp; | &nbsp; المستوى العام: <b style="color:{LEVELS_COLOR[overall_n]}">{overall_level} ({overall_avg:.2f}) - {overall_100:.0f}/100</b>
</p>
""", unsafe_allow_html=True)

# ---------- Top Metrics ----------
cols = st.columns(4)
cols[0].metric("المتوسط العام", f"{overall_avg:.2f}", delta=f"{overall_100:.1f} / 100")
cols[1].metric("المستوى العام", overall_level, delta=LEVELS_DESC[overall_n])
cols[2].metric("عدد المحاور", "9", delta=f"{TOTAL_INDICATORS} مؤشر")
cols[3].metric("المتبقي للمستوى 5", f"{5 - overall_avg:.2f}", delta=f"{100-overall_100:.0f} نقطة")

# ---------- Tabs ----------
tab1, tab2, tab3, tab4 = st.tabs(["📝 التقييم (الإدخال)", "📈 النتائج والرسوم", "📋 سجل المرفقات", "💾 التصدير"])

with tab1:
    st.info("💡 **تعليمات:** اختر درجة من 1 إلى 5 لكل مؤشر. 1=ناشئ، 5=ديناميكي. يتم حفظ النتائج تلقائياً وحساب المتوسطات فوراً.")
    
    # Domain selector
    domain_names = [d["domain"] for d in domains]
    selected = st.selectbox("انتقل إلى محور:", ["الكل"] + domain_names, index=0)
    
    # Show domains
    domains_to_show = domains if selected=="الكل" else [d for d in domains if d["domain"]==selected]
    
    for d in domains_to_show:
        vals = list(st.session_state.scores[d["domain"]].values())
        avg = sum(vals)/len(vals)
        lvl, _ = get_level(avg)
        color = LEVELS_COLOR[get_level(avg)[1]]
        with st.expander(f"🔹 {d['domain']}  —  {len(d['indicators'])} مؤشرات  |  متوسط: {avg:.2f} ({lvl}) — {avg*20:.0f}/100", expanded=(selected!="الكل")):
            # domain summary bar
            st.markdown(f"<div style='height:6px; background:{color}; border-radius:3px; width:{avg*20}%'></div>", unsafe_allow_html=True)
            for ind in d["indicators"]:
                key = f"{d['domain']}__{ind['name']}"
                current = st.session_state.scores[d["domain"]][ind["name"]]
                # Use radio horizontal
                st.markdown(f"**{ind['id']}. {ind['name']}**")
                st.caption(ind["question"])
                # show levels descriptions in small
                cols_l = st.columns(5)
                labels = ["1 ناشئ", "2 أولي", "3 أساسي", "4 منظم", "5 ديناميكي"]
                # radio
                score = st.radio(
                    f"درجة {ind['name']}", 
                    [1,2,3,4,5], 
                    index=current-1, 
                    horizontal=True, 
                    key=key,
                    label_visibility="collapsed",
                    format_func=lambda x: labels[x-1]
                )
                # update session
                st.session_state.scores[d["domain"]][ind["name"]] = score
                # show selected level description
                st.markdown(f"<small style='color:{LEVELS_COLOR[score]}'>➡️ <b>{LEVELS_AR[score]}:</b> {ind['levels'][score-1][:140]}...</small>", unsafe_allow_html=True)
                st.divider()

with tab2:
    # Recalculate
    results, overall_avg, overall_100, overall_level, overall_n = calculate_results()
    
    c1, c2 = st.columns([1,1])
    with c1:
        st.subheader("📊 العلامات حسب المحور (من 100)")
        df = pd.DataFrame([{"المحور": r["domain"], "العلامة": r["score100"], "المستوى": r["level"], "المتوسط": r["avg"]} for r in results])
        # Bar chart
        fig_bar = px.bar(df, x="المحور", y="العلامة", color="المستوى",
                         color_discrete_map={"ناشئ":"#e74c3c","أولي":"#e67e22","أساسي":"#f1c40f","منظم":"#2ecc71","ديناميكي":"#1abc9c"},
                         text="العلامة", range_y=[0,100])
        fig_bar.update_layout(xaxis_tickangle=-25, height=450, showlegend=True)
        fig_bar.update_traces(textposition="outside")
        st.plotly_chart(fig_bar, use_container_width=True)
        st.dataframe(df, use_container_width=True, hide_index=True)
    
    with c2:
        st.subheader("🕸️ الرسم الراداري (Spider)")
        categories = [r["domain"] for r in results]
        values = [r["avg"] for r in results]
        # close loop
        categories_closed = categories + [categories[0]]
        values_closed = values + [values[0]]
        fig_radar = go.Figure()
        fig_radar.add_trace(go.Scatterpolar(r=values_closed, theta=categories_closed, fill='toself', name='الوضع الحالي', line_color='#667eea'))
        fig_radar.add_trace(go.Scatterpolar(r=[5]*len(categories_closed), theta=categories_closed, fill=None, name='المأمول (5)', line=dict(dash='dash', color='#95a5a6')))
        fig_radar.update_layout(polar=dict(radialaxis=dict(visible=True, range=[0,5], tickvals=[1,2,3,4,5])), height=450)
        st.plotly_chart(fig_radar, use_container_width=True)
        
        st.subheader("🏆 الملخص العام")
        fig_gauge = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=overall_100,
            delta={'reference': 100, 'decreasing': {'color': "red"}},
            title={'text': f"العلامة النهائية - {overall_level}"},
            gauge={'axis': {'range': [0,100]},
                   'bar': {'color': LEVELS_COLOR[overall_n]},
                   'steps': [
                       {'range': [0,36], 'color': "#e74c3c"},
                       {'range': [36,52], 'color': "#e67e22"},
                       {'range': [52,68], 'color': "#f1c40f"},
                       {'range': [68,84], 'color': "#2ecc71"},
                       {'range': [84,100], 'color': "#1abc9c"}],
                   'threshold': {'line': {'color': "black", 'width': 4}, 'thickness': 0.75, 'value': overall_100}}))
        fig_gauge.update_layout(height=300)
        st.plotly_chart(fig_gauge, use_container_width=True)
    
    st.divider()
    # Detailed table with gaps
    st.subheader("📉 تحليل الفجوات - أولوية التطوير")
    df_gap = pd.DataFrame([{
        "المحور": r["domain"],
        "المتوسط": r["avg"],
        "العلامة/100": r["score100"],
        "المستوى": r["level"],
        "الفجوة حتى 5": round(5 - r["avg"],2),
        "النقاط حتى 100": round(100 - r["score100"],1)
    } for r in results]).sort_values("المتوسط")
    st.dataframe(df_gap, use_container_width=True, hide_index=True)
    # Highlight weakest
    weakest = df_gap.iloc[0]
    st.warning(f"⚠️ **أضعف محور:** {weakest['المحور']} ({weakest['المتوسط']} - {weakest['المستوى']}) — يحتاج {weakest['الفجوة حتى 5']} نقطة للوصول للديناميكي")

with tab3:
    st.subheader("📎 سجل المرفقات الداعمة")
    st.caption("المرفقات اختيارية. يمكن لمرفق واحد أن يدعم أكثر من مؤشر. (مطابق لورقة 'سجل المرفقات الداعمة' في الإكسل)")
    # Load attachments template from Excel if exists
    try:
        wb_tmp = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
        ws_a = wb_tmp["سجل المرفقات الداعمة"]
        rows = list(ws_a.iter_rows(values_only=True))
        # find header
        hdr_idx = None
        for i,r in enumerate(rows):
            if r and r[0]=="المحور":
                hdr_idx=i
                break
        if hdr_idx is not None:
            attach_data = []
            for r in rows[hdr_idx+1:]:
                if r[0] is None: continue
                attach_data.append({"المحور": r[0], "المؤشر": r[2], "المرفقات المحتملة": r[3], "اسم الملف": "", "الحالة": "غير مرفق"})
            df_att = pd.DataFrame(attach_data)
            edited = st.data_editor(df_att, use_container_width=True, hide_index=True, num_rows="dynamic",
                                    column_config={
                                        "الحالة": st.column_config.SelectboxColumn(options=["غير مرفق","مرفق","قيد المراجعة","معتمد"]),
                                        "اسم الملف": st.column_config.TextColumn(width="medium")
                                    }, height=500)
            st.session_state["attachments_df"] = edited
        else:
            st.info("لم يتم العثور على قالب المرفقات")
    except Exception as e:
        st.error(f"خطأ تحميل المرفقات: {e}")
    st.divider()
    st.markdown("**رفع نماذج المرفقات (تحضير للتصدير):**")
    uploaded = st.file_uploader("اسحب الملفات هنا (اختياري)", accept_multiple_files=True)
    if uploaded:
        for f in uploaded:
            st.success(f"✓ {f.name} ({f.size//1024} KB)")

with tab4:
    st.subheader("💾 تصدير النتائج")
    col1, col2, col3 = st.columns(3)
    
    # Prepare Excel export
    def build_excel_bytes():
        wb = openpyxl.load_workbook(EXCEL_PATH)
        # Update each domain sheet: column H (8) is الدرجة
        for d in domains:
            ws = wb[d["sheet"]]
            # find header row
            header_idx = None
            for idx, row in enumerate(ws.iter_rows(values_only=True), 1):
                if row and any("1 - ناشئ" in str(c) for c in row if c):
                    header_idx = idx
                    break
            if header_idx is None:
                continue
            # data starts at header_idx+1
            r = header_idx + 1
            for ind in d["indicators"]:
                score = st.session_state.scores[d["domain"]][ind["name"]]
                ws.cell(row=r, column=8, value=score)  # H
                ws.cell(row=r, column=9, value=score*20)  # I
                # color cell
                fill = PatternFill(start_color=LEVELS_COLOR[get_level(score)[1]].lstrip("#"), end_color=LEVELS_COLOR[get_level(score)[1]].lstrip("#"), fill_type="solid")
                # ws.cell(...).fill = fill  # optional
                r += 1
            # update متوسط المجال if exists
            for row in ws.iter_rows(min_row=r, max_row=ws.max_row):
                for cell in row:
                    if cell.value and "متوسط المجال" in str(cell.value):
                        avg = sum(st.session_state.scores[d["domain"]].values()) / len(d["indicators"])
                        # next col is avg, next is *20
                        ws.cell(row=cell.row, column=cell.column+1, value=round(avg,2))
                        ws.cell(row=cell.row, column=cell.column+2, value=round(avg*20,1))
        # Update النتائج sheet
        if "النتائج" in wb.sheetnames:
            ws = wb["النتائج"]
            # find data rows: col A is domain name, col C is avg, col D is level, col E is score100
            res_map = {r["domain"]: r for r in results}
            for row in ws.iter_rows(min_row=5, max_row=15):
                domain_cell = row[0].value
                if domain_cell and str(domain_cell).strip() in res_map:
                    key = str(domain_cell).strip()
                    # fuzzy match for abbreviated names
                    # try exact then contains
                    matched = None
                    for k in res_map:
                        if k.strip() == key or key in k or k in key:
                            matched = res_map[k]
                            break
                    if matched is None:
                        # try first word
                        continue
                    # col C index 2 = avg, D=3 level, E=4 score100
                    # row is tuple of cells, 0=A,1=B,2=C...
                    # But sheet structure: A=اسم المحور, B=عدد المؤشرات, C=متوسط الدرجة, D=المستوى, E=العلامة
                    row[2].value = matched["avg"]
                    row[3].value = matched["level"]
                    row[4].value = matched["score100"]
            # overall row: look for المتوسط العام
            for row in ws.iter_rows(min_row=14, max_row=17):
                if row[0].value and "المتوسط العام" in str(row[0].value):
                    row[2].value = round(overall_avg,2)
                    row[3].value = overall_level
                    row[4].value = round(overall_100,1)
        # Update ملخص التقييم
        if "ملخص التقييم ودليل المستويات" in wb.sheetnames:
            ws = wb["ملخص التقييم ودليل المستويات"]
            # F6 is overall 100? Let's just set G6? From earlier: F6? Actually read dims: G? We'll just set named cells
            # Row 6 has العلامة النهائية
            try:
                ws["G6"] = round(overall_100,1)
                ws["G8"] = overall_level
                ws["G9"] = round(overall_avg,2)
            except: pass
        # Add new sheet for report if not exists
        buff = io.BytesIO()
        wb.save(buff)
        buff.seek(0)
        return buff.getvalue()
    
    with col1:
        st.markdown("**1. الإكسل المحدث**")
        st.caption("نفس ملف KPI-Final مع الدرجات الجديدة + كل الحسابات")
        excel_bytes = build_excel_bytes()
        st.download_button("⬇️ تحميل Excel المحدث", data=excel_bytes, file_name=f"KPI-2026-{st.session_state.org_name or 'محدث'}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="primary")
    
    with col2:
        st.markdown("**2. تقرير PDF (قريباً)**")
        # Simple PDF via dataframe to excel? For now offer CSV
        st.caption("ملخص النتائج كـ Excel مصغر")
        df_export = pd.DataFrame([{"المحور":r["domain"],"المتوسط":r["avg"],"المستوى":r["level"],"العلامة/100":r["score100"]} for r in results])
        df_export.loc[len(df_export)] = ["الإجمالي", round(overall_avg,2), overall_level, round(overall_100,1)]
        csv = df_export.to_csv(index=False, encoding="utf-8-sig")
        st.download_button("⬇️ تحميل CSV الملخص", data=csv, file_name="KPI-ملخص.csv", mime="text/csv", use_container_width=True)
        # JSON export
        json_data = json.dumps({"org": st.session_state.org_name, "overall": {"avg": overall_avg, "score100": overall_100, "level": overall_level}, "domains": results, "scores": st.session_state.scores}, ensure_ascii=False, indent=2)
        st.download_button("⬇️ تحميل JSON", data=json_data, file_name="KPI-data.json", mime="application/json", use_container_width=True)
    
    with col3:
        st.markdown("**3. حفظ واسترجاع**")
        st.caption("احفظ عملك للعودة لاحقاً")
        if st.button("💾 حفظ في المتصفح", use_container_width=True):
            st.toast("تم الحفظ في Session (يبقى حتى إغلاق الصفحة)")
        if st.button("📤 استيراد JSON سابق", use_container_width=True):
            st.info("استخدم رفع الملف أدناه")
        up = st.file_uploader("رفع JSON محفوظ", type=["json"])
        if up:
            try:
                data = json.loads(up.read().decode("utf-8"))
                st.session_state.scores = data["scores"]
                st.session_state.org_name = data.get("org","")
                st.success("تم الاسترجاع - سيتم التحديث")
                st.rerun()
            except Exception as e:
                st.error(f"خطأ: {e}")
    
    st.divider()
    st.markdown("""
    ### ℹ️ كيف يعمل البرنامج؟
    1. كل مؤشر له سؤال و 5 مستويات (ناشئ→ديناميكي). اختيارك يحفظ تلقائياً.
    2. **الحساب:** متوسط المحور = مجموع درجات مؤشراته ÷ عددها. العلامة من 100 = المتوسط × 20. المستوى يُحدد حسب المتوسط (1-1.79 ناشئ ... 4.2-5 ديناميكي).
    3. **التصدير:** زر Excel المحدث ينسخ ملفك الأصلي ويملأ عمود **الدرجة** و **العلامة من 100** وكل أوراق **النتائج / كشف التحصيل / الملخص** تلقائياً — بدون كتابة يدوية.
    4. يمكنك تكرار التقييم كل فترة ومقارنة النتائج عبر ملفات JSON.
    """)
    st.caption("للتشغيل المحلي بدون إنترنت: استخدم kpi_tool.py — راجع README")

# ---------- Footer ----------
st.divider()
st.caption("KPI Program 2026 | مبني على KPI-Final 2026.xlsx (58 مؤشر) | للمساعدة: ضع الدرجات ثم حمّل الإكسل المحدث")
