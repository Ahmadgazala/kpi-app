#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
KPI 2026 - Local Desktop App (Tkinter)
Runs offline, no browser. Double-click or: python3 local_app.py
"""
import json, pathlib, io, sys, os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import openpyxl

# matplotlib for charts (embedded)
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

def get_base():
    # Works for normal python and for PyInstaller frozen exe (.app / .exe)
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        return pathlib.Path(sys._MEIPASS)
    # When running as single-file exe, EXCEL/JSON may be next to exe, not in _MEIPASS
    exe_dir = pathlib.Path(sys.executable).parent if getattr(sys, 'frozen', False) else pathlib.Path(__file__).parent
    # Prefer exe_dir if files there, else file parent
    if (exe_dir / "KPI-Final 2026.xlsx").exists():
        return exe_dir
    return pathlib.Path(__file__).parent

BASE = get_base()
EXCEL = BASE / "KPI-Final 2026.xlsx"
JSON_PATH = BASE / "kpi_domains.json"
# Fallback: if not found in BASE, try current working dir + app dir
if not EXCEL.exists():
    for p in [pathlib.Path.cwd(), pathlib.Path(__file__).parent, pathlib.Path(sys.executable).parent]:
        if (p / "KPI-Final 2026.xlsx").exists():
            EXCEL = p / "KPI-Final 2026.xlsx"
            JSON_PATH = p / "kpi_domains.json"
            BASE = p
            break

LEVELS_AR = {1:"ناشئ",2:"أولي",3:"أساسي",4:"منظم",5:"ديناميكي"}
COLORS = {1:"#e74c3c",2:"#e67e22",3:"#f1c40f",4:"#2ecc71",5:"#1abc9c"}
LEVEL_HEX = ["#e74c3c","#e67e22","#f1c40f","#2ecc71","#1abc9c"]

def get_level(avg):
    if avg < 1.8: return "ناشئ",1
    if avg < 2.6: return "أولي",2
    if avg < 3.4: return "أساسي",3
    if avg < 4.2: return "منظم",4
    return "ديناميكي",5

def load_domains():
    return json.loads(JSON_PATH.read_text(encoding="utf-8"))

domains = load_domains()

class ScrollableFrame(ttk.Frame):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        canvas = tk.Canvas(self, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.scrollable = ttk.Frame(canvas)
        self.scrollable.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0,0), window=self.scrollable, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        # mouse wheel
        def _on_wheel(e):
            canvas.yview_scroll(int(-1*(e.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_wheel)

class KPIApp:
    def __init__(self, root):
        self.root = root
        self.root.title("KPI 2026 - نظام التقييم المؤسسي (Local) - Mac & Windows")
        self.root.geometry("1300x750")
        # High DPI for Windows
        try:
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
        except: pass
        try: self.root.iconbitmap("") # no icon
        except: pass
        # style - cross-platform fonts: Helvetica (Mac), Segoe UI (Windows), fallback
        import platform
        is_win = platform.system() == "Windows"
        font_base = "Segoe UI" if is_win else "Helvetica"
        style = ttk.Style()
        try: style.theme_use("clam" if not is_win else "vista")
        except: pass
        style.configure("Title.TLabel", font=(font_base, 16, "bold"))
        style.configure("Metric.TLabel", font=(font_base, 11, "bold"))
        style.configure("Domain.TButton", font=(font_base, 10))
        self.font_base = font_base
        
        # data
        self.scores = {d["domain"]: {ind["name"]: 1 for ind in d["indicators"]} for d in domains}
        self.org_var = tk.StringVar(value="")
        self.vars = {} # (domain, name) -> IntVar
        
        # --- Top bar ---
        top = ttk.Frame(root, padding=8)
        top.pack(fill="x")
        ttk.Label(top, text="📊 KPI 2026 - التقييم المؤسسي", style="Title.TLabel").pack(side="left")
        ttk.Label(top, text=" اسم المؤسسة: ").pack(side="left", padx=(30,2))
        ttk.Entry(top, textvariable=self.org_var, width=28).pack(side="left")
        # overall metrics
        self.lbl_overall = ttk.Label(top, text="المتوسط: 1.00 (ناشئ)  20/100", font=("Helvetica", 11, "bold"), foreground="#e74c3c")
        self.lbl_overall.pack(side="right", padx=10)
        self.progress = ttk.Progressbar(top, length=160, mode="determinate", maximum=100, value=20)
        self.progress.pack(side="right")
        ttk.Label(top, text=" التقدم ").pack(side="right")

        # --- Notebook ---
        nb = ttk.Notebook(root)
        nb.pack(fill="both", expand=True, padx=6, pady=4)
        self.tab_eval = ttk.Frame(nb)
        self.tab_results = ttk.Frame(nb)
        nb.add(self.tab_eval, text="  📝 التقييم  ")
        nb.add(self.tab_results, text="  📈 النتائج والتصدير  ")

        self.build_eval_tab()
        self.build_results_tab()
        self.update_overall()

    def build_eval_tab(self):
        # Paned: left domains, right indicators
        paned = ttk.Panedwindow(self.tab_eval, orient="horizontal")
        paned.pack(fill="both", expand=True)

        # Left: domain list
        left = ttk.Frame(paned, width=280)
        paned.add(left, weight=0)
        ttk.Label(left, text="المحاور (9)", font=("Helvetica", 11, "bold")).pack(pady=6)
        self.domain_list = tk.Listbox(left, font=("Helvetica", 11), activestyle="dotbox", bg="#f8f9fa", selectbackground="#667eea")
        for d in domains:
            self.domain_list.insert("end", f"{d['domain']} ({len(d['indicators'])})")
        self.domain_list.pack(fill="both", expand=True, padx=6, pady=4)
        self.domain_list.selection_set(0)
        self.domain_list.bind("<<ListboxSelect>>", lambda e: self.show_domain())
        # quick actions
        btnf = ttk.Frame(left)
        btnf.pack(fill="x", pady=6, padx=6)
        ttk.Button(btnf, text="تصفير (1)", command=lambda: self.set_all(1)).pack(side="left", expand=True, padx=2)
        ttk.Button(btnf, text="الكل 5", command=lambda: self.set_all(5)).pack(side="left", expand=True, padx=2)
        # presets
        ttk.Label(left, text="سيناريو سريع:").pack()
        self.preset = ttk.Combobox(left, values=["— اختر —","ناشئ (1)","أساسي (3)","منظم (4)","ديناميكي (5)"], state="readonly", width=18)
        self.preset.current(0)
        self.preset.pack(pady=2)
        self.preset.bind("<<ComboboxSelected>>", self.apply_preset)

        # Right: scrollable indicators
        right = ttk.Frame(paned)
        paned.add(right, weight=1)
        self.scroll = ScrollableFrame(right)
        self.scroll.pack(fill="both", expand=True)
        self.show_domain()

    def show_domain(self):
        # clear
        for w in self.scroll.scrollable.winfo_children():
            w.destroy()
        idxs = self.domain_list.curselection()
        idx = idxs[0] if idxs else 0
        d = domains[idx]
        # domain header
        vals = list(self.scores[d["domain"]].values())
        avg = sum(vals)/len(vals)
        lvl, n = get_level(avg)
        hdr = ttk.Frame(self.scroll.scrollable, padding=8)
        hdr.pack(fill="x", pady=(6,8))
        ttk.Label(hdr, text=d["domain"], font=("Helvetica", 13, "bold")).pack(anchor="w")
        ttk.Label(hdr, text=f"{len(d['indicators'])} مؤشر  |  متوسط: {avg:.2f} ({lvl})  |  {avg*20:.0f}/100", foreground=COLORS[n]).pack(anchor="w")
        # bar
        bar = tk.Canvas(hdr, height=8, bg="#ecf0f1", highlightthickness=0)
        bar.pack(fill="x", pady=4)
        bar.create_rectangle(0,0, int(avg*20*6), 8, fill=COLORS[n], outline="")

        for ind in d["indicators"]:
            card = ttk.Frame(self.scroll.scrollable, padding=8, relief="flat")
            card.pack(fill="x", padx=8, pady=6)
            # separator line
            ttk.Separator(self.scroll.scrollable, orient="horizontal").pack(fill="x", padx=12)
            ttk.Label(card, text=f"{ind['id']}. {ind['name']}", font=("Helvetica", 11, "bold")).pack(anchor="w")
            lbl_q = tk.Label(card, text=ind["question"], wraplength=820, justify="left", fg="#555", font=("Helvetica", 9))
            lbl_q.pack(anchor="w", pady=2)

            # radio vars
            key = (d["domain"], ind["name"])
            if key not in self.vars:
                self.vars[key] = tk.IntVar(value=self.scores[d["domain"]][ind["name"]])
            var = self.vars[key]

            radio_frame = ttk.Frame(card)
            radio_frame.pack(anchor="w", pady=4)
            labels = ["1 ناشئ","2 أولي","3 أساسي","4 منظم","5 ديناميكي"]
            for i in range(1,6):
                rb = ttk.Radiobutton(radio_frame, text=labels[i-1], variable=var, value=i, command=lambda k=key, v=var: self.on_score_change(k, v))
                rb.pack(side="left", padx=6)
            # description
            desc = tk.Label(card, text="", wraplength=820, justify="left", font=("Helvetica", 8), fg=COLORS[var.get()])
            desc.pack(anchor="w")
            # store desc label to update
            var._desc = desc
            var._levels = ind["levels"]
            self.update_desc(var)

    def update_desc(self, var):
        v = var.get()
        var._desc.config(text=f"➡ {LEVELS_AR[v]}: {var._levels[v-1][:160]}...", fg=COLORS[v])

    def on_score_change(self, key, var):
        domain, name = key
        self.scores[domain][name] = var.get()
        self.update_desc(var)
        self.update_overall()
        # refresh header bar without rebuilding all? just update overall - header will update on next domain switch; also update listbox text? optional
        # update domain header avg live: rebuild only header? simpler keep as is

    def set_all(self, val):
        for d in domains:
            for ind in d["indicators"]:
                self.scores[d["domain"]][ind["name"]] = val
                k=(d["domain"], ind["name"])
                if k in self.vars:
                    self.vars[k].set(val)
                    self.update_desc(self.vars[k])
        self.update_overall()
        self.show_domain()
        messagebox.showinfo("تم", f"تم تعيين كل المؤشرات إلى {val} - {LEVELS_AR[val]}")

    def apply_preset(self, e):
        mp={"ناشئ (1)":1,"أساسي (3)":3,"منظم (4)":4,"ديناميكي (5)":5}
        v=mp.get(self.preset.get())
        if v: self.set_all(v)

    def update_overall(self):
        all_vals=[]
        for d in domains:
            all_vals.extend(self.scores[d["domain"]].values())
        avg = sum(all_vals)/len(all_vals) if all_vals else 0
        lvl,n = get_level(avg)
        score100 = avg*20
        self.lbl_overall.config(text=f"المتوسط: {avg:.2f} ({lvl})  {score100:.0f}/100", foreground=COLORS[n])
        self.progress["value"]=score100
        # update title
        self.root.title(f"KPI 2026 - {lvl} {avg:.2f} - {score100:.0f}/100 - {self.org_var.get() or 'Local'}")

    # ---------- Results tab ----------
    def build_results_tab(self):
        # top controls
        ctrl = ttk.Frame(self.tab_results, padding=6)
        ctrl.pack(fill="x")
        ttk.Button(ctrl, text="🔄 تحديث الرسوم", command=self.refresh_charts).pack(side="left", padx=4)
        ttk.Button(ctrl, text="💾 تصدير Excel", command=self.export_excel).pack(side="left", padx=4)
        ttk.Button(ctrl, text="💾 تصدير CSV", command=self.export_csv).pack(side="left", padx=4)
        ttk.Button(ctrl, text="💾 حفظ JSON", command=self.save_json).pack(side="left", padx=4)
        ttk.Button(ctrl, text="📂 تحميل JSON", command=self.load_json).pack(side="left", padx=4)
        ttk.Button(ctrl, text="🎲 عشوائي للتجربة", command=self.random_demo).pack(side="right", padx=4)

        # paned: table left, charts right
        paned = ttk.Panedwindow(self.tab_results, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=6, pady=4)

        left = ttk.Frame(paned)
        paned.add(left, weight=0)
        ttk.Label(left, text="جدول النتائج", font=("Helvetica", 11, "bold")).pack()
        cols=("المحور","المتوسط","المستوى","من 100","الفجوة")
        self.tree = ttk.Treeview(left, columns=cols, show="headings", height=12)
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=95 if c!="المحور" else 220, anchor="center")
        self.tree.pack(fill="both", expand=True, pady=4)
        # scrollbar
        sb = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.lbl_gap = ttk.Label(left, text="", foreground="#e74c3c", font=("Helvetica", 9, "bold"))
        self.lbl_gap.pack(pady=4)

        right = ttk.Frame(paned)
        paned.add(right, weight=1)
        # matplotlib figure with 2 subplots
        self.fig, (self.ax1, self.ax2) = plt.subplots(1, 2, figsize=(7,3.2), dpi=100)
        self.fig.tight_layout(pad=2.5)
        self.canvas = FigureCanvasTkAgg(self.fig, master=right)
        self.canvas.get_tk_widget().pack(fill="both", expand=True)
        self.refresh_charts()

    def calc_results(self):
        res=[]
        all_vals=[]
        for d in domains:
            vals=list(self.scores[d["domain"]].values())
            avg=sum(vals)/len(vals)
            res.append({"domain":d["domain"],"avg":round(avg,2),"score100":round(avg*20,1),"level":get_level(avg)[0],"level_n":get_level(avg)[1],"gap":round(5-avg,2)})
            all_vals.extend(vals)
        overall=sum(all_vals)/len(all_vals)
        return res, round(overall,2), round(overall*20,1), get_level(overall)[0], get_level(overall)[1]

    def refresh_charts(self):
        results, avg, score100, lvl, n = self.calc_results()
        # table
        for i in self.tree.get_children():
            self.tree.delete(i)
        for r in results:
            self.tree.insert("", "end", values=(r["domain"], r["avg"], r["level"], r["score100"], r["gap"]))
        self.tree.insert("", "end", values=("الإجمالي", avg, lvl, score100, round(5-avg,2)), tags=("total",))
        self.tree.tag_configure("total", background="#eaf2ff", font=("Helvetica", 10, "bold"))
        weakest = min(results, key=lambda x: x["avg"])
        self.lbl_gap.config(text=f"أضعف محور: {weakest['domain']} ({weakest['avg']} - {weakest['level']})")

        # charts
        self.ax1.clear(); self.ax2.clear()
        # bar
        names=[r["domain"][:16] for r in results]
        scores=[r["score100"] for r in results]
        colors=[COLORS[r["level_n"]] for r in results]
        self.ax1.bar(names, scores, color=colors, edgecolor="black", linewidth=0.5)
        self.ax1.set_ylim(0,100)
        self.ax1.set_ylabel("من 100")
        self.ax1.set_title("العلامات حسب المحور", fontsize=9)
        self.ax1.tick_params(axis='x', rotation=25, labelsize=6)
        for i,v in enumerate(scores):
            self.ax1.text(i, v+1, f"{v:.0f}", ha="center", fontsize=7)

        # radar (polar) - approximate with barh if polar not available? use polar
        # Use polar subplot replacement: create simple line radar via ax2 polar
        # To avoid complexity, draw horizontal bar + gauge
        # Gauge-like: donut for overall
        self.ax2.pie([score100, 100-score100], colors=[COLORS[n], "#ecf0f1"], startangle=90, counterclock=False, wedgeprops=dict(width=0.35))
        self.ax2.text(0,0, f"{score100:.0f}\n{lvl}", ha="center", va="center", fontsize=12, fontweight="bold", color=COLORS[n])
        self.ax2.set_title(f"العلامة النهائية {avg:.2f}/5", fontsize=9)
        self.fig.tight_layout()
        self.canvas.draw()
        self.update_overall()

    def export_excel(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=f"KPI-2026-{self.org_var.get() or 'result'}.xlsx", filetypes=[("Excel","*.xlsx")])
        if not path: return
        try:
            self._build_excel(path)
            messagebox.showinfo("تم", f"تم الحفظ:\n{path}")
        except Exception as e:
            messagebox.showerror("خطأ", str(e))

    def _build_excel(self, out_path):
        wb=openpyxl.load_workbook(EXCEL)
        results, avg, score100, lvl, n = self.calc_results()
        res_map={r["domain"]:r for r in results}
        for d in domains:
            ws=wb[d["sheet"]]
            hdr=None
            for idx,row in enumerate(ws.iter_rows(values_only=True),1):
                if row and any("1 - ناشئ" in str(c) for c in row if c):
                    hdr=idx; break
            if not hdr: continue
            r=hdr+1
            for ind in d["indicators"]:
                sc=self.scores[d["domain"]][ind["name"]]
                ws.cell(row=r,column=8,value=sc)
                ws.cell(row=r,column=9,value=sc*20)
                r+=1
            for row in ws.iter_rows(min_row=r,max_row=ws.max_row):
                for cell in row:
                    if cell.value and "متوسط المجال" in str(cell.value):
                        v=sum(self.scores[d["domain"]].values())/len(d["indicators"])
                        ws.cell(row=cell.row,column=cell.column+1,value=round(v,2))
                        ws.cell(row=cell.row,column=cell.column+2,value=round(v*20,1))
        if "النتائج" in wb.sheetnames:
            ws=wb["النتائج"]
            for row in ws.iter_rows(min_row=5,max_row=15):
                dc=row[0].value
                if dc and str(dc).strip():
                    for k,v in res_map.items():
                        if str(dc).strip() in k or k in str(dc).strip():
                            row[2].value=v["avg"]; row[3].value=v["level"]; row[4].value=v["score100"]; break
                    if "الترويج" in str(dc) and res_map.get("الترويج والخبرة في تعلم الكبار"):
                        m=res_map["الترويج والخبرة في تعلم الكبار"]
                        row[2].value=m["avg"]; row[3].value=m["level"]; row[4].value=m["score100"]
            for row in ws.iter_rows(min_row=14,max_row=17):
                if row[0].value and "المتوسط العام" in str(row[0].value):
                    row[2].value=avg; row[3].value=lvl; row[4].value=score100
        wb.save(out_path)

    def export_csv(self):
        import csv
        path=filedialog.asksaveasfilename(defaultextension=".csv", initialfile="KPI-ملخص.csv", filetypes=[("CSV","*.csv")])
        if not path: return
        results, avg, score100, lvl, _ = self.calc_results()
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w=csv.writer(f)
            w.writerow(["المحور","المتوسط","المستوى","العلامة/100","الفجوة"])
            for r in results: w.writerow([r["domain"],r["avg"],r["level"],r["score100"],r["gap"]])
            w.writerow(["الإجمالي",avg,lvl,score100,round(5-avg,2)])
        messagebox.showinfo("تم","تم حفظ CSV")

    def save_json(self):
        path=filedialog.asksaveasfilename(defaultextension=".json", initialfile="KPI-scores.json", filetypes=[("JSON","*.json")])
        if not path: return
        data={"org":self.org_var.get(),"scores":self.scores}
        pathlib.Path(path).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        messagebox.showinfo("تم", f"حفظ JSON:\n{path}")

    def load_json(self):
        path=filedialog.askopenfilename(filetypes=[("JSON","*.json")])
        if not path: return
        data=json.loads(pathlib.Path(path).read_text(encoding="utf-8"))
        sc=data.get("scores", data)
        self.scores=sc
        self.org_var.set(data.get("org",""))
        # update vars
        for k,var in self.vars.items():
            dom,name=k
            if dom in sc and name in sc[dom]:
                var.set(sc[dom][name])
                self.update_desc(var)
        self.refresh_charts()
        self.show_domain()
        messagebox.showinfo("تم","تم تحميل JSON")

    def random_demo(self):
        import random
        for d in domains:
            for ind in d["indicators"]:
                v=random.randint(1,5)
                self.scores[d["domain"]][ind["name"]]=v
                k=(d["domain"],ind["name"])
                if k in self.vars:
                    self.vars[k].set(v)
                    self.update_desc(self.vars[k])
        self.refresh_charts()
        self.show_domain()

if __name__=="__main__":
    root=tk.Tk()
    app=KPIApp(root)
    root.mainloop()
