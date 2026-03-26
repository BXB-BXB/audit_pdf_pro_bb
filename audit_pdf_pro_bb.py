import fitz  # PyMuPDF
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, colorchooser
import threading
import os

class AuditAppBB:
    def __init__(self, root):
        self.root = root
        self.root.title("Audit PDF Pro BB - Bogdan Bahrim")
        self.root.geometry("1300x850")
        self.root.configure(bg="#f0f2f5")

        self.full_results = []
        self.pdf_path = ""
        self.highlight_color = (1, 1, 0) 

        # --- UI PANEL ---
        top = tk.Frame(root, bg="#2c3e50", pady=20, padx=20)
        top.pack(fill=tk.X)
        
        btn_style = {"width": 18, "bg": "#34495e", "fg": "white", "relief": "flat", "pady": 5}
        tk.Button(top, text="📁 1. Load Excel", command=self.load_excel, **btn_style).pack(side=tk.LEFT, padx=5)
        tk.Button(top, text="📄 2. Load PDF", command=self.load_pdf, **btn_style).pack(side=tk.LEFT, padx=5)
        
        tk.Label(top, text="Skip Pgs:", fg="white", bg="#2c3e50").pack(side=tk.LEFT, padx=(15,0))
        self.exclude_entry = tk.Entry(top, width=10); self.exclude_entry.pack(side=tk.LEFT, padx=5)
        
        tk.Button(top, text="🎨 Color", command=self.pick_color, bg="#34495e", fg="white").pack(side=tk.LEFT, padx=5)

        self.run_btn = tk.Button(top, text="⚡ START AUDIT", command=self.start_thread, 
                                 state=tk.DISABLED, bg="#27ae60", fg="white", font=("Arial", 10, "bold"), width=18)
        self.run_btn.pack(side=tk.RIGHT, padx=5)

        # --- EXPORT SETTINGS ---
        name_frame = tk.Frame(root, bg="#dee2e6", pady=10)
        name_frame.pack(fill=tk.X)
        tk.Label(name_frame, text="Output Name:", bg="#dee2e6", font=("Arial", 9, "bold")).pack(side=tk.LEFT, padx=10)
        self.base_name = tk.Entry(name_frame, width=30); self.base_name.insert(0, "Audit_Report"); self.base_name.pack(side=tk.LEFT)
        self.suffix_name = tk.Entry(name_frame, width=15); self.suffix_name.insert(0, "_v1"); self.suffix_name.pack(side=tk.LEFT)

        # --- TABLE VIEW ---
        self.tree_frame = tk.Frame(root)
        self.tree_frame.pack(pady=10, fill=tk.BOTH, expand=True, padx=20)
        cols = ("Sheet", "Identifier", "Description", "QTY_BOM", "Found", "Verdict", "Pages")
        self.tree = ttk.Treeview(self.tree_frame, columns=cols, show='headings')
        for c in cols:
            self.tree.heading(c, text=c); self.tree.column(c, width=100, anchor=tk.CENTER)
        self.tree.column("Description", width=350, anchor=tk.W)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.progress = ttk.Progressbar(root, orient=tk.HORIZONTAL, length=1200, mode='determinate')
        self.progress.pack(pady=10)

    def pick_color(self):
        c = colorchooser.askcolor()[0]
        if c: self.highlight_color = (c[0]/255, c[1]/255, c[2]/255)

    def load_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsm *.xlsx")])
        if not path: return
        xl = pd.ExcelFile(path)
        sheets = [s for s in xl.sheet_names if "PIPI" in s]
        
        if not sheets:
            messagebox.showerror("Error", "No sheets found containing 'PIPI' in name!")
            return

        pop = tk.Toplevel(self.root); pop.title("Select Sheets"); pop.grab_set()
        lb = tk.Listbox(pop, selectmode="multiple", width=50, height=10); [lb.insert(tk.END, s) for s in sheets]; lb.pack(padx=20, pady=10)

        def confirm():
            self.full_results = []
            for i in lb.curselection():
                sn = lb.get(i)
                df = pd.read_excel(path, sheet_name=sn)
                
                # Logică de detecție coloane
                col_map = {"tag": 0, "desc": 2, "qty": 3} # Defaults
                for idx, col in enumerate(df.columns):
                    c_name = str(col).upper()
                    if any(x in c_name for x in ["TAG", "SPOOL", "IDENTIFIER"]): col_map["tag"] = idx
                    if "DESC" in c_name: col_map["desc"] = idx
                    if "QTY" in c_name: col_map["qty"] = idx

                for _, row in df.iterrows():
                    val = str(row.iloc[col_map["tag"]]).strip()
                    if val and val != "nan" and "TOTAL" not in val.upper():
                        q_val = row.iloc[col_map["qty"]]
                        self.full_results.append({
                            "sheet": sn, "term": val, "desc": str(row.iloc[col_map["desc"]]), 
                            "target": int(q_val) if pd.notnull(q_val) and str(q_val).isdigit() else 1,
                            "hits": 0, "pages": [], "verdict": "Pending"
                        })
            self.refresh_table()
            if self.pdf_path: self.run_btn.config(state=tk.NORMAL)
            pop.destroy()
            messagebox.showinfo("Success", f"Loaded {len(self.full_results)} items.")

        tk.Button(pop, text="Confirm Selection", command=confirm, bg="#27ae60", fg="white").pack(pady=10)

    def load_pdf(self):
        self.pdf_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if self.pdf_path and self.full_results: self.run_btn.config(state=tk.NORMAL)

    def refresh_table(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        for item in self.full_results:
            self.tree.insert("", "end", values=(item["sheet"], item["term"], item["desc"], item["target"], item["hits"], item["verdict"], ""))

    def start_thread(self):
        self.run_btn.config(state=tk.DISABLED)
        threading.Thread(target=self.process, daemon=True).start()

    def process(self):
        try:
            doc = fitz.open(self.pdf_path)
            excl = set()
            raw = self.exclude_entry.get().replace(" ", "")
            if raw:
                for p in raw.split(","):
                    try:
                        if "-" in p:
                            s, e = map(int, p.split("-")); [excl.add(x-1) for x in range(s, e+1)]
                        else: excl.add(int(p)-1)
                    except: pass

            self.progress["maximum"] = len(self.full_results)
            for i, item in enumerate(self.full_results):
                count, pgs = 0, []
                for p_idx in range(len(doc)):
                    if p_idx in excl: continue
                    page = doc[p_idx]
                    m = page.search_for(item["term"])
                    if m:
                        count += len(m); pgs.append(p_idx+1)
                        for r in m:
                            try:
                                a = page.add_highlight_annot(r)
                                a.set_colors(stroke=self.highlight_color); a.update()
                            except: continue
                item["hits"], item["pages"] = count, list(set(pgs))
                item["verdict"] = "✅ MATCH" if count == item["target"] else f"❌ {count}/{item['target']}"
                self.progress["value"] = i+1; self.root.update_idletasks()
                if i % 5 == 0: self.refresh_table()

            out_dir = os.path.dirname(self.pdf_path)
            base = f"{self.base_name.get()}{self.suffix_name.get()}"
            doc.save(os.path.join(out_dir, f"{base}.pdf"))
            pd.DataFrame(self.full_results).to_excel(os.path.join(out_dir, f"{base}.xlsx"), index=False)
            messagebox.showinfo("Success", f"Files saved as: {base}")
        except Exception as e: messagebox.showerror("Error", str(e))
        finally: self.run_btn.config(state=tk.NORMAL)

if __name__ == "__main__":
    root = tk.Tk(); AuditAppBB(root); root.mainloop()
