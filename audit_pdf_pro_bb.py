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
        self.root.geometry("1450x850")
        self.root.configure(bg="#f0f2f5")

        self.full_results = []
        self.pdf_path = ""
        self.highlight_color = (1, 1, 0) 

        # --- PANEL CONTROL ---
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
        self.base_name = tk.Entry(name_frame, width=40); self.base_name.insert(0, "Audit_Result"); self.base_name.pack(side=tk.LEFT)
        self.suffix_name = tk.Entry(name_frame, width=15); self.suffix_name.insert(0, "_REV00"); self.suffix_name.pack(side=tk.LEFT)

        # --- TABEL ---
        self.tree_frame = tk.Frame(root)
        self.tree_frame.pack(pady=10, fill=tk.BOTH, expand=True, padx=20)
        
        cols = ("Sheet", "Identifier", "Description", "QTY_BOM", "Found", "Verdict", "Pages")
        self.tree = ttk.Treeview(self.tree_frame, columns=cols, show='headings')
        
        cw = {"Sheet": 100, "Identifier": 180, "Description": 350, "QTY_BOM": 80, "Found": 80, "Verdict": 100, "Pages": 300}
        for c, w in cw.items():
            self.tree.heading(c, text=c)
            self.tree.column(c, width=w, anchor=tk.W if "Desc" in c or "Iden" in c else tk.CENTER)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        vsb.pack(side=tk.RIGHT, fill=tk.Y); self.tree.configure(yscrollcommand=vsb.set)
        
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
        lb = tk.Listbox(pop, selectmode="multiple", width=60, height=12); [lb.insert(tk.END, s) for s in sheets]; lb.pack(padx=20, pady=10)

        def confirm():
            self.full_results = []
            for i in lb.curselection():
                sn = lb.get(i)
                df = pd.read_excel(path, sheet_name=sn, header=1)
                
                col_tag = next((c for c in df.columns if any(x in str(c).upper() for x in ["TAG", "SPOOL", "ITEM"])), df.columns[0])
                col_desc = next((c for c in df.columns if "DESC" in str(c).upper()), df.columns[min(2, len(df.columns)-1)])
                col_qty = next((c for c in df.columns if "QTY" in str(c).upper()), df.columns[min(3, len(df.columns)-1)])

                for _, row in df.iterrows():
                    val = str(row[col_tag]).strip()
                    if val and val != "nan" and "TOTAL" not in val.upper():
                        desc_val = str(row[col_desc]) if pd.notnull(row[col_desc]) else "-"
                        qty_val = row[col_qty]
                        try: target = int(qty_val) if pd.notnull(qty_val) else 1
                        except: target = 1
                        
                        self.full_results.append({
                            "sheet": sn, "term": val, "desc": desc_val, "target": target, 
                            "hits": 0, "pages": [], "verdict": "Pending"
                        })
            self.refresh_table()
            if self.pdf_path: self.run_btn.config(state=tk.NORMAL)
            pop.destroy()
            messagebox.showinfo("Success", f"Loaded {len(self.full_results)} items.")

        tk.Button(pop, text="Confirm & Load Data", command=confirm, bg="#27ae60", fg="white", pady=5, width=20).pack(pady=10)

    def load_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if path:
            self.pdf_path = path
            # --- AICI PRELUĂM NUMELE DIN RĂDĂCINĂ ---
            filename_root = os.path.splitext(os.path.basename(path))[0]
            self.base_name.delete(0, tk.END)
            self.base_name.insert(0, f"{filename_root}_AUDIT") # Pune numele PDF-ului + sufix
            
            if self.full_results: self.run_btn.config(state=tk.NORMAL)

    def refresh_table(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        for item in self.full_results:
            pg_str = ", ".join(map(str, item["pages"]))
            self.tree.insert("", "end", values=(item["sheet"], item["term"], item["desc"], item["target"], item["hits"], item["verdict"], pg_str))

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
                        count += len(m)
                        pgs.append(p_idx + 1)
                        for r in m:
                            try:
                                a = page.add_highlight_annot(r)
                                a.set_colors(stroke=self.highlight_color); a.update()
                            except: continue
                
                item["hits"], item["pages"] = count, sorted(list(set(pgs)))
                item["verdict"] = "✅ OK" if count == item["target"] else f"❌ {count}/{item['target']}"
                
                self.progress["value"] = i+1
                if i % 5 == 0: self.root.after(0, self.refresh_table)
            
            out_dir = os.path.dirname(self.pdf_path)
            base = f"{self.base_name.get()}{self.suffix_name.get()}"
            
            doc.save(os.path.join(out_dir, f"{base}.pdf"))
            
            df_export = pd.DataFrame(self.full_results)
            df_export["pages"] = df_export["pages"].apply(lambda x: ", ".join(map(str, x)))
            df_export.to_excel(os.path.join(out_dir, f"{base}.xlsx"), index=False)
            
            self.root.after(0, lambda: messagebox.showinfo("Success", f"Files saved in the same folder as the PDF!"))
        except Exception as e: 
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
        finally:
            self.root.after(0, lambda: self.run_btn.config(state=tk.NORMAL))
            self.root.after(0, self.refresh_table)

if __name__ == "__main__":
    root = tk.Tk(); AuditAppBB(root); root.mainloop()
