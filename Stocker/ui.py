import customtkinter as ctk
from tkinter import ttk, messagebox, filedialog
import tkinter as tk
import os
from database import Database
from pdf_manager import PDFGenerator

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.db = Database()
        self.pdf_gen = PDFGenerator()
        self.title("TEKSTİL ERP - FOB/EAN EDITION")
        self.geometry("1400x900")

        if not os.path.exists("urun_gorselleri"): os.makedirs("urun_gorselleri")
        self.selected_image_path = None;
        self.invoice_cart = [];
        self.editing_product_id = None;
        self.active_account_id = None;
        self.account_map = {}

        self.grid_columnconfigure(1, weight=1);
        self.grid_rowconfigure(0, weight=1)
        self.setup_sidebar()
        self.setup_frames()
        self.show_frame("product")

    def setup_sidebar(self):
        self.sidebar = ctk.CTkFrame(self, width=200, corner_radius=0);
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        ctk.CTkLabel(self.sidebar, text="ERP SİSTEMİ", font=("Arial", 20, "bold")).pack(pady=30)
        btns = [("Ürün & Maliyet", "product"), ("Faturalar", "invoice"), ("Cari Hesaplar", "account"),
                ("📦 STOK RAPORU", "stock"), ("💰 FİNANS RAPORU", "finance")]
        for text, mode in btns:
            fg = "blue" if "RAPOR" not in text else "#2980B9"
            ctk.CTkButton(self.sidebar, text=text, height=40, fg_color=fg,
                          command=lambda m=mode: self.show_frame(m)).pack(pady=10, padx=20)

    def setup_frames(self):
        self.frames = {
            "product": self.create_product_frame(),
            "invoice": self.create_invoice_frame(),
            "account": self.create_account_frame(),
            "stock": self.create_stock_frame(),
            "finance": self.create_finance_frame()
        }

    def show_frame(self, name):
        for f in self.frames.values(): f.grid_forget()
        self.frames[name].grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
        if name == "product": self.refresh_prod_list(); self.update_product_combos()
        if name == "invoice": self.refresh_invoice_history(); self.update_account_combo()
        if name == "stock": self.refresh_stock_report()
        if name == "finance": self.refresh_finance_report()
        if name == "account": self.refresh_account_list()

    # --- 1. ÜRÜN EKRANI ---
    def create_product_frame(self):
        f = ctk.CTkFrame(self, fg_color="transparent")
        col_info = ctk.CTkFrame(f);
        col_info.pack(side="left", fill="both", expand=True, padx=5)
        col_cost = ctk.CTkFrame(f);
        col_cost.pack(side="left", fill="both", expand=True, padx=5)
        col_list = ctk.CTkFrame(f);
        col_list.pack(side="right", fill="both", expand=True, padx=5)

        ctk.CTkLabel(col_info, text="ÜRÜN TANIMLAMA").pack(pady=5)
        brand_frame = ctk.CTkFrame(col_info, fg_color="transparent");
        brand_frame.pack(fill="x", pady=2)
        self.cb_brand = ctk.CTkComboBox(brand_frame, width=200, values=["Seçiniz"]);
        self.cb_brand.pack(side="left", padx=5, fill="x", expand=True)
        ctk.CTkButton(brand_frame, text="+", width=30, command=self.open_brand_manager, fg_color="#E67E22").pack(
            side="left")
        self.e_name = ctk.CTkEntry(col_info, placeholder_text="Model Adı");
        self.e_name.pack(fill="x", pady=2, padx=5)
        self.cb_supp = ctk.CTkComboBox(col_info, values=["-"]);
        self.cb_supp.pack(fill="x", pady=2, padx=5)
        self.e_fabric = ctk.CTkEntry(col_info, placeholder_text="Kumaş");
        self.e_fabric.pack(fill="x", pady=2, padx=5)

        self.cat_sub_map = {"Apparel": ["Tshirt", "Pants", "Shirt", "Sweatshirt", "Jacket", "Hoodie"],
                            "Accessories": ["Bag", "Cap", "Bandana", "Socks", "Belt"],
                            "Shoes": ["Sneakers", "Classic", "Boots"]}
        self.cb_seg = ctk.CTkComboBox(col_info, values=["Fresh", "Ready"]);
        self.cb_seg.pack(fill="x", pady=2, padx=5)
        self.cb_cat = ctk.CTkComboBox(col_info, values=list(self.cat_sub_map.keys()), command=self.update_sub_cats);
        self.cb_cat.pack(fill="x", pady=2, padx=5)
        self.cb_sub = ctk.CTkComboBox(col_info, values=["-"]);
        self.cb_sub.pack(fill="x", pady=2, padx=5)
        self.cb_gen = ctk.CTkComboBox(col_info, values=["Man", "Woman", "Unisex"]);
        self.cb_gen.pack(fill="x", pady=2, padx=5)
        self.e_colors = ctk.CTkEntry(col_info, placeholder_text="Renkler");
        self.e_colors.pack(fill="x", pady=2, padx=5)
        self.e_sizes = ctk.CTkEntry(col_info, placeholder_text="Bedenler");
        self.e_sizes.pack(fill="x", pady=2, padx=5)
        ctk.CTkButton(col_info, text="Resim Seç", command=self.sel_img).pack(pady=5)

        self.default_values = {"e_fob": "0", "e_freight": "0.5", "e_local": "0.2", "e_test": "0.1", "e_parity": "1.1"}
        self.entries = {}
        for txt, key in [("A) FOB:", "e_fob"), ("B) Freight:", "e_freight"), ("D) Local:", "e_local"),
                         ("F) Test:", "e_test"), ("I) Parity:", "e_parity")]:
            fr = ctk.CTkFrame(col_cost);
            fr.pack(fill="x")
            ctk.CTkLabel(fr, text=txt, width=80).pack(side="left");
            e = ctk.CTkEntry(fr, width=80);
            e.pack(side="left");
            e.insert(0, self.default_values[key]);
            self.entries[key] = e

        ctk.CTkButton(col_cost, text="HESAPLA", command=self.calc_cost).pack(pady=5)
        self.lbl_res_eur = ctk.CTkLabel(col_cost, text="CPT EUR: 0.00", text_color="#F1C40F",
                                        font=("Arial", 16, "bold"));
        self.lbl_res_eur.pack()
        sf = ctk.CTkFrame(col_cost);
        sf.pack(pady=5)
        self.e_sales = ctk.CTkEntry(sf, width=80);
        self.e_sales.pack(side="left")
        self.cb_curr = ctk.CTkComboBox(sf, values=["EUR", "USD"], width=60);
        self.cb_curr.pack(side="left")
        self.btn_save_prod = ctk.CTkButton(col_cost, text="KAYDET", command=self.save_product, fg_color="green");
        self.btn_save_prod.pack(pady=10, fill="x", padx=10)
        ctk.CTkButton(col_cost, text="TEMİZLE / YENİ", command=self.clear_form, fg_color="orange").pack(fill="x",
                                                                                                        padx=10)

        ctk.CTkLabel(col_list, text="LİSTE").pack(pady=5)
        # TABLO DÜZELTİLDİ: ID - MODEL - CPT - RETAIL
        self.tree_prod = ttk.Treeview(col_list, columns=('ID', 'Model', 'FOB', 'CPT', 'Retail'), show='headings')
        self.tree_prod.heading('ID', text='ID');
        self.tree_prod.column('ID', width=30)
        self.tree_prod.heading('Model', text='Model');
        self.tree_prod.column('Model', width=100)
        self.tree_prod.heading('FOB', text='FOB ($)');
        self.tree_prod.column('FOB', width=60)
        self.tree_prod.heading('CPT', text='CPT (€)');
        self.tree_prod.column('CPT', width=60)
        self.tree_prod.heading('Retail', text='Retail');
        self.tree_prod.column('Retail', width=60)
        self.tree_prod.pack(fill="both", expand=True)
        self.tree_prod.bind("<Double-1>", self.load_edit)
        ctk.CTkButton(col_list, text="📥 EXCEL IMPORT", fg_color="#2980B9", command=self.import_excel).pack(fill="x",
                                                                                                           pady=10)
        return f

    def update_sub_cats(self, choice):
        self.cb_sub.configure(values=self.cat_sub_map.get(choice, ["-"]));
        self.cb_sub.set(self.cat_sub_map.get(choice, ["-"])[0])

    # --- 2. FATURA EKRANI ---
    def create_invoice_frame(self):
        f = ctk.CTkFrame(self, fg_color="transparent");
        tab = ctk.CTkTabview(f);
        tab.pack(fill="both", expand=True)
        t1 = tab.add("Yeni Fatura");
        t2 = tab.add("Geçmiş")

        h = ctk.CTkFrame(t1);
        h.pack(fill="x", pady=10)
        self.var_type = ctk.StringVar(value="ALIS")
        ctk.CTkRadioButton(h, text="ALIS", variable=self.var_type, value="ALIS",
                           command=self.update_account_combo).pack(side="left", padx=10)
        ctk.CTkRadioButton(h, text="SATIS", variable=self.var_type, value="SATIS",
                           command=self.update_account_combo).pack(side="left", padx=10)
        self.e_inv = ctk.CTkEntry(h, placeholder_text="Fatura No");
        self.e_inv.pack(side="left", padx=10)
        self.cb_acc = ctk.CTkComboBox(h, width=200);
        self.cb_acc.pack(side="left", padx=10)
        ctk.CTkButton(h, text="🔍 ÜRÜN SEÇ", command=self.open_multi_select, fg_color="#8E44AD").pack(side="left",
                                                                                                     padx=20)

        # TABLO DÜZELTİLDİ: SKU - MODEL - ADET - FİYAT - TOPLAM
        self.tree_cart = ttk.Treeview(t1, columns=('SKU', 'Model', 'Qty', 'Price', 'Tot'), show='headings')
        self.tree_cart.heading('SKU', text='SKU');
        self.tree_cart.column('SKU', width=100)
        self.tree_cart.heading('Model', text='Model');
        self.tree_cart.column('Model', width=120)
        self.tree_cart.heading('Qty', text='Adet');
        self.tree_cart.column('Qty', width=50)
        self.tree_cart.heading('Price', text='Fiyat');
        self.tree_cart.column('Price', width=70)
        self.tree_cart.heading('Tot', text='Toplam');
        self.tree_cart.column('Tot', width=80)
        self.tree_cart.pack(fill="both", expand=True)

        b = ctk.CTkFrame(t1);
        b.pack(fill="x", pady=10)
        self.lbl_tot = ctk.CTkLabel(b, text="TOPLAM: 0.00", font=("Arial", 18, "bold"));
        self.lbl_tot.pack(side="left", padx=20)
        ctk.CTkButton(b, text="KAYDET", command=self.save_inv, fg_color="green").pack(side="right", padx=10)
        ctk.CTkButton(b, text="TEMİZLE", command=self.clear_cart, fg_color="red").pack(side="right", padx=10)

        split = ctk.CTkFrame(t2, fg_color="transparent");
        split.pack(fill="both", expand=True)
        left = ctk.CTkFrame(split, width=400);
        left.pack(side="left", fill="y", padx=5)
        right = ctk.CTkFrame(split);
        right.pack(side="left", fill="both", expand=True, padx=5)

        ctk.CTkButton(left, text="YENİLE", command=self.refresh_invoice_history).pack(pady=5)
        self.tree_hist = ttk.Treeview(left, columns=('No', 'Tip', 'Tarih', 'Tutar'), show='headings');
        self.tree_hist.pack(fill="both", expand=True)
        for c in ('No', 'Tip', 'Tarih', 'Tutar'): self.tree_hist.heading(c, text=c)
        self.tree_hist.bind("<<TreeviewSelect>>", self.load_invoice_detail)

        act = ctk.CTkFrame(right);
        act.pack(fill="x", pady=5)
        ctk.CTkButton(act, text="✏️ DÜZENLE", command=self.edit_invoice_btn, fg_color="orange").pack(side="left",
                                                                                                     padx=5)
        ctk.CTkButton(act, text="🗑️ SİL", command=self.delete_invoice_btn, fg_color="red").pack(side="left", padx=5)
        ctk.CTkButton(act, text="📄 PDF", command=self.print_invoice_btn, fg_color="#3498DB").pack(side="left", padx=5)
        ctk.CTkButton(act, text="📤 EXCEL", command=lambda: self.export_excel("INVOICES"), fg_color="#27AE60").pack(
            side="left", padx=5)

        self.lbl_detail_title = ctk.CTkLabel(right, text="İçerik", font=("Arial", 16, "bold"));
        self.lbl_detail_title.pack(pady=10)

        # DETAY TABLOSUNDA EAN GÖRÜNECEK
        self.tree_detail = ttk.Treeview(right, columns=('SKU', 'EAN', 'Model', 'Qty', 'Price', 'Tot'), show='headings')
        self.tree_detail.heading('SKU', text='SKU');
        self.tree_detail.heading('EAN', text='EAN')
        self.tree_detail.heading('Model', text='Model');
        self.tree_detail.heading('Qty', text='Adet')
        self.tree_detail.heading('Price', text='Fiyat');
        self.tree_detail.heading('Tot', text='Toplam')
        self.tree_detail.pack(fill="both", expand=True)
        self.lbl_detail_total = ctk.CTkLabel(right, text="GENEL TOPLAM: 0.00", font=("Arial", 20, "bold"),
                                             text_color="#E67E22");
        self.lbl_detail_total.pack(pady=10)
        return f

    # --- POPUP ÜRÜN SEÇİCİ (ALIS=FOB, SATIS=RETAIL) ---
    def open_multi_select(self):
        variants = self.db.get_variants_for_combo()
        if not variants: messagebox.showwarning("Uyarı", "Ürün yok."); return
        top = ctk.CTkToplevel(self);
        top.geometry("600x500");
        top.attributes("-topmost", True);
        scroll = ctk.CTkScrollableFrame(top);
        scroll.pack(fill="both", expand=True)
        self.chks = [];
        trans_type = self.var_type.get()

        for v in variants:
            # v index: 0:sku, 1:model, 2:color, 3:size, 4:cpt_eur, 5:sales, 6:id, 7:curr_sell, 8:stock, 9:ean, 10:fob
            stk = v[8]
            if trans_type == "SATIS" and stk <= 0: continue

            # Alış Faturası ise FOB USD, Satış Faturası ise Sales Price kullanılır
            price = v[10] if trans_type == "ALIS" else v[5]

            row = ctk.CTkFrame(scroll);
            row.pack(fill="x", pady=2)
            c = ctk.BooleanVar()
            ctk.CTkCheckBox(row, text=f"{v[0]} | {v[1]} | Stok:{stk}", variable=c, width=400).pack(side="left")

            # Fiyatı göster (Alış ise FOB, Satış ise Retail)
            lbl_text = f"FOB: ${price}" if trans_type == "ALIS" else f"Retail: {price}"
            ctk.CTkLabel(row, text=lbl_text).pack(side="left", padx=5)

            q = ctk.CTkEntry(row, width=50);
            q.pack(side="left");
            q.insert(0, "1")
            self.chks.append({'c': c, 'q': q, 'd': {'sku': v[0], 'var_id': v[6], 'model': v[1], 'var': f"{v[2]}/{v[3]}",
                                                    'price': price}})
        ctk.CTkButton(top, text="EKLE", command=lambda: self.add_sel(top)).pack(pady=10)

    def add_sel(self, w):
        for i in self.chks:
            if i['c'].get():
                try:
                    q = int(i['q'].get()); d = i['d']; t = q * d['price']; self.invoice_cart.append(
                        {**d, 'qty': q, 'total': t}); self.tree_cart.insert("", "end", values=(d['sku'], d['model'], q,
                                                                                               f"{d['price']:.2f}",
                                                                                               f"{t:.2f}"))
                except:
                    pass
        self.lbl_tot.configure(text=f"TOPLAM: {sum(x['total'] for x in self.invoice_cart):.2f}");
        w.destroy()

    # --- DİĞERLERİ ---
    def create_account_frame(self):
        return self.base_account_frame()

    def create_stock_frame(self):
        return self.base_stock_frame()

    def create_finance_frame(self):
        return self.base_finance_frame()

    # Mantıksal Fonksiyonlar
    def calc_cost(self):
        try:
            fob = float(self.entries['e_fob'].get());
            par = float(self.entries['e_parity'].get())
            cpt_usd = fob + float(self.entries['e_freight'].get()) + float(self.entries['e_local'].get()) + float(
                self.entries['e_test'].get()) + (fob * 0.22)
            cpt_eur = cpt_usd / par if par else 0
            self.lbl_res_eur.configure(text=f"CPT EUR: {round(cpt_eur, 2)} €")
            return {'fob': fob, 'freight': 0, 'local': 0, 'test': 0, 'parity': par, 'vat': fob * 0.21,
                    'broker': fob * 0.005, 'bank': fob * 0.005, 'cpt_usd': round(cpt_usd, 2),
                    'cpt_eur': round(cpt_eur, 2)}
        except:
            return None

    def save_product(self):
        c = self.calc_cost();
        if not c: return
        data = {'name': self.e_name.get(), 'brand': self.cb_brand.get(), 'supp': self.cb_supp.get(),
                'fabric': self.e_fabric.get(), 'segment': self.cb_seg.get(), 'cat': self.cb_cat.get(),
                'sub_cat': self.cb_sub.get(), 'gender': self.cb_gen.get(), 'sales': float(self.e_sales.get() or 0),
                'curr_sell': self.cb_curr.get(), **c}
        if self.editing_product_id:
            self.db.update_product(self.editing_product_id, data); messagebox.showinfo("OK", "Güncellendi")
        else:
            self.db.add_product(data, [
                {'color': co.strip(), 'size': si.strip(), 'ean': '', 'img_path': self.selected_image_path or ''} for co
                in self.e_colors.get().split(',') for si in self.e_sizes.get().split(',')]); messagebox.showinfo("OK",
                                                                                                                 "Kaydedildi")
        self.refresh_prod_list();
        self.clear_form()

    def clear_form(self):
        self.editing_product_id = None;
        self.e_name.delete(0, tk.END);
        self.e_colors.delete(0, tk.END);
        self.e_sizes.delete(0, tk.END)
        for key, val in self.default_values.items(): self.entries[key].delete(0, tk.END); self.entries[key].insert(0,
                                                                                                                   val)
        self.btn_save_prod.configure(text="KAYDET", fg_color="green")

    def load_edit(self, e):
        s = self.tree_prod.selection();
        if not s: return
        pid = self.tree_prod.item(s[0], "values")[0];
        self.editing_product_id = pid;
        d = self.db.get_product_details(pid)
        self.e_name.delete(0, tk.END);
        self.e_name.insert(0, d[1]);
        self.cb_brand.set(d[2]);
        self.cb_supp.set(d[7]);
        self.entries['e_fob'].delete(0, tk.END);
        self.entries['e_fob'].insert(0, d[9])
        self.calc_cost();
        self.btn_save_prod.configure(text="GÜNCELLE", fg_color="#D35400")

    def open_brand_manager(self):
        t = ctk.CTkToplevel(self); e = ctk.CTkEntry(t); e.pack(pady=10); ctk.CTkButton(t, text="EKLE", command=lambda: [
            self.db.add_brand(e.get()), self.update_product_combos(), t.destroy()]).pack()

    def import_excel(self):
        p = filedialog.askopenfilename(); self.db.import_products_from_excel(p) if p else None; self.refresh_prod_list()

    def sel_img(self):
        self.selected_image_path = filedialog.askopenfilename()

    def update_account_combo(self):
        v = [r[1] for r in
             self.db.get_accounts("SUPPLIER" if self.var_type.get() == "ALIS" else "CUSTOMER")]; self.account_map = {
            r[1]: r[0] for r in
            self.db.get_accounts("SUPPLIER" if self.var_type.get() == "ALIS" else "CUSTOMER")}; self.cb_acc.configure(
            values=v)

    def save_inv(self):
        if not self.cb_acc.get() or self.cb_acc.get() not in self.account_map: messagebox.showerror("Hata",
                                                                                                    "Cari Seçin"); return
        self.db.save_invoice_items(self.e_inv.get(), self.var_type.get(), self.invoice_cart,
                                   self.account_map.get(self.cb_acc.get()));
        self.invoice_cart = [];
        self.clear_cart();
        self.refresh_invoice_history();
        messagebox.showinfo("OK", "Kaydedildi")

    def clear_cart(self):
        self.invoice_cart = []; [self.tree_cart.delete(i) for i in
                                 self.tree_cart.get_children()]; self.lbl_tot.configure(text="TOPLAM: 0.00")

    def delete_invoice_btn(self):
        sel = self.tree_hist.selection();
        inv = self.tree_hist.item(sel[0], "values")[0] if sel else None
        if inv and messagebox.askyesno("Sil", "Emin misin?"): self.db.delete_invoice(
            inv); self.refresh_invoice_history()

    def edit_invoice_btn(self):
        sel = self.tree_hist.selection();
        inv = self.tree_hist.item(sel[0], "values")[0] if sel else None
        if inv:
            self.clear_cart();
            self.e_inv.delete(0, tk.END);
            self.e_inv.insert(0, inv);
            acc_id = self.db.get_invoice_account_id(inv)
            if acc_id:
                for n, i in self.account_map.items():
                    if i == acc_id: self.cb_acc.set(n)
            for d in self.db.get_invoice_details(inv): self.invoice_cart.append(
                {'sku': d[0], 'model': d[1], 'var': f"{d[3]}/{d[4]}", 'qty': d[5], 'price': d[6], 'total': d[7],
                 'var_id': d[8]}); self.tree_cart.insert("", "end",
                                                         values=(d[0], d[1], d[5], f"{d[6]:.2f}", f"{d[7]:.2f}"))
            self.lbl_tot.configure(text=f"TOPLAM: {sum(x['total'] for x in self.invoice_cart):.2f}");
            self.tab_inv.set("Yeni Fatura")

    def print_invoice_btn(self):
        sel = self.tree_hist.selection();
        inv = self.tree_hist.item(sel[0], "values")[0] if sel else None
        if inv:
            items = [{'sku': d[0], 'model': d[1], 'var': f"{d[3]}/{d[4]}", 'qty': d[5], 'price': d[6], 'total': d[7]}
                     for d in self.db.get_invoice_details(inv)]
            f = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile=f"Invoice_{inv}.pdf")
            if f: self.pdf_gen.create_invoice(f, inv, self.tree_hist.item(sel[0], "values")[1],
                                              items); messagebox.showinfo("OK", "PDF Oluşturuldu")

    def refresh_prod_list(self):
        [self.tree_prod.delete(i) for i in self.tree_prod.get_children()]; [
            self.tree_prod.insert("", "end", values=(r[0], r[1], f"{r[2]:.2f} $", f"{r[3]:.2f} €", f"{r[4]:.2f}")) for r
            in self.db.get_all_products_simple()]

    def refresh_invoice_history(self):
        [self.tree_hist.delete(i) for i in self.tree_hist.get_children()]; [
            self.tree_hist.insert("", "end", values=(r[0], r[2], r[1], f"{r[3]:.2f}")) for r in
            self.db.get_invoice_history()]

    def refresh_stock_report(self):
        [self.tree_stock.delete(i) for i in self.tree_stock.get_children()]; [self.tree_stock.insert("", "end",
                                                                                                     values=(r['sku'],
                                                                                                             r[
                                                                                                                 'ean_code'],
                                                                                                             r[
                                                                                                                 'model_name'],
                                                                                                             r[
                                                                                                                 'total_in'],
                                                                                                             r[
                                                                                                                 'total_out'],
                                                                                                             r[
                                                                                                                 'total_in'] -
                                                                                                             r[
                                                                                                                 'total_out']))
                                                                              for i, r in
                                                                              self.db.get_stock_only_report().iterrows()]

    def refresh_finance_report(self):
        [self.tree_fin.delete(i) for i in
         self.tree_fin.get_children()]; df = self.db.get_finance_only_report(); net = 0; [
            self.tree_fin.insert("", "end",
                                 values=(r['sku'], r['model_name'], f"{r['total_cost']:.2f}", f"{r['total_sales']:.2f}",
                                         f"{r['total_sales'] - r['total_cost']:.2f}")) for i, r in
            df.iterrows()]; self.lbl_fin_net.configure(text=f"NET: {sum(df['total_sales'] - df['total_cost']):.2f}")

    def refresh_account_list(self):
        [self.tree_acc.delete(i) for i in self.tree_acc.get_children()]; [
            self.tree_acc.insert("", "end", values=(r[0], r[1], r[2], self.db.get_account_balance(r[0]))) for r in
            self.db.get_accounts()]

    def update_product_combos(self):
        self.cb_brand.configure(values=self.db.get_brands() or ["-"]); self.cb_supp.configure(
            values=self.db.get_suppliers_list() or ["-"])

    def save_acc(self):
        self.db.add_account(self.e_an.get(), self.cb_at.get(), self.e_ac.get(), "-"); self.refresh_account_list()

    def load_acc(self, e):
        s = self.tree_acc.selection(); v = self.tree_acc.item(s[0],
                                                              "values") if s else None; self.lbl_acc_title.configure(
            text=v[1]) if v else None; self.active_account_id = v[0] if v else None; self.lbl_bal.configure(
            text=v[3]) if v else None; [self.tree_led.delete(i) for i in self.tree_led.get_children()]; [
            self.tree_led.insert("", "end", values=r) for r in self.db.get_account_statement(v[0])] if v else None

    def add_trans(self, t, m):
        self.db.add_financial_transaction(self.active_account_id, "PAYMENT", "Nakit",
                                          float(self.e_pay.get()) * m); self.load_acc(None); self.refresh_account_list()

    def load_invoice_detail(self, e):
        sel = self.tree_hist.selection();
        inv = self.tree_hist.item(sel[0], "values")[0] if sel else None;
        self.lbl_detail_title.configure(text=f"İçerik: {inv}");
        [self.tree_detail.delete(i) for i in self.tree_detail.get_children()];
        gt = 0
        for d in self.db.get_invoice_details(inv): self.tree_detail.insert("", "end", values=(d[0], d[9], d[1], d[5],
                                                                                              f"{d[6]:.2f}",
                                                                                              f"{d[7]:.2f}")); gt += d[
            7]
        self.lbl_detail_total.configure(text=f"GENEL TOPLAM: {gt:.2f}")

    def export_excel(self, t):
        p = filedialog.asksaveasfilename(defaultextension=".xlsx"); self.db.export_data_to_excel(t, p) if p else None

    # Base Frames
    def base_account_frame(self):
        f = ctk.CTkFrame(self, fg_color="transparent");
        l = ctk.CTkFrame(f, width=350);
        l.pack(side="left", fill="y", padx=10)
        self.tree_acc = ttk.Treeview(l, columns=('ID', 'Ad', 'Tip', 'Bal'), show='headings');
        self.tree_acc.pack(fill="both", expand=True);
        self.tree_acc.bind("<<TreeviewSelect>>", self.load_acc)
        ctk.CTkButton(l, text="EXCEL EXPORT", command=lambda: self.export_excel("ACCOUNTS")).pack(fill="x", pady=5)
        self.e_an = ctk.CTkEntry(l, placeholder_text="Ad");
        self.e_an.pack(fill="x");
        self.cb_at = ctk.CTkComboBox(l, values=["CUSTOMER", "SUPPLIER"]);
        self.cb_at.pack(fill="x");
        self.e_ac = ctk.CTkEntry(l, placeholder_text="İletişim");
        self.e_ac.pack(fill="x");
        ctk.CTkButton(l, text="EKLE", command=self.save_acc, fg_color="green").pack(fill="x", pady=5)
        r = ctk.CTkFrame(f);
        r.pack(side="left", fill="both", expand=True, padx=10);
        self.lbl_acc_title = ctk.CTkLabel(r, text="-", font=("Arial", 18));
        self.lbl_acc_title.pack(pady=10);
        self.lbl_bal = ctk.CTkLabel(r, text="0.00", font=("Arial", 24), text_color="orange");
        self.lbl_bal.pack();
        act = ctk.CTkFrame(r);
        act.pack(pady=10);
        self.e_pay = ctk.CTkEntry(act, width=100);
        self.e_pay.pack(side="left");
        ctk.CTkButton(act, text="TAHSİLAT", command=lambda: self.add_trans("P", -1), fg_color="#27AE60").pack(
            side="left", padx=5);
        ctk.CTkButton(act, text="ÖDEME", command=lambda: self.add_trans("P", -1), fg_color="#C0392B").pack(side="left",
                                                                                                           padx=5);
        self.tree_led = ttk.Treeview(r, columns=('Tarih', 'İşlem', 'Açıklama', 'Tutar'), show='headings');
        self.tree_led.pack(fill="both", expand=True);
        return f

    def base_stock_frame(self):
        f = ctk.CTkFrame(self, fg_color="transparent"); ctk.CTkButton(f, text="EXCEL EXPORT",
                                                                      command=lambda: self.export_excel("STOCK")).pack(
            pady=10); self.tree_stock = ttk.Treeview(f, columns=('SKU', 'EAN', 'Model', 'Alinan', 'Cikan', 'Kalan'),
                                                     show='headings'); self.tree_stock.heading('EAN',
                                                                                               text='EAN'); self.tree_stock.pack(
            fill="both", expand=True); ctk.CTkButton(f, text="YENİLE",
                                                     command=self.refresh_stock_report).pack(); return f

    def base_finance_frame(self):
        f = ctk.CTkFrame(self, fg_color="transparent"); ctk.CTkButton(f, text="EXCEL EXPORT",
                                                                      command=lambda: self.export_excel(
                                                                          "FINANCE")).pack(
            pady=10); self.tree_fin = ttk.Treeview(f, columns=('SKU', 'Model', 'Maliyet', 'Satis', 'KarZarar'),
                                                   show='headings'); self.tree_fin.pack(fill="both",
                                                                                        expand=True); ctk.CTkButton(f,
                                                                                                                    text="YENİLE",
                                                                                                                    command=self.refresh_finance_report).pack(); self.lbl_fin_net = ctk.CTkLabel(
            f, text="NET: 0", font=("Arial", 16, "bold")); self.lbl_fin_net.pack(); return f


if __name__ == "__main__":
    app = App()
    app.mainloop()