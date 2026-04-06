import sqlite3
import pandas as pd
from datetime import datetime
import random


class Database:
    def __init__(self, db_name="tekstil_erp_v16.db"):
        self.conn = sqlite3.connect(db_name)
        self.cursor = self.conn.cursor()
        self.create_tables()

    def create_tables(self):
        self.cursor.execute(
            """CREATE TABLE IF NOT EXISTS products (id INTEGER PRIMARY KEY AUTOINCREMENT, model_name TEXT, brand TEXT, segment TEXT, category TEXT, sub_category TEXT, gender TEXT, supplier TEXT, fabric_info TEXT, cost_fob REAL, cost_freight REAL, cost_vat REAL, cost_local_transport REAL, cost_broker REAL, cost_test REAL, cost_bank REAL, cost_cpt_usd REAL, parity_eur_usd REAL, cost_cpt_eur REAL, sales_price REAL, currency_sell TEXT)""")
        self.cursor.execute(
            """CREATE TABLE IF NOT EXISTS variants (id INTEGER PRIMARY KEY AUTOINCREMENT, product_id INTEGER, sku TEXT, color TEXT, size TEXT, ean_code TEXT, image_path TEXT, FOREIGN KEY(product_id) REFERENCES products(id))""")
        self.cursor.execute(
            """CREATE TABLE IF NOT EXISTS transactions (id INTEGER PRIMARY KEY AUTOINCREMENT, variant_id INTEGER, invoice_no TEXT, trans_type TEXT, quantity INTEGER, unit_price REAL, date TEXT, FOREIGN KEY(variant_id) REFERENCES variants(id))""")
        self.cursor.execute(
            """CREATE TABLE IF NOT EXISTS accounts (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, acc_type TEXT, contact TEXT, address TEXT, balance REAL DEFAULT 0)""")
        self.cursor.execute(
            """CREATE TABLE IF NOT EXISTS ledgers (id INTEGER PRIMARY KEY AUTOINCREMENT, account_id INTEGER, date TEXT, doc_type TEXT, description TEXT, amount REAL, FOREIGN KEY(account_id) REFERENCES accounts(id))""")
        self.cursor.execute(
            """CREATE TABLE IF NOT EXISTS brands (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE)""")
        self.conn.commit()

    def generate_ean13(self):
        prefix = "869"
        body = "".join([str(random.randint(0, 9)) for _ in range(9)])
        code_12 = prefix + body
        total = 0
        for i, digit in enumerate(code_12):
            n = int(digit)
            total += n * 1 if i % 2 == 0 else n * 3
        remainder = total % 10
        check_digit = 0 if remainder == 0 else 10 - remainder
        return code_12 + str(check_digit)

    def add_product(self, d, vl):
        self.cursor.execute(
            """INSERT INTO products (model_name, brand, segment, category, sub_category, gender, supplier, fabric_info, cost_fob, cost_freight, cost_vat, cost_local_transport, cost_broker, cost_test, cost_bank, cost_cpt_usd, parity_eur_usd, cost_cpt_eur, sales_price, currency_sell) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (d['name'], d['brand'], d['segment'], d['cat'], d['sub_cat'], d['gender'], d['supp'], d['fabric'], d['fob'],
             d['freight'], d['vat'], d['local'], d['broker'], d['test'], d['bank'], d['cpt_usd'], d['parity'],
             d['cpt_eur'], d['sales'], d['curr_sell']))
        pid = self.cursor.lastrowid
        for var in vl:
            sku = f"{d['brand'][:3]}-{d['name']}-{var['color']}-{var['size']}".upper().replace(" ", "")
            ean = self.generate_ean13()
            self.cursor.execute(
                "INSERT INTO variants (product_id, sku, color, size, ean_code, image_path) VALUES (?,?,?,?,?,?)",
                (pid, sku, var['color'], var['size'], ean, var['img_path']))
        self.conn.commit()

    def update_product(self, pid, d):
        self.cursor.execute(
            "UPDATE products SET model_name=?, brand=?, segment=?, category=?, sub_category=?, gender=?, supplier=?, fabric_info=?, cost_fob=?, cost_freight=?, cost_vat=?, cost_local_transport=?, cost_broker=?, cost_test=?, cost_bank=?, cost_cpt_usd=?, parity_eur_usd=?, cost_cpt_eur=?, sales_price=?, currency_sell=? WHERE id=?",
            (d['name'], d['brand'], d['segment'], d['cat'], d['sub_cat'], d['gender'], d['supp'], d['fabric'], d['fob'],
             d['freight'], d['vat'], d['local'], d['broker'], d['test'], d['bank'], d['cpt_usd'], d['parity'],
             d['cpt_eur'], d['sales'], d['curr_sell'], pid))
        self.conn.commit()

    def get_all_products_simple(self):
        # UI Güncellemesi: FOB USD'yi de gösterelim listede
        self.cursor.execute("SELECT id, model_name, cost_fob, cost_cpt_eur, sales_price FROM products")
        return self.cursor.fetchall()

    def get_product_details(self, pid):
        self.cursor.execute("SELECT * FROM products WHERE id=?", (pid,)); return self.cursor.fetchone()

    def get_variants_for_combo(self):
        # KRİTİK GÜNCELLEME: En sona 'p.cost_fob' eklendi (Index 10)
        query = """
            SELECT v.sku, p.model_name, v.color, v.size, p.cost_cpt_eur, p.sales_price, v.id, p.currency_sell,
            (SELECT COALESCE(SUM(CASE WHEN trans_type='ALIS' THEN quantity ELSE -quantity END), 0)
             FROM transactions t WHERE t.variant_id = v.id) as current_stock,
            v.ean_code,
            p.cost_fob
            FROM variants v JOIN products p ON v.product_id = p.id
        """
        self.cursor.execute(query)
        return self.cursor.fetchall()

    def save_invoice_items(self, invoice_no, trans_type, cart_items, account_id):
        self.delete_invoice(invoice_no);
        date = datetime.now().strftime("%Y-%m-%d %H:%M:%S");
        total = sum([item['total'] for item in cart_items])
        for item in cart_items: self.cursor.execute(
            "INSERT INTO transactions (variant_id, invoice_no, trans_type, quantity, unit_price, date) VALUES (?,?,?,?,?,?)",
            (item['var_id'], invoice_no, trans_type, item['qty'], item['price'], date))
        self.cursor.execute("INSERT INTO ledgers (account_id, date, doc_type, description, amount) VALUES (?,?,?,?,?)",
                            (account_id, date, "INVOICE", f"Fatura: {invoice_no}", total));
        self.conn.commit()

    def delete_invoice(self, invoice_no):
        try:
            self.cursor.execute("DELETE FROM transactions WHERE invoice_no = ?", (invoice_no,)); self.cursor.execute(
                "DELETE FROM ledgers WHERE description LIKE ?", (f"%{invoice_no}%",)); self.conn.commit(); return True
        except:
            return False

    def get_invoice_history(self):
        self.cursor.execute(
            "SELECT invoice_no, date, trans_type, SUM(quantity * unit_price), COUNT(*) FROM transactions GROUP BY invoice_no ORDER BY date DESC"); return self.cursor.fetchall()

    def get_invoice_details(self, invoice_no):
        # Fatura detayında EAN gösteriliyor
        self.cursor.execute(
            """SELECT v.sku, p.model_name, p.brand, v.color, v.size, t.quantity, t.unit_price, (t.quantity * t.unit_price), t.variant_id, v.ean_code FROM transactions t JOIN variants v ON t.variant_id = v.id JOIN products p ON v.product_id = p.id WHERE t.invoice_no = ?""",
            (invoice_no,))
        return self.cursor.fetchall()

    def get_invoice_account_id(self, invoice_no):
        self.cursor.execute("SELECT account_id FROM ledgers WHERE description LIKE ?",
                            (f"%{invoice_no}%",)); res = self.cursor.fetchone(); return res[0] if res else None

    # --- RAPORLAR ---
    # Stok raporunda hala CPT EUR kullanılır (Envanter Değeri)
    def get_detailed_stock_report(self):
        return pd.read_sql_query(
            """SELECT p.brand, p.model_name, v.sku, v.color, v.size, (SELECT COALESCE(SUM(CASE WHEN trans_type='ALIS' THEN quantity ELSE -quantity END), 0) FROM transactions t WHERE t.variant_id = v.id) as current_stock, p.cost_cpt_eur, p.sales_price, p.currency_sell FROM variants v JOIN products p ON v.product_id = p.id""",
            self.conn)

    def get_stock_only_report(self):
        return pd.read_sql_query(
            "SELECT v.sku, v.ean_code, p.model_name, (SELECT COALESCE(SUM(quantity), 0) FROM transactions t WHERE t.variant_id = v.id AND t.trans_type = 'ALIS') as total_in, (SELECT COALESCE(SUM(quantity), 0) FROM transactions t WHERE t.variant_id = v.id AND t.trans_type = 'SATIS') as total_out FROM variants v JOIN products p ON v.product_id = p.id",
            self.conn)

    def get_finance_only_report(self):
        return pd.read_sql_query(
            "SELECT v.sku, p.model_name, (SELECT COALESCE(SUM(quantity * unit_price), 0) FROM transactions t WHERE t.variant_id = v.id AND t.trans_type = 'ALIS') as total_cost, (SELECT COALESCE(SUM(quantity * unit_price), 0) FROM transactions t WHERE t.variant_id = v.id AND t.trans_type = 'SATIS') as total_sales FROM variants v JOIN products p ON v.product_id = p.id",
            self.conn)

    # --- DİĞERLERİ ---
    def add_brand(self, name):
        self.cursor.execute("INSERT INTO brands (name) VALUES (?)", (name,)); self.conn.commit(); return True

    def get_brands(self):
        self.cursor.execute("SELECT name FROM brands ORDER BY name ASC"); return [r[0] for r in self.cursor.fetchall()]

    def add_account(self, name, acc_type, contact, address):
        self.cursor.execute("INSERT INTO accounts (name, acc_type, contact, address) VALUES (?,?,?,?)",
                            (name, acc_type, contact, address)); self.conn.commit()

    def get_accounts(self, acc_type=None):
        self.cursor.execute("SELECT * FROM accounts WHERE acc_type=?" if acc_type else "SELECT * FROM accounts",
                            (acc_type,) if acc_type else ()); return self.cursor.fetchall()

    def get_suppliers_list(self):
        return [r[1] for r in self.get_accounts("SUPPLIER")]

    def get_account_balance(self, acc_id):
        self.cursor.execute("SELECT SUM(amount) FROM ledgers WHERE account_id=?", (acc_id,)); r = \
        self.cursor.fetchone()[0]; return r if r else 0.0

    def add_financial_transaction(self, acc_id, doc_type, desc, amount):
        self.cursor.execute("INSERT INTO ledgers (account_id, date, doc_type, description, amount) VALUES (?,?,?,?,?)",
                            (acc_id, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), doc_type, desc,
                             amount)); self.conn.commit()

    def get_account_statement(self, acc_id):
        self.cursor.execute(
            "SELECT date, doc_type, description, amount FROM ledgers WHERE account_id=? ORDER BY date DESC",
            (acc_id,)); return self.cursor.fetchall()

    def import_products_from_excel(self, path):
        try:
            df = pd.read_excel(path);
            count = 0
            for i, r in df.iterrows():
                b, m = str(r.get('Brand', '')), str(r.get('Model', ''));
                fob = float(r.get('FOB', 0));
                p = float(r.get('Parity', 1.05));
                c_usd = fob * 1.22 + float(r.get('Freight', 0)) + float(r.get('Local', 0)) + float(r.get('Test', 0))
                self.cursor.execute("SELECT id FROM products WHERE model_name=? AND brand=?", (m, b));
                res = self.cursor.fetchone()
                if res:
                    pid = res[0]
                else:
                    self.cursor.execute(
                        "INSERT INTO products (model_name, brand, cost_fob, parity_eur_usd, cost_cpt_usd, cost_cpt_eur, sales_price) VALUES (?,?,?,?,?,?,?)",
                        (m, b, fob, p, c_usd, c_usd / p if p else 0,
                         float(r.get('SalesPrice', 0)))); pid = self.cursor.lastrowid
                sku = f"{b[:3]}-{m}-{r.get('Color')}-{r.get('Size')}".upper().replace(" ", "");
                self.cursor.execute("SELECT id FROM variants WHERE sku=?", (sku,))
                if not self.cursor.fetchone(): self.cursor.execute(
                    "INSERT INTO variants (product_id, sku, color, size, ean_code, image_path) VALUES (?,?,?,?,?,?)",
                    (pid, sku, str(r.get('Color')), str(r.get('Size')), self.generate_ean13(), '')); count += 1
            self.conn.commit();
            return True, f"{count} eklendi"
        except Exception as e:
            return False, str(e)

    def export_data_to_excel(self, r_type, path):
        try:
            if r_type == "STOCK":
                df = self.get_stock_only_report()
            elif r_type == "FINANCE":
                df = self.get_finance_only_report()
            elif r_type == "ACCOUNTS":
                df = pd.read_sql_query("SELECT * FROM accounts", self.conn)
            elif r_type == "INVOICES":
                df = pd.read_sql_query(
                    "SELECT t.invoice_no, t.date, v.sku, v.ean_code, p.model_name, t.quantity, t.unit_price, (t.quantity*t.unit_price) as Total FROM transactions t JOIN variants v ON t.variant_id=v.id JOIN products p ON v.product_id=p.id ORDER BY t.date DESC",
                    self.conn)
            else:
                return False
            df.to_excel(path, index=False);
            return True
        except:
            return False