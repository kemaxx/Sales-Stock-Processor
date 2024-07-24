import pandas as pd
import numpy as np
import re
from collections import defaultdict
import Levenshtein
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import gspread as gc
import os
from dotenv import load_dotenv

load_dotenv()
#import seaborn as sns
#import matplotlib.pyplot as plt


class SalesStockProcessor:
    def __init__(self, root):
        self.root = root
        self.root.title("Sales Stock Processor")
        self.root.iconbitmap("favicon.ico")

        # Set a theme for the application
        self.style = ttk.Style()
        self.style.theme_use("clam")

        # Set styles for widgets
        self.style.configure("TLabel", font=("Helvetica", 12), padding=5)
        self.style.configure("TButton", font=("Helvetica", 12), padding=5)
        self.style.configure("TEntry", padding=5)
        self.style.configure("TCombobox", padding=5)

        # Initialize the UI
        self.create_widgets()
        self.create_menu()

    def create_widgets(self):
        # Create and place the labels and file entry fields with browse buttons
        self.store_record_label = ttk.Label(self.root, text="Store Record")
        self.store_record_label.grid(row=1, column=0, padx=10, pady=5, sticky="e")
        self.store_record_entry = ttk.Entry(self.root, width=40, state="readonly")
        self.store_record_entry.grid(row=1, column=1, padx=10, pady=5)
        self.store_record_button = ttk.Button(self.root, text="Browse",
                                              command=lambda: self.browse_file(self.store_record_entry))
        self.store_record_button.grid(row=1, column=2, padx=10, pady=5)

        self.department_label = ttk.Label(self.root, text="Department Record")
        self.department_label.grid(row=2, column=0, padx=10, pady=5, sticky="e")
        self.department_entry = ttk.Entry(self.root, width=40, state="readonly")
        self.department_entry.grid(row=2, column=1, padx=10, pady=5)
        self.department_button = ttk.Button(self.root, text="Browse",
                                            command=lambda: self.browse_file(self.department_entry))
        self.department_button.grid(row=2, column=2, padx=10, pady=5)

        # Create and place the Combo Box
        self.combobox_label = ttk.Label(self.root, text="Department Name")
        self.combobox_label.grid(row=1, column=3, padx=10, pady=5, sticky="e")
        self.combobox = ttk.Combobox(self.root, width=30, values=["KITCHEN & RESTAURANTANT DRINKS"], state='readonly')
        self.combobox.grid(row=1, column=4, padx=10, pady=5)

        # Create and place the Process Transactions button
        self.process_button = ttk.Button(self.root, text="Process Transactions", command=self.process_data_file)
        self.process_button.grid(row=3, column=1, columnspan=3, pady=20)

        # Additional Widgets
        # Widget 8: Info Label
        self.info_label = ttk.Label(self.root, text="", font=("Helvetica", 12))
        self.info_label.grid(row=4, column=0, columnspan=5, pady=10)

        # Widget 9: Treeview
        self.columns = ("Stock Name", "Opening Balance", "New Stock", "Qty Sold", "Expected Balance")
        self.tree = ttk.Treeview(self.root, columns=self.columns, show="headings")
        for col in self.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor='center')  # Center align the contents

        # Add vertical scrollbar
        self.scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=self.scrollbar.set)
        self.tree.grid(row=5, column=0, columnspan=5, padx=10, pady=5, sticky="nsew")
        self.scrollbar.grid(row=5, column=5, sticky='ns')

        # Widget 10: Export Report Button
        self.export_button = ttk.Button(self.root, text="Export Report", state=tk.DISABLED, command=self.export_data)
        self.export_button.grid(row=6, column=4, padx=10, pady=20)

        # Check treeview and update UI
        self.tree.bind("<ButtonRelease-1>", lambda e: self.check_treeview())

        # Set the window's main background color
        self.root.configure(bg="#f0f0f0")

    def create_menu(self):
        # Menu Bar
        self.menu_bar = tk.Menu(self.root)
        self.root.config(menu=self.menu_bar)

        # Add About Menu
        self.help_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Help", menu=self.help_menu)
        self.help_menu.add_command(label="About", command=self.show_about)

    def browse_file(self, entry):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")]
        )
        if file_path:
            entry.config(state=tk.NORMAL)  # Temporarily make the entry normal to insert the file path
            entry.delete(0, tk.END)
            entry.insert(0, file_path)
            entry.config(state="readonly")  # Make the entry readonly again

    def check_treeview(self):
        if len(self.tree.get_children()) > 0:
            self.export_button.config(state=tk.NORMAL)
            self.info_label.config(text="Data available for export.")
        else:
            self.export_button.config(state=tk.DISABLED)
            self.info_label.config(text="")

    def export_data(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")]
        )
        if file_path:
            try:
                # Collect data from treeview
                data = []
                for item in self.tree.get_children():
                    data.append(self.tree.item(item)['values'])

                # Convert to DataFrame
                df = pd.DataFrame(data, columns=self.columns)

                # Save to file
                if file_path.endswith('.xlsx'):
                    df.to_excel(file_path, index=False)
                else:
                    df.to_csv(file_path, index=False)

                messagebox.showinfo("Export Successful", f"Data exported successfully to {file_path}")
            except Exception as e:
                messagebox.showerror("Export Failed", f"An error occurred while exporting data: {str(e)}")

    def show_about(self):
        messagebox.showinfo("About", "Sales Stock Processor\n\nAll Rights Reserved Kenneth Mark\n©2024")

    def convert_stock_to_pos_unit(self, stock_name, qty):

        converted_new_stock = list()
        unit_of_measures = {
            "YAM": lambda qty: -0.43 + 2.36 * qty,
            "INDOMIE": lambda qty: qty * 40,
            "CHICKEN": lambda qty: qty * 4,
            "SAUSAGE": lambda qty: qty * 5,
            "SPAGHETTE": lambda qty: qty * 4,
            "PLANTAIN": lambda qty: qty * 2 + 1.4,
            "EGG": lambda qty: qty * 15,
        }

        if stock_name in unit_of_measures:
            conversion_function = unit_of_measures.get(stock_name)
            return conversion_function(qty)
        else:
            return np.nan
    def levenshtein_distance_percentage(self, word1, word2):
        m, n = len(word1), len(word2)
        dp = [[0] * (n + 1) for _ in range(m + 1)]

        for i in range(m + 1):
            dp[i][0] = i
        for j in range(n + 1):
            dp[0][j] = j

        for i in range(1, m + 1):
            for j in range(1, n + 1):
                if word1[i - 1] == word2[j - 1]:
                    dp[i][j] = dp[i - 1][j - 1]
                else:
                    dp[i][j] = min(dp[i - 1][j], dp[i][j - 1], dp[i - 1][j - 1]) + 1

        distance = dp[m][n]
        max_len = max(m, n)
        similarity = (1 - distance / max_len) * 100  # Convert to percentage
        return similarity

    def highest_levenshtein_similarity(self, word, word_list):
        max_similarity = 0
        most_similar_word = ""
        for w in word_list:
            similarity = self.levenshtein_distance_percentage(word, w)
            if similarity > max_similarity:
                max_similarity = similarity
                most_similar_word = w
        return most_similar_word, max_similarity

        ## Project New Stock

    def process_new_stock(self):
        # We'll need to categorize our new stocks for ease of analysis using gspread
        gspread_path = os.getenv("MY_GSPREAD_JSON_FILE")

        sales_df = pd.read_excel(str(self.department_entry.get()))
        rest_tokenized_menu = self.rest_tokenized_menu

        store_df = pd.read_excel(str(self.store_record_entry.get()), usecols=range(1, 8))
        store_df.columns = ["Date", "Item", "Sn", "Dept", "Qty_Bf", "Qty_Issued", "Qty_Af", ]

        # Categorizing our issued stocks from store
        account = gc.service_account(gspread_path)
        wkbook = account.open_by_key(os.getenv("MY_SECRET_GOOGLE_SHEET_KEY"))
        wk_sheet = wkbook.worksheet("My Stock")

        data = wk_sheet.get_all_values()
        data = [[str(cell).replace('"', '') for cell in row] for row in data]

        all_stock_db = pd.DataFrame(data=data[1:], columns=data[0])
        all_stock_category_dict = dict(zip(all_stock_db["Stock Name"].tolist(), all_stock_db["Category"].tolist()))

        # return store_df

        # Selecting only restaurant and kitchen requisitions
        store_df = store_df.loc[store_df["Dept"].isin(["KITCHEN", "RESTAURANT DRINKS"])]
        store_df = store_df.groupby(["Item", "Dept"])["Qty_Issued"].sum().reset_index()

        store_df = store_df.set_index("Item")
        for item in store_df.index.tolist():
            store_df.loc[item, "Category"] = all_stock_category_dict.get(item, np.nan)

        store_df = store_df.reset_index()

        stock_db_item_list = rest_tokenized_menu["Items"].str.strip().unique().tolist()

        sales_df = sales_df.mask(sales_df['ITEMS'].isin(["REC", "TOTAL", "TOTAL BY", "GROUND TOTAL"])).drop("TOT",
                                                                                                            axis='columns').dropna(
            subset="ITEMS")
        sales_df["QTY"] = sales_df["QTY"].astype(str).str.replace("R\d+|,[a-zA-Z]+|\W", "", regex=True)

        list_of_transations = list()
        sales_record_dict = defaultdict(float)
        for sold_item, qty in zip(sales_df["ITEMS"], sales_df["QTY"]):
            for stock in stock_db_item_list:

                if stock == "CHIPS":

                    exact_match = re.search(r'^\bCHIPS\b$', sold_item)
                    if exact_match:
                        #print("True match", sold_item)

                        sales_record_dict[stock] += float(qty)
                elif stock == "CHICKEN" and "B.B.Q." in sold_item:
                    continue


                elif stock in sold_item:
                    sales_record_dict[stock] += float(qty)

        processed_sales_df = pd.DataFrame(list(dict(sales_record_dict).items()), columns=["Item", "Qty"])


        ## Really???
        processed_sales_df = processed_sales_df.merge(store_df.loc[store_df["Category"].isin(["BEVERAGE"]), :],
                                                      on="Item", how="outer")


        ### Really????

        replaceables = {
            "PLAIN OMELLETE": "EGG",
            "CHIPS": "IRISH POTATOES",
            "SPANISH OMELLETE": "EGG"
        }

        processed_sales_df["Item"] = processed_sales_df["Item"].replace(replaceables)

        # Need to get the values counts of each item on this df
        grp_by_qty = processed_sales_df.groupby("Item")["Qty"].sum().reset_index()
        item_dict = dict(zip(grp_by_qty["Item"],grp_by_qty["Qty"]))
        print(item_dict)
        processed_sales_df = processed_sales_df.drop_duplicates(subset=["Item"])
        processed_sales_df["Qty"] = processed_sales_df["Item"].apply(lambda x:item_dict.get(x))


        # Since the stock naming in the store requisitions diffs mostly from those of POS (i.e Restaurant), we'll need to rename
        # the latter to feat with the former. We'll use the Levenshtein Distance algorithm for this task

        levenshtein_stock_score_list = list()
        levenshtein_stock_name_list = list()
        store_stock_list = store_df["Item"].unique().tolist()

        for item in processed_sales_df['Item'].tolist():
            computed_stock, similarity_score = self.highest_levenshtein_similarity(item, store_stock_list)
            levenshtein_stock_score_list.append(similarity_score)
            levenshtein_stock_name_list.append(computed_stock)

        processed_sales_df["Lev Stock"] = levenshtein_stock_name_list
        processed_sales_df["Lev Score"] = levenshtein_stock_score_list
        processed_sales_df["Splitted"] = processed_sales_df["Lev Stock"].str.split(" ")

        verified_list = list()

        for item, lev, splitted, score in zip(processed_sales_df["Item"].tolist(),
                                              processed_sales_df["Lev Stock"].tolist()
                , processed_sales_df["Splitted"],
                                              processed_sales_df["Lev Score"]):

            if (item == lev) | (splitted[0] == item) | (round(score, 0) >= 80):
                verified_list.append('Yes')
            else:
                verified_list.append("No")

        processed_sales_df["Verified"] = verified_list

        # Drop duplicated entries
        processed_sales_df = processed_sales_df.drop_duplicates(subset=["Item"], keep="first")
        # processed_sales_df["Qty"] = new_qty

        # Aggregate any duplicated entries
        processed_sales_df_grouped = processed_sales_df.groupby("Item")["Qty"].sum()
        real_qty_dict = dict(processed_sales_df_grouped.items())

        real_stock_list = processed_sales_df["Item"].tolist()

        processed_sales_df.set_index("Item", inplace=True)

        for item in real_stock_list:
            processed_sales_df.loc[item, "Qty"] = real_qty_dict[item]
        processed_sales_df = processed_sales_df.reset_index()

        # Getting the actual new stocks from store issues

        processed_store_stock_list = processed_sales_df["Lev Stock"].tolist()
        verified_stock_with_pos = processed_sales_df["Verified"].tolist()
        store_df_stock_qty_dict = dict(zip(store_df["Item"], store_df["Qty_Issued"]))

        new_stock_from_store_list = list()
        for lev_stock, verified_with_pos in zip(processed_store_stock_list, verified_stock_with_pos):
            if verified_with_pos == "Yes":
                new_stock_from_store_list.append(store_df_stock_qty_dict.get(lev_stock))
            else:
                new_stock_from_store_list.append(np.nan)

        processed_sales_df["New Stock"] = new_stock_from_store_list

        return processed_sales_df

        # f = processed_sales_df.merge(store_df.loc[store_df["Category"].isin(["BEVERAGE"]),:],on="Item",how="outer")

        # return f

        # processed_sales_df["Item"] = processed_sales_df["Item"].replace(replaceables)

        # Checking the similarity score of store stocks against POS stock
        #     store_stock_against_pos_stock_score_list = list()
        #     data = {}
        #     store_item_list = store_df["Item"].unique().tolist()
        #     for item in store_item_list:
        #         stock, score = highest_levenshtein_similarity(item,stock_db_item_list)
        #         #print(stock,score)

        #         data.update({item:[stock,score]})

    def process_data_file(self):

        self.info_label.config(text="Processing data, please wait...")
        self.root.update_idletasks()
        # Get records
        store_record_file = str(self.store_record_entry.get())
        dept_record_file = str(self.department_entry.get())

        # Validate records
        if not store_record_file:
            messagebox.showerror("Empty Store Rcord File", "Please, kindly make a valid selection for store record")
        elif not dept_record_file:
            messagebox.showerror("Empty Department Rcord File",
                                 "Please, kindly make a valid selection for department record")
            return

        sales_df = pd.read_excel(dept_record_file)
        stock_db = pd.read_excel(store_record_file)

        self.rest_tokenized_menu = pd.read_excel("tokenized restaurant menu.xlsx")
        stock_db_item_list = self.rest_tokenized_menu["Items"].str.strip().unique().tolist()

        #         sales_df = sales_df.mask(sales_df['ITEMS'].isin(["REC","TOTAL","TOTAL BY","GROUND TOTAL"])).drop("TOT",axis='columns').dropna(subset="ITEMS")
        #         sales_df["QTY"] = sales_df["QTY"].astype(str).str.replace("R\d+|,[a-zA-Z]+|\W","",regex=True)

        #         # Reworked
        #         list_of_transations = list()
        #         sales_record_dict = defaultdict(float)
        #         for sold_item, qty in zip(sales_df["ITEMS"],sales_df["QTY"]):
        #             for stock in stock_db_item_list:
        #                 if stock=="CHIPS":
        #                     exact_match = re.search(r'^\bCHIPS\b$',sold_item)
        #                     if exact_match:
        #                         print("True match",sold_item)

        #                         sales_record_dict[stock]+=float(qty)
        #                 elif stock=="CHICKEN" and "B.B.Q." in sold_item:
        #                     continue

        #                 elif stock in sold_item:
        #                     sales_record_dict[stock]+=float(qty)

        #         # Solutions
        #         ## Replace differences
        #         replaceables = {
        #             "PLAIN OMELLETE":"EGG"
        #         }

        #         processed_sales_df = pd.DataFrame(list(dict(sales_record_dict).items()),columns=["Item","Qty"])

        #         processed_sales_df = self.process_new_stock(processed_sales_df)

        # input("Hi")
        processed_sales_df = self.process_new_stock()

        processed_sales_df["Opening Balance"] = np.random.randint(5, 20, size=(processed_sales_df.shape[0]), )
        # processed_sales_df["New Stock"] = np.random.randint(5,20,size=(processed_sales_df.shape[0]),)
        processed_sales_df["Expected Balance"] = (processed_sales_df["Opening Balance"] + processed_sales_df[
            "New Stock"]) - processed_sales_df["Qty"]
        # processed_sales_df = processed_sales_df[["Item","Opening Balance","New Stock","Qty","Expected Balance"]]

        processed_sales_df = processed_sales_df[["Item", "Opening Balance", "New Stock", "Qty", "Expected Balance"]]

        processed_sales_df.sort_values(by="Item", ascending=True, inplace=True)
        replaceables = {
            "PLAIN OMELLETE": "EGG",
            "CHIPS": "IRISH POTATOES",
            "SPANISH OMELLETE": "EGG",

        }
        processed_sales_df["Item"] = processed_sales_df["Item"].replace(replaceables)


        # This section focuses on making the necessary conversions of new stocks from the store
        converted_new_stock_list = list()
        # processed_sales_df["New Stock"] = new_stock_from_store_list

        for stock, qty in zip(processed_sales_df["Item"].tolist(), processed_sales_df["New Stock"].tolist()):

            value = self.convert_stock_to_pos_unit(stock, qty)
            if np.isnan(value):
                converted_new_stock_list.append(qty)
            else:
                converted_new_stock_list.append(value)
            # prin(converted_new_stock_list)

        processed_sales_df["New Stock"] = converted_new_stock_list
        
        # This part is for drill only
        processed_sales_df["Opening Balance"] = processed_sales_df["Opening Balance"].astype(str).replace(r'^.*$', 'X',
                                                                                                          regex=True)
        processed_sales_df["Expected Balance"] = processed_sales_df["Expected Balance"].astype(str).replace(r'^.*$', 'X',
                                                                                                          regex=True)
        # populate tree widget
        for row in processed_sales_df.itertuples(index=False):
            self.tree.insert("", tk.END, values=row)
        self.tree.focus()

# Run the application
root = tk.Tk()
app = SalesStockProcessor(root)
root.mainloop()
