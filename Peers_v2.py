# https://chatgpt.com/share/681f5e34-186c-8010-8fec-01d663ee5dfd

import pandas as pd
from xbbg import blp
import customtkinter as ctk
from tkinter import messagebox, ttk
import tkinter as tk
import pyperclip

sort_column = None
sort_reverse = False

# Mapping Bloomberg fields to user-friendly labels
column_name_map = {
    "ticker": "Ticker",
    "name": "Company",
    "market_cap": "Market Cap",
    "gics_sector_name": "Sector",
    "gics_industry_group_name": "Industry Group",
    "gics_industry_name": "Industry",
    "gics_sub_industry_name": "Sub-Industry",
    "px_last": "Price",
    "eqy_sh_out": "Shares Outstanding",
    "eqy_dvd_yld_ind": "Dividend Yield"
}

def fetch_and_display(event=None):
    ticker = entry.get().strip().upper()
    if not ticker.endswith(" EQUITY"):
        ticker += " EQUITY"

    try:
        # Step 1: Fetch GICS Classification
        fields = [
            'GICS_SECTOR_NAME', 
            'GICS_INDUSTRY_GROUP_NAME', 
            'GICS_INDUSTRY_NAME', 
            'GICS_SUB_INDUSTRY_NAME'
        ]
        data = blp.bdp(ticker, fields)

        if data.empty:
            messagebox.showerror("Error", f"No data found for {ticker}.")
            return

        sub_industry = data.iloc[0]['gics_sub_industry_name']
        if not sub_industry:
            messagebox.showerror("Error", f"No Sub-Industry data for {ticker}.")
            return

        # Step 2: Fetch Peers in Same Sub-Industry
        peers = blp.screen(
            screen='MOST_ACTIVE',  # Replace with your custom Bloomberg screen if available
            overrides={'GICS_SUB_INDUSTRY_NAME': sub_industry},
            fields=['ticker', 'name', 'market_cap', 'gics_sub_industry_name', 'px_last', 'eqy_sh_out', 'eqy_dvd_yld_ind']
        )

        if peers.empty:
            messagebox.showinfo("Info", f"No peers found for Sub-Industry: {sub_industry}.")
            return

        df = peers.copy()
        df = df[df['gics_sub_industry_name'] == sub_industry]

        if 'market_cap' in df.columns:
            df['marketCap_raw'] = df['market_cap']
            df['market_cap'] = (df['market_cap'] / 1e6).round(0).astype(int).map('${:,.0f} M'.format)

        if 'eqy_sh_out' in df.columns:
            df['eqy_sh_out'] = df['eqy_sh_out'].fillna(0).astype(int).map('{:,}'.format)

        if 'marketCap_raw' in df.columns:
            df.sort_values(by='marketCap_raw', ascending=False, inplace=True)

        display_df(df)

    except Exception as e:
        messagebox.showerror("Error", str(e))

def display_df(df):
    for row in tree.get_children():
        tree.delete(row)

    tree["columns"] = [col for col in df.columns if not col.endswith("_raw")]
    tree["show"] = "headings"

    for col in tree["columns"]:
        display_name = column_name_map.get(col, col)
        tree.heading(col, text=display_name, command=lambda _col=col: sort_by_column(_col))
        tree.column(col, anchor="w", width=150)

    for _, row in df.iterrows():
        values = [row[col] for col in tree["columns"]]
        tree.insert("", "end", values=values)

def copy_selected_symbols(event=None):
    selected_items = tree.selection()
    symbols = []
    for item in selected_items:
        values = tree.item(item, "values")
        symbol_index = tree["columns"].index("ticker")
        symbols.append(values[symbol_index])
    
    if symbols:
        pyperclip.copy("\r\n".join(symbols))

def show_context_menu(event):
    selected_item = tree.identify_row(event.y)
    if selected_item:
        tree.selection_set(selected_item)
        context_menu.post(event.x_root, event.y_root)

def sort_by_column(col):
    global sort_column, sort_reverse
    rows = [(tree.set(k, col), k) for k in tree.get_children('')]
    try:
        rows.sort(key=lambda x: float(str(x[0]).replace('$', '').replace(',', '').replace(' M', '')), reverse=sort_column == col and not sort_reverse)
    except:
        rows.sort(reverse=sort_column == col and not sort_reverse)
    for index, (_, k) in enumerate(rows):
        tree.move(k, '', index)
    sort_column = col
    sort_reverse = not (sort_column == col and sort_reverse)

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

root = ctk.CTk()
root.geometry("1100x700")
root.title("Peers (Bloomberg)")

frame = ctk.CTkFrame(root)
frame.pack(pady=20, fill="x")

entry_label = ctk.CTkLabel(frame, text="Enter Ticker:")
entry_label.pack(side="left", padx=(10, 5))

entry = ctk.CTkEntry(frame, width=100)
entry.pack(side="left")
entry.bind("<Return>", fetch_and_display)

button_column = ctk.CTkFrame(frame)
button_column.pack(side="left", padx=(20, 0))

submit_btn = ctk.CTkButton(button_column, text="Get Peers", command=fetch_and_display)
submit_btn.pack(pady=(0, 10))

copy_btn = ctk.CTkButton(
    button_column,
    text="Copy Selected Tickers",
    command=copy_selected_symbols,
    fg_color="#FFD700",
    hover_color="#E6BE00",
    text_color="black"
)
copy_btn.pack()

tree_frame = tk.Frame(root)
tree_frame.pack(pady=10, fill="both", expand=True)

tree_scrollbar = tk.Scrollbar(tree_frame, width=30)
tree_scrollbar.pack(side="right", fill="y")

tree = ttk.Treeview(tree_frame, selectmode="extended", yscrollcommand=tree_scrollbar.set)
tree.pack(fill="both", expand=True)
tree.bind("<Control-c>", copy_selected_symbols)
tree.bind("<Button-3>", show_context_menu)

tree_scrollbar.config(command=tree.yview)

style = ttk.Style()
style.configure("Treeview", font=("Segoe UI", 11))
style.configure("Treeview.Heading", font=("Segoe UI", 11))

context_menu = tk.Menu(root, tearoff=0)
context_menu.add_command(label="Copy Selected Tickers", command=copy_selected_symbols)

root.mainloop()