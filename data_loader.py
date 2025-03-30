import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
import os
import sqlite3
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# --- Configuration ---
default_excel_file_name = 'mutual_funds.xlsx'
db_file_name = 'mf_data.db'
table_name = 'mutual_funds'
script_dir = os.path.dirname(os.path.abspath(__file__))
db_file_path = os.path.join(script_dir, db_file_name)

# --- Mapping for Ranking Criteria ---
RANKING_OPTIONS = {
    "Best 3Y CAGR": ("CAGR_3Y", "DESC"),
    "Best 5Y CAGR": ("CAGR_5Y", "DESC"),
    "Best 1Y Return": ("Absolute_Returns_1Y", "DESC"),
    "Lowest Expense Ratio": ("Expense_Ratio", "ASC"),
    "Highest Sharpe Ratio": ("Sharpe_Ratio", "DESC"),
    "Highest Alpha": ("Alpha", "DESC"),
    "Largest AUM": ("AUM", "DESC"),
}

# --- CORE DATA FUNCTIONS (Paste latest working versions here) ---
# --- [Paste load_mutual_fund_data_from_excel here] ---
# --- [Paste save_data_to_sqlite here] ---
# --- [Paste query_data_from_sqlite here] ---
# --- [Paste show_category_chart here] ---

# Placeholder functions if needed - REPLACE these with your actual working functions
# --- CORE DATA FUNCTIONS ---

def load_mutual_fund_data_from_excel(file_path):
    """ Loads data from Excel, cleans column names, and adds a timestamp. """
    status_update(f"Attempting to load data from Excel: {os.path.basename(file_path)}")
    try:
        if not os.path.exists(file_path):
            messagebox.showerror("Error", f"Excel file not found at {file_path}")
            status_update("Error: Excel file not found.")
            return None
        df = pd.read_excel(file_path)
        status_update(f"Successfully loaded {len(df)} rows from Excel!")
        # Clean column names for SQL compatibility
        original_columns = list(df.columns)
        df.columns = df.columns.str.replace(r'[^A-Za-z0-9_]+', '_', regex=True).str.replace(r'_+', '_', regex=True).str.strip('_')
        cleaned_columns = list(df.columns)
        if original_columns != cleaned_columns:
             print("Column name changes:")
             for orig, clean in zip(original_columns, cleaned_columns):
                  if orig != clean: print(f"  '{orig}' -> '{clean}'")
        status_update("Cleaned column names.")
        # Add 'Date_Loaded' column
        load_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        df['Date_Loaded'] = load_time
        status_update(f"Added 'Date_Loaded' column: {load_time}")
        return df
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred loading Excel:\n{e}")
        status_update(f"Error loading Excel: {e}")
        return None

def save_data_to_sqlite(df, db_path, tbl_name):
    """ Saves the DataFrame to SQLite, APPENDING data. """
    if df is None or df.empty:
        status_update("No data provided to save.")
        return False
    status_update(f"Appending {len(df)} rows to SQLite table '{tbl_name}'...")
    conn = None
    try:
        conn = sqlite3.connect(db_path)
        df.to_sql(tbl_name, conn, if_exists='append', index=False)
        conn.commit() # Ensure data is saved
        conn.close()
        status_update(f"Successfully appended data to table '{tbl_name}'.")
        return True
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred saving to database:\n{e}")
        status_update(f"Error saving/appending to DB: {e}")
        if conn:
            conn.rollback() # Rollback changes on error
            conn.close()
        return False

def query_data_from_sqlite(db_path, query, params=None):
    """ Generic function to query data from SQLite. Returns DataFrame or None on error. """
    # Note: Removed table_name argument as it's part of the query string
    conn = None
    df_from_db = None
    try:
        if not os.path.exists(db_path):
             status_update("Error: Database file not found.")
             return None
        conn = sqlite3.connect(db_path)
        print(f"Executing query: {query}") # Keep console log
        if params:
            print(f"With parameters: {params}")
            df_from_db = pd.read_sql_query(query, conn, params=params)
        else:
            df_from_db = pd.read_sql_query(query, conn)
        conn.close()
        return df_from_db
    except Exception as e:
        print(f"Error querying DB: {e}")
        status_update(f"Error querying DB: {e}")
        if conn: conn.close()
        return None # Ensure None is returned on error

def show_category_chart():
    """ Queries category counts for the LATEST data and displays a bar chart. """
    status_update("Generating category chart...")
    # 1. Find latest timestamp
    latest_date_query = f"SELECT MAX(Date_Loaded) FROM {table_name}"
    df_latest_date = query_data_from_sqlite(db_path=db_file_path, query=latest_date_query) # Use corrected call signature
    if df_latest_date is None or df_latest_date.empty or pd.isna(df_latest_date.iloc[0,0]):
         messagebox.showerror("Error", "Could not determine the latest data load time.")
         status_update("Error: Cannot find latest load time for chart.")
         return
    latest_timestamp = df_latest_date.iloc[0,0]
    status_update(f"Latest data timestamp for chart: {latest_timestamp}")
    # 2. Query category counts for that timestamp
    chart_query = f"""
        SELECT Sub_Category, COUNT(*) as Count
        FROM {table_name} WHERE Date_Loaded = ?
        GROUP BY Sub_Category ORDER BY Count DESC LIMIT 15 """
    df_chart = query_data_from_sqlite(db_path=db_file_path, query=chart_query, params=(latest_timestamp,)) # Use corrected call signature
    if df_chart is None:
         status_update("Could not query chart data."); return
    if df_chart.empty:
        messagebox.showinfo("Info", "No category data found for the latest timestamp.")
        status_update("No category data found for chart."); return
    status_update("Chart data queried successfully.")
    # 3. Create and display plot
    try:
        fig, ax = plt.subplots(figsize=(10, 6))
        ax.bar(df_chart['Sub_Category'], df_chart['Count'])
        ax.set_xlabel('Sub Category'); ax.set_ylabel('Number of Funds')
        ax.set_title(f'Top 15 Fund Categories by Count (as of {latest_timestamp})')
        plt.xticks(rotation=90); plt.tight_layout()
        chart_window = tk.Toplevel(root)
        chart_window.title("Fund Category Chart"); chart_window.geometry("800x600")
        canvas = FigureCanvasTkAgg(fig, master=chart_window)
        canvas_widget = canvas.get_tk_widget(); canvas_widget.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        canvas.draw()
        status_update("Chart displayed.")
    except Exception as e:
        messagebox.showerror("Chart Error", f"Could not generate or display chart:\n{e}")
        status_update(f"Chart error: {e}")

# --- [The rest of your code (GUI HELPER FUNCTIONS, Main Application Setup) should follow here] ---
# --- [Make sure it's the latest version with the advanced filters etc.] ---


# --- GUI HELPER FUNCTIONS ---

def browse_excel_file():
    file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path: excel_path_var.set(file_path); status_update(f"Selected: {os.path.basename(file_path)}")

def process_data():
    excel_file = excel_path_var.get()
    if not excel_file: messagebox.showwarning("Missing Input", "Select Excel file."); return
    clear_treeview()
    df = load_mutual_fund_data_from_excel(excel_file)
    if df is not None:
        save_successful = save_data_to_sqlite(df, db_file_path, table_name)
        if save_successful:
            status_update("Pipeline Complete: Load & Append successful!"); messagebox.showinfo("Success", "Data loaded & appended!")
            populate_category_filter()
        else: status_update("Pipeline Failed: Could not append.")
    else: status_update("Pipeline Failed: Could not load.")

# --- MODIFIED: Display Data - Now with MORE Filters ---
def display_ranked_data():
    """ Queries latest data based on ALL filters and ranking, shows in Treeview. """
    selected_category = category_filter_var.get()
    selected_ranking_key = ranking_criteria_var.get()

    # --- Get Filter Values ---
    try: # Use try-except for number conversions
        min_aum_str = min_aum_var.get()
        min_aum = float(min_aum_str) if min_aum_str else None

        max_aum_str = max_aum_var.get()
        max_aum = float(max_aum_str) if max_aum_str else None

        min_exp_str = min_exp_var.get()
        min_exp = float(min_exp_str) if min_exp_str else None

        max_exp_str = max_exp_var.get()
        max_exp = float(max_exp_str) if max_exp_str else None
    except ValueError:
        messagebox.showerror("Input Error", "Invalid number entered in AUM or Expense Ratio filters.")
        return

    if not selected_ranking_key or selected_ranking_key not in RANKING_OPTIONS:
        messagebox.showwarning("Input Needed", "Please select a ranking criterion."); return

    ranking_col, sort_order = RANKING_OPTIONS[selected_ranking_key]
    status_update(f"Querying data: Cat='{selected_category}', Rank='{selected_ranking_key}', AUM=({min_aum}-{max_aum}), Exp=({min_exp}-{max_exp})")

    # --- Build Query Dynamically ---
    columns_to_select = f"Name, Sub_Category, AUM, NAV, Expense_Ratio, CAGR_3Y, CAGR_5Y, Absolute_Returns_1Y, Sharpe_Ratio, Alpha, Date_Loaded"
    base_query = f""" SELECT {columns_to_select} FROM {table_name}
                      WHERE Date_Loaded = (SELECT MAX(Date_Loaded) FROM {table_name}) """
    params = [] # Parameters for SQL query

    # Add filters dynamically
    if selected_category and selected_category != "All Categories":
        base_query += " AND Sub_Category = ? "
        params.append(selected_category)
    if min_aum is not None:
        base_query += " AND AUM >= ? "
        params.append(min_aum)
    if max_aum is not None:
        base_query += " AND AUM <= ? "
        params.append(max_aum)
    if min_exp is not None:
        base_query += " AND Expense_Ratio >= ? "
        params.append(min_exp)
    if max_exp is not None:
        base_query += " AND Expense_Ratio <= ? "
        params.append(max_exp)

    base_query += f" ORDER BY {ranking_col} {sort_order} LIMIT 100 " # Limit results

    # --- Execute Query ---
    df_filtered_ranked = query_data_from_sqlite(db_path=db_file_path, query=base_query, params=params if params else None)

    # --- Update Treeview ---
    clear_treeview()
    if df_filtered_ranked is not None:
        if not df_filtered_ranked.empty:
            tree["columns"] = list(df_filtered_ranked.columns)
            tree["show"] = "headings"
            for col in df_filtered_ranked.columns:
                heading_text = col.replace('_', ' ')
                tree.heading(col, text=heading_text)
                width = 100; anchor = 'center' # Defaults
                if "Name" in col: width=250; anchor='w'
                elif "Category" in col: width=150; anchor='w'
                elif "Date" in col: width=130; anchor='center'
                elif "AUM" in col or "NAV" in col: width=100; anchor='e'
                elif "Ratio" in col or "CAGR" in col or "Return" in col or "Alpha" in col: width=80; anchor='e'
                tree.column(col, width=width, anchor=anchor)
            for index, row in df_filtered_ranked.iterrows():
                tree.insert("", tk.END, values=tuple(str(v) if pd.notna(v) else "" for v in row))
            status_update(f"Displayed {len(df_filtered_ranked)} funds matching filters, ranked by '{selected_ranking_key}'.")
        else:
             status_update(f"No data found matching filters.")
    else:
        status_update("Failed to retrieve filtered/ranked data.")


def clear_treeview():
     try: tree.delete(*tree.get_children())
     except Exception as e: print(f"Error clearing treeview: {e}")

def status_update(message):
    try: status_var.set(message)
    except Exception as e: print(f"Error updating status bar: {e}")
    print(message)

def populate_category_filter():
    status_update("Populating category filter...")
    query = f""" SELECT DISTINCT Sub_Category FROM {table_name}
                 WHERE Date_Loaded = (SELECT MAX(Date_Loaded) FROM {table_name}) ORDER BY Sub_Category """
    df_categories = query_data_from_sqlite(db_path=db_file_path, query=query)
    category_list = ["All Categories"]
    if df_categories is not None and not df_categories.empty:
        category_list.extend(df_categories['Sub_Category'].tolist())
        status_update(f"Found {len(category_list)-1} categories.")
    else: status_update("No categories found in DB yet.")
    try:
        category_filter_combobox['values'] = category_list
        if category_list: category_filter_var.set(category_list[0])
    except Exception as e: print(f"Error populating combobox: {e}")


# --- Main Application Setup ---
if __name__ == "__main__":
    root = tk.Tk()
    root.title("MF Analyzer V7 (Advanced Filters)") # Updated Title
    root.geometry("1100x750") # Wider window again

    # --- Tkinter Variables ---
    excel_path_var = tk.StringVar(value=os.path.join(script_dir, default_excel_file_name))
    status_var = tk.StringVar(value="Initializing...")
    category_filter_var = tk.StringVar()
    ranking_criteria_var = tk.StringVar()
    # New variables for filter entries
    min_aum_var = tk.StringVar()
    max_aum_var = tk.StringVar()
    min_exp_var = tk.StringVar()
    max_exp_var = tk.StringVar()


    # --- Style ---
    style = ttk.Style()
    style.theme_use('clam')

    # --- Control Frame (Top) ---
    control_frame = ttk.Frame(root, padding="10")
    control_frame.pack(side=tk.TOP, fill=tk.X, pady=5)

    # Configure columns for controls
    control_frame.columnconfigure(1, weight=1) # Allow category/AUM filters to expand
    control_frame.columnconfigure(3, weight=1) # Allow ranking/Exp filters to expand

    # Row 0: File Selection
    ttk.Label(control_frame, text="Excel File:").grid(row=0, column=0, padx=5, pady=2, sticky="w")
    excel_entry = ttk.Entry(control_frame, textvariable=excel_path_var, width=60)
    excel_entry.grid(row=0, column=1, columnspan=3, padx=5, pady=2, sticky="ew")
    browse_button = ttk.Button(control_frame, text="Browse...", command=browse_excel_file)
    browse_button.grid(row=0, column=4, padx=5, pady=2)

    # Row 1: Category & Ranking Filters
    ttk.Label(control_frame, text="Filter Category:").grid(row=1, column=0, padx=5, pady=2, sticky="w")
    category_filter_combobox = ttk.Combobox(control_frame, textvariable=category_filter_var, state='readonly', width=35)
    category_filter_combobox.grid(row=1, column=1, padx=5, pady=2, sticky="w")

    ttk.Label(control_frame, text="Rank By:").grid(row=1, column=2, padx=(15, 5), pady=2, sticky="w") # More space before Rank
    ranking_criteria_combobox = ttk.Combobox(control_frame, textvariable=ranking_criteria_var, state='readonly', width=30, values=list(RANKING_OPTIONS.keys()))
    ranking_criteria_combobox.grid(row=1, column=3, columnspan=2, padx=5, pady=2, sticky="w")
    if RANKING_OPTIONS: ranking_criteria_combobox.current(0)

    # --- NEW Row 2: Numerical Filters ---
    ttk.Label(control_frame, text="Min AUM:").grid(row=2, column=0, padx=5, pady=2, sticky="w")
    min_aum_entry = ttk.Entry(control_frame, textvariable=min_aum_var, width=10)
    min_aum_entry.grid(row=2, column=1, padx=5, pady=2, sticky="w")

    ttk.Label(control_frame, text="Max AUM:").grid(row=2, column=2, padx=(15,5), pady=2, sticky="w")
    max_aum_entry = ttk.Entry(control_frame, textvariable=max_aum_var, width=10)
    max_aum_entry.grid(row=2, column=3, padx=5, pady=2, sticky="w")

    # --- NEW Row 3: Numerical Filters (Continued) ---
    ttk.Label(control_frame, text="Min Exp Ratio:").grid(row=3, column=0, padx=5, pady=2, sticky="w")
    min_exp_entry = ttk.Entry(control_frame, textvariable=min_exp_var, width=10)
    min_exp_entry.grid(row=3, column=1, padx=5, pady=2, sticky="w")

    ttk.Label(control_frame, text="Max Exp Ratio:").grid(row=3, column=2, padx=(15,5), pady=2, sticky="w")
    max_exp_entry = ttk.Entry(control_frame, textvariable=max_exp_var, width=10)
    max_exp_entry.grid(row=3, column=3, padx=5, pady=2, sticky="w")


    # --- Row 4: Action Buttons ---
    load_button = ttk.Button(control_frame, text="Load Excel & Append", command=process_data, width=20)
    load_button.grid(row=4, column=0, padx=5, pady=10, sticky="w")

    display_button = ttk.Button(control_frame, text="Show Filtered & Ranked Data", command=display_ranked_data, width=25)
    display_button.grid(row=4, column=1, columnspan=2, padx=5, pady=10, sticky="w") # Span 2 cols

    chart_button = ttk.Button(control_frame, text="Show Category Chart", command=show_category_chart, width=20)
    chart_button.grid(row=4, column=3, columnspan=2, padx=5, pady=10, sticky="w") # Span 2 cols


    # --- Treeview Frame (Middle) ---
    tree_frame = ttk.Frame(root, padding="10")
    tree_frame.pack(fill=tk.BOTH, expand=True)
    tree = ttk.Treeview(tree_frame)
    vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
    hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    tree.grid(row=0, column=0, sticky='nsew')
    vsb.grid(row=0, column=1, sticky='ns')
    hsb.grid(row=1, column=0, sticky='ew')
    tree_frame.grid_rowconfigure(0, weight=1)
    tree_frame.grid_columnconfigure(0, weight=1)

    # --- Status Bar (Bottom) ---
    status_bar = ttk.Label(root, textvariable=status_var, relief=tk.SUNKEN, anchor=tk.W, padding="2 5")
    status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    # --- Initial Population ---
    root.after(100, populate_category_filter) # Populate dropdown after window loads

    # --- Run ---
    status_update("Ready.")
    root.mainloop()