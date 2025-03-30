import tkinter as tk
from tkinter import ttk # Using themed widgets
from tkinter import filedialog, messagebox, scrolledtext # scrolledtext might be removed if not needed elsewhere
import pandas as pd
import os
import sqlite3
from datetime import datetime

# Matplotlib Imports (Keep for chart functionality)
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# --- Configuration ---
default_excel_file_name = 'mutual_funds.xlsx'
db_file_name = 'mf_data.db'
table_name = 'mutual_funds'

script_dir = os.path.dirname(__file__)
db_file_path = os.path.join(script_dir, db_file_name)

# --- Database/Excel/Chart Functions (Largely unchanged from previous version) ---

def load_mutual_fund_data_from_excel(file_path):
    """ Loads data from Excel, cleans column names, and adds a timestamp. """
    status_update(f"Attempting to load data from Excel: {os.path.basename(file_path)}")
    try:
        if not os.path.exists(file_path):
            messagebox.showerror("Error", f"Excel file not found at {file_path}")
            status_update("Error: Excel file not found.")
            return None
        df = pd.read_excel(file_path)
        status_update("Successfully loaded data from Excel!")
        df.columns = df.columns.str.replace(r'[^A-Za-z0-9_]+', '_', regex=True).str.replace(r'_+', '_', regex=True).str.strip('_')
        status_update("Cleaned column names.")
        df['Date_Loaded'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        status_update(f"Added 'Date_Loaded' column: {df['Date_Loaded'].iloc[0]}")
        return df
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred loading Excel:\n{e}")
        status_update(f"Error loading Excel: {e}")
        return None

def save_data_to_sqlite(df, db_path, table_name):
    """ Saves the DataFrame to SQLite, APPENDING data. """
    if df is None:
        status_update("No data to save.")
        return False
    status_update(f"Appending data to SQLite table '{table_name}'...")
    conn = None
    try:
        conn = sqlite3.connect(db_path)
        df.to_sql(table_name, conn, if_exists='append', index=False)
        conn.close()
        status_update(f"Successfully appended data to table '{table_name}'.")
        return True
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred saving to database:\n{e}")
        status_update(f"Error saving/appending to DB: {e}")
        if conn:
            conn.close()
        return False

def query_data_from_sqlite(db_path, table_name, query, params=None):
    """ Generic function to query data from SQLite using a provided query string and parameters. """
    conn = None
    try:
        if not os.path.exists(db_path):
             messagebox.showerror("Error", f"Database file not found at {db_path}")
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
        messagebox.showerror("Error", f"DB Query Error: {e}")
        status_update(f"Error querying DB: {e}")
        if conn: conn.close()
        return None

def show_category_chart():
    """ Queries category counts for the LATEST data and displays a bar chart. """
    status_update("Generating category chart...")
    latest_date_query = f"SELECT MAX(Date_Loaded) FROM {table_name}"
    df_latest_date = query_data_from_sqlite(db_path=db_file_path, table_name=table_name, query=latest_date_query)
    if df_latest_date is None or df_latest_date.empty or df_latest_date.iloc[0,0] is None:
         messagebox.showerror("Error", "Could not determine the latest data load time.")
         status_update("Error: Cannot find latest load time for chart.")
         return
    latest_timestamp = df_latest_date.iloc[0,0]
    status_update(f"Latest data timestamp for chart: {latest_timestamp}")
    chart_query = f"""
        SELECT Sub_Category, COUNT(*) as Count
        FROM {table_name}
        WHERE Date_Loaded = ?
        GROUP BY Sub_Category ORDER BY Count DESC LIMIT 15 """
    df_chart = query_data_from_sqlite(db_path=db_file_path, table_name=table_name, query=chart_query, params=(latest_timestamp,))
    if df_chart is None or df_chart.empty:
        messagebox.showinfo("Info", "No category data found for the latest timestamp.")
        status_update("No category data found for chart.")
        return
    status_update("Chart data queried successfully.")
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.bar(df_chart['Sub_Category'], df_chart['Count'])
    ax.set_xlabel('Sub Category')
    ax.set_ylabel('Number of Funds')
    ax.set_title(f'Top 15 Fund Categories by Count (as of {latest_timestamp})')
    plt.xticks(rotation=90)
    plt.tight_layout()
    chart_window = tk.Toplevel(root)
    chart_window.title("Fund Category Chart")
    chart_window.geometry("800x600")
    canvas = FigureCanvasTkAgg(fig, master=chart_window)
    canvas_widget = canvas.get_tk_widget()
    canvas_widget.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
    canvas.draw()
    status_update("Chart displayed.")


# --- GUI Functions ---

def browse_excel_file():
    """ Opens dialog to select Excel file. """
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if file_path:
        excel_path_var.set(file_path)
        status_update(f"Selected Excel file: {os.path.basename(file_path)}")

def process_data():
    """ Loads Excel and appends to DB. """
    excel_file = excel_path_var.get()
    if not excel_file:
        messagebox.showwarning("Missing Input", "Please select an Excel file first.")
        return
    # Clear Treeview before processing
    clear_treeview()
    df = load_mutual_fund_data_from_excel(excel_file)
    if df is not None:
        save_successful = save_data_to_sqlite(df, db_file_path, table_name)
        if save_successful:
            status_update("--- Pipeline Complete: Excel to SQLite append successful! ---")
            messagebox.showinfo("Success", "Data loaded from Excel and appended to database successfully!")
        else:
             status_update("--- Pipeline Failed: Could not append to SQLite. ---")
    else:
        status_update("--- Pipeline Failed: Could not load from Excel. ---")


# --- MODIFIED: Function to display data in Treeview ---
def display_sample_data():
    """ Queries latest data and shows it in the Treeview table. """
    query = f"""
        SELECT Name, Sub_Category, AUM, NAV, Expense_Ratio, Date_Loaded
        FROM {table_name}
        ORDER BY Date_Loaded DESC
        LIMIT 50
    """ # Increased limit slightly
    status_update("Querying latest sample data for table...")
    df_sample = query_data_from_sqlite(db_path=db_file_path, table_name=table_name, query=query)

    # Clear existing data in the Treeview
    clear_treeview()

    if df_sample is not None:
        if not df_sample.empty:
            # Configure Treeview columns (do this only once? No, needs columns from df)
            # Get columns from the DataFrame - they MUST match the query
            tree["columns"] = list(df_sample.columns)
            tree["show"] = "headings" # Hide the default first empty column

            # Configure headings
            for col in df_sample.columns:
                tree.heading(col, text=col)
                # Set column width (optional, adjust as needed)
                if col == "Name":
                    tree.column(col, width=250, anchor='w')
                elif col == "Sub_Category":
                     tree.column(col, width=150, anchor='w')
                elif col == "Date_Loaded":
                     tree.column(col, width=130, anchor='center')
                else:
                    tree.column(col, width=80, anchor='e') # Align numbers right

            # Insert data rows
            for index, row in df_sample.iterrows():
                # Convert row to tuple for insertion
                tree.insert("", tk.END, values=tuple(row))

            status_update(f"Displayed latest {len(df_sample)} data rows in table.")
        else:
             status_update("No data found in database table.")
             # Optionally display a message in the treeview area?
    else:
        status_update("Failed to retrieve data for table.")

def clear_treeview():
     """ Clears all items from the Treeview. """
     tree.delete(*tree.get_children())
     # You might want to clear column definitions too if queries change columns
     # tree["columns"] = []


def status_update(message):
    """ Updates the status bar label. """
    status_var.set(message)
    print(message) # Also print to console


# --- Main Application Setup (Tkinter Window) ---
if __name__ == "__main__":
    root = tk.Tk()
    root.title("MF Analyzer - Loader V4 (Table View)") # Updated Title
    root.geometry("900x700") # Increased size slightly

    excel_path_var = tk.StringVar(value=os.path.join(script_dir, default_excel_file_name))
    status_var = tk.StringVar(value="Ready.")

    style = ttk.Style()
    style.theme_use('clam') # Or 'alt', 'default', 'classic'

    # --- Control Frame (Top) ---
    control_frame = ttk.Frame(root, padding="10")
    control_frame.pack(side=tk.TOP, fill=tk.X, pady=5)

    ttk.Label(control_frame, text="Excel File:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
    excel_entry = ttk.Entry(control_frame, textvariable=excel_path_var, width=50)
    excel_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
    browse_button = ttk.Button(control_frame, text="Browse...", command=browse_excel_file)
    browse_button.grid(row=0, column=2, padx=5, pady=5)

    load_button = ttk.Button(control_frame, text="Load Excel & Append to DB", command=process_data)
    load_button.grid(row=1, column=0, padx=5, pady=10, sticky="w")

    display_button = ttk.Button(control_frame, text="Show Latest DB Data", command=display_sample_data)
    display_button.grid(row=1, column=1, padx=5, pady=10, sticky="w")

    chart_button = ttk.Button(control_frame, text="Show Category Chart", command=show_category_chart)
    chart_button.grid(row=1, column=2, padx=5, pady=10, sticky="w")

    control_frame.columnconfigure(1, weight=1) # Make entry expand

    # --- Treeview Frame (Middle) ---
    tree_frame = ttk.Frame(root, padding="10")
    tree_frame.pack(fill=tk.BOTH, expand=True)

    # Create the Treeview widget
    tree = ttk.Treeview(tree_frame)

    # Add Scrollbars
    vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
    hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

    # Grid layout for Treeview and scrollbars
    tree.grid(row=0, column=0, sticky='nsew')
    vsb.grid(row=0, column=1, sticky='ns')
    hsb.grid(row=1, column=0, sticky='ew')

    # Configure resizing behavior
    tree_frame.grid_rowconfigure(0, weight=1)
    tree_frame.grid_columnconfigure(0, weight=1)


    # --- Status Bar (Bottom) ---
    status_bar = ttk.Label(root, textvariable=status_var, relief=tk.SUNKEN, anchor=tk.W, padding="2 5")
    status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    # --- Run ---
    root.mainloop()