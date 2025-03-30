import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
import os
import sqlite3
from datetime import datetime

# --- Matplotlib Imports ---
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# --- Configuration ---
default_excel_file_name = 'mutual_funds.xlsx'
db_file_name = 'mf_data.db'
table_name = 'mutual_funds'

script_dir = os.path.dirname(__file__)
db_file_path = os.path.join(script_dir, db_file_name)

# --- Database/Excel Functions (load, save, query - slightly modified query) ---

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

def query_data_from_sqlite(db_path, table_name, query):
    """ Generic function to query data from SQLite using a provided query string. """
    conn = None
    try:
        if not os.path.exists(db_path):
             messagebox.showerror("Error", f"Database file not found at {db_path}")
             status_update("Error: Database file not found.")
             return None
        conn = sqlite3.connect(db_path)
        print(f"Executing query: {query}") # Keep console log
        df_from_db = pd.read_sql_query(query, conn)
        conn.close()
        return df_from_db
    except Exception as e:
        messagebox.showerror("Error", f"DB Query Error: {e}")
        status_update(f"Error querying DB: {e}")
        if conn: conn.close()
        return None


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
    display_area.config(state=tk.NORMAL)
    display_area.delete('1.0', tk.END)
    display_area.config(state=tk.DISABLED)
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


def display_sample_data():
    """ Queries latest data and shows in text area. """
    # Query latest 20 rows
    query = f"""
        SELECT Name, Sub_Category, AUM, NAV, Expense_Ratio, Date_Loaded
        FROM {table_name}
        ORDER BY Date_Loaded DESC
        LIMIT 20
    """
    status_update("Querying latest sample data...")
    df_sample = query_data_from_sqlite(db_path=db_file_path, table_name=table_name, query=query)

    display_area.config(state=tk.NORMAL)
    display_area.delete('1.0', tk.END)
    if df_sample is not None:
        if not df_sample.empty:
            pd.set_option('display.max_rows', None)
            display_area.insert(tk.END, df_sample.to_string())
            pd.reset_option('display.max_rows')
            status_update("Displayed latest sample data.")
        else:
             display_area.insert(tk.END,"No data found.")
             status_update("No data found in table.")
    else:
        display_area.insert(tk.END, "Failed to retrieve data.")
    display_area.config(state=tk.DISABLED)

# --- NEW CHARTING FUNCTION ---
def show_category_chart():
    """ Queries category counts for the LATEST data and displays a bar chart. """
    status_update("Generating category chart...")

    # 1. Find the latest Date_Loaded timestamp
    latest_date_query = f"SELECT MAX(Date_Loaded) FROM {table_name}"
    df_latest_date = query_data_from_sqlite(db_path=db_file_path, table_name=table_name, query=latest_date_query)

    if df_latest_date is None or df_latest_date.empty or df_latest_date.iloc[0,0] is None:
         messagebox.showerror("Error", "Could not determine the latest data load time.")
         status_update("Error: Cannot find latest load time for chart.")
         return
    latest_timestamp = df_latest_date.iloc[0,0]
    status_update(f"Latest data timestamp for chart: {latest_timestamp}")

    # 2. Query category counts for that specific timestamp
    chart_query = f"""
        SELECT Sub_Category, COUNT(*) as Count
        FROM {table_name}
        WHERE Date_Loaded = ?
        GROUP BY Sub_Category
        ORDER BY Count DESC
        LIMIT 15
    """ # Using parameterized query placeholder '?'

    conn = None
    try:
        conn = sqlite3.connect(db_file_path)
        # Use params argument in read_sql_query for safety
        df_chart = pd.read_sql_query(chart_query, conn, params=(latest_timestamp,))
        conn.close()

        if df_chart is None or df_chart.empty:
            messagebox.showinfo("Info", "No category data found for the latest timestamp.")
            status_update("No category data found for chart.")
            return

        status_update("Chart data queried successfully.")

        # 3. Create the plot using Matplotlib
        fig, ax = plt.subplots(figsize=(10, 6)) # Create figure and axes
        ax.bar(df_chart['Sub_Category'], df_chart['Count'])
        ax.set_xlabel('Sub Category')
        ax.set_ylabel('Number of Funds')
        ax.set_title(f'Top 15 Fund Categories by Count (as of {latest_timestamp})')
        plt.xticks(rotation=90) # Rotate x-axis labels for readability
        plt.tight_layout() # Adjust layout

        # 4. Display plot in a new Tkinter window
        chart_window = tk.Toplevel(root) # Create a new window
        chart_window.title("Fund Category Chart")
        chart_window.geometry("800x600")

        canvas = FigureCanvasTkAgg(fig, master=chart_window) # Create canvas
        canvas_widget = canvas.get_tk_widget()
        canvas_widget.pack(side=tk.TOP, fill=tk.BOTH, expand=True) # Pack canvas widget
        canvas.draw() # Draw the chart

        status_update("Chart displayed.")

    except Exception as e:
        messagebox.showerror("Chart Error", f"Could not generate chart:\n{e}")
        status_update(f"Chart error: {e}")
        if conn: conn.close()


def status_update(message):
    """ Updates the status bar label. """
    status_var.set(message)
    print(message) # Also print to console


# --- Main Application Setup (Tkinter Window) ---
if __name__ == "__main__":
    root = tk.Tk()
    root.title("MF Analyzer - Loader V3 (Chart)")
    root.geometry("800x600")

    excel_path_var = tk.StringVar(value=os.path.join(script_dir, default_excel_file_name))
    status_var = tk.StringVar(value="Ready.")

    style = ttk.Style()
    style.theme_use('clam')

    # --- Control Frame ---
    control_frame = ttk.Frame(root, padding="10")
    control_frame.pack(side=tk.TOP, fill=tk.X)

    ttk.Label(control_frame, text="Excel File:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
    excel_entry = ttk.Entry(control_frame, textvariable=excel_path_var, width=50) # Adjusted width
    excel_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
    browse_button = ttk.Button(control_frame, text="Browse...", command=browse_excel_file)
    browse_button.grid(row=0, column=2, padx=5, pady=5)

    load_button = ttk.Button(control_frame, text="Load Excel & Append to DB", command=process_data)
    load_button.grid(row=1, column=0, padx=5, pady=10, sticky="w")

    display_button = ttk.Button(control_frame, text="Show Latest DB Data", command=display_sample_data)
    display_button.grid(row=1, column=1, padx=5, pady=10, sticky="w")

    # --- NEW CHART BUTTON ---
    chart_button = ttk.Button(control_frame, text="Show Category Chart", command=show_category_chart)
    chart_button.grid(row=1, column=2, padx=5, pady=10, sticky="w") # Placed next to display button

    control_frame.columnconfigure(1, weight=1)

    # --- Display Frame ---
    display_frame = ttk.Frame(root, padding="10")
    display_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
    display_area = scrolledtext.ScrolledText(display_frame, wrap=tk.WORD, height=15, state=tk.DISABLED) # Reduced height slightly
    display_area.pack(fill=tk.BOTH, expand=True)

    # --- Status Bar ---
    status_bar = ttk.Label(root, textvariable=status_var, relief=tk.SUNKEN, anchor=tk.W, padding="2 5")
    status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    root.mainloop()