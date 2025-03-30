import tkinter as tk
from tkinter import ttk # For themed widgets (looks slightly better)
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
import os
import sqlite3

# --- Configuration ---
# We might select the Excel file via dialog now, but keep default/db names
default_excel_file_name = 'mutual_funds.xlsx'
db_file_name = 'mf_data.db'
table_name = 'mutual_funds'

script_dir = os.path.dirname(__file__)
db_file_path = os.path.join(script_dir, db_file_name)

# --- Database/Excel Functions (Modified slightly for GUI context) ---

def load_mutual_fund_data_from_excel(file_path):
    """ Loads data from Excel and cleans column names. """
    status_update(f"Attempting to load data from Excel: {os.path.basename(file_path)}")
    try:
        if not os.path.exists(file_path):
            messagebox.showerror("Error", f"Excel file not found at {file_path}")
            status_update("Error: Excel file not found.")
            return None
        df = pd.read_excel(file_path)
        status_update("Successfully loaded data from Excel!")
        # Clean column names
        df.columns = df.columns.str.replace(r'[^A-Za-z0-9_]+', '_', regex=True).str.replace(r'_+', '_', regex=True).str.strip('_')
        status_update("Cleaned column names.")
        return df
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred loading Excel:\n{e}")
        status_update(f"Error loading Excel: {e}")
        return None

def save_data_to_sqlite(df, db_path, table_name):
    """ Saves the DataFrame to SQLite, replacing the table. """
    if df is None:
        status_update("No data to save.")
        return False
    status_update(f"Saving data to SQLite table '{table_name}'...")
    conn = None
    try:
        conn = sqlite3.connect(db_path)
        df.to_sql(table_name, conn, if_exists='replace', index=False)
        conn.close()
        status_update(f"Successfully saved data to table '{table_name}'.")
        return True
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred saving to database:\n{e}")
        status_update(f"Error saving to DB: {e}")
        if conn:
            conn.close()
        return False

def query_data_from_sqlite(db_path, table_name, num_rows=5):
    """ Queries sample data from SQLite. """
    status_update(f"Querying sample data from table '{table_name}'...")
    conn = None
    try:
        if not os.path.exists(db_path):
             messagebox.showerror("Error", f"Database file not found at {db_path}")
             status_update("Error: Database file not found.")
             return None
        conn = sqlite3.connect(db_path)
        # Query first few rows, limited columns
        query = f"SELECT Name, Sub_Category, AUM, NAV, Expense_Ratio FROM {table_name} LIMIT {num_rows}"
        df_from_db = pd.read_sql_query(query, conn)
        conn.close()
        status_update(f"Successfully queried {len(df_from_db)} rows.")
        return df_from_db
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred querying the database:\n{e}")
        status_update(f"Error querying DB: {e}")
        if conn:
            conn.close()
        return None

# --- GUI Functions ---

def browse_excel_file():
    """ Opens a dialog to select an Excel file and updates the entry field. """
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if file_path:
        excel_path_var.set(file_path)
        status_update(f"Selected Excel file: {os.path.basename(file_path)}")

def process_data():
    """ Main function called by the 'Load & Save' button. """
    excel_file = excel_path_var.get()
    if not excel_file:
        messagebox.showwarning("Missing Input", "Please select an Excel file first.")
        return

    # Clear previous results
    display_area.config(state=tk.NORMAL)
    display_area.delete('1.0', tk.END)
    display_area.config(state=tk.DISABLED)

    # Run the pipeline
    df = load_mutual_fund_data_from_excel(excel_file)
    if df is not None:
        save_successful = save_data_to_sqlite(df, db_file_path, table_name)
        if save_successful:
            status_update("--- Pipeline Complete: Excel to SQLite successful! ---")
            messagebox.showinfo("Success", "Data loaded from Excel and saved to database successfully!")
        else:
             status_update("--- Pipeline Failed: Could not save to SQLite. ---")
             messagebox.showerror("Error", "Failed to save data to the database.")
    else:
        status_update("--- Pipeline Failed: Could not load from Excel. ---")
        # Error message shown within load function

def display_sample_data():
    """ Queries data from DB and shows it in the text area. """
    df_sample = query_data_from_sqlite(db_file_path, table_name, num_rows=10)

    display_area.config(state=tk.NORMAL) # Enable editing
    display_area.delete('1.0', tk.END) # Clear previous content

    if df_sample is not None:
        if not df_sample.empty:
            display_area.insert(tk.END, df_sample.to_string()) # Insert DataFrame as string
            status_update("Displayed sample data from database.")
        else:
             display_area.insert(tk.END,"No data found in the database table.")
             status_update("Database table exists but is empty.")
    else:
        display_area.insert(tk.END, "Failed to retrieve data from database.")
        # Status updated within query function

    display_area.config(state=tk.DISABLED) # Disable editing

def status_update(message):
    """ Updates the status bar label. """
    status_var.set(message)
    print(message) # Also print to console for debugging

# --- Main Application Setup ---
if __name__ == "__main__":
    root = tk.Tk()
    root.title("MF Analyzer - Basic Loader")
    root.geometry("800x600") # Set initial window size

    # --- Variables ---
    excel_path_var = tk.StringVar(value=os.path.join(script_dir, default_excel_file_name)) # Default path
    status_var = tk.StringVar(value="Ready.")

    # --- Style --- (Optional, makes it look slightly less basic)
    style = ttk.Style()
    style.theme_use('clam') # Try 'clam', 'alt', 'default', 'classic'

    # --- Widgets ---
    # Frame for controls
    control_frame = ttk.Frame(root, padding="10")
    control_frame.pack(side=tk.TOP, fill=tk.X)

    # Excel file selection
    ttk.Label(control_frame, text="Excel File:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
    excel_entry = ttk.Entry(control_frame, textvariable=excel_path_var, width=60)
    excel_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
    browse_button = ttk.Button(control_frame, text="Browse...", command=browse_excel_file)
    browse_button.grid(row=0, column=2, padx=5, pady=5)

    # Action buttons
    load_button = ttk.Button(control_frame, text="Load Excel & Save to DB", command=process_data)
    load_button.grid(row=1, column=0, columnspan=2, padx=5, pady=10, sticky="w")

    display_button = ttk.Button(control_frame, text="Show Sample DB Data", command=display_sample_data)
    display_button.grid(row=1, column=2, padx=5, pady=10, sticky="w")

    control_frame.columnconfigure(1, weight=1) # Make entry field expand

    # Frame for display area
    display_frame = ttk.Frame(root, padding="10")
    display_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    # Display area (ScrolledText)
    display_area = scrolledtext.ScrolledText(display_frame, wrap=tk.WORD, height=20, state=tk.DISABLED) # Start disabled
    display_area.pack(fill=tk.BOTH, expand=True)

    # Status bar
    status_bar = ttk.Label(root, textvariable=status_var, relief=tk.SUNKEN, anchor=tk.W, padding="2 5")
    status_bar.pack(side=tk.BOTTOM, fill=tk.X)


    # --- Run the Application ---
    root.mainloop() # Start the Tkinter event loop