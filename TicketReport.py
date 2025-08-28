import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
import csv
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from collections import defaultdict, Counter
from datetime import datetime, timedelta

# Open and read CSV file, store data globally
def open_file():
    global data, x_var, series_var, x_menu, series_menu, owner_var, owner_menu

    file = fd.askopenfilename(title="Open CSV File", filetypes=[("CSV files", "*.csv")])
    with open(str(file), newline='') as f:
        data = list(csv.reader(f))

    # Update dropdowns with new column names once file is loaded
    columns = data[0]
    x_var.set(columns[0])
    series_var.set(columns[0])
    x_menu['menu'].delete(0, 'end')
    series_menu['menu'].delete(0, 'end')
    for col in columns:
        x_menu['menu'].add_command(label=col, command=tk._setit(x_var, col))
        series_menu['menu'].add_command(label=col, command=tk._setit(series_var, col))

    # Populate owner dropdown
    if "Owner" in columns:
        owner_idx = columns.index("Owner")
        owners = sorted(set(row[owner_idx] for row in data[1:] if len(row) > owner_idx))
        owner_var.set(owners[0] if owners else "")
        owner_menu['menu'].delete(0, 'end')
        for owner in owners:
            owner_menu['menu'].add_command(label=owner, command=tk._setit(owner_var, owner))
    else:
        owner_var.set("")
        owner_menu['menu'].delete(0, 'end')

# Function for Custom Graph output, selecing X axis and Series
def custom_graph():
    header = data[0]
    try:
        x_idx = header.index(x_var.get())
        series_idx = header.index(series_var.get())
    except ValueError:
        tk.messagebox.showerror("Error", "Selected column not found in CSV.")
        return

    rows = [row for row in data[1:] if len(row) > max(x_idx, series_idx)]
    counts = defaultdict(lambda: defaultdict(int))
    x_values = set()
    series_values = set()

    for row in rows:
        x_val = row[x_idx]
        series_val = row[series_idx]
        counts[series_val][x_val] += 1
        x_values.add(x_val)
        series_values.add(series_val)

    x_values = sorted(x_values)
    series_values = sorted(series_values)

    plt.figure(figsize=(10, 6))
    for s in series_values:
        y_counts = [counts[s][x] for x in x_values]
        plt.plot(x_values, y_counts, marker='o', label=str(s))
    plt.xlabel(x_var.get())
    plt.ylabel("Count")
    plt.title(f"Count by {x_var.get()} and {series_var.get()}")
    plt.legend(title=series_var.get())
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.show()

# Placeholder for function
def quick_output_example():
    # Example: Show total number of tickets
    total = len(data) - 1  # Exclude header
    tk.messagebox.showinfo("Quick Output", f"Total tickets: {data[1]}")

# Function for Owner Last Modified output
def owner_last_modified():
    header = data[0]
    try:
        owner_idx = header.index("Owner")
        date_idx = header.index("Modified On")
    except ValueError:
        tk.messagebox.showerror("Error", "Required columns not found in CSV.")
        return

    selected_owner = owner_var.get()
    date_grouping = date_group_var.get()

    # Filter rows for the selected owner
    rows = [row for row in data[1:] if len(row) > max(owner_idx, date_idx) and row[owner_idx] == selected_owner]

    # Parse dates using the known format
    date_list = []
    for row in rows:
        date_str = row[date_idx].strip()
        try:
            date = datetime.strptime(date_str, "%m/%d/%Y %H:%M")
            date_list.append(date)
        except Exception:
            continue

    if not date_list:
        tk.messagebox.showerror("Error", "No valid dates found for this owner.")
        return

    # Group dates
    if date_grouping == "Daily":
        grouped = [d.date() for d in date_list]
        counts = Counter(grouped)
        x_vals = sorted(counts)
        y_vals = [counts[x] for x in x_vals]
        x_fmt = mdates.DateFormatter('%Y-%m-%d')
        plt.figure(figsize=(10, 6))
        plt.bar(x_vals, y_vals, width=0.8)
        plt.gca().xaxis.set_major_formatter(x_fmt)
        plt.gcf().autofmt_xdate()
    elif date_grouping == "Weekly":
        grouped = [d.date() - timedelta(days=d.weekday()) for d in date_list]
        counts = Counter(grouped)
        x_vals = sorted(counts)
        y_vals = [counts[x] for x in x_vals]
        x_fmt = mdates.DateFormatter('Week of %Y-%m-%d')
        plt.figure(figsize=(10, 6))
        plt.bar(x_vals, y_vals, width=4)
        plt.gca().xaxis.set_major_formatter(x_fmt)
        plt.gcf().autofmt_xdate()
    elif date_grouping == "Monthly":
        grouped = [d.replace(day=1).date() for d in date_list]
        counts = Counter(grouped)
        x_vals = sorted(counts)
        y_vals = [counts[x] for x in x_vals]
        x_fmt = mdates.DateFormatter('%Y-%m')
        plt.figure(figsize=(10, 6))
        plt.bar(x_vals, y_vals, width=15)
        plt.gca().xaxis.set_major_formatter(x_fmt)
        plt.gcf().autofmt_xdate()
    else:
        grouped = [d.date() for d in date_list]
        counts = Counter(grouped)
        x_vals = sorted(counts)
        y_vals = [counts[x] for x in x_vals]
        x_fmt = mdates.DateFormatter('%Y-%m-%d')
        plt.figure(figsize=(10, 6))
        plt.bar(x_vals, y_vals, width=0.8)
        plt.gca().xaxis.set_major_formatter(x_fmt)
        plt.gcf().autofmt_xdate()

    plt.xlabel(f"{date_grouping}")
    plt.ylabel("Count")
    plt.title(f"Entries for {selected_owner} by {date_grouping}")
    plt.tight_layout()
    plt.show()

def average_ticket_age():
    header = data[0]
    try:
        owner_idx = header.index("Owner")
        created_idx = header.index("Created On")
    except ValueError:
        tk.messagebox.showerror("Error", "Required columns not found in CSV.")
        return

    today = datetime.now()
    owner_ages = defaultdict(list)

    for row in data[1:]:
        if len(row) > max(owner_idx, created_idx):
            owner = row[owner_idx]
            date_str = row[created_idx].strip()
            try:
                created = datetime.strptime(date_str, "%m/%d/%Y %H:%M")
                age_days = (today - created).days
                owner_ages[owner].append(age_days)
            except Exception:
                continue

    avg_ages = {owner: (sum(ages) / len(ages)) for owner, ages in owner_ages.items() if ages}
    if not avg_ages:
        tk.messagebox.showinfo("Info", "No valid ticket ages found.")
        return

    owners = list(avg_ages.keys())
    avg_days = list(avg_ages.values())

    plt.figure(figsize=(10, 6))
    plt.bar(owners, avg_days)
    plt.xlabel("Owner")
    plt.ylabel("Average Ticket Age (days)")
    plt.title("Average Ticket Age per Owner")
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.show()

def average_ticket_modified():
    header = data[0]
    try:
        owner_idx = header.index("Owner")
        modified_idx = header.index("Modified On")
    except ValueError:
        tk.messagebox.showerror("Error", "Required columns not found in CSV.")
        return

    today = datetime.now()
    owner_ages = defaultdict(list)

    for row in data[1:]:
        if len(row) > max(owner_idx, modified_idx):
            owner = row[owner_idx]
            date_str = row[modified_idx].strip()
            try:
                modified = datetime.strptime(date_str, "%m/%d/%Y %H:%M")
                age_days = (today - modified).days
                owner_ages[owner].append(age_days)
            except Exception:
                continue

    avg_ages = {owner: (sum(ages) / len(ages)) for owner, ages in owner_ages.items() if ages}
    if not avg_ages:
        tk.messagebox.showinfo("Info", "No valid modified ages found.")
        return

    owners = list(avg_ages.keys())
    avg_days = list(avg_ages.values())

    plt.figure(figsize=(10, 6))
    plt.bar(owners, avg_days)
    plt.xlabel("Owner")
    plt.ylabel("Average Age Since Last Modified (days)")
    plt.title("Average Age Since Last Modified per Owner")
    plt.xticks(rotation=45, ha='right')
    plt.tight_layout()
    plt.show()

# Function to run the selected tool
def run_tool():
    method = output_var.get()
    if method == "Custom Graph":
        custom_graph()
    elif method == "Average Ticket Age":
        average_ticket_age()
    elif method == "Owner Last Modified":
        owner_last_modified()
    elif method == "Average Ticket Modified":
        average_ticket_modified()
    # Add more elifs for other output methods

# Function to show/hide input fields based on output method
def update_input_visibility(*args):
    method = output_var.get()
    if method == "Custom Graph":
        x_label.place(relx=0.33, rely=0.66, anchor=tk.CENTER)
        x_menu.place(relx=0.33, rely=0.76, anchor=tk.CENTER)
        series_label.place(relx=0.66, rely=0.66, anchor=tk.CENTER)
        series_menu.place(relx=0.66, rely=0.78, anchor=tk.CENTER)
        owner_label.place_forget()
        owner_menu.place_forget()
        date_group_label.place_forget()
        date_group_menu.place_forget()
    elif method == "Owner Last Modified":
        x_label.place_forget()
        x_menu.place_forget()
        series_label.place_forget()
        series_menu.place_forget()
        owner_label.place(relx=0.25, rely=0.66, anchor=tk.CENTER)
        owner_menu.place(relx=0.25, rely=0.76, anchor=tk.CENTER)
        date_group_label.place(relx=0.66, rely=0.66, anchor=tk.CENTER)
        date_group_menu.place(relx=0.66, rely=0.76, anchor=tk.CENTER)
    else:
        x_label.place_forget()
        x_menu.place_forget()
        series_label.place_forget()
        series_menu.place_forget()
        owner_label.place_forget()
        owner_menu.place_forget()
        date_group_label.place_forget()
        date_group_menu.place_forget()

# Initialize main window
root = tk.Tk()
root.title('Ticket Report Tool')
root.geometry('600x200')
frame = ttk.Frame(root, width=500, height=300)
frame.grid(row=0, column=0)

# Title and Version
title = ttk.Label(text='Ticket Report Tool', font=('Arial', 14, 'bold'))
title.place(relx=0.5, rely=0.15, anchor=tk.CENTER)
version = ttk.Label(text='V1.0', font=('Arial', 11))
version.place(relx=0.5, rely=0.26, anchor=tk.CENTER)

# File selection button
fileButton = ttk.Button(text='Open a file', command=open_file)
fileButton.place(relx=0.66, rely=0.42, anchor=tk.CENTER)

# Run button
runButton = ttk.Button(text='Run tool', command=run_tool)
runButton.place(relx=0.5, rely=0.90, anchor=tk.CENTER)

# Dropdowns for X axis and Series
x_var = tk.StringVar()
series_var = tk.StringVar()
x_menu = ttk.OptionMenu(root, x_var, "")
series_menu = ttk.OptionMenu(root, series_var, "")
x_label = ttk.Label(text='X Axis:')
x_label.place(relx=0.33, rely=0.66, anchor=tk.CENTER)
x_menu.place(relx=0.33, rely=0.76, anchor=tk.CENTER)
series_label = ttk.Label(text='Line Series:')
series_label.place(relx=0.66, rely=0.66, anchor=tk.CENTER)
series_menu.place(relx=0.66, rely=0.78, anchor=tk.CENTER)

# Dropdown for Output Method
output_var = tk.StringVar(value="Custom Graph")
output_methods = ["Custom Graph", "Average Ticket Age", "Average Ticket Modified", "Owner Last Modified"] # Add more methods as needed
output_menu = ttk.OptionMenu(root, output_var, output_methods[0], *output_methods)
output_label = ttk.Label(text='Output Method:')
output_label.place(relx=0.33, rely=0.36, anchor=tk.CENTER)
output_menu.place(relx=0.33, rely=0.48, anchor=tk.CENTER)

# Owner selection (for Owner Last Modified)
owner_var = tk.StringVar()
owner_menu = ttk.OptionMenu(root, owner_var, "")
owner_label = ttk.Label(text='Owner:')

# Date grouping (for Owner Last Modified)
date_group_var = tk.StringVar(value="Daily")
date_group_menu = ttk.OptionMenu(root, date_group_var, "Daily", "Daily", "Weekly", "Monthly")
date_group_label = ttk.Label(text='Date Grouping:')

# Place owner and date grouping menus (initially hidden)
owner_label.place(relx=0.25, rely=0.66, anchor=tk.CENTER)
owner_menu.place(relx=0.25, rely=0.76, anchor=tk.CENTER)
date_group_label.place(relx=0.66, rely=0.66, anchor=tk.CENTER)
date_group_menu.place(relx=0.66, rely=0.76, anchor=tk.CENTER)

# Hide Custom Graph inputs when not selected
output_var.trace_add("write", update_input_visibility)
update_input_visibility()

# Start the GUI event loop
root.mainloop()