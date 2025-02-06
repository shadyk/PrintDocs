import pandas as pd
from docxtpl import DocxTemplate
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os

# Global DataFrame for storing Excel data
df = None

# Function to load the Excel file
def load_excel(file_path):
    global df
    try:
        df = pd.read_excel(file_path)
        messagebox.showinfo("Success", "Excel file loaded successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load Excel file: {e}")

# Function to search for data in the DataFrame
def search_data(query):
    if df is None:
        messagebox.showwarning("Warning", "Please load an Excel file first.")
        return None
    try:
        results = df[df.apply(lambda row: row.astype(str).str.contains(query, case=False).any(), axis=1)]
        return results
    except Exception as e:
        messagebox.showerror("Error", f"Search failed: {e}")

# Function to set RTL alignment for a paragraph in the Word document
def set_rtl_alignment(paragraph):
    # Set alignment to right
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    # Set text direction to RTL
    paragraph_format = paragraph.paragraph_format
    bidi = OxmlElement('w:bidi')
    bidi.set('w:val', '1')
    paragraph_format.element.append(bidi)

# Function to fill the Word template with Arabic data
def fill_template(template_path, output_path, context):
    try:
        doc = DocxTemplate(template_path)
        doc.render(context)

        # Apply RTL alignment to all paragraphs in the document
        for paragraph in doc.docx.paragraphs:
            set_rtl_alignment(paragraph)

        doc.save(output_path)
        messagebox.showinfo("Success", f"Document saved as: {output_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to fill template: {e}")

# Function to print the document
def print_document(file_path):
    try:
        if os.name == 'nt':  # Windows
            os.startfile(file_path, 'print')
        else:  # macOS/Linux
            os.system(f'lpr {file_path}')
        messagebox.showinfo("Success", "Document sent to printer.")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to print document: {e}")

# Function to handle document generation
def generate_document(selected_row):
    if selected_row is None:
        messagebox.showwarning("Warning", "Please select a row.")
        return

    # Define the Word template and output file paths
    template_path = filedialog.askopenfilename(title="Select Word Template", filetypes=[("Word Files", "*.docx")])
    if not template_path:
        return
    output_path = filedialog.asksaveasfilename(title="Save Document As", defaultextension=".docx",
                                               filetypes=[("Word Files", "*.docx")])
    if not output_path:
        return

    # Create a context dictionary from the selected row
    context = dict(zip(df.columns, selected_row))
    fill_template(template_path, output_path, context)

    # Option to print
    if messagebox.askyesno("Print", "Do you want to print the document?"):
        print_document(output_path)

# GUI Application
def create_gui():
    def browse_excel():
        file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            load_excel(file_path)
            update_table()

    def search():
        query = search_entry.get()
        if query:
            results = search_data(query)
            update_table(results)

    def update_table(data=None):
        for row in table.get_children():
            table.delete(row)
        if data is None and df is not None:
            data = df
        if data is not None:
            for idx, row in data.iterrows():
                table.insert("", "end", values=list(row))

    def generate():
        selected_item = table.selection()
        if selected_item:
            selected_row = table.item(selected_item)['values']
            if selected_row:
                generate_document(selected_row)

    root = tk.Tk()
    root.title("Document Generator (Arabic Support)")

    # Search bar
    tk.Label(root, text="Search:").pack(pady=5)
    search_frame = tk.Frame(root)
    search_frame.pack()
    search_entry = tk.Entry(search_frame, width=40)
    search_entry.pack(side="left", padx=5)
    tk.Button(search_frame, text="Search", command=search).pack(side="left")

    # Table for displaying Excel data
    table = ttk.Treeview(root, show="headings", height=10)
    table.pack(padx=10, pady=10, fill="both", expand=True)
    scrollbar = ttk.Scrollbar(root, orient="vertical", command=table.yview)
    scrollbar.pack(side="right", fill="y")
    table.configure(yscroll=scrollbar.set)

    # Buttons
    button_frame = tk.Frame(root)
    button_frame.pack(pady=10)
    tk.Button(button_frame, text="Load Excel", command=browse_excel).pack(side="left", padx=5)
    tk.Button(button_frame, text="Generate Document", command=generate).pack(side="left", padx=5)
    tk.Button(button_frame, text="Quit", command=root.quit).pack(side="left", padx=5)

    root.mainloop()

# Run the GUI
if __name__ == "__main__":
    create_gui()