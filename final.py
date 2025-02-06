import os
import pandas as pd
from docx import Document
import tkinter as tk
from tkinter import ttk, messagebox, font

# Function to fill the Word template
def fill_template(template_path, data_row, output_path):
    """Fill the Word document template with data from the selected row."""
    if not os.path.exists(template_path):
        messagebox.showerror("Error", f"Template file not found: {template_path}")
        return

    try:
        doc = Document(template_path)
    except Exception as e:
        messagebox.showerror("Error", f"Error loading document: {e}")
        return

    for key, value in data_row.items():
        key = key.strip()
        for paragraph in doc.paragraphs:
            if f'{{{key}}}' in paragraph.text:
                paragraph.text = paragraph.text.replace(f'{{{key}}}', str(value))
    
    doc.save(output_path)
    messagebox.showinfo("Success", f"Document saved to {output_path}")

# Function to generate a document for the selected row
def generate_release_document():
    selected_index = listbox.curselection()
    if not selected_index:
        messagebox.showwarning("Warning", "Please select a row!")
        return
    selectedTemplate = release_situation_f
    selected_row = df.iloc[selected_index[0]]
    output_path = os.path.join(output_directory, f"اطلاق حال{selected_row.name + 1}.docx")
    for key, value in selected_row.items():
       key = key.strip()
       if key == 'Gender':
           if value == 'M':
               selectedTemplate = release_situation_m

    fill_template(selectedTemplate, selected_row, output_path)

def generate_baptisim_document():
    selected_index = listbox.curselection()
    if not selected_index:
        messagebox.showwarning("Warning", "Please select a row!")
        return
    selectedTemplate = baptisim_template_f
    selected_row = df.iloc[selected_index[0]]
    output_path = os.path.join(output_directory, f"معمودية{selected_row.name + 1}.docx")
    for key, value in selected_row.items():
       key = key.strip()
       if key == 'Gender':
           if value == 'M':
               selectedTemplate = baptisim_template_m
    fill_template(selectedTemplate, selected_row, output_path)

# Function to filter rows based on search input
def search_rows():
    search_term = search_var.get().strip().lower()
    if not search_term:
        listbox.delete(0, tk.END)
        for index, row in df.iterrows():
            listbox.insert(tk.END, " - ".join(map(str, row.values)))
        return

    filtered_rows = df.apply(lambda row: row.astype(str).str.contains(search_term, case=False).any(), axis=1)
    listbox.delete(0, tk.END)
    for index, row in df[filtered_rows].iterrows():
            listbox.insert(tk.END, " - ".join(map(str, row.values)))
# path = ''
path = 'Desktop/print/'

# Load Excel data
excel_file = path + 'sample.xlsx'
baptisim_template_m = path + 'baptisim_template_m.docx'  
baptisim_template_f = path +'baptisim_template_f.docx'  
release_situation_m = path +'release_situation_m.docx'  
release_situation_f = path + 'release_situation_f.docx'  
output_directory = path + 'output_docs'

if not os.path.exists(excel_file):
    raise FileNotFoundError(f"Excel file not found: {excel_file}")


if not os.path.exists(output_directory):
    os.makedirs(output_directory)

df = pd.read_excel(excel_file, header=1)
df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

# Create the Tkinter UI
root = tk.Tk()
root.title("اطلاق حال وورقات معموديات")

# Search bar
search_var = tk.StringVar()
search_entry = ttk.Entry(root, textvariable=search_var, width=50)
search_entry.pack(pady=10)
search_entry.bind("<KeyRelease>", lambda event: search_rows())

# Listbox to display rows
custom_font = font.Font(family="Traditional Arabic", size=2)  # Adjust size and font family as needed

listbox = tk.Listbox(root, width=100, height=40, font=custom_font, justify="right")
listbox.pack(pady=20)

# Populate the listbox with all rows initially
for index, row in df.iterrows():
    listbox.insert(tk.END, " - ".join(map(str, row.values)))

# Generate document button
generate_button = ttk.Button(root, text="اطلاق حال", command=generate_release_document)
generate_button.pack(pady=10)
generate_button = ttk.Button(root, text="ورقة معمودية", command=generate_baptisim_document)
generate_button.pack(pady=10)
# Run the Tkinter event loop
root.mainloop()
