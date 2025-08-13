#!/usr/bin/env python3
"""
Quick Script made with Claude 4

PDF Form Auto-Filler
A simple GUI application to automatically fill multiple PDF forms with the same data.

Requirements:
- Python 3.6+
- tkinter (install with: pip install tk)
- PyPDF2 (install with: pip install PyPDF2)

Usage:
1. Run the script
2. Select the folder containing your PDF forms
3. Choose one PDF to use as a template for field mapping
4. Fill in the form fields in the GUI
5. Click "Fill All PDFs" to process all forms in the folder
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import sys
from pathlib import Path
import traceback

try:
    import PyPDF2
except ImportError:
    print("PyPDF2 is required. Install it with: pip install PyPDF2")
    sys.exit(1)


class PDFFormFiller:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Form Auto-Filler")
        self.root.geometry("800x600")
        
        # Variables
        self.pdf_folder = tk.StringVar()
        self.template_pdf = tk.StringVar()
        self.field_entries = {}
        self.pdf_files = []
        
        self.create_widgets()
        
    def create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Folder selection
        ttk.Label(main_frame, text="1. Select PDF Folder:").grid(row=0, column=0, sticky=tk.W, pady=5)
        
        folder_frame = ttk.Frame(main_frame)
        folder_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        folder_frame.columnconfigure(0, weight=1)
        
        ttk.Entry(folder_frame, textvariable=self.pdf_folder, width=50).grid(row=0, column=0, sticky=(tk.W, tk.E))
        ttk.Button(folder_frame, text="Browse", command=self.select_folder).grid(row=0, column=1, padx=(5, 0))
        
        # Template PDF selection
        ttk.Label(main_frame, text="2. Select Template PDF:").grid(row=2, column=0, sticky=tk.W, pady=(20, 5))
        
        template_frame = ttk.Frame(main_frame)
        template_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        template_frame.columnconfigure(0, weight=1)
        
        self.template_combo = ttk.Combobox(template_frame, textvariable=self.template_pdf, state="readonly")
        self.template_combo.grid(row=0, column=0, sticky=(tk.W, tk.E))
        self.template_combo.bind('<<ComboboxSelected>>', self.load_fields)
        
        ttk.Button(template_frame, text="Load Fields", command=self.load_fields).grid(row=0, column=1, padx=(5, 0))
        
        # Fields frame
        ttk.Label(main_frame, text="3. Fill Form Fields:").grid(row=4, column=0, sticky=tk.W, pady=(20, 5))
        
        # Create scrollable frame for fields
        self.fields_canvas = tk.Canvas(main_frame)
        self.fields_scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.fields_canvas.yview)
        self.fields_frame = ttk.Frame(self.fields_canvas)
        
        self.fields_frame.bind(
            "<Configure>",
            lambda e: self.fields_canvas.configure(scrollregion=self.fields_canvas.bbox("all"))
        )
        
        self.fields_canvas.create_window((0, 0), window=self.fields_frame, anchor="nw")
        self.fields_canvas.configure(yscrollcommand=self.fields_scrollbar.set)
        
        self.fields_canvas.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        self.fields_scrollbar.grid(row=5, column=2, sticky=(tk.N, tk.S), pady=5)
        
        main_frame.rowconfigure(5, weight=1)
        
        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=6, column=0, columnspan=3, pady=(20, 0))
        
        ttk.Button(button_frame, text="Fill All PDFs", command=self.fill_all_pdfs, 
                  style="Accent.TButton").pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="Clear Fields", command=self.clear_fields).pack(side=tk.LEFT)
        
        # Status bar
        self.status_var = tk.StringVar(value="Ready - Select a folder to begin")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # Bind mouse wheel to canvas
        self.fields_canvas.bind("<MouseWheel>", self._on_mousewheel)
        
    def _on_mousewheel(self, event):
        self.fields_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        
    def select_folder(self):
        folder = filedialog.askdirectory(title="Select folder containing PDF forms")
        if folder:
            self.pdf_folder.set(folder)
            self.scan_pdf_files()
            
    def scan_pdf_files(self):
        folder = self.pdf_folder.get()
        if not folder:
            return
            
        try:
            # Find all PDF files in the folder
            pdf_files = [f for f in os.listdir(folder) if f.lower().endswith('.pdf')]
            
            if not pdf_files:
                messagebox.showwarning("No PDFs Found", "No PDF files found in the selected folder.")
                return
                
            self.pdf_files = pdf_files
            self.template_combo['values'] = pdf_files
            
            if pdf_files:
                self.template_combo.set(pdf_files[0])
                
            self.status_var.set(f"Found {len(pdf_files)} PDF files")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error scanning folder: {str(e)}")
            
    def load_fields(self, event=None):
        template_file = self.template_pdf.get()
        if not template_file or not self.pdf_folder.get():
            return
            
        try:
            pdf_path = os.path.join(self.pdf_folder.get(), template_file)
            
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                
                # Extract form fields with better error handling
                fields = {}
                
                # Method 1: Try AcroForm fields
                try:
                    if "/AcroForm" in pdf_reader.trailer["/Root"]:
                        acro_form = pdf_reader.trailer["/Root"]["/AcroForm"]
                        if hasattr(acro_form, 'get_object'):
                            acro_form = acro_form.get_object()
                        
                        if "/Fields" in acro_form:
                            form_fields = acro_form["/Fields"]
                            if hasattr(form_fields, 'get_object'):
                                form_fields = form_fields.get_object()
                                
                            for field in form_fields:
                                field_obj = field
                                if hasattr(field_obj, 'get_object'):
                                    field_obj = field_obj.get_object()
                                    
                                if "/T" in field_obj:
                                    field_name = field_obj["/T"]
                                    if hasattr(field_name, 'get_object'):
                                        field_name = field_name.get_object()
                                    
                                    field_value = field_obj.get("/V", "")
                                    if hasattr(field_value, 'get_object'):
                                        field_value = field_value.get_object()
                                    
                                    fields[str(field_name)] = str(field_value) if field_value else ""
                except Exception as e:
                    print(f"Error reading AcroForm fields: {e}")
                
                # Method 2: Try annotation-based fields if AcroForm failed
                if not fields:
                    try:
                        for page_num, page in enumerate(pdf_reader.pages):
                            if "/Annots" in page:
                                annotations = page["/Annots"]
                                if hasattr(annotations, 'get_object'):
                                    annotations = annotations.get_object()
                                    
                                for annotation in annotations:
                                    annot_obj = annotation
                                    if hasattr(annot_obj, 'get_object'):
                                        annot_obj = annot_obj.get_object()
                                        
                                    if (annot_obj.get("/Subtype") == "/Widget" or 
                                        str(annot_obj.get("/Subtype", "")).endswith("Widget")) and "/T" in annot_obj:
                                        
                                        field_name = annot_obj["/T"]
                                        if hasattr(field_name, 'get_object'):
                                            field_name = field_name.get_object()
                                            
                                        field_value = annot_obj.get("/V", "")
                                        if hasattr(field_value, 'get_object'):
                                            field_value = field_value.get_object()
                                        
                                        fields[str(field_name)] = str(field_value) if field_value else ""
                    except Exception as e:
                        print(f"Error reading annotation fields: {e}")
                
                # Method 3: Try using PyPDF2's built-in form field extraction
                if not fields:
                    try:
                        if hasattr(pdf_reader, 'get_form_text_fields'):
                            form_fields = pdf_reader.get_form_text_fields()
                            if form_fields:
                                fields = {k: str(v) if v else "" for k, v in form_fields.items()}
                    except Exception as e:
                        print(f"Error using built-in form field extraction: {e}")
                
                if fields:
                    self.create_field_entries(fields)
                    self.status_var.set(f"Loaded {len(fields)} form fields from {template_file}")
                else:
                    messagebox.showinfo("No Fields", "No form fields found in the selected PDF. Make sure it's a fillable form.")
                    self.status_var.set("No form fields found in template")
                    
        except Exception as e:
            messagebox.showerror("Error", f"Error loading PDF fields: {str(e)}")
            self.status_var.set("Error loading fields")
            
    def create_field_entries(self, fields):
        # Clear existing field entries
        for widget in self.fields_frame.winfo_children():
            widget.destroy()
        self.field_entries.clear()
        
        # Create entry widgets for each field
        for i, (field_name, field_value) in enumerate(fields.items()):
            # Field label
            label = ttk.Label(self.fields_frame, text=f"{field_name}:")
            label.grid(row=i, column=0, sticky=tk.W, padx=(0, 10), pady=2)
            
            # Field entry
            entry_var = tk.StringVar(value=str(field_value) if field_value else "")
            entry = ttk.Entry(self.fields_frame, textvariable=entry_var, width=40)
            entry.grid(row=i, column=1, sticky=(tk.W, tk.E), pady=2)
            
            self.field_entries[field_name] = entry_var
            
        # Configure column weight
        self.fields_frame.columnconfigure(1, weight=1)
        
    def clear_fields(self):
        for var in self.field_entries.values():
            var.set("")
        self.status_var.set("Fields cleared")
        
    def fill_all_pdfs(self):
        if not self.pdf_folder.get() or not self.field_entries:
            messagebox.showwarning("Incomplete Setup", "Please select a folder and load form fields first.")
            return
            
        # Get field values
        field_values = {name: var.get() for name, var in self.field_entries.items()}
        
        # Check if any fields have values
        if not any(field_values.values()):
            messagebox.showwarning("No Data", "Please fill in at least one field before processing.")
            return
            
        try:
            folder = self.pdf_folder.get()
            output_folder = os.path.join(folder, "filled_forms")
            os.makedirs(output_folder, exist_ok=True)
            
            success_count = 0
            error_count = 0
            
            for pdf_file in self.pdf_files:
                try:
                    input_path = os.path.join(folder, pdf_file)
                    output_path = os.path.join(output_folder, f"filled_{pdf_file}")
                    
                    self.fill_single_pdf(input_path, output_path, field_values)
                    success_count += 1
                    self.status_var.set(f"Processing: {pdf_file} ({success_count}/{len(self.pdf_files)})")
                    self.root.update()
                    
                except Exception as e:
                    error_count += 1
                    print(f"Error processing {pdf_file}: {str(e)}")
                    
            # Show completion message
            message = f"Completed!\n\nSuccessfully filled: {success_count} PDFs"
            if error_count > 0:
                message += f"\nErrors: {error_count} PDFs"
            message += f"\n\nFilled PDFs saved to:\n{output_folder}"
            
            messagebox.showinfo("Process Complete", message)
            self.status_var.set(f"Complete - {success_count} PDFs filled, {error_count} errors")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error during processing: {str(e)}")
            self.status_var.set("Error during processing")
            
    def fill_single_pdf(self, input_path, output_path, field_values):
        with open(input_path, 'rb') as input_file:
            pdf_reader = PyPDF2.PdfReader(input_file)
            pdf_writer = PyPDF2.PdfWriter()
            
            # Copy all pages
            for page in pdf_reader.pages:
                pdf_writer.add_page(page)
            
            # Try multiple methods to fill form fields
            try:
                # Method 1: Use update_page_form_field_values (newer PyPDF2)
                if hasattr(pdf_writer, 'update_page_form_field_values'):
                    for page_num in range(len(pdf_writer.pages)):
                        try:
                            pdf_writer.update_page_form_field_values(
                                pdf_writer.pages[page_num], 
                                field_values
                            )
                        except Exception as e:
                            print(f"Error updating page {page_num}: {e}")
                            continue
                            
                # Method 2: Try updatePageFormFieldValues (older PyPDF2)
                elif hasattr(pdf_writer, 'updatePageFormFieldValues'):
                    for page_num in range(len(pdf_writer.pages)):
                        try:
                            pdf_writer.updatePageFormFieldValues(
                                pdf_writer.pages[page_num], 
                                field_values
                            )
                        except Exception as e:
                            print(f"Error updating page {page_num} (legacy method): {e}")
                            continue
                            
                # Method 3: Manual field filling
                else:
                    print("Using manual field filling method")
                    for page in pdf_writer.pages:
                        if "/Annots" in page:
                            annotations = page["/Annots"]
                            if hasattr(annotations, 'get_object'):
                                annotations = annotations.get_object()
                                
                            for annotation in annotations:
                                annot_obj = annotation
                                if hasattr(annot_obj, 'get_object'):
                                    annot_obj = annot_obj.get_object()
                                    
                                if "/T" in annot_obj:
                                    field_name = annot_obj["/T"]
                                    if hasattr(field_name, 'get_object'):
                                        field_name = field_name.get_object()
                                    
                                    field_name_str = str(field_name)
                                    if field_name_str in field_values and field_values[field_name_str]:
                                        try:
                                            annot_obj.update({"/V": field_values[field_name_str]})
                                        except Exception as e:
                                            print(f"Error setting field {field_name_str}: {e}")
                                            
            except Exception as e:
                print(f"Error filling form fields: {e}")
            
            # Write the filled PDF
            with open(output_path, 'wb') as output_file:
                pdf_writer.write(output_file)


def main():
    # Check for PyPDF2
    try:
        import PyPDF2
    except ImportError:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "Missing Dependency", 
            "PyPDF2 is required to run this application.\n\n"
            "Please install it using:\npip install PyPDF2\n\n"
            "Then run this program again."
        )
        return
    
    root = tk.Tk()
    app = PDFFormFiller(root)
    
    # Center window
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")
    
    root.mainloop()


if __name__ == "__main__":
    main()