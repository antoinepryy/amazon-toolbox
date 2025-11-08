# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import threading
import io
from tkinterdnd2 import DND_FILES, TkinterDnD
import shutil

class CSVToExcelGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV to Excel Converter")
        self.root.geometry("600x400")
        self.root.configure(bg='#f0f0f0')
        
        # Variables
        self.converted_file = None
        self.is_converting = False
        
        self.setup_ui()
        
    def setup_ui(self):
        # Main frame
        main_frame = tk.Frame(self.root, bg='#f0f0f0')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Title
        title = tk.Label(
            main_frame,
            text="CSV to Excel Converter",
            font=("Arial", 24, "bold"),
            bg='#f0f0f0',
            fg='#2c3e50'
        )
        title.pack(pady=(0, 30))
        
        # Drag and drop area
        self.drop_frame = tk.Frame(
            main_frame,
            bg='#ecf0f1',
            relief=tk.RAISED,
            bd=2,
            height=200,
            width=500
        )
        self.drop_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        self.drop_frame.pack_propagate(False)
        
        # Enable drag and drop
        self.drop_frame.drop_target_register(DND_FILES)
        self.drop_frame.dnd_bind('<<Drop>>', self.handle_drop)
        
        # Drop zone content
        self.setup_drop_zone()
        
        # Button frame
        button_frame = tk.Frame(main_frame, bg='#f0f0f0')
        button_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Browse button
        self.browse_btn = tk.Button(
            button_frame,
            text="📁 Browse Files",
            command=self.browse_file,
            font=("Arial", 12),
            bg='#3498db',
            fg='white',
            relief=tk.FLAT,
            padx=20,
            pady=10,
            cursor='hand2'
        )
        self.browse_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Convert button
        self.convert_btn = tk.Button(
            button_frame,
            text="🔄 Convert to Excel",
            command=self.start_conversion,
            font=("Arial", 12),
            bg='#27ae60',
            fg='white',
            relief=tk.FLAT,
            padx=20,
            pady=10,
            cursor='hand2',
            state=tk.DISABLED
        )
        self.convert_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Download button
        self.download_btn = tk.Button(
            button_frame,
            text="💾 Save Excel File",
            command=self.download_file,
            font=("Arial", 12),
            bg='#e74c3c',
            fg='white',
            relief=tk.FLAT,
            padx=20,
            pady=10,
            cursor='hand2',
            state=tk.DISABLED
        )
        self.download_btn.pack(side=tk.LEFT)
        
        # Progress bar
        self.progress = ttk.Progressbar(
            main_frame,
            mode='indeterminate',
            length=400
        )
        self.progress.pack(pady=(10, 0))
        self.progress.pack_forget()  # Hide initially
        
        # Status label
        self.status_label = tk.Label(
            main_frame,
            text="Ready to convert CSV files",
            font=("Arial", 10),
            bg='#f0f0f0',
            fg='#7f8c8d'
        )
        self.status_label.pack(pady=(10, 0))
        
    def setup_drop_zone(self):
        # Clear drop frame
        for widget in self.drop_frame.winfo_children():
            widget.destroy()
            
        # Drop zone icon and text
        if not hasattr(self, 'file_path') or not self.file_path:
            icon_label = tk.Label(
                self.drop_frame,
                text="📄",
                font=("Arial", 48),
                bg='#ecf0f1',
                fg='#bdc3c7'
            )
            icon_label.pack(expand=True)
            
            text_label = tk.Label(
                self.drop_frame,
                text="Drag & Drop your CSV file here\nor click Browse Files",
                font=("Arial", 14),
                bg='#ecf0f1',
                fg='#7f8c8d',
                justify=tk.CENTER
            )
            text_label.pack(expand=True)
        else:
            # Show file info
            file_name = os.path.basename(self.file_path)
            icon_label = tk.Label(
                self.drop_frame,
                text="✅",
                font=("Arial", 36),
                bg='#ecf0f1',
                fg='#27ae60'
            )
            icon_label.pack(expand=True)
            
            name_label = tk.Label(
                self.drop_frame,
                text=file_name,
                font=("Arial", 14, "bold"),
                bg='#ecf0f1',
                fg='#2c3e50'
            )
            name_label.pack(expand=True)
            
            info_label = tk.Label(
                self.drop_frame,
                text="Ready to convert",
                font=("Arial", 10),
                bg='#ecf0f1',
                fg='#27ae60'
            )
            info_label.pack(expand=True)
    
    def handle_drop(self, event):
        file_path = event.data
        # Remove curly braces if present
        if file_path.startswith('{') and file_path.endswith('}'):
            file_path = file_path[1:-1]
        
        if file_path.lower().endswith('.csv'):
            self.set_file(file_path)
        else:
            messagebox.showerror("Invalid File", "Please select a CSV file.")
    
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select CSV File",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if file_path:
            self.set_file(file_path)
    
    def set_file(self, file_path):
        self.file_path = file_path
        self.converted_file = None
        self.setup_drop_zone()
        self.convert_btn.config(state=tk.NORMAL)
        self.download_btn.config(state=tk.DISABLED)
        self.status_label.config(text=f"File loaded: {os.path.basename(file_path)}")
    
    def start_conversion(self):
        if self.is_converting:
            return
            
        self.is_converting = True
        self.convert_btn.config(state=tk.DISABLED, text="Converting...")
        self.progress.pack(pady=(10, 0))
        self.progress.start(10)
        self.status_label.config(text="Converting CSV to Excel...")
        
        # Run conversion in separate thread
        thread = threading.Thread(target=self.convert_file)
        thread.daemon = True
        thread.start()
    
    def convert_file(self):
        try:
            # Convert CSV to Excel (same logic as main.py)
            with open(self.file_path, 'r', encoding='utf-8-sig') as file:
                lines = file.readlines()
            
            # Fix malformed CSV
            fixed_lines = []
            for i, line in enumerate(lines):
                line = line.strip()
                if i == 0:  # Header row
                    fixed_lines.append(line)
                else:  # Data rows
                    if line.startswith('"') and line.endswith('"'):
                        line = line[1:-1]
                        line = line.replace('""', '"')
                    fixed_lines.append(line)
            
            # Parse with pandas
            csv_string = '\n'.join(fixed_lines)
            df = pd.read_csv(io.StringIO(csv_string), sep=',')
            
            # Create output file path
            base_name = os.path.splitext(self.file_path)[0]
            self.converted_file = f"{base_name}_converted.xlsx"
            
            # Write to Excel
            with pd.ExcelWriter(self.converted_file, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Amazon Data', index=False)
                
                # Auto-adjust column widths
                worksheet = writer.sheets['Amazon Data']
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width
            
            # Update UI on main thread
            self.root.after(0, self.conversion_success)
            
        except Exception as e:
            self.root.after(0, lambda: self.conversion_error(str(e)))
    
    def conversion_success(self):
        self.is_converting = False
        self.progress.stop()
        self.progress.pack_forget()
        self.convert_btn.config(state=tk.NORMAL, text="🔄 Convert to Excel")
        self.download_btn.config(state=tk.NORMAL)
        self.status_label.config(text="✅ Conversion completed successfully!")
        
        messagebox.showinfo("Success", "CSV file converted to Excel successfully!")
    
    def conversion_error(self, error_msg):
        self.is_converting = False
        self.progress.stop()
        self.progress.pack_forget()
        self.convert_btn.config(state=tk.NORMAL, text="🔄 Convert to Excel")
        self.status_label.config(text="❌ Conversion failed")
        
        messagebox.showerror("Error", f"Conversion failed:\n{error_msg}")
    
    def download_file(self):
        if not self.converted_file or not os.path.exists(self.converted_file):
            messagebox.showerror("Error", "No converted file available.")
            return
        
        # Ask user where to save
        save_path = filedialog.asksaveasfilename(
            title="Save Excel File",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialname=os.path.basename(self.converted_file)
        )
        
        if save_path:
            try:
                shutil.copy2(self.converted_file, save_path)
                self.status_label.config(text=f"📁 File saved: {os.path.basename(save_path)}")
                messagebox.showinfo("Success", f"File saved successfully!\n{save_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save file:\n{str(e)}")

def main():
    root = TkinterDnD.Tk()
    app = CSVToExcelGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()