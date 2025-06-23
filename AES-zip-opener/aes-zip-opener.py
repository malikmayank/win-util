# -*- coding: utf-8 -*-
"""
AES Encrypted ZIP File Opener for Windows 10
Handles AES encrypted ZIP files with proper encoding support
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import zipfile
import os
import tempfile
import subprocess
import sys
from pathlib import Path
import locale

class AESZipOpener:
    def __init__(self, root):
        self.root = root
        self.root.title("AES Encrypted ZIP File Opener")
        self.root.geometry("600x400")
        self.root.resizable(True, True)
        
        # Variables
        self.zip_file_path = tk.StringVar()
        self.password = tk.StringVar()
        
        self.setup_ui()
        
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="AES Encrypted ZIP File Opener", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # File selection
        ttk.Label(main_frame, text="ZIP File:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.zip_file_path, width=50).grid(
            row=1, column=1, sticky=(tk.W, tk.E), pady=5, padx=(5, 5))
        ttk.Button(main_frame, text="Browse", command=self.browse_file).grid(
            row=1, column=2, pady=5)
        
        # Password entry
        ttk.Label(main_frame, text="Password:").grid(row=2, column=0, sticky=tk.W, pady=5)
        password_entry = ttk.Entry(main_frame, textvariable=self.password, show="*", width=50)
        password_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5, padx=(5, 5))
        
        # Show/Hide password checkbox
        self.show_password = tk.BooleanVar()
        ttk.Checkbutton(main_frame, text="Show", variable=self.show_password,
                       command=lambda: password_entry.config(show="" if self.show_password.get() else "*")).grid(
            row=2, column=2, pady=5)
        
        # Buttons frame
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.grid(row=3, column=0, columnspan=3, pady=20)
        
        ttk.Button(buttons_frame, text="Extract All", command=self.extract_all).pack(
            side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="View Contents", command=self.view_contents).pack(
            side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Extract Selected", command=self.extract_selected).pack(
            side=tk.LEFT, padx=5)
        
        # File list frame
        list_frame = ttk.LabelFrame(main_frame, text="Archive Contents", padding="5")
        list_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(4, weight=1)
        
        # Treeview for file list
        self.file_tree = ttk.Treeview(list_frame, columns=("Size", "Modified"), show="tree headings")
        self.file_tree.heading("#0", text="File Name")
        self.file_tree.heading("Size", text="Size")
        self.file_tree.heading("Modified", text="Modified")
        self.file_tree.column("#0", width=300)
        self.file_tree.column("Size", width=100)
        self.file_tree.column("Modified", width=150)
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.file_tree.yview)
        h_scrollbar = ttk.Scrollbar(list_frame, orient=tk.HORIZONTAL, command=self.file_tree.xview)
        self.file_tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        self.file_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready - Select a ZIP file to begin")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        
    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Select AES Encrypted ZIP File",
            filetypes=[("ZIP files", "*.zip"), ("All files", "*.*")]
        )
        if filename:
            self.zip_file_path.set(filename)
            self.status_var.set(f"Selected: {os.path.basename(filename)}")
    
    def format_size(self, size_bytes):
        """Convert bytes to human readable format"""
        if size_bytes == 0:
            return "0 B"
        size_names = ["B", "KB", "MB", "GB"]
        i = 0
        while size_bytes >= 1024 and i < len(size_names) - 1:
            size_bytes /= 1024.0
            i += 1
        return f"{size_bytes:.1f} {size_names[i]}"
    
    def view_contents(self):
        if not self.zip_file_path.get():
            messagebox.showerror("Error", "Please select a ZIP file first.")
            return
        
        if not self.password.get():
            messagebox.showerror("Error", "Please enter the password.")
            return
        
        try:
            # Clear existing items
            for item in self.file_tree.get_children():
                self.file_tree.delete(item)
            
            with zipfile.ZipFile(self.zip_file_path.get(), 'r') as zip_file:
                # Set password
                zip_file.setpassword(self.password.get().encode('utf-8'))
                
                # Get file list
                file_list = zip_file.infolist()
                
                for file_info in file_list:
                    # Skip directories in the display (they'll be shown as folders)
                    if not file_info.filename.endswith('/'):
                        size = self.format_size(file_info.file_size)
                        # Convert date_time tuple to string
                        try:
                            modified = f"{file_info.date_time[0]}-{file_info.date_time[1]:02d}-{file_info.date_time[2]:02d} {file_info.date_time[3]:02d}:{file_info.date_time[4]:02d}"
                        except:
                            modified = "Unknown"
                        
                        self.file_tree.insert("", tk.END, text=file_info.filename, 
                                            values=(size, modified))
                
                self.status_var.set(f"Archive contents loaded - {len(file_list)} items")
                
        except zipfile.BadZipFile:
            messagebox.showerror("Error", "Invalid ZIP file.")
        except RuntimeError as e:
            if "Bad password" in str(e):
                messagebox.showerror("Error", "Incorrect password.")
            else:
                messagebox.showerror("Error", f"Error reading ZIP file: {str(e)}")
        except Exception as e:
            messagebox.showerror("Error", f"Unexpected error: {str(e)}")
    
    def extract_all(self):
        if not self.zip_file_path.get():
            messagebox.showerror("Error", "Please select a ZIP file first.")
            return
        
        if not self.password.get():
            messagebox.showerror("Error", "Please enter the password.")
            return
        
        # Select extraction directory
        extract_dir = filedialog.askdirectory(title="Select Extraction Directory")
        if not extract_dir:
            return
        
        try:
            with zipfile.ZipFile(self.zip_file_path.get(), 'r') as zip_file:
                password_bytes = self.password.get().encode('utf-8')
                zip_file.setpassword(password_bytes)
                
                # Extract all files with proper encoding handling
                for file_info in zip_file.infolist():
                    try:
                        zip_file.extract(file_info, extract_dir)
                    except UnicodeDecodeError:
                        # Handle files with encoding issues
                        data = zip_file.read(file_info)
                        file_path = os.path.join(extract_dir, file_info.filename)
                        os.makedirs(os.path.dirname(file_path), exist_ok=True)
                        with open(file_path, 'wb') as f:
                            f.write(data)
                
                messagebox.showinfo("Success", f"All files extracted to:\n{extract_dir}")
                self.status_var.set(f"Extraction completed to: {extract_dir}")
                
                # Open the extraction directory
                if messagebox.askyesno("Open Folder", "Would you like to open the extraction folder?"):
                    os.startfile(extract_dir)
                    
        except zipfile.BadZipFile:
            messagebox.showerror("Error", "Invalid ZIP file.")
        except RuntimeError as e:
            if "Bad password" in str(e):
                messagebox.showerror("Error", "Incorrect password.")
            else:
                messagebox.showerror("Error", f"Error extracting ZIP file: {str(e)}")
        except Exception as e:
            messagebox.showerror("Error", f"Unexpected error: {str(e)}")
    
    def extract_selected(self):
        selected_items = self.file_tree.selection()
        if not selected_items:
            messagebox.showwarning("Warning", "Please select files to extract.")
            return
        
        if not self.zip_file_path.get():
            messagebox.showerror("Error", "Please select a ZIP file first.")
            return
        
        if not self.password.get():
            messagebox.showerror("Error", "Please enter the password.")
            return
        
        # Select extraction directory
        extract_dir = filedialog.askdirectory(title="Select Extraction Directory")
        if not extract_dir:
            return
        
        try:
            with zipfile.ZipFile(self.zip_file_path.get(), 'r') as zip_file:
                password_bytes = self.password.get().encode('utf-8')
                zip_file.setpassword(password_bytes)
                
                extracted_files = []
                for item in selected_items:
                    filename = self.file_tree.item(item, "text")
                    try:
                        zip_file.extract(filename, extract_dir)
                        extracted_files.append(filename)
                    except UnicodeDecodeError:
                        # Handle files with encoding issues
                        try:
                            data = zip_file.read(filename)
                            file_path = os.path.join(extract_dir, filename)
                            os.makedirs(os.path.dirname(file_path), exist_ok=True)
                            with open(file_path, 'wb') as f:
                                f.write(data)
                            extracted_files.append(filename)
                        except Exception as inner_e:
                            messagebox.showwarning("Warning", f"Could not extract {filename}: {str(inner_e)}")
                            continue
                
                messagebox.showinfo("Success", 
                                  f"Extracted {len(extracted_files)} file(s) to:\n{extract_dir}")
                self.status_var.set(f"Extracted {len(extracted_files)} selected files")
                
                # Open the extraction directory
                if messagebox.askyesno("Open Folder", "Would you like to open the extraction folder?"):
                    os.startfile(extract_dir)
                    
        except Exception as e:
            messagebox.showerror("Error", f"Error extracting selected files: {str(e)}")

def main():
    # Set proper encoding for the console
    if sys.platform.startswith('win'):
        try:
            # Try to set UTF-8 encoding for Windows console
            sys.stdout.reconfigure(encoding='utf-8')
            sys.stderr.reconfigure(encoding='utf-8')
        except:
            pass
    
    # Check if required modules are available
    try:
        root = tk.Tk()
        app = AESZipOpener(root)
        root.mainloop()
    except ImportError as e:
        print(f"Required module not found: {e}")
        print("Please install required dependencies:")
        print("pip install tkinter")
        sys.exit(1)

if __name__ == "__main__":
    main()