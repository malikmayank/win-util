# -*- coding: utf-8 -*-
"""
AES Encrypted ZIP File Opener for Windows 10
Enhanced version with 7-Zip backend support for all compression methods
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
import json
import shutil
import threading
from datetime import datetime

class AESZipOpener:
    def __init__(self, root):
        self.root = root
        self.root.title("AES Encrypted ZIP File Opener - Enhanced")
        self.root.geometry("700x500")
        self.root.resizable(True, True)
        
        # Variables
        self.zip_file_path = tk.StringVar()
        self.password = tk.StringVar()
        self.seven_zip_path = self.find_7zip()
        
        self.setup_ui()
        
        # Check for 7-Zip availability
        if not self.seven_zip_path:
            self.show_7zip_warning()
    
    def find_7zip(self):
        """Find 7-Zip installation on the system"""
        possible_paths = [
            r"C:\Program Files\7-Zip\7z.exe",
            r"C:\Program Files (x86)\7-Zip\7z.exe",
            r"C:\Tools\7-Zip\7z.exe",
            "7z.exe"  # If in PATH
        ]
        
        for path in possible_paths:
            if shutil.which(path) or os.path.exists(path):
                return path
        return None
    
    def show_7zip_warning(self):
        """Show warning if 7-Zip is not found"""
        warning_msg = """7-Zip not found on your system!

This tool works best with 7-Zip for maximum compatibility with all ZIP compression methods.

Options:
1. Download and install 7-Zip from https://www.7-zip.org/
2. Continue with basic Python support (limited compression methods)

The application will try Python's built-in ZIP support first, but may fail with advanced compression methods."""
        
        messagebox.showwarning("7-Zip Not Found", warning_msg)
        
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
        
        # 7-Zip status
        status_color = "green" if self.seven_zip_path else "red"
        status_text = "7-Zip: Available" if self.seven_zip_path else "7-Zip: Not Found"
        ttk.Label(main_frame, text=status_text, foreground=status_color).grid(
            row=1, column=0, columnspan=3, pady=(0, 10))
        
        # File selection
        ttk.Label(main_frame, text="ZIP File:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.zip_file_path, width=50).grid(
            row=2, column=1, sticky=(tk.W, tk.E), pady=5, padx=(5, 5))
        ttk.Button(main_frame, text="Browse", command=self.browse_file).grid(
            row=2, column=2, pady=5)
        
        # Password entry
        ttk.Label(main_frame, text="Password:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.password_entry = ttk.Entry(main_frame, textvariable=self.password, show="*", width=50)
        self.password_entry.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=5, padx=(5, 5))
        
        # Show/Hide password checkbox
        self.show_password = tk.BooleanVar()
        ttk.Checkbutton(main_frame, text="Show", variable=self.show_password,
                       command=self.toggle_password_visibility).grid(row=3, column=2, pady=5)
        
        # Method selection
        ttk.Label(main_frame, text="Method:").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.method_var = tk.StringVar(value="auto")
        method_frame = ttk.Frame(main_frame)
        method_frame.grid(row=4, column=1, columnspan=2, sticky=tk.W, pady=5)
        
        ttk.Radiobutton(method_frame, text="Auto", variable=self.method_var, value="auto").pack(side=tk.LEFT)
        ttk.Radiobutton(method_frame, text="7-Zip", variable=self.method_var, value="7zip").pack(side=tk.LEFT, padx=(10, 0))
        ttk.Radiobutton(method_frame, text="Python", variable=self.method_var, value="python").pack(side=tk.LEFT, padx=(10, 0))
        
        # Buttons frame
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.grid(row=5, column=0, columnspan=3, pady=20)
        
        ttk.Button(buttons_frame, text="View Contents", command=self.view_contents).pack(
            side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Extract All", command=self.extract_all).pack(
            side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Extract Selected", command=self.extract_selected).pack(
            side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Test Archive", command=self.test_archive).pack(
            side=tk.LEFT, padx=5)
        
        # File list frame
        list_frame = ttk.LabelFrame(main_frame, text="Archive Contents", padding="5")
        list_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(6, weight=1)
        
        # Treeview for file list
        self.file_tree = ttk.Treeview(list_frame, columns=("Size", "Compressed", "Method", "Modified"), show="tree headings")
        self.file_tree.heading("#0", text="File Name")
        self.file_tree.heading("Size", text="Size")
        self.file_tree.heading("Compressed", text="Compressed")
        self.file_tree.heading("Method", text="Method")
        self.file_tree.heading("Modified", text="Modified")
        self.file_tree.column("#0", width=250)
        self.file_tree.column("Size", width=80)
        self.file_tree.column("Compressed", width=80)
        self.file_tree.column("Method", width=80)
        self.file_tree.column("Modified", width=130)
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.file_tree.yview)
        h_scrollbar = ttk.Scrollbar(list_frame, orient=tk.HORIZONTAL, command=self.file_tree.xview)
        self.file_tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        self.file_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready - Select a ZIP file to begin")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.grid(row=8, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(5, 0))
    
    def toggle_password_visibility(self):
        """Toggle password visibility"""
        self.password_entry.config(show="" if self.show_password.get() else "*")
        
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
    
    def get_compression_method_name(self, method_id):
        """Get human-readable compression method name"""
        methods = {
            0: "Stored",
            8: "Deflate",
            9: "Deflate64",
            12: "BZIP2",
            14: "LZMA",
            95: "XZ",
            99: "AES"
        }
        return methods.get(method_id, f"Method {method_id}")
    
    def run_with_progress(self, func, *args, **kwargs):
        """Run a function with progress indication"""
        self.progress.start()
        self.root.update()
        
        def worker():
            try:
                result = func(*args, **kwargs)
                self.root.after(0, lambda: self.progress.stop())
                return result
            except Exception as e:
                self.root.after(0, lambda: self.progress.stop())
                raise e
        
        return worker()
    
    def list_with_7zip(self):
        """List archive contents using 7-Zip"""
        if not self.seven_zip_path:
            raise Exception("7-Zip not available")
        
        cmd = [self.seven_zip_path, "l", "-slt", self.zip_file_path.get()]
        if self.password.get():
            cmd.extend([f"-p{self.password.get()}"])
        
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, 
                                  encoding='utf-8', errors='replace', timeout=30)
            
            if result.returncode != 0:
                if "Wrong password" in result.stderr or "Cannot open encrypted archive" in result.stderr:
                    raise Exception("Incorrect password")
                else:
                    raise Exception(f"7-Zip error: {result.stderr}")
            
            return self.parse_7zip_listing(result.stdout)
        except subprocess.TimeoutExpired:
            raise Exception("Operation timed out")
    
    def parse_7zip_listing(self, output):
        """Parse 7-Zip listing output"""
        files = []
        current_file = {}
        
        for line in output.split('\n'):
            line = line.strip()
            if line.startswith('Path = '):
                if current_file and not current_file.get('Path', '').endswith('/'):
                    files.append(current_file)
                current_file = {'Path': line[7:]}
            elif line.startswith('Size = '):
                current_file['Size'] = int(line[7:]) if line[7:].isdigit() else 0
            elif line.startswith('Packed Size = '):
                current_file['Packed Size'] = int(line[14:]) if line[14:].isdigit() else 0
            elif line.startswith('Method = '):
                current_file['Method'] = line[9:]
            elif line.startswith('Modified = '):
                current_file['Modified'] = line[11:]
        
        if current_file and not current_file.get('Path', '').endswith('/'):
            files.append(current_file)
        
        return files
    
    def list_with_python(self):
        """List archive contents using Python's zipfile"""
        files = []
        with zipfile.ZipFile(self.zip_file_path.get(), 'r') as zip_file:
            if self.password.get():
                zip_file.setpassword(self.password.get().encode('utf-8'))
            
            for file_info in zip_file.infolist():
                if not file_info.filename.endswith('/'):
                    try:
                        filename = file_info.filename
                        if isinstance(filename, bytes):
                            filename = filename.decode('utf-8')
                    except UnicodeDecodeError:
                        filename = str(file_info.filename)
                    
                    try:
                        modified = f"{file_info.date_time[0]}-{file_info.date_time[1]:02d}-{file_info.date_time[2]:02d} {file_info.date_time[3]:02d}:{file_info.date_time[4]:02d}"
                    except:
                        modified = "Unknown"
                    
                    files.append({
                        'Path': filename,
                        'Size': file_info.file_size,
                        'Packed Size': file_info.compress_size,
                        'Method': self.get_compression_method_name(file_info.compress_type),
                        'Modified': modified
                    })
        return files
    
    def view_contents(self):
        if not self.zip_file_path.get():
            messagebox.showerror("Error", "Please select a ZIP file first.")
            return
        
        # Clear existing items
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)
        
        try:
            self.progress.start()
            self.status_var.set("Reading archive contents...")
            self.root.update()
            
            files = None
            method_used = ""
            
            if self.method_var.get() == "7zip" or (self.method_var.get() == "auto" and self.seven_zip_path):
                try:
                    files = self.list_with_7zip()
                    method_used = "7-Zip"
                except Exception as e:
                    if self.method_var.get() == "7zip":
                        raise e
                    # Fall back to Python method if auto mode
                    pass
            
            if files is None:
                files = self.list_with_python()
                method_used = "Python"
            
            # Populate the tree
            for file_data in files:
                self.file_tree.insert("", tk.END, 
                                    text=file_data['Path'],
                                    values=(
                                        self.format_size(file_data.get('Size', 0)),
                                        self.format_size(file_data.get('Packed Size', 0)),
                                        file_data.get('Method', 'Unknown'),
                                        file_data.get('Modified', 'Unknown')
                                    ))
            
            self.status_var.set(f"Archive contents loaded using {method_used} - {len(files)} files")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error reading archive: {str(e)}")
            self.status_var.set("Error reading archive")
        finally:
            self.progress.stop()
    
    def extract_with_7zip(self, extract_dir, selected_files=None):
        """Extract files using 7-Zip"""
        if not self.seven_zip_path:
            raise Exception("7-Zip not available")
        
        if selected_files:
            # Extract selected files
            for filename in selected_files:
                cmd = [self.seven_zip_path, "e", self.zip_file_path.get(), 
                      f"-o{extract_dir}", filename]
                if self.password.get():
                    cmd.extend([f"-p{self.password.get()}"])
                
                result = subprocess.run(cmd, capture_output=True, text=True,
                                      encoding='utf-8', errors='replace')
                if result.returncode != 0:
                    raise Exception(f"7-Zip extraction error: {result.stderr}")
        else:
            # Extract all files
            cmd = [self.seven_zip_path, "x", self.zip_file_path.get(), 
                  f"-o{extract_dir}"]
            if self.password.get():
                cmd.extend([f"-p{self.password.get()}"])
            
            result = subprocess.run(cmd, capture_output=True, text=True,
                                  encoding='utf-8', errors='replace')
            if result.returncode != 0:
                if "Wrong password" in result.stderr:
                    raise Exception("Incorrect password")
                else:
                    raise Exception(f"7-Zip extraction error: {result.stderr}")
    
    def extract_all(self):
        if not self.zip_file_path.get():
            messagebox.showerror("Error", "Please select a ZIP file first.")
            return
        
        extract_dir = filedialog.askdirectory(title="Select Extraction Directory")
        if not extract_dir:
            return
        
        try:
            self.progress.start()
            self.status_var.set("Extracting all files...")
            self.root.update()
            
            if self.method_var.get() == "7zip" or (self.method_var.get() == "auto" and self.seven_zip_path):
                try:
                    self.extract_with_7zip(extract_dir)
                    method_used = "7-Zip"
                except Exception as e:
                    if self.method_var.get() == "7zip":
                        raise e
                    # Fall back to Python method
                    self.extract_with_python(extract_dir)
                    method_used = "Python"
            else:
                self.extract_with_python(extract_dir)
                method_used = "Python"
            
            messagebox.showinfo("Success", f"All files extracted to:\n{extract_dir}\n\nMethod used: {method_used}")
            self.status_var.set(f"Extraction completed using {method_used}")
            
            if messagebox.askyesno("Open Folder", "Would you like to open the extraction folder?"):
                os.startfile(extract_dir)
                
        except Exception as e:
            messagebox.showerror("Error", f"Error extracting files: {str(e)}")
            self.status_var.set("Extraction failed")
        finally:
            self.progress.stop()
    
    def extract_with_python(self, extract_dir, selected_files=None):
        """Extract files using Python's zipfile"""
        with zipfile.ZipFile(self.zip_file_path.get(), 'r') as zip_file:
            if self.password.get():
                zip_file.setpassword(self.password.get().encode('utf-8'))
            
            if selected_files:
                for filename in selected_files:
                    zip_file.extract(filename, extract_dir)
            else:
                zip_file.extractall(extract_dir)
    
    def extract_selected(self):
        selected_items = self.file_tree.selection()
        if not selected_items:
            messagebox.showwarning("Warning", "Please select files to extract.")
            return
        
        if not self.zip_file_path.get():
            messagebox.showerror("Error", "Please select a ZIP file first.")
            return
        
        extract_dir = filedialog.askdirectory(title="Select Extraction Directory")
        if not extract_dir:
            return
        
        try:
            self.progress.start()
            self.status_var.set("Extracting selected files...")
            self.root.update()
            
            selected_files = [self.file_tree.item(item, "text") for item in selected_items]
            
            if self.method_var.get() == "7zip" or (self.method_var.get() == "auto" and self.seven_zip_path):
                try:
                    self.extract_with_7zip(extract_dir, selected_files)
                    method_used = "7-Zip"
                except Exception as e:
                    if self.method_var.get() == "7zip":
                        raise e
                    # Fall back to Python method
                    self.extract_with_python(extract_dir, selected_files)
                    method_used = "Python"
            else:
                self.extract_with_python(extract_dir, selected_files)
                method_used = "Python"
            
            messagebox.showinfo("Success", 
                              f"Extracted {len(selected_files)} file(s) to:\n{extract_dir}\n\nMethod used: {method_used}")
            self.status_var.set(f"Extracted {len(selected_files)} files using {method_used}")
            
            if messagebox.askyesno("Open Folder", "Would you like to open the extraction folder?"):
                os.startfile(extract_dir)
                
        except Exception as e:
            messagebox.showerror("Error", f"Error extracting selected files: {str(e)}")
            self.status_var.set("Extraction failed")
        finally:
            self.progress.stop()
    
    def test_archive(self):
        """Test archive integrity"""
        if not self.zip_file_path.get():
            messagebox.showerror("Error", "Please select a ZIP file first.")
            return
        
        try:
            self.progress.start()
            self.status_var.set("Testing archive...")
            self.root.update()
            
            if self.seven_zip_path:
                cmd = [self.seven_zip_path, "t", self.zip_file_path.get()]
                if self.password.get():
                    cmd.extend([f"-p{self.password.get()}"])
                
                result = subprocess.run(cmd, capture_output=True, text=True,
                                      encoding='utf-8', errors='replace')
                
                if result.returncode == 0:
                    messagebox.showinfo("Test Result", "Archive test passed! All files are OK.")
                    self.status_var.set("Archive test passed")
                else:
                    if "Wrong password" in result.stderr:
                        messagebox.showerror("Test Result", "Incorrect password.")
                    else:
                        messagebox.showerror("Test Result", f"Archive test failed:\n{result.stderr}")
                    self.status_var.set("Archive test failed")
            else:
                # Basic test with Python
                with zipfile.ZipFile(self.zip_file_path.get(), 'r') as zip_file:
                    if self.password.get():
                        zip_file.setpassword(self.password.get().encode('utf-8'))
                    
                    bad_files = zip_file.testzip()
                    if bad_files:
                        messagebox.showerror("Test Result", f"Archive test failed. Corrupt file: {bad_files}")
                        self.status_var.set("Archive test failed")
                    else:
                        messagebox.showinfo("Test Result", "Archive test passed! (Basic Python test)")
                        self.status_var.set("Archive test passed")
                        
        except Exception as e:
            messagebox.showerror("Error", f"Error testing archive: {str(e)}")
            self.status_var.set("Archive test error")
        finally:
            self.progress.stop()

def main():
    # Set proper encoding for the console
    if sys.platform.startswith('win'):
        try:
            sys.stdout.reconfigure(encoding='utf-8')
            sys.stderr.reconfigure(encoding='utf-8')
        except:
            pass
    
    try:
        root = tk.Tk()
        app = AESZipOpener(root)
        root.mainloop()
    except ImportError as e:
        print(f"Required module not found: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()