# app.py - FIXED VERSION
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path
from datetime import datetime
import threading
import queue
import os
import sys
import re
import traceback
from processor import process_presentations

if getattr(sys, 'frozen', False):
    # Running as compiled executable
    BASE_DIR = os.path.dirname(sys.executable)
    
    # Add config path to system path for imports
    config_dir = os.path.join(BASE_DIR, 'config')
    if os.path.exists(config_dir) and config_dir not in sys.path:
        sys.path.insert(0, config_dir)
else:
    # Running as script
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    
    # Add config path for script mode
    config_dir = os.path.join(BASE_DIR, 'config')
    if os.path.exists(config_dir) and config_dir not in sys.path:
        sys.path.insert(0, config_dir)
        
class SocialMediaExtractorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("📊 Social Media Report Extractor")
        self.root.geometry("900x750")
        self.root.minsize(800, 600)
        
        # Queue for thread-safe updates
        self.queue = queue.Queue()
        
        # Store selected files
        self.selected_files = []
        
        # Initialize variables that will be created in setup_ui
        self.status_var = None
        
        # Setup UI
        self.setup_ui()
        
        # Start checking queue
        self.check_queue()
        
        # Center window
        self.center_window()
        
    def center_window(self):
        """Center the window on screen"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
        
    def setup_ui(self):
        # Configure style
        self.setup_styles()
        
        # Main container
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(4, weight=1)
        
        # ==================== TITLE ====================
        title_frame = ttk.Frame(main_frame)
        title_frame.grid(row=0, column=0, columnspan=3, pady=(0, 20), sticky=(tk.W, tk.E))
        title_frame.columnconfigure(0, weight=1)
        
        title = ttk.Label(
            title_frame,
            text="📊 Social Media Report Extractor",
            font=("Arial", 18, "bold"),
            foreground="#2C3E50"
        )
        title.grid(row=0, column=0, sticky=tk.W)
        
        version = ttk.Label(
            title_frame,
            text="v1.0",
            font=("Arial", 9),
            foreground="gray"
        )
        version.grid(row=0, column=1, sticky=tk.E)
        
        # ==================== SECTION 1: INPUT FILES ====================
        input_frame = ttk.LabelFrame(main_frame, text="1. Select PowerPoint Files", padding="10")
        input_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        input_frame.columnconfigure(0, weight=1)
        
        # File list with scrollbar
        file_list_frame = ttk.Frame(input_frame)
        file_list_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        file_list_frame.columnconfigure(0, weight=1)
        file_list_frame.rowconfigure(0, weight=1)
        
        self.file_listbox = tk.Listbox(
            file_list_frame, 
            height=6, 
            selectmode=tk.EXTENDED, 
            font=("Consolas", 9),
            bg="white",
            relief=tk.SUNKEN,
            borderwidth=1
        )
        self.file_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 5))
        
        # Scrollbar for file list
        scrollbar = ttk.Scrollbar(file_list_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.file_listbox.config(yscrollcommand=scrollbar.set)
        
        # File buttons
        btn_frame = ttk.Frame(input_frame)
        btn_frame.grid(row=1, column=0, columnspan=3, pady=(0, 5))
        
        ttk.Button(
            btn_frame, 
            text="📁 Add Files", 
            command=self.add_files, 
            width=15,
            style="Action.TButton"
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            btn_frame, 
            text="🗑️ Remove Selected", 
            command=self.remove_files, 
            width=15
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            btn_frame, 
            text="🧹 Clear All", 
            command=self.clear_files, 
            width=15
        ).pack(side=tk.LEFT, padx=5)
        
        # File count
        self.file_count_var = tk.StringVar(value="0 files selected")
        ttk.Label(
            input_frame, 
            textvariable=self.file_count_var, 
            foreground="gray",
            font=("Arial", 9)
        ).grid(row=2, column=0, sticky=tk.W, pady=(5, 0))
        
        # ==================== SECTION 2: OUTPUT SETTINGS ====================
        output_frame = ttk.LabelFrame(main_frame, text="2. Output Settings", padding="10")
        output_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        output_frame.columnconfigure(1, weight=1)
        
        # Output folder
        ttk.Label(
            output_frame, 
            text="Output Folder:",
            font=("Arial", 9, "bold")
        ).grid(row=0, column=0, sticky=tk.W, padx=(0, 10), pady=5)
        
        self.folder_var = tk.StringVar(value=os.path.join(BASE_DIR, "Social_Reports"))
        folder_entry = ttk.Entry(
            output_frame, 
            textvariable=self.folder_var, 
            width=40,
            font=("Arial", 9)
        )
        folder_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10), pady=5)
        
        ttk.Button(
            output_frame, 
            text="📂 Browse", 
            command=self.browse_folder, 
            width=10
        ).grid(row=0, column=2, sticky=tk.W, pady=5)
        
        # File naming
        ttk.Label(
            output_frame, 
            text="File Name:",
            font=("Arial", 9, "bold")
        ).grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=5)
        
        name_frame = ttk.Frame(output_frame)
        name_frame.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(0, 10), pady=5)
        name_frame.columnconfigure(0, weight=1)
        
        self.filename_var = tk.StringVar(value="social_media_report")
        self.filename_entry = ttk.Entry(
            name_frame, 
            textvariable=self.filename_var, 
            width=35,
            font=("Arial", 9)
        )
        self.filename_entry.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        ttk.Label(
            name_frame, 
            text=".xlsx", 
            foreground="gray",
            font=("Arial", 9)
        ).grid(row=0, column=1, sticky=tk.W, padx=(5, 0))
        
        # Bind to update preview in real-time
        self.filename_var.trace_add("write", lambda *args: self.safe_update_preview())
        self.folder_var.trace_add("write", lambda *args: self.safe_update_preview())
        
        # Date options
        ttk.Label(
            output_frame, 
            text="Date Suffix:",
            font=("Arial", 9, "bold")
        ).grid(row=2, column=0, sticky=tk.W, padx=(0, 10), pady=5)
        
        date_frame = ttk.Frame(output_frame)
        date_frame.grid(row=2, column=1, columnspan=2, sticky=tk.W, pady=5)
        
        self.date_option_var = tk.StringVar(value="month")
        ttk.Radiobutton(
            date_frame, 
            text="None", 
            variable=self.date_option_var, 
            value="none", 
            command=self.safe_update_preview
        ).pack(side=tk.LEFT, padx=(0, 15))
        
        ttk.Radiobutton(
            date_frame, 
            text="Month (YYYY-MM)", 
            variable=self.date_option_var, 
            value="month", 
            command=self.safe_update_preview
        ).pack(side=tk.LEFT, padx=(0, 15))
        
        ttk.Radiobutton(
            date_frame, 
            text="Quarter (YYYY-QQ)", 
            variable=self.date_option_var, 
            value="quarter", 
            command=self.safe_update_preview
        ).pack(side=tk.LEFT)
        
        # Preview - with better formatting
        ttk.Label(
            output_frame, 
            text="Output Preview:",
            font=("Arial", 9, "bold")
        ).grid(row=3, column=0, sticky=tk.W, padx=(0, 10), pady=(10, 5))
        
        preview_frame = ttk.Frame(output_frame)
        preview_frame.grid(row=3, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 5))
        preview_frame.columnconfigure(0, weight=1)
        
        self.fullpath_var = tk.StringVar()
        self.preview_label = ttk.Label(
            preview_frame, 
            textvariable=self.fullpath_var, 
            foreground="#2C3E50", 
            wraplength=450, 
            font=("Consolas", 9, "bold"),
            background="#F8F9FA",
            padding=5,
            relief=tk.SUNKEN,
            borderwidth=1
        )
        self.preview_label.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        # ==================== SECTION 3: PROCESS ====================
        process_frame = ttk.LabelFrame(main_frame, text="3. Process Files", padding="10")
        process_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        
        # Progress frame
        progress_frame = ttk.Frame(process_frame)
        progress_frame.pack(fill=tk.X, pady=5)
        
        self.progress = ttk.Progressbar(
            progress_frame, 
            mode='indeterminate', 
            length=400,
            style="Green.Horizontal.TProgressbar"
        )
        self.progress.pack(side=tk.LEFT, padx=(0, 10))
        
        self.process_btn = ttk.Button(
            progress_frame,
            text="🚀 Start Processing",
            command=self.start_processing,
            width=20,
            style="Primary.TButton"
        )
        self.process_btn.pack(side=tk.RIGHT)
        
        # Status label - CREATE THIS BEFORE CALLING update_preview
        self.status_var = tk.StringVar(value="Ready to process")
        self.status_label = ttk.Label(
            process_frame, 
            textvariable=self.status_var, 
            font=("Arial", 9), 
            foreground="gray"
        )
        self.status_label.pack(pady=(5, 0))
        
        # ==================== SECTION 4: LOG OUTPUT ====================
        log_frame = ttk.LabelFrame(main_frame, text="Processing Log", padding="10")
        log_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(
            log_frame, 
            height=12, 
            wrap=tk.WORD,
            font=("Consolas", 9),
            bg="white",
            relief=tk.SUNKEN,
            borderwidth=1
        )
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Log buttons frame
        log_buttons_frame = ttk.Frame(log_frame)
        log_buttons_frame.grid(row=1, column=0, sticky=tk.E, pady=(5, 0))
        
        ttk.Button(
            log_buttons_frame, 
            text="📋 Copy Log", 
            command=self.copy_log, 
            width=10
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            log_buttons_frame, 
            text="🧹 Clear Log", 
            command=self.clear_log, 
            width=10
        ).pack(side=tk.LEFT, padx=5)
        
        # Initial preview update
        self.safe_update_preview()
        
    def setup_styles(self):
        """Configure custom styles for the application"""
        style = ttk.Style()
        
        # Configure button styles
        style.configure("Primary.TButton", 
                       font=("Arial", 10, "bold"), 
                       padding=8,
                       background="#3498db",
                       foreground="white")
        
        style.configure("Action.TButton",
                       font=("Arial", 9),
                       padding=5)
        
        # Configure progress bar
        style.configure("Green.Horizontal.TProgressbar",
                       background="#2ecc71")
        
        # Configure label frames
        style.configure("TLabelframe", 
                       font=("Arial", 10, "bold"),
                       foreground="#2C3E50")
        
        style.configure("TLabelframe.Label",
                       font=("Arial", 10, "bold"),
                       foreground="#2C3E50")
    
    def safe_update_preview(self):
        """Safe version of update_preview that checks if status_var exists"""
        if hasattr(self, 'status_var'):
            self.update_preview()
        else:
            # If status_var doesn't exist yet, just update the path
            self.update_path_only()
    
    def update_path_only(self):
        """Update just the path without status_var"""
        folder = self.folder_var.get().strip()
        filename = self.filename_var.get().strip()
        
        if not folder:
            folder = os.path.join(os.getcwd(), "Social_Reports")
            self.folder_var.set(folder)
        
        if not filename:
            filename = "social_media_report"
            self.filename_var.set(filename)
            return
        
        # Clean filename
        cleaned_filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
        cleaned_filename = cleaned_filename.strip(' .')
        
        if cleaned_filename != filename:
            self.filename_var.set(cleaned_filename)
            return
        
        # Get date suffix
        date_option = self.date_option_var.get()
        if date_option == "month":
            date_suffix = datetime.now().strftime("%Y-%m")
        elif date_option == "quarter":
            month = datetime.now().month
            quarter = (month - 1) // 3 + 1
            date_suffix = f"{datetime.now().year}-Q{quarter}"
        else:  # none
            date_suffix = ""
        
        # Construct full path
        if date_suffix:
            full_path = os.path.join(folder, f"{cleaned_filename}_{date_suffix}.xlsx")
        else:
            full_path = os.path.join(folder, f"{cleaned_filename}.xlsx")
        
        # Update the display
        self.fullpath_var.set(full_path)
    
    def update_preview(self):
        """Update the full path preview in real-time"""
        folder = self.folder_var.get().strip()
        filename = self.filename_var.get().strip()
        
        # Ensure we have valid values
        if not folder:
            folder = os.path.join(os.getcwd(), "Social_Reports")
            self.folder_var.set(folder)
        
        if not filename:
            filename = "social_media_report"
            self.filename_var.set(filename)
            return
        
        # Clean filename (remove invalid characters)
        cleaned_filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
        cleaned_filename = cleaned_filename.strip(' .')
        
        if cleaned_filename != filename:
            self.filename_var.set(cleaned_filename)
            return
        
        # Get date suffix
        date_option = self.date_option_var.get()
        if date_option == "month":
            date_suffix = datetime.now().strftime("%Y-%m")
        elif date_option == "quarter":
            month = datetime.now().month
            quarter = (month - 1) // 3 + 1
            date_suffix = f"{datetime.now().year}-Q{quarter}"
        else:  # none
            date_suffix = ""
        
        # Construct full path
        if date_suffix:
            full_path = os.path.join(folder, f"{cleaned_filename}_{date_suffix}.xlsx")
        else:
            full_path = os.path.join(folder, f"{cleaned_filename}.xlsx")
        
        # Update the display
        self.fullpath_var.set(full_path)
        
        # Check if file exists (only if status_var exists)
        if hasattr(self, 'status_var'):
            if os.path.exists(full_path):
                try:
                    file_size = os.path.getsize(full_path)
                    size_text = self.format_file_size(file_size)
                    self.status_var.set(f"⚠️ File exists ({size_text}) - will be overwritten")
                    self.status_label.config(foreground="orange")
                except:
                    self.status_var.set("⚠️ File exists - will be overwritten")
                    self.status_label.config(foreground="orange")
            else:
                self.status_var.set("Ready to process")
                self.status_label.config(foreground="gray")
    
    def format_file_size(self, size_bytes):
        """Format file size in human readable format"""
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size_bytes < 1024.0:
                return f"{size_bytes:.1f} {unit}"
            size_bytes /= 1024.0
        return f"{size_bytes:.1f} TB"
    
    def add_files(self):
        """Add PowerPoint files to the list"""
        files = filedialog.askopenfilenames(
            title="Select PowerPoint Files",
            filetypes=[
                ("PowerPoint Files", "*.pptx *.ppt"),
                ("All Files", "*.*")
            ]
        )
        if files:
            for f in files:
                if f not in self.file_listbox.get(0, tk.END):
                    display_name = os.path.basename(f)
                    self.file_listbox.insert(tk.END, f"{display_name}")
                    self.selected_files.append(f)
            
            self.update_file_count()
            self.log_message(f"✅ Added {len(files)} file(s)")
            
            # Auto-update filename based on first file
            if len(self.selected_files) == 1:
                self.suggest_filename()
    
    def suggest_filename(self):
        """Suggest a filename based on selected files"""
        if not self.selected_files:
            return
            
        first_file = self.selected_files[0]
        base_name = os.path.splitext(os.path.basename(first_file))[0]
        
        # Clean up the name
        clean_name = re.sub(r'[^a-zA-Z0-9\s]', ' ', base_name)
        clean_name = re.sub(r'\s+', '_', clean_name.strip()).lower()
        
        if clean_name and clean_name != "social_media_report":
            self.filename_var.set(clean_name)
            self.log_message(f"📝 Suggested filename: {clean_name}")
    
    def remove_files(self):
        """Remove selected files from the list"""
        selected = self.file_listbox.curselection()
        if not selected:
            messagebox.showinfo("No Selection", "Please select files to remove.")
            return
            
        # Confirm removal
        if len(selected) == 1:
            msg = "Remove 1 selected file?"
        else:
            msg = f"Remove {len(selected)} selected files?"
            
        if messagebox.askyesno("Confirm Removal", msg):
            for i in reversed(selected):
                file_path = self.selected_files[i]
                self.file_listbox.delete(i)
                self.selected_files.pop(i)
            
            self.update_file_count()
            self.log_message(f"🗑️ Removed {len(selected)} file(s)")
    
    def clear_files(self):
        """Clear all files from the list"""
        if not self.selected_files:
            return
            
        if messagebox.askyesno("Clear All", "Clear all files from the list?"):
            self.file_listbox.delete(0, tk.END)
            self.selected_files.clear()
            self.update_file_count()
            self.log_message("🧹 Cleared all files")
    
    def update_file_count(self):
        """Update the file count display"""
        count = len(self.selected_files)
        if count == 0:
            self.file_count_var.set("No files selected")
            self.process_btn.config(state=tk.DISABLED)
        else:
            self.file_count_var.set(f"{count} file{'s' if count != 1 else ''} selected")
            self.process_btn.config(state=tk.NORMAL)
    
    def browse_folder(self):
        """Browse for output folder"""
        folder = filedialog.askdirectory(
            title="Select Output Folder",
            initialdir=self.folder_var.get()
        )
        if folder:
            self.folder_var.set(folder)
            self.log_message(f"📂 Output folder: {folder}")
    
    def clear_log(self):
        """Clear the log window"""
        self.log_text.delete(1.0, tk.END)
        self.log_message("Log cleared")
    
    def copy_log(self):
        """Copy log contents to clipboard"""
        log_content = self.log_text.get(1.0, tk.END).strip()
        if log_content:
            self.root.clipboard_clear()
            self.root.clipboard_append(log_content)
            self.log_message("📋 Log copied to clipboard")
    
    def log_message(self, msg):
        """Add a timestamped message to the log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {msg}\n")
        self.log_text.see(tk.END)
        self.root.update()
    
    def check_queue(self):
        """Check for messages from processing thread"""
        try:
            while True:
                msg_type, data = self.queue.get_nowait()
                if msg_type == "log":
                    self.log_message(data)
                elif msg_type == "progress_start":
                    self.progress.start()
                    self.process_btn.config(state=tk.DISABLED)
                    self.status_var.set("Processing...")
                elif msg_type == "progress_stop":
                    self.progress.stop()
                elif msg_type == "enable_ui":
                    self.process_btn.config(state=tk.NORMAL)
                    self.progress.stop()
                elif msg_type == "error":
                    messagebox.showerror("Error", data)
                    self.status_var.set("Error occurred")
                    self.status_label.config(foreground="red")
                elif msg_type == "success":
                    self.processing_success(data)
        except queue.Empty:
            pass
        self.root.after(100, self.check_queue)
    
    def start_processing(self):
        """Start the processing of selected files"""
        if not self.selected_files:
            messagebox.showwarning("No Files", "Please select at least one PowerPoint file.")
            return
        
        # Update preview one last time
        self.update_preview()
        output_file = self.fullpath_var.get()
        
        # Check for overwrite
        if os.path.exists(output_file):
            response = messagebox.askyesno(
                "Overwrite File?",
                f"The file already exists:\n\n{output_file}\n\nDo you want to overwrite it?"
            )
            if not response:
                return
        
        # Create output folder if it doesn't exist
        output_dir = os.path.dirname(output_file)
        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
                self.log_message(f"📁 Created folder: {output_dir}")
            except Exception as e:
                messagebox.showerror("Error", f"Cannot create folder: {e}")
                return
        
        # Start processing in thread
        self.queue.put(("progress_start", None))
        self.log_text.delete(1.0, tk.END)
        
        # Show file summary
        self.log_message("=" * 60)
        self.log_message("📊 PROCESSING SUMMARY")
        self.log_message("=" * 60)
        self.log_message(f"📁 Input Files: {len(self.selected_files)}")
        for i, file_path in enumerate(self.selected_files, 1):
            file_name = os.path.basename(file_path)
            self.log_message(f"  {i}. {file_name}")
        self.log_message(f"📄 Output File: {os.path.basename(output_file)}")
        self.log_message("=" * 60)
        self.log_message("🚀 Starting processing...")
        
        thread = threading.Thread(
            target=self.process_files_thread,
            args=(self.selected_files, output_file),
            daemon=True
        )
        thread.start()
    
    def process_files_thread(self, files, output_file):
        """Thread function to process files"""
        try:
            # Process all files
            count = process_presentations(files, output_file)
            
            if count > 0:
                self.queue.put(("success", {"count": count, "file": output_file}))
            else:
                self.queue.put(("error", "No posts were extracted from any file"))
                
        except Exception as e:
            error_msg = f"Error: {str(e)}"
            if "debug" in str(e).lower():
                error_msg += f"\n\n{traceback.format_exc()}"
            self.queue.put(("error", error_msg))
        finally:
            self.queue.put(("enable_ui", None))
    
    def processing_success(self, data):
        """Handle successful processing"""
        count = data["count"]
        output_file = data["file"]
        
        self.log_message("=" * 60)
        self.log_message("✅ PROCESSING COMPLETE")
        self.log_message("=" * 60)
        self.log_message(f"📊 Total Posts Extracted: {count:,}")
        self.log_message(f"📁 Output File: {output_file}")
        self.log_message("✨ Features: Filters, Sorting, Color Coding, Charts")
        
        # Update status
        self.status_var.set(f"✅ Extracted {count:,} posts")
        self.status_label.config(foreground="green")
        
        # Show success dialog
        features = [
            "✅ Auto-filter dropdowns in every column",
            "✅ One-click sorting (click column headers)",
            "✅ Color-coded post types",
            "✅ Conditional formatting for metrics",
            "✅ Summary statistics section",
            "✅ Interactive charts",
            "✅ Clickable social media links"
        ]
        
        features_text = "\n".join(features)
        
        response = messagebox.askyesno(
            "🎉 Processing Complete!",
            f"✅ Successfully extracted {count:,} posts\n\n"
            f"📁 Saved to:\n{output_file}\n\n"
            f"📊 Features included:\n{features_text}\n\n"
            "Open Excel file now?"
        )
        
        if response:
            try:
                os.startfile(output_file)
                self.log_message("📂 Opening Excel file...")
            except:
                self.log_message("⚠️ Could not open file automatically")

def main():
    """Main entry point"""
    root = tk.Tk()
    app = SocialMediaExtractorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()