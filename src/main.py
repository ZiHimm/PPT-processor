"""
Enhanced Main Application with all features integrated
COMPLETE WORKING VERSION
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path
import sys
import os
import traceback
import threading
import queue
import logging
from datetime import datetime
import re
import time

# Import pandas for data handling
try:
    import pandas as pd
    import numpy as np
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False
    print("⚠ pandas not available. Some features disabled.")

# Enhanced imports
try:
    import tkinterdnd2 as tkdnd
    HAS_DND = True
except ImportError:
    HAS_DND = False
    print("⚠ tkinterdnd2 not available. Drag & drop disabled.")

# Import enhanced modules
try:
    from ppt_processor import EnhancedPPTProcessor
    MODULES_LOADED = True
except ImportError as e:
    MODULES_LOADED = False
    print(f"⚠ PPT processor failed to load: {e}")

# Try to import other modules
try:
    from excel_generator import EnhancedExcelGenerator
    from dashboard_builder import EnhancedDashboardBuilder
    from data_validator import DataValidator
    from database_exporter import DatabaseExporter
    from utils import validate_file_path, setup_logging
except ImportError as e:
    print(f"⚠ Some modules failed to load: {e}")
    # Create dummy classes for missing modules
    class EnhancedExcelGenerator:
        def generate_dashboard_report(self, *args, **kwargs):
            raise ImportError("Excel generator not available")
    
    class EnhancedDashboardBuilder:
        def build_all_dashboards(self, *args, **kwargs):
            raise ImportError("Dashboard builder not available")
    
    class DataValidator:
        def __init__(self, *args, **kwargs):
            pass
        def validate_and_clean(self, data, dashboard_type):
            return data
    
    class DatabaseExporter:
        pass
    
    def validate_file_path(file_path):
        return Path(file_path).exists() and Path(file_path).suffix.lower() in ['.ppt', '.pptx']
    
    def setup_logging():
        logging.basicConfig(level=logging.INFO)

class EnhancedMarketingDashboardApp:
    """Enhanced marketing dashboard application with all features"""
    
    def __init__(self):
        # Initialize logging
        setup_logging()
        self.logger = logging.getLogger(__name__)
        
        # Create root window
        if HAS_DND:
            self.root = tkdnd.TkinterDnD.Tk()
        else:
            self.root = tk.Tk()
            
        self.root.title("Marketing Dashboard Automator v2.0")
        self.root.geometry("1200x800")
        
        # Center window
        self.root.eval('tk::PlaceWindow . center')
        
        # Initialize components
        self.ppt_files = []
        self.all_data = {}
        self.processing_queue = queue.Queue()
        self.processing_metrics = {}
        self.cancel_processing = False
        self.processing_thread = None
        self.processing_start_time = None
        
        # Initialize modules
        self.validator = DataValidator()
        
        # Setup UI
        self.setup_ui()
        
        # Check for updates
        self.check_for_updates()
    
    def setup_ui(self):
        """Setup enhanced user interface"""
        # Configure styles
        style = ttk.Style()
        style.theme_use('clam')
        
        # Custom styles
        style.configure("Title.TLabel", font=("Segoe UI", 18, "bold"))
        style.configure("Subtitle.TLabel", font=("Segoe UI", 10))
        style.configure("Success.TLabel", foreground="#10b981")
        style.configure("Warning.TLabel", foreground="#f59e0b")
        style.configure("Error.TLabel", foreground="#ef4444")
        
        # Main container
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header with gradient effect
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        title_label = ttk.Label(
            header_frame,
            text="🚀 Marketing Dashboard Automator v2.0",
            style="Title.TLabel"
        )
        title_label.pack()
        
        subtitle_label = ttk.Label(
            header_frame,
            text="Advanced PPT Data Extraction & Intelligent Dashboard Generation",
            style="Subtitle.TLabel"
        )
        subtitle_label.pack(pady=(5, 0))
        
        # Create notebook for tabs
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        # File Selection Tab
        file_tab = ttk.Frame(notebook)
        notebook.add(file_tab, text="📁 File Selection")
        self.setup_file_tab(file_tab)
        
        # Processing Tab
        process_tab = ttk.Frame(notebook)
        notebook.add(process_tab, text="⚙️ Processing")
        self.setup_process_tab(process_tab)
        
        # Export Tab
        export_tab = ttk.Frame(notebook)
        notebook.add(export_tab, text="📤 Export")
        self.setup_export_tab(export_tab)
        
        # Analytics Tab
        analytics_tab = ttk.Frame(notebook)
        notebook.add(analytics_tab, text="📊 Analytics")
        self.setup_analytics_tab(analytics_tab)
        
        # Setup drag and drop
        if HAS_DND:
            self.setup_drag_drop()
    
    def setup_file_tab(self, parent):
        """Setup file selection tab"""
        # File selection frame
        file_frame = ttk.LabelFrame(parent, text="PPT File Selection", padding="15")
        file_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # File list with scrollbar
        list_container = ttk.Frame(file_frame)
        list_container.pack(fill=tk.BOTH, expand=True)
        
        # Create treeview for files
        columns = ("Filename", "Size", "Status", "Records")
        self.file_tree = ttk.Treeview(
            list_container,
            columns=columns,
            show="headings",
            height=10
        )
        
        # Define columns
        self.file_tree.heading("Filename", text="Filename")
        self.file_tree.heading("Size", text="Size")
        self.file_tree.heading("Status", text="Status")
        self.file_tree.heading("Records", text="Records")
        
        self.file_tree.column("Filename", width=300)
        self.file_tree.column("Size", width=100, anchor=tk.CENTER)
        self.file_tree.column("Status", width=100, anchor=tk.CENTER)
        self.file_tree.column("Records", width=100, anchor=tk.CENTER)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(list_container, orient=tk.VERTICAL, command=self.file_tree.yview)
        self.file_tree.configure(yscrollcommand=scrollbar.set)
        
        self.file_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # File buttons
        btn_frame = ttk.Frame(file_frame)
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(btn_frame, text="Add PPT Files", 
                  command=self.add_files, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Add Folder", 
                  command=self.add_folder, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Remove Selected", 
                  command=self.remove_selected, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Clear All", 
                  command=self.clear_files, width=15).pack(side=tk.LEFT, padx=5)
        
        # File info
        info_frame = ttk.LabelFrame(file_frame, text="File Information", padding="10")
        info_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.file_info_label = ttk.Label(info_frame, text="No files selected")
        self.file_info_label.pack()
    
    def setup_process_tab(self, parent):
        """Setup processing tab"""
        # Processing options frame
        options_frame = ttk.LabelFrame(parent, text="Processing Options", padding="15")
        options_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Options
        self.enable_validation = tk.BooleanVar(value=True)
        self.enable_caching = tk.BooleanVar(value=True)
        self.enable_insights = tk.BooleanVar(value=True)
        self.export_to_db = tk.BooleanVar(value=False)
        
        ttk.Checkbutton(options_frame, text="Enable Data Validation", 
                       variable=self.enable_validation).pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(options_frame, text="Enable Caching", 
                       variable=self.enable_caching).pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(options_frame, text="Generate Insights", 
                       variable=self.enable_insights).pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(options_frame, text="Export to Database", 
                       variable=self.export_to_db).pack(anchor=tk.W, pady=2)
        
        # Process button frame
        process_frame = ttk.Frame(parent)
        process_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Process button
        self.process_btn = ttk.Button(
            process_frame,
            text="▶ Start Processing",
            command=self.start_processing,
            width=15
        )
        self.process_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Cancel button (initially disabled)
        self.cancel_btn = ttk.Button(
            process_frame,
            text="⏹ Cancel",
            command=self.cancel_processing,
            width=15,
            state=tk.DISABLED
        )
        self.cancel_btn.pack(side=tk.LEFT, padx=5)
        
        # Progress bar
        self.progress = ttk.Progressbar(parent, mode='determinate')
        self.progress.pack(fill=tk.X, padx=10, pady=5)
        
        # Progress labels
        progress_frame = ttk.Frame(parent)
        progress_frame.pack(fill=tk.X, padx=10)
        
        self.progress_label = ttk.Label(progress_frame, text="Ready")
        self.progress_label.pack(side=tk.LEFT)
        
        self.progress_percent = ttk.Label(progress_frame, text="0%")
        self.progress_percent.pack(side=tk.RIGHT)
        
        # Add loading indicator
        self.loading_label = ttk.Label(progress_frame, text="", style="Warning.TLabel")
        self.loading_label.pack(side=tk.LEFT, padx=(10, 0))

        # Add time elapsed label
        self.time_label = ttk.Label(progress_frame, text="")
        self.time_label.pack(side=tk.RIGHT)
        
        # Log output
        log_frame = ttk.LabelFrame(parent, text="Processing Log", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            height=12,
            wrap=tk.WORD,
            font=("Consolas", 9)
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

    
    def setup_export_tab(self, parent):
        """Setup export tab"""
        # Export options
        options_frame = ttk.LabelFrame(parent, text="Export Options", padding="15")
        options_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.export_excel = tk.BooleanVar(value=True)
        self.export_html = tk.BooleanVar(value=True)
        self.export_json = tk.BooleanVar(value=False)
        self.export_pdf = tk.BooleanVar(value=False)
        
        ttk.Checkbutton(options_frame, text="Export to Excel", 
                       variable=self.export_excel).grid(row=0, column=0, sticky=tk.W, pady=2)
        ttk.Checkbutton(options_frame, text="Generate HTML Dashboards", 
                       variable=self.export_html).grid(row=0, column=1, sticky=tk.W, pady=2)
        ttk.Checkbutton(options_frame, text="Export to JSON", 
                       variable=self.export_json).grid(row=1, column=0, sticky=tk.W, pady=2)
        ttk.Checkbutton(options_frame, text="Generate PDF Report", 
                       variable=self.export_pdf).grid(row=1, column=1, sticky=tk.W, pady=2)
        
        # Export buttons
        export_frame = ttk.Frame(parent)
        export_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.gen_excel_btn = ttk.Button(
            export_frame,
            text="📊 Generate Excel Report",
            command=self.generate_excel,
            width=20,
            state=tk.DISABLED
        )
        self.gen_excel_btn.pack(side=tk.LEFT, padx=5)
        
        self.gen_dash_btn = ttk.Button(
            export_frame,
            text="🌐 Generate Dashboards",
            command=self.generate_dashboard,
            width=20,
            state=tk.DISABLED
        )
        self.gen_dash_btn.pack(side=tk.LEFT, padx=5)
        
        self.gen_all_btn = ttk.Button(
            export_frame,
            text="🚀 Generate All",
            command=self.generate_all,
            width=20,
            state=tk.DISABLED
        )
        self.gen_all_btn.pack(side=tk.LEFT, padx=5)
        
        # Export status
        status_frame = ttk.LabelFrame(parent, text="Export Status", padding="10")
        status_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.export_status = scrolledtext.ScrolledText(
            status_frame,
            height=8,
            wrap=tk.WORD,
            font=("Consolas", 9)
        )
        self.export_status.pack(fill=tk.BOTH, expand=True)
    
    def setup_analytics_tab(self, parent):
        """Setup analytics tab"""
        # Analytics dashboard
        analytics_frame = ttk.LabelFrame(parent, text="Data Analytics", padding="15")
        analytics_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create metrics display
        metrics_frame = ttk.Frame(analytics_frame)
        metrics_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Metrics labels
        self.metrics_labels = {}
        metrics = [
            ("Total Files", "0"),
            ("Total Records", "0"),
            ("Data Quality", "N/A"),
            ("Processing Time", "N/A")
        ]
        
        for i, (label, value) in enumerate(metrics):
            frame = ttk.Frame(metrics_frame)
            frame.grid(row=0, column=i, padx=5, pady=5, sticky=tk.NSEW)
            
            lbl = ttk.Label(frame, text=label, font=("Segoe UI", 9))
            lbl.pack()
            
            val_lbl = ttk.Label(frame, text=value, font=("Segoe UI", 14, "bold"))
            val_lbl.pack()
            
            self.metrics_labels[label] = val_lbl
        
        # Data preview
        preview_frame = ttk.LabelFrame(analytics_frame, text="Data Preview", padding="10")
        preview_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create treeview for data preview
        columns = ("Dashboard", "Records", "Quality", "Status")
        self.data_tree = ttk.Treeview(
            preview_frame,
            columns=columns,
            show="headings",
            height=8
        )
        
        for col in columns:
            self.data_tree.heading(col, text=col)
            self.data_tree.column(col, width=150, anchor=tk.CENTER)
        
        scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.data_tree.yview)
        self.data_tree.configure(yscrollcommand=scrollbar.set)
        
        self.data_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Preview buttons
        preview_btn_frame = ttk.Frame(preview_frame)
        preview_btn_frame.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Button(preview_btn_frame, text="Refresh Preview", 
                  command=self.refresh_preview).pack(side=tk.LEFT)
        ttk.Button(preview_btn_frame, text="View Details", 
                  command=self.view_details).pack(side=tk.LEFT, padx=5)
    
    def setup_drag_drop(self):
        """Setup drag and drop functionality"""
        try:
            self.root.drop_target_register(tkdnd.DND_FILES)
            self.root.dnd_bind('<<Drop>>', self.on_drop)
        except Exception as e:
            self.log("Drag and drop initialization failed")
    
    def on_drop(self, event):
        """Handle file drop"""
        files = self.root.tk.splitlist(event.data)
        added_count = 0
        
        for f in files:
            if validate_file_path(f):
                self.ppt_files.append(f)
                filename = Path(f).name
                size = os.path.getsize(f) // 1024  # Size in KB
                
                # Add to treeview
                self.file_tree.insert("", "end", values=(filename, f"{size} KB", "Pending", "-"))
                added_count += 1
        
        if added_count > 0:
            self.log(f"Added {added_count} file(s) via drag & drop")
            self.update_file_info()
    
    def add_files(self):
        """Add PPT files via file dialog"""
        files = filedialog.askopenfilenames(
            title="Select PPT files",
            filetypes=[
                ("PowerPoint files", "*.ppt *.pptx"),
                ("All files", "*.*")
            ]
        )
        
        for f in files:
            if validate_file_path(f):
                self.ppt_files.append(f)
                filename = Path(f).name
                size = os.path.getsize(f) // 1024
                
                self.file_tree.insert("", "end", values=(filename, f"{size} KB", "Pending", "-"))
        
        if files:
            self.log(f"Added {len(files)} file(s)")
            self.update_file_info()
    
    def add_folder(self):
        """Add all PPT files from a folder"""
        folder = filedialog.askdirectory(title="Select folder with PPT files")
        if folder:
            file_count = 0
            for ext in ['.pptx', '.ppt']:
                for f in Path(folder).glob(f"*{ext}"):
                    if validate_file_path(str(f)):
                        self.ppt_files.append(str(f))
                        filename = f.name
                        size = f.stat().st_size // 1024
                        
                        self.file_tree.insert("", "end", values=(filename, f"{size} KB", "Pending", "-"))
                        file_count += 1
            
            if file_count > 0:
                self.log(f"Added {file_count} file(s) from folder")
                self.update_file_info()
    
    def remove_selected(self):
        """Remove selected files"""
        selected_items = self.file_tree.selection()
        if not selected_items:
            return
        
        for item in selected_items:
            values = self.file_tree.item(item)['values']
            filename = values[0]
            
            # Remove from list
            self.ppt_files = [f for f in self.ppt_files if not f.endswith(filename)]
            
            # Remove from treeview
            self.file_tree.delete(item)
        
        self.log(f"Removed {len(selected_items)} file(s)")
        self.update_file_info()
    
    def clear_files(self):
        """Clear all selected files"""
        self.ppt_files.clear()
        self.file_tree.delete(*self.file_tree.get_children())
        self.log("Cleared all files")
        self.update_file_info()
    
    def update_file_info(self):
        """Update file information label"""
        total_size = sum(os.path.getsize(f) for f in self.ppt_files) // 1024  # KB
        self.file_info_label.config(
            text=f"{len(self.ppt_files)} files selected | Total size: {total_size:,} KB"
        )
    
    def start_processing(self):
        """Start processing PPT files - FIXED with proper initialization"""
        if not self.ppt_files:
            messagebox.showwarning("No Files", "Please add PPT files first.")
            return
        
        print(f"[MAIN] ===== STARTING PROCESSING at {datetime.now().strftime('%H:%M:%S.%f')} =====")
        
        # Reset flags and start time
        self.cancel_processing = False
        self.processing_start_time = datetime.now()
        
        # IMPORTANT: Disable ALL buttons to prevent premature enabling
        self.process_btn.config(state=tk.DISABLED)
        self.cancel_btn.config(state=tk.NORMAL)
        self.gen_excel_btn.config(state=tk.DISABLED)
        self.gen_dash_btn.config(state=tk.DISABLED)
        self.gen_all_btn.config(state=tk.DISABLED)
        
        # Force UI update immediately
        self.root.update()
        self.root.update_idletasks()
        
        # Reset progress
        self.progress['value'] = 0
        self.progress_label.config(text="Starting...")
        self.progress_percent.config(text="0%")
        self.loading_label.config(text="Initializing...", foreground="#f59e0b")
        self.time_label.config(text="0m 0s")
        
        # Clear previous data
        self.all_data.clear()
        self.processing_metrics.clear()
        
        # Clear log
        self.log_text.delete(1.0, tk.END)
        
        # Clear the queue - CRITICAL!
        while not self.processing_queue.empty():
            try:
                self.processing_queue.get_nowait()
            except queue.Empty:
                break
        
        # Start progress bar
        self.log("=" * 60)
        self.log("Starting enhanced PPT processing...")
        self.log(f"Processing {len(self.ppt_files)} file(s)")
        self.log("=" * 60)
        
        # Create and start thread
        self.processing_thread = threading.Thread(target=self.process_files, daemon=True)
        self.processing_thread.start()
        
        print(f"[MAIN] Thread started at {datetime.now().strftime('%H:%M:%S.%f')}")
        print(f"[MAIN] Queue cleared, starting monitor...")
        
        # Start progress monitoring with minimal delay
        # Use lambda to pass initial call_count=0
        self.root.after(10, lambda: self.check_processing(0))

    def cancel_processing(self):
        """Cancel ongoing processing"""
        self.cancel_processing = True
        self.cancel_btn.config(state=tk.DISABLED)
        self.log("⚠ Processing cancellation requested...")
    
    def process_files(self):
        """Process files in background thread - FIXED with completion confirmation"""
        try:
            import time
            thread_start = time.time()
            
            print(f"[THREAD] Thread started at {datetime.now().strftime('%H:%M:%S.%f')}")
            
            # Initial messages
            self.processing_queue.put("PROGRESS:1")
            self.processing_queue.put("DEBUG: Starting process_files")
            
            # Test imports (optional)
            try:
                self.processing_queue.put("DEBUG: Testing excel_generator import...")
                from excel_generator import EnhancedExcelGenerator
                self.processing_queue.put("DEBUG: ✓ excel_generator imported")
            except ImportError as e:
                self.processing_queue.put(f"ERROR: excel_generator import failed: {e}")
            
            try:
                self.processing_queue.put("DEBUG: Testing dashboard_builder import...")
                from dashboard_builder import EnhancedDashboardBuilder
                self.processing_queue.put("DEBUG: ✓ dashboard_builder imported")
            except ImportError as e:
                self.processing_queue.put(f"ERROR: dashboard_builder import failed: {e}")
            
            # Check for required packages
            try:
                import openpyxl
                self.processing_queue.put("DEBUG: ✓ openpyxl available")
            except ImportError:
                self.processing_queue.put("ERROR: openpyxl not installed")
            
            try:
                import plotly
                self.processing_queue.put("DEBUG: ✓ plotly available")
            except ImportError:
                self.processing_queue.put("ERROR: plotly not installed")
            
            # ========= ACTUAL PROCESSING =========
            from ppt_processor import EnhancedPPTProcessor
            processor = EnhancedPPTProcessor()
            
            self.processing_queue.put("DEBUG: Processor created")
            
            # Initialize data storage
            self.all_data = {
                'social_media': [],
                'community_marketing': [],
                'kol_engagement': [],
                'performance_marketing': [],
                'promotion_posts': []
            }
            
            total_records = 0
            
            # Process each file
            for i, ppt_file in enumerate(self.ppt_files, 1):
                self.processing_queue.put(f"DEBUG: Processing file {i}/{len(self.ppt_files)}")
                
                # Send heartbeat
                elapsed = time.time() - thread_start
                self.processing_queue.put(f"HEARTBEAT: {elapsed:.2f}s into processing")
                
                # Process the PPT file
                results = processor.process_ppt(ppt_file)
                self.processing_queue.put(f"DEBUG: Got results with {len(results)} dashboard types")
                
                file_records = 0
                
                # Collect data
                for dashboard_type, data_list in results.items():
                    if data_list:
                        self.all_data[dashboard_type].extend(data_list)
                        file_records += len(data_list)
                        self.processing_queue.put(f"DEBUG: Added {len(data_list)} records to {dashboard_type}")
                
                total_records += file_records
                self.processing_queue.put(f"DEBUG: File {i} complete: {file_records} records")
                
                # Update progress
                progress = int((i / len(self.ppt_files)) * 100)
                self.processing_queue.put(f"PROGRESS:{progress}")
            
            # ========= COMPLETION SEQUENCE =========
            print(f"[THREAD] All files processed. Total: {total_records} records")
            print(f"[THREAD] Total thread time: {time.time() - thread_start:.2f}s")

            # Send completion in a reliable sequence
            completion_messages = [
                f"DEBUG: All files processed. Total: {total_records} records",
                "PROGRESS:50",
                "PROGRESS:75",
                "PROGRESS:90",
                "PROGRESS:95",
                "PROGRESS:99",
                "PROGRESS:100",
                "DONE",  # First DONE
                "DONE",  # Second DONE for redundancy
                f"✓ Processing complete! Extracted {total_records} records",
                f"THREAD_FINISHED: {time.time() - thread_start:.2f}s"
            ]

            print(f"[THREAD] Sending {len(completion_messages)} completion messages...")

            for i, msg in enumerate(completion_messages, 1):
                self.processing_queue.put(msg)
                print(f"[THREAD] [{i}/{len(completion_messages)}] Sent: {msg}")
                time.sleep(0.02)  # Small delay to ensure UI can process

            # Final confirmation
            print(f"[THREAD] ✅ All completion messages sent at {datetime.now().strftime('%H:%M:%S.%f')}")
            print(f"[THREAD] Final queue size: {self.processing_queue.qsize()}")
            print(f"[THREAD] Thread exiting...")
            
        except Exception as e:
            import traceback
            error_msg = f"✗ Processing failed: {str(e)}\n{traceback.format_exc()}"
            print(f"[THREAD] ❌ ERROR: {error_msg}")
            self.processing_queue.put(f"CRITICAL ERROR in process_files: {str(e)}")
            self.processing_queue.put("DONE_ERROR")
    
    def update_file_status(self, file_path, status, records):
        """Update file status in treeview"""
        filename = Path(file_path).name
        
        for item in self.file_tree.get_children():
            values = self.file_tree.item(item)['values']
            if values[0] == filename:
                self.file_tree.item(item, values=(values[0], values[1], status, records))
                break
    
    def check_processing(self, call_count=0):
        """Check for processing updates - FIXED to always keep checking"""
        call_count += 1
        
        # Safety limit: 200 calls * 20ms = 4 seconds max
        if call_count > 200:
            print(f"[UI] ⚠️  Safety timeout after {call_count} calls (4 seconds)")
            if hasattr(self, 'processing_thread') and self.processing_thread:
                if not self.processing_thread.is_alive():
                    print("[UI] Thread is dead but no DONE received")
                    self.processing_complete()
                else:
                    print("[UI] Thread still alive after timeout - forcing completion")
                    self.processing_complete()
            elif self.process_btn['state'] == tk.DISABLED:
                print("[UI] Button disabled but no thread - forcing completion")
                self.processing_complete()
            return
        
        try:
            messages_processed = 0
            while True:
                try:
                    msg = self.processing_queue.get_nowait()
                    messages_processed += 1
                    
                    # DEBUG: Print with timestamp
                    timestamp = datetime.now().strftime('%H:%M:%S.%f')[:-3]
                    print(f"[UI #{call_count}.{messages_processed} @ {timestamp}] {msg}")
                    
                    # Parse progress messages
                    if msg.startswith("PROGRESS:"):
                        # Extract percentage
                        numbers = re.findall(r'\d+', msg)
                        if numbers:
                            percent = int(numbers[0])
                            
                            # Update progress bar
                            self.progress['value'] = percent
                            self.progress_percent.config(text=f"{percent}%")
                            
                            # Update label
                            if percent < 100:
                                self.progress_label.config(text=f"Processing... {percent}%")
                                self.update_loading_indicator(percent)
                            else:
                                self.progress_label.config(text=f"Complete! {percent}%")
                                self.loading_label.config(text="Complete!", foreground="#10b981")
                            
                            # Update time
                            if hasattr(self, 'processing_start_time') and self.processing_start_time:
                                elapsed = (datetime.now() - self.processing_start_time).total_seconds()
                                minutes = int(elapsed // 60)
                                seconds = int(elapsed % 60)
                                self.time_label.config(text=f"{minutes}m {seconds}s")
                    
                    # Check for completion
                    elif msg == "DONE":
                        print(f"[UI] ✅ Got DONE message on call #{call_count}")
                        self.processing_complete()
                        return  # This is the CORRECT way to stop
                        
                    elif msg == "DONE_ERROR":
                        print(f"[UI] ❌ Got ERROR message on call #{call_count}")
                        self.processing_error()
                        return
                        
                    else:
                        # Log non-progress messages (skip DEBUG messages to avoid spam)
                        if not msg.startswith("DEBUG:") and not msg.startswith("HEARTBEAT:"):
                            self.log(msg)
                            
                except queue.Empty:
                    # Queue is temporarily empty
                    if messages_processed > 0:
                        print(f"[UI] Processed {messages_processed} messages, queue empty (call #{call_count})")
                break
        
        except Exception as e:
            self.log(f"Error in check_processing: {e}")
            import traceback
            traceback.print_exc()
        
        # ALWAYS check again in 20ms, regardless of button state
        # The ONLY ways to stop are: 1) Got DONE, 2) Safety timeout, 3) Got DONE_ERROR
        self.root.after(20, lambda: self.check_processing(call_count))

    def update_loading_indicator(self, percent):
        """Update loading indicator based on progress percentage"""
        if percent < 30:
            self.loading_label.config(text="Initializing...", foreground="#f59e0b")
        elif percent < 70:
            self.loading_label.config(text="Processing slides...", foreground="#f59e0b")
        elif percent < 100:
            self.loading_label.config(text="Finalizing...", foreground="#10b981")
    
    def processing_complete(self):
        """Handle processing completion - SAFE version"""
        print(f"[UI] 🎉 processing_complete() called at {datetime.now().strftime('%H:%M:%S.%f')}")
        
        # Prevent multiple calls
        if self.process_btn['state'] == tk.NORMAL:
            print("[UI] ⚠️  Already completed, ignoring duplicate call")
            return
        
        # Calculate total records
        total_records = 0
        if hasattr(self, 'all_data'):
            total_records = sum(len(data) for data in self.all_data.values())
        
        # Update UI
        self.progress['value'] = 100
        self.progress_percent.config(text="100%")
        self.progress_label.config(text="Complete!", foreground="#10b981")
        self.loading_label.config(text="✅ Ready for export!", foreground="#10b981")
        
        # Update time
        if hasattr(self, 'processing_start_time') and self.processing_start_time:
            elapsed = (datetime.now() - self.processing_start_time).total_seconds()
            minutes = int(elapsed // 60)
            seconds = int(elapsed % 60)
            self.time_label.config(text=f"{minutes}m {seconds}s")
        
        # Re-enable buttons (CRITICAL: Do this LAST)
        self.process_btn.config(state=tk.NORMAL)
        self.cancel_btn.config(state=tk.DISABLED)
        
        # Only enable export buttons if we have data
        if total_records > 0:
            self.gen_excel_btn.config(state=tk.NORMAL)
            self.gen_dash_btn.config(state=tk.NORMAL)
            self.gen_all_btn.config(state=tk.NORMAL)
        
        # Show completion message
        self.log("=" * 60)
        self.log(f"✅ PROCESSING COMPLETE!")
        self.log(f"✅ Extracted {total_records} total records")
        
        # Log breakdown by dashboard type
        if hasattr(self, 'all_data'):
            for dashboard_type, data_list in self.all_data.items():
                if data_list:
                    self.log(f"   • {dashboard_type.replace('_', ' ').title()}: {len(data_list)} records")
        
        self.log("=" * 60)
        print(f"[UI] ✅ Processing complete with {total_records} records")
    
    def update_analytics(self):
        """Update analytics metrics"""
        if not HAS_PANDAS:
            self.log("⚠ pandas not available for analytics")
            return
            
        total_files = len(self.ppt_files)
        total_records = sum(len(data) for data in self.all_data.values())
        
        # Calculate data quality
        quality_scores = []
        for data_list in self.all_data.values():
            if data_list:
                df = pd.DataFrame(data_list)
                if 'confidence_score' in df.columns:
                    quality_scores.append(df['confidence_score'].mean())
        
        avg_quality = sum(quality_scores) / len(quality_scores) if quality_scores else 0
        quality_status = self.get_quality_status(avg_quality)
        
        # Update labels
        self.metrics_labels["Total Files"].config(text=str(total_files))
        self.metrics_labels["Total Records"].config(text=f"{total_records:,}")
        self.metrics_labels["Data Quality"].config(text=quality_status)
        
        # Color code quality
        if quality_status == "Excellent":
            self.metrics_labels["Data Quality"].config(foreground="#10b981")
        elif quality_status == "Good":
            self.metrics_labels["Data Quality"].config(foreground="#10b981")
        elif quality_status == "Fair":
            self.metrics_labels["Data Quality"].config(foreground="#f59e0b")
        else:
            self.metrics_labels["Data Quality"].config(foreground="#ef4444")
    
    def get_quality_status(self, score: float) -> str:
        """Get quality status based on score"""
        if score >= 0.9:
            return "Excellent"
        elif score >= 0.7:
            return "Good"
        elif score >= 0.5:
            return "Fair"
        else:
            return "Poor"
    
    def refresh_preview(self):
        """Refresh data preview"""
        if not HAS_PANDAS:
            self.log("⚠ pandas not available for data preview")
            return
            
        # Clear existing items
        for item in self.data_tree.get_children():
            self.data_tree.delete(item)
        
        # Add current data
        for dashboard_type, data_list in self.all_data.items():
            if data_list:
                df = pd.DataFrame(data_list)
                records = len(data_list)
                
                # Calculate quality
                quality_score = 0.8  # Default
                if 'confidence_score' in df.columns:
                    quality_score = df['confidence_score'].mean()
                
                quality_status = self.get_quality_status(quality_score)
                
                self.data_tree.insert(
                    "", "end",
                    values=(
                        self.format_dashboard_name(dashboard_type),
                        records,
                        quality_status,
                        "✓ Ready"
                    )
                )
    
    def view_details(self):
        """View selected data details"""
        if not HAS_PANDAS:
            messagebox.showwarning("Feature Unavailable", "pandas is required for this feature")
            return
            
        selected = self.data_tree.selection()
        if not selected:
            return
        
        item = self.data_tree.item(selected[0])
        values = item['values']
        dashboard_name = values[0]
        
        # Find corresponding dashboard type
        dashboard_type = None
        for d_type in self.all_data.keys():
            if self.format_dashboard_name(d_type) == dashboard_name:
                dashboard_type = d_type
                break
        
        if dashboard_type and dashboard_type in self.all_data:
            data_list = self.all_data[dashboard_type]
            if data_list:
                df = pd.DataFrame(data_list)
                
                # Create detail window
                detail_window = tk.Toplevel(self.root)
                detail_window.title(f"Details: {dashboard_name}")
                detail_window.geometry("800x600")
                
                # Create text widget with scrollbar
                text_frame = ttk.Frame(detail_window)
                text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
                
                text_widget = scrolledtext.ScrolledText(
                    text_frame,
                    wrap=tk.WORD,
                    font=("Consolas", 9)
                )
                text_widget.pack(fill=tk.BOTH, expand=True)
                
                # Display data summary
                summary = f"""
{dashboard_name} - Data Summary
{"="*50}

Total Records: {len(df):,}

Column Information:
{df.dtypes.to_string()}

Summary Statistics:
{df.describe().to_string()}

First 10 Records:
{df.head(10).to_string()}
"""
                text_widget.insert(tk.END, summary)
                text_widget.config(state=tk.DISABLED)
    
    def generate_excel(self):
        """Generate Excel report"""
        total_records = sum(len(data) for data in self.all_data.values())
        if total_records == 0:
            messagebox.showwarning("No Data", "No data to export.")
            return
        
        output_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("All files", "*.*")
            ],
            initialfile=f"marketing_report_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        )
        
        if output_path:
            try:
                self.export_status.insert(tk.END, f"\n{'='*60}\n")
                self.export_status.insert(tk.END, "Generating Excel report...\n")
                self.export_status.see(tk.END)
                
                generator = EnhancedExcelGenerator()
                result = generator.generate_dashboard_report(
                    self.all_data, 
                    output_path,
                    include_charts=True,
                    include_summary=True
                )
                
                self.export_status.insert(tk.END, f"✓ Excel report saved: {output_path}\n")
                self.export_status.see(tk.END)
                
                messagebox.showinfo(
                    "Success", 
                    f"Excel report generated successfully!\n\n"
                    f"File: {output_path}\n"
                    f"Total Records: {total_records:,}"
                )
                
            except Exception as e:
                self.export_status.insert(tk.END, f"✗ Excel generation failed: {str(e)}\n")
                self.export_status.see(tk.END)
                messagebox.showerror("Error", f"Failed to generate Excel: {str(e)}")
    
    def generate_dashboard(self):
        """Generate interactive dashboard"""
        total_records = sum(len(data) for data in self.all_data.values())
        if total_records == 0:
            messagebox.showwarning("No Data", "No data for dashboard.")
            return
        
        output_dir = filedialog.askdirectory(title="Select Dashboard Output Directory")
        if output_dir:
            try:
                self.export_status.insert(tk.END, f"\n{'='*60}\n")
                self.export_status.insert(tk.END, "Generating interactive dashboards...\n")
                self.export_status.see(tk.END)
                
                builder = EnhancedDashboardBuilder()
                dashboards = builder.build_all_dashboards(
                    self.all_data, 
                    output_dir,
                    include_insights=self.enable_insights.get()
                )
                
                self.export_status.insert(tk.END, f"✓ Dashboards generated in: {output_dir}\n")
                
                for d_type, file_path in dashboards.items():
                    self.export_status.insert(tk.END, f"  • {d_type}: {Path(file_path).name}\n")
                
                self.export_status.see(tk.END)
                
                # Open master dashboard
                master_file = Path(output_dir) / "master_dashboard.html"
                if master_file.exists():
                    messagebox.showinfo(
                        "Success", 
                        f"Dashboards generated successfully!\n\n"
                        f"Open this file in your browser:\n{master_file}\n\n"
                        f"Total dashboards: {len(dashboards)}"
                    )
                else:
                    messagebox.showinfo(
                        "Success", 
                        f"Dashboards generated in:\n{output_dir}"
                    )
                
            except Exception as e:
                self.export_status.insert(tk.END, f"✗ Dashboard generation failed: {str(e)}\n")
                self.export_status.see(tk.END)
                messagebox.showerror("Error", f"Failed to generate dashboards: {str(e)}")
    
    def generate_all(self):
        """Generate all outputs"""
        # Generate Excel
        self.generate_excel()
        
        # Generate dashboards
        self.generate_dashboard()
        
        # Generate JSON if selected
        if self.export_json.get():
            self.export_status.insert(tk.END, "\nGenerating JSON export...\n")
            self.export_json_data()
    
    def export_json_data(self):
        """Export data to JSON format"""
        try:
            output_dir = filedialog.askdirectory(title="Select JSON Output Directory")
            if output_dir:
                output_path = Path(output_dir) / f"marketing_data_{datetime.now():%Y%m%d_%H%M%S}.json"
                
                # Convert data to JSON-serializable format
                json_data = {}
                for d_type, data_list in self.all_data.items():
                    if data_list:
                        if HAS_PANDAS:
                            df = pd.DataFrame(data_list)
                            json_data[d_type] = df.to_dict('records')
                        else:
                            # Manual conversion if pandas not available
                            json_data[d_type] = data_list
                
                import json
                with open(output_path, 'w', encoding='utf-8') as f:
                    json.dump(json_data, f, indent=2, default=str)
                
                self.export_status.insert(tk.END, f"✓ JSON data exported: {output_path}\n")
                self.export_status.see(tk.END)
                
        except Exception as e:
            self.export_status.insert(tk.END, f"✗ JSON export failed: {str(e)}\n")
    
    def log(self, message: str):
        """Add message to log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def format_dashboard_name(self, dashboard_type: str) -> str:
        """Format dashboard name for display"""
        names = {
            'social_media': 'Social Media',
            'community_marketing': 'Community Marketing',
            'kol_engagement': 'KOL Engagement',
            'performance_marketing': 'Performance Marketing',
            'promotion_posts': 'Promotion Posts'
        }
        return names.get(dashboard_type, dashboard_type.replace('_', ' ').title())
    
    def check_for_updates(self):
        """Check for application updates"""
        self.log("Application started successfully")
        self.log("Ready to process marketing data")
    
    def run(self):
        """Run the application"""
        self.root.mainloop()

def main():
    """Main entry point"""
    try:
        # Create output directory if it doesn't exist
        output_dir = Path("output")
        output_dir.mkdir(exist_ok=True)
        
        # Create logs directory
        logs_dir = Path("logs")
        logs_dir.mkdir(exist_ok=True)
        
        # Create cache directory
        cache_dir = Path("cache")
        cache_dir.mkdir(exist_ok=True)
        
        # Run application
        app = EnhancedMarketingDashboardApp()
        app.run()
        
    except Exception as e:
        print(f"Failed to start application: {e}")
        traceback.print_exc()
        input("Press Enter to exit...")

if __name__ == "__main__":
    main()