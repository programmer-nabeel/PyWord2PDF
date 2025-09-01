import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import threading
import queue
from pathlib import Path
import time
from datetime import datetime
try:
    from docx2pdf import convert
    DOCX2PDF_AVAILABLE = True
except ImportError:
    DOCX2PDF_AVAILABLE = False

class WordToPDFConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Professional Word to PDF Converter")
        self.root.geometry("800x600")
        self.root.minsize(600, 500)
        
        # Configure style
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Variables
        self.source_folder = tk.StringVar()
        self.dest_folder = tk.StringVar()
        self.is_converting = False
        self.conversion_queue = queue.Queue()
        self.log_queue = queue.Queue()
        
        # Setup UI
        self.setup_ui()
        
        # Start queue processing
        self.process_queues()
        
        # Check dependencies
        self.check_dependencies()
    
    def check_dependencies(self):
        if not DOCX2PDF_AVAILABLE:
            self.log_message("ERROR: docx2pdf library not found!")
            self.log_message("Please install it using: pip install docx2pdf")
            messagebox.showerror("Missing Dependency", 
                               "docx2pdf library is required.\nInstall it using: pip install docx2pdf")
    
    def setup_ui(self):
        # Main container with padding
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(4, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Word to PDF Converter", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Source folder selection
        ttk.Label(main_frame, text="Source Folder:").grid(row=1, column=0, sticky=tk.W, pady=5)
        source_entry = ttk.Entry(main_frame, textvariable=self.source_folder, width=50)
        source_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=5)
        ttk.Button(main_frame, text="Browse", 
                  command=self.browse_source_folder).grid(row=1, column=2, padx=(0, 0), pady=5)
        
        # Destination folder selection
        ttk.Label(main_frame, text="Destination Folder:").grid(row=2, column=0, sticky=tk.W, pady=5)
        dest_entry = ttk.Entry(main_frame, textvariable=self.dest_folder, width=50)
        dest_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=5)
        ttk.Button(main_frame, text="Browse", 
                  command=self.browse_dest_folder).grid(row=2, column=2, padx=(0, 0), pady=5)
        
        # Control buttons frame
        control_frame = ttk.Frame(main_frame)
        control_frame.grid(row=3, column=0, columnspan=3, pady=20)
        
        self.convert_button = ttk.Button(control_frame, text="Start Conversion", 
                                        command=self.start_conversion, style='Accent.TButton')
        self.convert_button.pack(side=tk.LEFT, padx=(0, 10))
        
        self.cancel_button = ttk.Button(control_frame, text="Cancel", 
                                       command=self.cancel_conversion, state=tk.DISABLED)
        self.cancel_button.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(control_frame, text="Clear Log", 
                  command=self.clear_log).pack(side=tk.LEFT, padx=(0, 10))
        
        # Progress frame
        progress_frame = ttk.LabelFrame(main_frame, text="Progress", padding="10")
        progress_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        progress_frame.columnconfigure(0, weight=1)
        progress_frame.rowconfigure(1, weight=1)
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, 
                                          maximum=100, length=400)
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        
        self.progress_label = ttk.Label(progress_frame, text="Ready to convert")
        self.progress_label.grid(row=0, column=1, padx=(10, 0))
        
        # Log area
        log_frame = ttk.Frame(progress_frame)
        log_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, width=70, 
                                                 font=('Consolas', 9))
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, 
                              relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
    
    def browse_source_folder(self):
        folder = filedialog.askdirectory(title="Select Source Folder with Word Files")
        if folder:
            self.source_folder.set(folder)
            word_files = self.get_word_files(folder)
            self.log_message(f"Source folder selected: {folder}")
            self.log_message(f"Found {len(word_files)} Word files")
    
    def browse_dest_folder(self):
        folder = filedialog.askdirectory(title="Select Destination Folder for PDF Files")
        if folder:
            self.dest_folder.set(folder)
            self.log_message(f"Destination folder selected: {folder}")
    
    def get_word_files(self, folder):
        """Get all Word files from the source folder"""
        word_extensions = ['.doc', '.docx']
        word_files = []
        
        try:
            for file_path in Path(folder).iterdir():
                if file_path.is_file() and file_path.suffix.lower() in word_extensions:
                    word_files.append(file_path)
        except Exception as e:
            self.log_message(f"Error scanning folder: {str(e)}")
        
        return word_files
    
    def validate_inputs(self):
        """Validate user inputs before conversion"""
        if not self.source_folder.get():
            messagebox.showerror("Error", "Please select a source folder")
            return False
        
        if not self.dest_folder.get():
            messagebox.showerror("Error", "Please select a destination folder")
            return False
        
        if not os.path.exists(self.source_folder.get()):
            messagebox.showerror("Error", "Source folder does not exist")
            return False
        
        if not os.path.exists(self.dest_folder.get()):
            messagebox.showerror("Error", "Destination folder does not exist")
            return False
        
        word_files = self.get_word_files(self.source_folder.get())
        if not word_files:
            messagebox.showwarning("Warning", "No Word files found in source folder")
            return False
        
        if not DOCX2PDF_AVAILABLE:
            messagebox.showerror("Error", "docx2pdf library is not installed")
            return False
        
        return True
    
    def start_conversion(self):
        """Start the conversion process"""
        if not self.validate_inputs():
            return
        
        if self.is_converting:
            return
        
        self.is_converting = True
        self.convert_button.config(state=tk.DISABLED)
        self.cancel_button.config(state=tk.NORMAL)
        self.progress_var.set(0)
        self.status_var.set("Converting...")
        
        # Start conversion in separate thread
        conversion_thread = threading.Thread(target=self.convert_files, daemon=True)
        conversion_thread.start()
    
    def cancel_conversion(self):
        """Cancel the ongoing conversion"""
        self.is_converting = False
        self.log_message("Conversion cancelled by user")
        self.status_var.set("Cancelled")
        self.reset_ui()
    
    def convert_files(self):
        """Main conversion logic - runs in separate thread"""
        try:
            source_folder = self.source_folder.get()
            dest_folder = self.dest_folder.get()
            
            word_files = self.get_word_files(source_folder)
            total_files = len(word_files)
            
            self.log_message(f"Starting conversion of {total_files} files...")
            self.log_message(f"Source: {source_folder}")
            self.log_message(f"Destination: {dest_folder}")
            self.log_message("-" * 60)
            
            successful_conversions = 0
            failed_conversions = 0
            
            for index, word_file in enumerate(word_files):
                if not self.is_converting:
                    break
                
                try:
                    # Update progress
                    progress = (index / total_files) * 100
                    self.conversion_queue.put(('progress', progress, f"Converting {word_file.name}..."))
                    
                    # Convert file
                    pdf_name = word_file.stem + '.pdf'
                    pdf_path = os.path.join(dest_folder, pdf_name)
                    
                    self.log_message(f"Converting: {word_file.name}")
                    
                    # Perform conversion
                    convert(str(word_file), pdf_path)
                    
                    # Verify the PDF was created
                    if os.path.exists(pdf_path):
                        file_size = os.path.getsize(pdf_path)
                        self.log_message(f"✓ Success: {pdf_name} ({file_size:,} bytes)")
                        successful_conversions += 1
                    else:
                        self.log_message(f"✗ Failed: {word_file.name} - PDF not created")
                        failed_conversions += 1
                
                except Exception as e:
                    self.log_message(f"✗ Error converting {word_file.name}: {str(e)}")
                    failed_conversions += 1
                
                # Small delay to prevent UI freezing
                time.sleep(0.1)
            
            # Final update
            if self.is_converting:
                self.conversion_queue.put(('progress', 100, "Conversion completed"))
                self.log_message("-" * 60)
                self.log_message(f"Conversion completed!")
                self.log_message(f"Successful: {successful_conversions}")
                self.log_message(f"Failed: {failed_conversions}")
                self.log_message(f"Total: {total_files}")
                
                if failed_conversions == 0:
                    self.conversion_queue.put(('status', "All files converted successfully!"))
                else:
                    self.conversion_queue.put(('status', f"Completed with {failed_conversions} errors"))
            
        except Exception as e:
            self.log_message(f"Critical error during conversion: {str(e)}")
            self.conversion_queue.put(('status', "Conversion failed"))
        
        finally:
            self.conversion_queue.put(('finished', None, None))
    
    def process_queues(self):
        """Process updates from the conversion thread"""
        # Process conversion updates
        try:
            while True:
                update_type, value, message = self.conversion_queue.get_nowait()
                
                if update_type == 'progress':
                    self.progress_var.set(value)
                    if message:
                        self.progress_label.config(text=message)
                elif update_type == 'status':
                    self.status_var.set(value)
                elif update_type == 'finished':
                    self.reset_ui()
                    break
        except queue.Empty:
            pass
        
        # Process log messages
        try:
            while True:
                message = self.log_queue.get_nowait()
                self.log_text.insert(tk.END, message + "\n")
                self.log_text.see(tk.END)
        except queue.Empty:
            pass
        
        # Schedule next update
        self.root.after(100, self.process_queues)
    
    def log_message(self, message):
        """Add a message to the log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}"
        self.log_queue.put(formatted_message)
    
    def clear_log(self):
        """Clear the log text area"""
        self.log_text.delete(1.0, tk.END)
        self.log_message("Log cleared")
    
    def reset_ui(self):
        """Reset UI to initial state"""
        self.is_converting = False
        self.convert_button.config(state=tk.NORMAL)
        self.cancel_button.config(state=tk.DISABLED)
        self.progress_label.config(text="Ready")

def main():
    root = tk.Tk()
    app = WordToPDFConverter(root)
    
    # Center the window
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")
    
    root.mainloop()

if __name__ == "__main__":
    main()