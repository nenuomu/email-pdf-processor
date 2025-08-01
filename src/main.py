import pandas as pd
import re
import pdfplumber
import os
import sys
import threading
import customtkinter as ctk
from datetime import datetime
import email
import tempfile
import shutil
from pathlib import Path

# Set the appearance of customtkinter
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("green")

class EmailPDFProcessor:
    def __init__(self):
        self.temp_pdf_folder = None
        
    def extract_pdfs_from_eml(self, eml_file_path, output_folder):
        """Extract PDFs from .eml email files"""
        pdfs_found = []
        
        try:
            with open(eml_file_path, 'rb') as f:
                msg = email.message_from_bytes(f.read())
            
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_disposition() == 'attachment':
                        filename = part.get_filename()
                        if filename and filename.lower().endswith('.pdf'):
                            pdf_data = part.get_payload(decode=True)
                            if pdf_data:
                                final_filename = self.get_unique_filename(output_folder, filename)
                                output_path = os.path.join(output_folder, final_filename)
                                with open(output_path, 'wb') as pdf_file:
                                    pdf_file.write(pdf_data)
                                pdfs_found.append(output_path)
            else:
                content_type = msg.get_content_type()
                if content_type == 'application/pdf':
                    filename = msg.get_filename() or 'attachment.pdf'
                    pdf_data = msg.get_payload(decode=True)
                    if pdf_data:
                        final_filename = self.get_unique_filename(output_folder, filename)
                        output_path = os.path.join(output_folder, final_filename)
                        with open(output_path, 'wb') as pdf_file:
                            pdf_file.write(pdf_data)
                        pdfs_found.append(output_path)
        
        except Exception as e:
            raise Exception(f"Error processing {os.path.basename(eml_file_path)}: {str(e)}")
        
        return pdfs_found
    
    def extract_pdfs_from_msg(self, msg_file_path, output_folder):
        """Extract PDFs from .msg Outlook files"""
        pdfs_found = []
        
        try:
            # Try extract-msg library first
            try:
                import extract_msg
                msg = extract_msg.Message(msg_file_path)
                
                for attachment in msg.attachments:
                    filename = None
                    if hasattr(attachment, 'longFilename') and attachment.longFilename:
                        filename = attachment.longFilename
                    elif hasattr(attachment, 'shortFilename') and attachment.shortFilename:
                        filename = attachment.shortFilename
                    
                    if filename and filename.lower().endswith('.pdf'):
                        final_filename = self.get_unique_filename(output_folder, filename)
                        output_path = os.path.join(output_folder, final_filename)
                        with open(output_path, 'wb') as pdf_file:
                            pdf_file.write(attachment.data)
                        pdfs_found.append(output_path)
                
                msg.close()
            
            except ImportError:
                # Alternative method without extract-msg
                with open(msg_file_path, 'rb') as f:
                    content = f.read()
                
                pdf_start = b'%PDF-'
                pdf_end = b'%%EOF'
                start_pos = 0
                pdf_count = 0
                
                while True:
                    pdf_start_pos = content.find(pdf_start, start_pos)
                    if pdf_start_pos == -1:
                        break
                    
                    pdf_end_pos = content.find(pdf_end, pdf_start_pos)
                    if pdf_end_pos == -1:
                        break
                    
                    pdf_end_pos += len(pdf_end)
                    pdf_data = content[pdf_start_pos:pdf_end_pos]
                    
                    pdf_count += 1
                    filename = f"extracted_pdf_{pdf_count}.pdf"
                    final_filename = self.get_unique_filename(output_folder, filename)
                    
                    output_path = os.path.join(output_folder, final_filename)
                    with open(output_path, 'wb') as pdf_file:
                        pdf_file.write(pdf_data)
                    pdfs_found.append(output_path)
                    
                    start_pos = pdf_end_pos
        
        except Exception as e:
            raise Exception(f"Error processing {os.path.basename(msg_file_path)}: {str(e)}")
        
        return pdfs_found
    
    def get_unique_filename(self, output_folder, filename):
        """Generate unique filename if file already exists"""
        base_name = Path(filename).stem
        extension = Path(filename).suffix
        counter = 1
        final_filename = filename
        
        while os.path.exists(os.path.join(output_folder, final_filename)):
            final_filename = f"{base_name}_{counter}{extension}"
            counter += 1
        
        return final_filename
    
    def extract_pdfs_from_emails(self, input_folder, recursive=True, update_callback=None):
        """Extract all PDFs from email files in the input folder"""
        # Create temporary folder for extracted PDFs
        self.temp_pdf_folder = tempfile.mkdtemp(prefix="extracted_pdfs_")
        
        extracted_pdfs = []
        processed_emails = 0
        
        # Collect all email files
        email_files = []
        if recursive:
            for root, dirs, files in os.walk(input_folder):
                for file in files:
                    if file.lower().endswith(('.eml', '.msg')):
                        email_files.append(os.path.join(root, file))
        else:
            for file in os.listdir(input_folder):
                if file.lower().endswith(('.eml', '.msg')):
                    email_files.append(os.path.join(input_folder, file))
        
        total_files = len(email_files)
        if update_callback:
            update_callback(f"Found {total_files} email files to process")
        
        # Process each email file
        for i, file_path in enumerate(email_files, 1):
            filename = os.path.basename(file_path)
            file_extension = Path(file_path).suffix.lower()
            
            if update_callback:
                update_callback(f"Processing email {i}/{total_files}: {filename}")
            
            pdfs_found = []
            
            try:
                if file_extension == '.eml':
                    pdfs_found = self.extract_pdfs_from_eml(file_path, self.temp_pdf_folder)
                elif file_extension == '.msg':
                    pdfs_found = self.extract_pdfs_from_msg(file_path, self.temp_pdf_folder)
                
                if pdfs_found:
                    extracted_pdfs.extend(pdfs_found)
                    if update_callback:
                        update_callback(f"  → Extracted {len(pdfs_found)} PDF(s)")
                
                processed_emails += 1
                
            except Exception as e:
                if update_callback:
                    update_callback(f"  → Error: {str(e)}")
        
        if update_callback:
            update_callback(f"Extraction complete: {len(extracted_pdfs)} PDFs from {processed_emails} emails")
        
        return extracted_pdfs
    
    def cleanup_temp_folder(self):
        """Clean up temporary PDF folder"""
        if self.temp_pdf_folder and os.path.exists(self.temp_pdf_folder):
            try:
                shutil.rmtree(self.temp_pdf_folder)
            except:
                pass

class PDFDataExtractor:
    @staticmethod
    def extract_store_sections(pdf_path):
        """Extract store transaction sections from the PDF"""
        store_sections = {}
        current_store = None
        store_content = []
        capture_mode = False
        
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                text = page.extract_text()
                lines = text.split('\n')
                
                for line in lines:
                    # Look for store section headers
                    store_header_match = re.search(r'(.*?)\s+t/a\s+-\s+(.*?)$', line)
                    if store_header_match and "REMITTANCE" in text[:text.find(line)]:
                        if current_store and store_content:
                            store_sections[current_store] = '\n'.join(store_content)
                        
                        current_store = store_header_match.group(2).strip()
                        store_content = []
                        capture_mode = False
                    
                    # Start capturing after we see the "Date Amount Total" line
                    if current_store and "Date Amount Total" in line:
                        capture_mode = True
                        continue
                    
                    # Stop capturing when we reach reconciling items section
                    if current_store and "TOTAL AS PER STATEMENT" in line:
                        capture_mode = False
                    
                    # Capture transaction lines when in capture mode
                    if current_store and capture_mode:
                        store_content.append(line)
        
        # Add the last store
        if current_store and store_content:
            store_sections[current_store] = '\n'.join(store_content)
        
        return store_sections
    
    @staticmethod
    def parse_store_transactions(store_name, store_content):
        """Parse transaction data from a store's content section"""
        transactions = []
        
        lines = store_content.split('\n')
        running_total = None
        
        for line in lines:
            if not line.strip() or "Prepared by" in line or "-" == line.strip():
                continue

            # Try different patterns to match various transaction formats
            pattern1_match = re.match(r'(\d{1,2}-\w{3}-\d{4})\s+(.*?)\s+([+-]?[\d,]+\.\d{2})\s+(\.\d{2}|\d*\.\d{2})$', line)
            pattern2_match = re.match(r'(\d{1,2}-\w{3}-\d{4})\s+(.*?)\s+([+-]?[\d,]+\.\d{2})\s+([+-]?[\d,]+\.\d{2})$', line)
            pattern3_match = re.match(r'(\d{1,2}-\w{3}-\d{4})\s+(.*?)\s+([+-]?[\d,]+\.\d{2})\s+(.*?)$', line)
            
            match = pattern1_match or pattern2_match or pattern3_match
            
            if match:
                date_str = match.group(1)
                description = match.group(2).strip()
                amount = match.group(3).replace(',', '')
                
                total_str = match.group(4).strip() if match.group(4) else None
                
                if total_str:
                    if total_str.startswith('.'):
                        total_str = '0' + total_str
                    total = total_str.replace(',', '')
                else:
                    total = None
                
                # Determine transaction type
                if "Revesal" in description or "Reversal" in description:
                    transaction_type = "Reversal"
                elif float(amount) < 0:
                    transaction_type = "Credit"
                else:
                    transaction_type = "Invoice"
                
                if "TRG - CA SALES - MONDELEZ" in description:
                    transaction_type = "Rebate"
                
                if (total is None or total == '0.00' or total == '0') and running_total is not None:
                    total = str(running_total + float(amount))
                
                if total and (total != '0.00' and total != '0'):
                    try:
                        running_total = float(total)
                    except ValueError:
                        numbers = re.findall(r'[+-]?[\d,]+\.\d+', total)
                        if numbers:
                            running_total = float(numbers[0].replace(',', ''))
                        else:
                            if running_total is not None:
                                running_total = running_total + float(amount)
                            else:
                                running_total = float(amount)
                
                transactions.append({
                    'Store': store_name,
                    'Date': date_str,
                    'Description': description,
                    'Amount': float(amount),
                    'Running_Total': running_total if running_total is not None else float(amount),
                    'Transaction_Type': transaction_type
                })
        
        return transactions
    
    @staticmethod
    def process_single_pdf(pdf_path):
        """Process a single PDF and return the transactions data"""
        try:
            store_sections = PDFDataExtractor.extract_store_sections(pdf_path)
            
            if not store_sections:
                return None, f"No store sections found in {os.path.basename(pdf_path)}"
                
            all_transactions = []
            
            for store_name, store_content in store_sections.items():
                store_transactions = PDFDataExtractor.parse_store_transactions(store_name, store_content)
                all_transactions.extend(store_transactions)
            
            if all_transactions:
                df = pd.DataFrame(all_transactions)
                
                column_order = ['Store', 'Date', 'Transaction_Type', 'Description', 'Amount']
                df = df[column_order]
                
                df['Date'] = pd.to_datetime(df['Date'], format='%d-%b-%Y', errors='coerce')
                df = df.sort_values(['Store', 'Date'])
                df['Date'] = df['Date'].dt.strftime('%d-%b-%Y')
                
                return df, None
            else:
                return None, f"No transactions found in {os.path.basename(pdf_path)}"
                
        except Exception as e:
            return None, f"Error processing {os.path.basename(pdf_path)}: {str(e)}"

class IntegratedApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Configure the window
        self.title("Email PDF to Excel Processor")
        self.geometry("800x700")
        self.minsize(750, 650)
        
        # Configure grid layout
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1)
        
        # Initialize processors
        self.email_processor = EmailPDFProcessor()
        
        # Create main frame
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)
        
        # Add title
        self.title_label = ctk.CTkLabel(
            self.main_frame, 
            text="Email PDF to Excel Processor",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        self.title_label.grid(row=0, column=0, padx=20, pady=(20, 30))
        
        # Create content frame
        self.content_frame = ctk.CTkFrame(self.main_frame)
        self.content_frame.grid(row=1, column=0, padx=20, pady=0, sticky="ew")
        self.content_frame.grid_columnconfigure(1, weight=1)
        
        # Input folder selection
        self.input_label = ctk.CTkLabel(
            self.content_frame, 
            text="Email Folder:",
            font=ctk.CTkFont(size=14)
        )
        self.input_label.grid(row=0, column=0, padx=(20, 10), pady=15, sticky="w")
        
        self.input_entry = ctk.CTkEntry(
            self.content_frame,
            height=32,
            font=ctk.CTkFont(size=13)
        )
        self.input_entry.grid(row=0, column=1, padx=10, pady=15, sticky="ew")
        
        self.input_button = ctk.CTkButton(
            self.content_frame, 
            text="Browse",
            width=80,
            height=32,
            command=self.browse_input_folder
        )
        self.input_button.grid(row=0, column=2, padx=(10, 20), pady=15)
        
        # Output file selection
        self.output_label = ctk.CTkLabel(
            self.content_frame, 
            text="Output Excel File:",
            font=ctk.CTkFont(size=14)
        )
        self.output_label.grid(row=1, column=0, padx=(20, 10), pady=15, sticky="w")
        
        self.output_entry = ctk.CTkEntry(
            self.content_frame,
            height=32,
            font=ctk.CTkFont(size=13)
        )
        self.output_entry.grid(row=1, column=1, padx=10, pady=15, sticky="ew")
        
        self.output_button = ctk.CTkButton(
            self.content_frame, 
            text="Browse",
            width=80,
            height=32,
            command=self.browse_output_file
        )
        self.output_button.grid(row=1, column=2, padx=(10, 20), pady=15)
        
        # Options frame
        self.options_frame = ctk.CTkFrame(self.main_frame)
        self.options_frame.grid(row=2, column=0, padx=20, pady=20, sticky="ew")
        
        self.options_title = ctk.CTkLabel(
            self.options_frame,
            text="Options",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.options_title.grid(row=0, column=0, padx=20, pady=(15, 10), sticky="w")
        
        self.recursive_var = ctk.BooleanVar(value=True)
        self.recursive_checkbox = ctk.CTkCheckBox(
            self.options_frame,
            text="Search subfolders recursively",
            variable=self.recursive_var
        )
        self.recursive_checkbox.grid(row=1, column=0, padx=20, pady=5, sticky="w")
        
        # Process button
        self.process_button = ctk.CTkButton(
            self.main_frame, 
            text="Process Emails & Generate Excel",
            height=40,
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self.process_files
        )
        self.process_button.grid(row=3, column=0, padx=20, pady=30)
        
        # Progress frame
        self.progress_frame = ctk.CTkFrame(self.main_frame)
        self.progress_frame.grid(row=4, column=0, padx=20, pady=(0, 20), sticky="ew")
        self.progress_frame.grid_columnconfigure(0, weight=1)
        
        # Progress bar
        self.progress_bar = ctk.CTkProgressBar(self.progress_frame)
        self.progress_bar.grid(row=0, column=0, padx=20, pady=10, sticky="ew")
        self.progress_bar.set(0)
        
        # Status text
        self.status_label = ctk.CTkLabel(
            self.progress_frame, 
            text="Ready to process emails",
            font=ctk.CTkFont(size=13)
        )
        self.status_label.grid(row=1, column=0, padx=20, pady=(0, 10), sticky="w")
        
        # Log frame
        self.log_frame = ctk.CTkScrollableFrame(self.main_frame, height=200)
        self.log_frame.grid(row=5, column=0, padx=20, pady=(0, 20), sticky="ew")
        self.log_frame.grid_columnconfigure(0, weight=1)
        
        self.log_title = ctk.CTkLabel(
            self.log_frame,
            text="Processing Log",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.log_title.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="w")
        
        # Footer
        self.appearance_mode_option = ctk.CTkOptionMenu(
            self, 
            values=["System", "Light", "Dark"],
            command=self.change_appearance_mode
        )
        self.appearance_mode_option.grid(row=6, column=0, padx=20, pady=(0, 10), sticky="w")
        self.appearance_mode_option.set("System")
        
        # Initialize variables
        self.process_thread = None
        self._is_running = False
        self.log_row_counter = 1
    
    def change_appearance_mode(self, new_appearance_mode):
        ctk.set_appearance_mode(new_appearance_mode)
    
    def browse_input_folder(self):
        folder = ctk.filedialog.askdirectory(title="Select folder containing email files")
        if folder:
            self.input_entry.delete(0, "end")
            self.input_entry.insert(0, folder)
            
            # Set default output filename if not set
            if not self.output_entry.get():
                default_output = os.path.join(folder, "extracted_data.xlsx")
                self.output_entry.delete(0, "end")
                self.output_entry.insert(0, default_output)
    
    def browse_output_file(self):
        filename = ctk.filedialog.asksaveasfilename(
            title="Save Excel file",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            defaultextension=".xlsx"
        )
        if filename:
            self.output_entry.delete(0, "end")
            self.output_entry.insert(0, filename)
    
    def add_log_message(self, message):
        """Add message to log"""
        if not self._is_running:
            return
            
        log_label = ctk.CTkLabel(
            self.log_frame,
            text=f"• {message}",
            font=ctk.CTkFont(size=11),
            anchor="w"
        )
        log_label.grid(row=self.log_row_counter, column=0, padx=20, pady=2, sticky="w")
        self.log_row_counter += 1
        
        # Auto-scroll to bottom
        self.log_frame._parent_canvas.yview_moveto(1.0)
        self.update_idletasks()
    
    def clear_log(self):
        """Clear log messages"""
        for widget in self.log_frame.winfo_children():
            if widget != self.log_title:
                widget.destroy()
        self.log_row_counter = 1
    
    def update_status(self, status_text):
        if not self._is_running:
            return
        self.status_label.configure(text=status_text)
        self.add_log_message(status_text)
        self.update_idletasks()
    
    def update_progress(self, progress_value):
        if not self._is_running:
            return
        self.progress_bar.set(progress_value / 100)
        self.update_idletasks()
    
    def process_complete(self, success, message=""):
        self.process_button.configure(state="normal")
        self._is_running = False
        
        if success:
            self.show_success_dialog(message)
        else:
            self.show_error_dialog(message)
        
        # Cleanup
        self.email_processor.cleanup_temp_folder()
    
    def show_success_dialog(self, message):
        output_path = self.output_entry.get()
        dialog = ctk.CTkToplevel(self)
        dialog.title("Success")
        dialog.geometry("500x250")
        dialog.resizable(False, False)
        dialog.grab_set()
        
        dialog.geometry(f"+{self.winfo_x() + self.winfo_width()//2 - 250}+{self.winfo_y() + self.winfo_height()//2 - 125}")
        
        success_text = f"Processing completed successfully!\n\n{message}\n\nOutput saved to:\n{output_path}"
        
        label = ctk.CTkLabel(
            dialog, 
            text=success_text,
            font=ctk.CTkFont(size=14),
            wraplength=460
        )
        label.pack(padx=20, pady=(20, 10))
        
        buttons_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        buttons_frame.pack(padx=20, pady=(10, 20), fill="x")
        
        open_button = ctk.CTkButton(
            buttons_frame, 
            text="Open File",
            command=lambda: self.open_file(output_path, dialog)
        )
        open_button.pack(side="left", padx=(80, 10), pady=10)
        
        close_button = ctk.CTkButton(
            buttons_frame, 
            text="Close",
            command=dialog.destroy
        )
        close_button.pack(side="right", padx=(10, 80), pady=10)
    
    def show_error_dialog(self, message):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Error")
        dialog.geometry("450x200")
        dialog.resizable(False, False)
        dialog.grab_set()
        
        dialog.geometry(f"+{self.winfo_x() + self.winfo_width()//2 - 225}+{self.winfo_y() + self.winfo_height()//2 - 100}")
        
        error_text = f"Processing failed:\n\n{message}\n\nCheck the processing log for details."
        
        label = ctk.CTkLabel(
            dialog, 
            text=error_text,
            font=ctk.CTkFont(size=14),
            wraplength=420
        )
        label.pack(padx=20, pady=(20, 10))
        
        close_button = ctk.CTkButton(
            dialog, 
            text="Close",
            command=dialog.destroy
        )
        close_button.pack(padx=20, pady=(10, 20))
    
    def open_file(self, file_path, dialog=None):
        if dialog:
            dialog.destroy()
            
        try:
            if sys.platform == 'win32':
                os.startfile(file_path)
            elif sys.platform == 'darwin':
                import subprocess
                subprocess.call(['open', file_path])
            else:
                import subprocess
                subprocess.call(['xdg-open', file_path])
        except Exception as e:
            self.show_error_dialog(f"Could not open the file: {str(e)}")
    
    def process_files(self):
        input_folder = self.input_entry.get()
        output_path = self.output_entry.get()
        
        # Validate inputs
        if not input_folder:
            self.show_error_dialog("Please select an input folder containing email files.")
            return
        
        if not output_path:
            self.show_error_dialog("Please specify an output Excel file.")
            return
        
        if not os.path.exists(input_folder):
            self.show_error_dialog("Input folder does not exist.")
            return
        
        # Clear log and reset progress
        self.clear_log()
        self.process_button.configure(state="disabled")
        self.progress_bar.set(0)
        self.status_label.configure(text="Starting process...")
        self._is_running = True
        
        # Start processing in a separate thread
        self.process_thread = threading.Thread(
            target=self.run_processing,
            args=(input_folder, output_path, self.recursive_var.get())
        )
        self.process_thread.daemon = True
        self.process_thread.start()
    
    def run_processing(self, input_folder, output_path, recursive):
        try:
            # Phase 1: Extract PDFs from emails (40% of progress)
            self.update_status("Phase 1: Extracting PDFs from emails...")
            self.update_progress(10)
            
            extracted_pdfs = self.email_processor.extract_pdfs_from_emails(
                input_folder, recursive, self.update_status
            )
            
            if not extracted_pdfs:
                self.after(0, lambda: self.process_complete(False, "No PDF files found in the email attachments."))
                return
            
            self.update_progress(40)
            self.update_status(f"Phase 1 complete: {len(extracted_pdfs)} PDFs extracted")
            
            # Phase 2: Process PDFs to Excel (60% of progress)
            self.update_status("Phase 2: Processing PDFs to generate Excel...")
            self.update_progress(50)
            
            success = self.process_pdfs_to_excel(extracted_pdfs, output_path)
            
            if success:
                total_transactions = getattr(self, '_total_transactions', 0)
                success_message = f"Successfully processed {len(extracted_pdfs)} PDF files with {total_transactions} total transactions."
                self.after(0, lambda: self.process_complete(True, success_message))
            else:
                self.after(0, lambda: self.process_complete(False, "Failed to process PDFs to Excel."))
                
        except Exception as e:
            error_msg = f"An error occurred during processing: {str(e)}"
            self.after(0, lambda: self.process_complete(False, error_msg))
    
    def process_pdfs_to_excel(self, pdf_paths, output_path):
        """Process extracted PDFs and create Excel file"""
        try:
            total_pdfs = len(pdf_paths)
            processed_count = 0
            
            with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
                successful_files = []
                failed_files = []
                
                for i, pdf_path in enumerate(pdf_paths):
                    filename = os.path.basename(pdf_path)
                    self.update_status(f"Processing PDF {i+1}/{total_pdfs}: {filename}")
                    
                    df, error_msg = PDFDataExtractor.process_single_pdf(pdf_path)
                    
                    if df is not None:
                        # Create worksheet name
                        worksheet_name = os.path.splitext(filename)[0]
                        worksheet_name = worksheet_name[:31]
                        
                        # Remove invalid characters
                        invalid_chars = ['\\', '/', '*', '[', ']', ':', '?']
                        for char in invalid_chars:
                            worksheet_name = worksheet_name.replace(char, '_')
                        
                        # Ensure unique worksheet name
                        original_name = worksheet_name
                        counter = 1
                        while worksheet_name in [sheet[0] for sheet in successful_files]:
                            worksheet_name = f"{original_name}_{counter}"
                            if len(worksheet_name) > 31:
                                worksheet_name = f"{original_name[:27]}_{counter}"
                            counter += 1
                        
                        df.to_excel(writer, sheet_name=worksheet_name, index=False)
                        successful_files.append((worksheet_name, len(df), filename))
                        
                    else:
                        failed_files.append((filename, error_msg))
                    
                    processed_count += 1
                    progress = 50 + (processed_count / total_pdfs) * 40
                    self.update_progress(progress)
                
                # Add summary worksheet
                if successful_files or failed_files:
                    summary_data = []
                    
                    for worksheet_name, transaction_count, filename in successful_files:
                        summary_data.append({
                            'Filename': filename,
                            'Status': 'Success',
                            'Worksheet': worksheet_name,
                            'Transactions': transaction_count,
                            'Notes': f'{transaction_count} transactions extracted'
                        })
                    
                    for filename, error_msg in failed_files:
                        summary_data.append({
                            'Filename': filename,
                            'Status': 'Failed',
                            'Worksheet': 'N/A',
                            'Transactions': 0,
                            'Notes': error_msg
                        })
                    
                    if summary_data:
                        summary_df = pd.DataFrame(summary_data)
                        summary_df.to_excel(writer, sheet_name='Summary', index=False)
            
            self.update_progress(100)
            
            if successful_files:
                self._total_transactions = sum(count for _, count, _ in successful_files)
                self.update_status(f"Excel file created successfully with {len(successful_files)} worksheets!")
                return True
            else:
                self.update_status("No PDF files were successfully processed.")
                return False
                
        except Exception as e:
            self.update_status(f"Error creating Excel file: {str(e)}")
            return False

if __name__ == "__main__":
    app = IntegratedApp()
    app.mainloop()
