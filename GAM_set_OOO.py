import tkinter as tk
from tkinter import filedialog, ttk
import csv
import subprocess
import os
import pandas as pd
from datetime import datetime

class OOOSetterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Set Out of Office Messages")
        
        # Configure main window - remove fixed geometry
        self.root.configure(padx=20, pady=20)
        
        # Variables
        self.csv_path = tk.StringVar()
        self.gam_path = tk.StringVar()  # We'll set this after widgets are created
        self.company_name = tk.StringVar()
        self.contact_email = tk.StringVar()
        self.ooo_subject = tk.StringVar(value="OOO: No longer with {company_name}")
        self.ooo_message = tk.StringVar(value="{name} is no longer with {company_name}.\nFor any enquiries or requests, please reach out to {contact_email}.")
        
        # Add status log variable
        self.status_log = []
        
        # Add preview variables
        self.preview_name = ""
        self.preview_email = ""
        
        # Create GUI elements first
        self.create_widgets()
        
        # Now that widgets exist, try to detect GAM path
        self.gam_path.set(self.detect_gam_path())
        
        # Update window size after widgets are created
        self.root.update_idletasks()
        self.root.pack_propagate(True)
        
    def create_widgets(self):
        # CSV Format Instructions
        instruction_frame = ttk.LabelFrame(self.root, text="CSV File Format Instructions", padding=10)
        instruction_frame.pack(fill="x", pady=(0, 10))
        
        instructions = (
            "CSV file must contain the following columns:\n"
            "- email: The user's email address\n"
            "- name: The user's full name\n\n"
            "Example CSV format:\n"
            "email,name\n"
            "user@company.com,John Smith"
        )
        
        instruction_label = ttk.Label(instruction_frame, text=instructions, justify="left", wraplength=550)
        instruction_label.pack(fill="x", padx=5)

        # CSV File Selection
        csv_frame = ttk.LabelFrame(self.root, text="CSV File Selection", padding=10)
        csv_frame.pack(fill="x", pady=(0, 10))
        
        ttk.Button(csv_frame, text="Select CSV File", command=self.browse_csv).pack(side="left", padx=5)
        ttk.Entry(csv_frame, textvariable=self.csv_path, width=50, state='readonly').pack(side="left", padx=5)
        
        # GAM Directory Selection
        gam_frame = ttk.LabelFrame(self.root, text="GAM Configuration Directory", padding=10)
        gam_frame.pack(fill="x", pady=(0, 10))
        
        ttk.Button(gam_frame, text="Select GAM Directory", command=self.browse_gam).pack(side="left", padx=5)
        ttk.Entry(gam_frame, textvariable=self.gam_path, width=50, state='readonly').pack(side="left", padx=5)
        
        # Company Settings
        settings_frame = ttk.LabelFrame(self.root, text="Company Settings", padding=10)
        settings_frame.pack(fill="x", pady=(0, 10))
        
        ttk.Label(settings_frame, text="Company Name:").grid(row=0, column=0, sticky="w", padx=5)
        ttk.Entry(settings_frame, textvariable=self.company_name).grid(row=0, column=1, sticky="ew", padx=5)
        
        ttk.Label(settings_frame, text="Contact Email:").grid(row=1, column=0, sticky="w", padx=5)
        ttk.Entry(settings_frame, textvariable=self.contact_email).grid(row=1, column=1, sticky="ew", padx=5)
        
        # OOO Message Settings
        message_frame = ttk.LabelFrame(self.root, text="Out of Office Settings", padding=10)
        message_frame.pack(fill="x", pady=(0, 10))
        
        ttk.Label(message_frame, text="Subject:").grid(row=0, column=0, sticky="w", padx=5)
        ttk.Entry(message_frame, textvariable=self.ooo_subject).grid(row=0, column=1, sticky="ew", padx=5)
        
        ttk.Label(message_frame, text="Message:").grid(row=1, column=0, sticky="w", padx=5)
        self.message_text = tk.Text(message_frame, height=5, width=50)
        self.message_text.insert('1.0', self.ooo_message.get())
        self.message_text.grid(row=1, column=1, sticky="ew", padx=5)
        
        # Bind the Text widget changes to update preview
        self.message_text.bind('<KeyRelease>', self.on_message_change)
        
        settings_frame.grid_columnconfigure(1, weight=1)
        message_frame.grid_columnconfigure(1, weight=1)
        
        # Add Preview Section
        preview_frame = ttk.LabelFrame(self.root, text="Message Preview", padding=10)
        preview_frame.pack(fill="x", pady=(0, 10))
        
        self.preview_text = tk.Text(preview_frame, height=8, width=50, wrap=tk.WORD)
        self.preview_text.pack(fill="x", padx=5, pady=5)
        self.preview_text.config(state='disabled')
        
        # Bind update events to all relevant fields
        self.company_name.trace_add('write', self.update_preview)
        self.contact_email.trace_add('write', self.update_preview)
        self.ooo_subject.trace_add('write', self.update_preview)
        self.ooo_message.trace_add('write', self.update_preview)
        
        # Move the existing CSV preview below the message preview
        csv_preview_frame = ttk.LabelFrame(self.root, text="CSV Preview", padding=10)
        csv_preview_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        # Create Treeview for CSV preview
        self.tree = ttk.Treeview(csv_preview_frame, columns=("Email", "Name"), show="headings")
        self.tree.heading("Email", text="Email Address")
        self.tree.heading("Name", text="Name")
        self.tree.pack(fill="both", expand=True)
        
        # Scrollbar for Treeview
        scrollbar = ttk.Scrollbar(csv_preview_frame, orient="vertical", command=self.tree.yview)
        scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # Add Status Log
        log_frame = ttk.LabelFrame(self.root, text="Status Log", padding=10)
        log_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        self.log_text = tk.Text(log_frame, height=10, width=50, wrap=tk.WORD)
        self.log_text.pack(fill="both", expand=True, side="left")
        
        # Add scrollbar for log
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.config(state='disabled')
        
        # Action Buttons
        button_frame = ttk.Frame(self.root)
        button_frame.pack(fill="x")
        
        ttk.Button(button_frame, text="Set OOO Messages", command=self.set_ooo_messages).pack(side="right")
        
    def browse_csv(self):
        filename = filedialog.askopenfilename(
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if filename:
            self.csv_path.set(filename)
            self.load_csv_preview()
    
    def load_csv_preview(self):
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        try:
            with open(self.csv_path.get(), 'r') as file:
                csv_reader = csv.reader(file)
                next(csv_reader)  # Skip header row
                first_row = next(csv_reader)  # Get first data row
                if len(first_row) >= 2:
                    self.preview_email = first_row[0]
                    self.preview_name = first_row[1]
                    self.update_preview()
                
                # Reset file pointer and skip header
                file.seek(0)
                next(csv_reader)
                # Load rest of CSV into preview
                for row in csv_reader:
                    if len(row) >= 2:
                        self.tree.insert("", "end", values=(row[0], row[1]))
        except Exception as e:
            tk.messagebox.showerror("Error", f"Failed to load CSV file: {str(e)}")
    
    def update_preview(self, *args):
        if not self.preview_name:  # If no CSV loaded yet
            return
            
        try:
            formatted_subject = self.ooo_subject.get().format(
                company_name=self.company_name.get() or "[Company Name]"
            )
            
            # Format contact email as HTML link for preview
            contact_email_html = f'<a href="mailto:{self.contact_email.get() or "[Contact Email]"}">{self.contact_email.get() or "[Contact Email]"}</a>'
            
            formatted_message = self.ooo_message.get().format(
                name=self.preview_name,
                company_name=self.company_name.get() or "[Company Name]",
                contact_email=contact_email_html
            )
            
            preview_text = f"Subject: {formatted_subject}\n\n{formatted_message}"
            
            self.preview_text.config(state='normal')
            self.preview_text.delete('1.0', tk.END)
            self.preview_text.insert('1.0', preview_text)
            self.preview_text.config(state='disabled')
        except Exception as e:
            self.preview_text.config(state='normal')
            self.preview_text.delete('1.0', tk.END)
            self.preview_text.insert('1.0', f"Preview Error: {str(e)}")
            self.preview_text.config(state='disabled')
    
    def browse_gam(self):
        directory = filedialog.askdirectory(
            title="Select GAM Configuration Directory"
        )
        if directory:
            self.gam_path.set(directory)
    
    def log_status(self, message, error=False):
        self.log_text.config(state='normal')
        timestamp = datetime.now().strftime("%H:%M:%S")
        prefix = "ERROR: " if error else "INFO: "
        self.log_text.insert('end', f"[{timestamp}] {prefix}{message}\n")
        self.log_text.see('end')  # Auto-scroll to bottom
        self.log_text.config(state='disabled')
    
    def on_message_change(self, event=None):
        # Update the StringVar with the current Text widget content
        self.ooo_message.set(self.message_text.get('1.0', 'end-1c'))
        self.update_preview()
    
    def detect_gam_path(self):
        try:
            # Check common installation paths
            possible_paths = [
                "/home/im-admin/bin/gamadv-xtd3/gam",  # Your known path
                os.path.expanduser("~/bin/gamadv-xtd3/gam"),  # Expanded home path
                "/usr/local/bin/gam",
                "/usr/bin/gam",
                "/opt/gam/gam",
            ]
            
            self.log_status("Checking known GAM installation paths...")
            
            for path in possible_paths:
                self.log_status(f"Checking path: {path}")
                if os.path.exists(path) and os.access(path, os.X_OK):
                    self.log_status(f"Found GAM at: {path}")
                    return os.path.dirname(path)
            
            # If we get here, GAM wasn't found in any of the common locations
            self.log_status("Could not find GAM in common installation paths.", error=True)
            self.log_status("Please select GAM directory manually.", error=True)
            
        except Exception as e:
            self.log_status(f"Error detecting GAM path: {str(e)}", error=True)
        
        return ""
    
    def set_ooo_messages(self):
        if not self.csv_path.get():
            tk.messagebox.showerror("Error", "Please select a CSV file first")
            return
        
        if not self.gam_path.get():
            tk.messagebox.showerror("Error", "GAM path not found. Please select GAM directory manually.")
            return
            
        try:
            df = pd.read_csv(self.csv_path.get())
            success_count = 0
            error_count = 0
            
            self.log_status(f"Starting to process {len(df)} users...")
            
            for index, row in df.iterrows():
                email_address = row['email'].strip()
                name = row['name'].strip()
                
                contact_email_html = f'<a href="mailto:{self.contact_email.get()}">{self.contact_email.get()}</a>'
                
                formatted_subject = self.ooo_subject.get().format(
                    company_name=self.company_name.get()
                )
                
                formatted_message = self.message_text.get('1.0', 'end-1c').format(
                    name=name,
                    company_name=self.company_name.get(),
                    contact_email=contact_email_html
                )
                
                # Use the detected GAM path
                gam_executable = os.path.join(self.gam_path.get(), "gam")
                
                cmd = [
                    gam_executable,
                    "user", email_address,
                    "vacation", "on",
                    "subject", formatted_subject,
                    "message", formatted_message,
                    "html"
                ]
                
                try:
                    result = subprocess.run(cmd, check=True, capture_output=True, text=True)
                    success_count += 1
                    self.log_status(f"Successfully set OOO for {email_address}")
                except subprocess.CalledProcessError as e:
                    error_count += 1
                    error_msg = f"Failed to set OOO for {email_address}: {e.stderr.strip()}"
                    self.log_status(error_msg, error=True)
            
            summary = f"Completed: {success_count} successful, {error_count} failed"
            self.log_status(summary)
            
            if error_count == 0:
                tk.messagebox.showinfo("Success", f"Successfully set OOO messages for all {success_count} users!")
            else:
                tk.messagebox.showwarning("Completed with Errors", 
                    f"Process completed with {error_count} errors.\n"
                    f"Successfully processed: {success_count}\n"
                    "Check the status log for details.")
                
        except Exception as e:
            self.log_status(f"Critical error: {str(e)}", error=True)
            tk.messagebox.showerror("Error", f"Failed to process CSV file: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = OOOSetterApp(root)
    root.mainloop()