import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import csv
import smtplib
import threading
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import codecs

class SimpleEmailSender:
    def __init__(self, root):
        self.root = root
        self.root.title("Simple Email Sender")
        self.root.geometry("700x600")
        self.root.resizable(True, True)
        
        self.csv_file_path = None
        self.template_content = None
        self.emails_to_send = []
        self.is_sending = False
        
        # Create main frame
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # SMTP Settings Frame
        smtp_frame = ttk.LabelFrame(main_frame, text="SMTP Settings", padding="10")
        smtp_frame.pack(fill=tk.X, pady=5)
        
        # SMTP Server
        ttk.Label(smtp_frame, text="SMTP Server:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.smtp_server = ttk.Entry(smtp_frame, width=30)
        self.smtp_server.insert(0, "localhost")  # Default to localhost for HMailServer
        self.smtp_server.grid(row=0, column=1, sticky=tk.W, pady=2)
        
        # SMTP Port
        ttk.Label(smtp_frame, text="SMTP Port:").grid(row=0, column=2, sticky=tk.W, pady=2, padx=(10, 0))
        self.smtp_port = ttk.Entry(smtp_frame, width=10)
        self.smtp_port.insert(0, "25")  # Default port for non-TLS SMTP
        self.smtp_port.grid(row=0, column=3, sticky=tk.W, pady=2)
        
        # Connection type
        ttk.Label(smtp_frame, text="Connection:").grid(row=0, column=4, sticky=tk.W, pady=2, padx=(10, 0))
        self.connection_type = ttk.Combobox(smtp_frame, width=10)
        self.connection_type['values'] = ('Plain', 'STARTTLS', 'SSL/TLS')
        self.connection_type.current(0)  # Default to Plain for HMailServer
        self.connection_type.grid(row=0, column=5, sticky=tk.W, pady=2)
        
        # Sender Email
        ttk.Label(smtp_frame, text="Sender Email:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.sender_email = ttk.Entry(smtp_frame, width=40)
        self.sender_email.grid(row=1, column=1, columnspan=3, sticky=tk.W, pady=2)
        
        # Password
        ttk.Label(smtp_frame, text="Password:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.sender_password = ttk.Entry(smtp_frame, width=40, show="*")
        self.sender_password.grid(row=2, column=1, columnspan=3, sticky=tk.W, pady=2)
        
        # Authentication checkbox
        self.use_auth = tk.BooleanVar(value=True)
        self.auth_checkbox = ttk.Checkbutton(smtp_frame, text="Use Authentication", variable=self.use_auth)
        self.auth_checkbox.grid(row=2, column=4, columnspan=2, sticky=tk.W, pady=2)
        
        # Email Content Frame
        content_frame = ttk.LabelFrame(main_frame, text="Email Content", padding="10")
        content_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Subject
        ttk.Label(content_frame, text="Subject:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.email_subject = ttk.Entry(content_frame, width=60)
        self.email_subject.grid(row=0, column=1, columnspan=2, sticky=tk.W, pady=2)
        
        # Delay
        ttk.Label(content_frame, text="Delay (seconds):").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.delay = ttk.Entry(content_frame, width=10)
        self.delay.insert(0, "5")
        self.delay.grid(row=1, column=1, sticky=tk.W, pady=2)
        
        # Template
        ttk.Label(content_frame, text="Email Template:").grid(row=2, column=0, sticky=tk.W, pady=2)
        template_frame = ttk.Frame(content_frame)
        template_frame.grid(row=2, column=1, columnspan=2, sticky=tk.W, pady=2)
        
        self.template_path = ttk.Entry(template_frame, width=50)
        self.template_path.pack(side=tk.LEFT)
        
        ttk.Button(template_frame, text="Browse", command=self.load_template).pack(side=tk.LEFT, padx=5)
        
        # Template Preview
        ttk.Label(content_frame, text="Template Preview:").grid(row=3, column=0, sticky=tk.NW, pady=2)
        self.template_preview = scrolledtext.ScrolledText(content_frame, width=60, height=10)
        self.template_preview.grid(row=3, column=1, columnspan=2, sticky=tk.NSEW, pady=2)
        
        # CSV File
        ttk.Label(content_frame, text="Recipients CSV:").grid(row=4, column=0, sticky=tk.W, pady=2)
        csv_frame = ttk.Frame(content_frame)
        csv_frame.grid(row=4, column=1, columnspan=2, sticky=tk.W, pady=2)
        
        self.csv_path = ttk.Entry(csv_frame, width=50)
        self.csv_path.pack(side=tk.LEFT)
        
        ttk.Button(csv_frame, text="Browse", command=self.load_csv).pack(side=tk.LEFT, padx=5)
        
        # CSV Encoding
        ttk.Label(content_frame, text="CSV Encoding:").grid(row=5, column=0, sticky=tk.W, pady=2)
        self.csv_encoding = ttk.Combobox(content_frame, width=20)
        self.csv_encoding['values'] = ('utf-8', 'utf-8-sig', 'latin-1', 'iso-8859-1', 'cp1252')
        self.csv_encoding.current(0)
        self.csv_encoding.grid(row=5, column=1, sticky=tk.W, pady=2)
        ttk.Button(content_frame, text="Reload CSV", command=self.reload_csv).grid(row=5, column=2, sticky=tk.W, pady=2)
        
        # Email/Name Column Selection
        csv_columns_frame = ttk.Frame(content_frame)
        csv_columns_frame.grid(row=6, column=0, columnspan=3, sticky=tk.W, pady=2)
        
        ttk.Label(csv_columns_frame, text="Email Column:").pack(side=tk.LEFT, padx=(0, 5))
        self.email_column = ttk.Entry(csv_columns_frame, width=15)
        self.email_column.insert(0, "email")
        self.email_column.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Label(csv_columns_frame, text="Name Column:").pack(side=tk.LEFT, padx=(0, 5))
        self.name_column = ttk.Entry(csv_columns_frame, width=15)
        self.name_column.insert(0, "name")
        self.name_column.pack(side=tk.LEFT)
        
        # Preview Recipients
        ttk.Label(content_frame, text="Recipients Preview:").grid(row=7, column=0, sticky=tk.NW, pady=2)
        self.recipients_preview = scrolledtext.ScrolledText(content_frame, width=60, height=5)
        self.recipients_preview.grid(row=7, column=1, columnspan=2, sticky=tk.NSEW, pady=2)
        
        # View CSV Structure button
        ttk.Button(content_frame, text="View CSV Structure", command=self.view_csv_structure).grid(row=8, column=1, sticky=tk.W, pady=2)
        
        # Configure grid weights for resizing
        content_frame.columnconfigure(1, weight=1)
        content_frame.rowconfigure(3, weight=2)
        content_frame.rowconfigure(7, weight=1)
        
        # Buttons Frame
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X, pady=10)
        
        # Test Connection button
        ttk.Button(buttons_frame, text="Test Connection", command=self.test_connection).pack(side=tk.LEFT, padx=(0, 10))
        
        # Progress bar
        self.progress = ttk.Progressbar(buttons_frame, orient=tk.HORIZONTAL, length=400, mode='determinate')
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        # Send button
        self.send_button = ttk.Button(buttons_frame, text="Send Emails", command=self.start_sending)
        self.send_button.pack(side=tk.RIGHT)
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        self.status_bar = ttk.Label(root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def test_connection(self):
        """Test SMTP connection with current settings"""
        try:
            self.status_var.set("Testing connection...")
            self.root.update_idletasks()
            
            server = self.get_smtp_connection()
            
            # Try to say HELLO to the server
            server.ehlo()
            
            if self.use_auth.get():
                server.login(self.sender_email.get(), self.sender_password.get())
            
            server.quit()
            self.status_var.set("Connection successful!")
            messagebox.showinfo("Success", "SMTP connection test successful!")
        except Exception as e:
            self.status_var.set(f"Connection failed: {str(e)}")
            messagebox.showerror("Connection Failed", f"Error: {str(e)}")
    
    def get_smtp_connection(self):
        """Create and return SMTP connection based on selected type"""
        connection_type = self.connection_type.get()
        server_address = self.smtp_server.get()
        port = int(self.smtp_port.get())
        
        if connection_type == "SSL/TLS":
            server = smtplib.SMTP_SSL(server_address, port)
        else:
            server = smtplib.SMTP(server_address, port)
            
            if connection_type == "STARTTLS":
                server.starttls()
                
        return server
    
    def load_template(self):
        file_path = filedialog.askopenfilename(
            title="Select Email Template",
            filetypes=(("HTML files", "*.html"), ("Text files", "*.txt"), ("All files", "*.*"))
        )
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as file:
                    self.template_content = file.read()
                    self.template_path.delete(0, tk.END)
                    self.template_path.insert(0, file_path)
                    self.template_preview.delete(1.0, tk.END)
                    self.template_preview.insert(tk.END, self.template_content[:1000] + 
                                                ("..." if len(self.template_content) > 1000 else ""))
            except Exception as e:
                messagebox.showerror("Error", f"Could not load template: {str(e)}")
    
    def load_csv(self):
        file_path = filedialog.askopenfilename(
            title="Select Recipients CSV",
            filetypes=(("CSV files", "*.csv"), ("All files", "*.*"))
        )
        if file_path:
            self.csv_file_path = file_path
            self.csv_path.delete(0, tk.END)
            self.csv_path.insert(0, file_path)
            self.reload_csv()
    
    def reload_csv(self):
        if not self.csv_file_path:
            messagebox.showerror("Error", "No CSV file selected")
            return
            
        try:
            # Try to read the CSV with the selected encoding
            recipients = self.read_email_list()
            
            email_col = self.email_column.get()
            name_col = self.name_column.get()
            
            preview_text = f"Total recipients: {len(recipients)}\n\nSample entries:\n"
            for i, recipient in enumerate(recipients[:5]):
                # Try to get email and name, show if there are missing fields
                email = recipient.get(email_col, "N/A")
                name = recipient.get(name_col, "N/A")
                preview_text += f"{i+1}. {email} - {name}\n"
                
                # Display all fields for the first recipient to help diagnose issues
                if i == 0:
                    preview_text += "   All fields in first row: " + str(recipient) + "\n"
            
            if len(recipients) > 5:
                preview_text += "...\n"
                
            self.recipients_preview.delete(1.0, tk.END)
            self.recipients_preview.insert(tk.END, preview_text)
        except Exception as e:
            messagebox.showerror("Error", f"Could not load CSV: {str(e)}")
    
    def view_csv_structure(self):
        """Show the structure and first few rows of the CSV file to help diagnosis"""
        if not self.csv_file_path:
            messagebox.showerror("Error", "No CSV file selected")
            return
            
        try:
            # Try different encodings to find one that works
            encodings = ['utf-8', 'utf-8-sig', 'latin-1', 'iso-8859-1', 'cp1252']
            data = None
            
            for encoding in encodings:
                try:
                    with open(self.csv_file_path, 'r', encoding=encoding) as f:
                        sample = f.read(1024)  # Read a sample to check encoding
                    
                    # If we got here, encoding worked for the sample
                    with open(self.csv_file_path, 'r', encoding=encoding) as f:
                        reader = csv.reader(f)
                        headers = next(reader, [])
                        rows = [row for _, row in zip(range(5), reader)]  # Get first 5 rows
                    
                    data = {
                        'encoding': encoding,
                        'headers': headers,
                        'rows': rows
                    }
                    break
                except UnicodeDecodeError:
                    continue
                except Exception as e:
                    messagebox.showerror("Error", f"Error reading CSV: {str(e)}")
                    return
            
            if not data:
                messagebox.showerror("Error", "Could not read CSV with any encoding")
                return
            
            # Create a new window to display the structure
            structure_window = tk.Toplevel(self.root)
            structure_window.title("CSV Structure")
            structure_window.geometry("800x600")
            
            info_text = f"File: {self.csv_file_path}\n"
            info_text += f"Detected Encoding: {data['encoding']}\n\n"
            info_text += f"Headers ({len(data['headers'])}): {', '.join(data['headers'])}\n\n"
            info_text += "First few rows:\n"
            
            for i, row in enumerate(data['rows']):
                info_text += f"Row {i+1}: {row}\n"
                
                # Show as dictionary mapping for clarity
                if data['headers'] and len(data['headers']) > 0:
                    row_dict = {}
                    for j, header in enumerate(data['headers']):
                        if j < len(row):
                            row_dict[header] = row[j]
                    info_text += f"       As Dict: {row_dict}\n"
            
            display = scrolledtext.ScrolledText(structure_window, width=90, height=30)
            display.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            display.insert(tk.END, info_text)
            
            # Add suggestions
            suggestions = "\n\nSuggestions:\n"
            suggestions += "1. Set 'CSV Encoding' to: " + data['encoding'] + "\n"
            
            if data['headers']:
                email_candidates = [h for h in data['headers'] if 'email' in h.lower() or 'mail' in h.lower()]
                name_candidates = [h for h in data['headers'] if 'name' in h.lower()]
                
                if email_candidates:
                    suggestions += f"2. Set 'Email Column' to: {email_candidates[0]}\n"
                if name_candidates:
                    suggestions += f"3. Set 'Name Column' to: {name_candidates[0]}\n"
            
            display.insert(tk.END, suggestions)
            
            # Create buttons to apply suggestions
            btn_frame = ttk.Frame(structure_window)
            btn_frame.pack(fill=tk.X, padx=10, pady=10)
            
            ttk.Button(btn_frame, text="Apply Suggested Settings", 
                      command=lambda: self.apply_suggested_settings(data)).pack(side=tk.LEFT, padx=5)
            
            ttk.Button(btn_frame, text="Close", 
                      command=structure_window.destroy).pack(side=tk.RIGHT, padx=5)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error analyzing CSV: {str(e)}")
    
    def apply_suggested_settings(self, data):
        """Apply the suggested settings from CSV analysis"""
        # Set encoding
        self.csv_encoding.set(data['encoding'])
        
        # Try to find email and name columns
        headers = data['headers']
        if headers:
            email_candidates = [h for h in headers if 'email' in h.lower() or 'mail' in h.lower()]
            name_candidates = [h for h in headers if 'name' in h.lower()]
            
            if email_candidates:
                self.email_column.delete(0, tk.END)
                self.email_column.insert(0, email_candidates[0])
            
            if name_candidates:
                self.name_column.delete(0, tk.END)
                self.name_column.insert(0, name_candidates[0])
        
        # Reload CSV with new settings
        self.reload_csv()
    
    def read_email_list(self):
        """Read email list from CSV file"""
        recipients = []
        try:
            with open(self.csv_file_path, 'r', encoding=self.csv_encoding.get()) as file:
                reader = csv.DictReader(file)
                for row in reader:
                    recipients.append(row)
            return recipients
        except UnicodeDecodeError:
            # If the selected encoding fails, try some common encodings
            encodings = ['utf-8', 'utf-8-sig', 'latin-1', 'iso-8859-1', 'cp1252']
            for encoding in encodings:
                if encoding == self.csv_encoding.get():
                    continue  # Skip the one we already tried
                
                try:
                    with open(self.csv_file_path, 'r', encoding=encoding) as file:
                        reader = csv.DictReader(file)
                        recipients = [row for row in reader]
                    
                    # If successful, update the encoding dropdown
                    self.csv_encoding.set(encoding)
                    self.status_var.set(f"Automatically switched to {encoding} encoding")
                    return recipients
                except UnicodeDecodeError:
                    continue
                
            # If we get here, none of the encodings worked
            raise Exception("Could not decode CSV file with common encodings. Try manually selecting an encoding.")
    
    def personalize_message(self, template, recipient_data):
        """Replace placeholders in template with recipient data"""
        personalized = template
        for key, value in recipient_data.items():
            placeholder = f"{{{{{key}}}}}"
            if placeholder in personalized:
                personalized = personalized.replace(placeholder, str(value))
        return personalized
    
    def send_email(self, sender_email, sender_password, recipient_email, subject, message):
        """Send an email using SMTP"""
        try:
            # Create message
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = recipient_email
            msg['Subject'] = subject
            
            # Add body to email
            msg.attach(MIMEText(message, 'html'))
            
            # Connect to SMTP server
            server = self.get_smtp_connection()
            
            # Login to sender email account if authentication is enabled
            if self.use_auth.get():
                server.login(sender_email, sender_password)
            
            # Send email
            server.sendmail(sender_email, recipient_email, msg.as_string())
            server.quit()
            return True
        except Exception as e:
            self.status_var.set(f"Error sending to {recipient_email}: {str(e)}")
            return False
    
    def start_sending(self):
        """Start sending emails in a separate thread"""
        if not self.template_content:
            messagebox.showerror("Error", "Please load an email template first.")
            return
        
        if not self.csv_file_path:
            messagebox.showerror("Error", "Please load a CSV file with recipients.")
            return
        
        if not self.email_subject.get():
            messagebox.showerror("Error", "Please enter an email subject.")
            return
        
        if self.use_auth.get() and (not self.sender_email.get() or not self.sender_password.get()):
            messagebox.showerror("Error", "Please enter sender email and password.")
            return
        
        if self.is_sending:
            messagebox.showinfo("Info", "Emails are already being sent.")
            return
        
        # Prepare for sending
        try:
            self.emails_to_send = self.read_email_list()
            if not self.emails_to_send:
                messagebox.showerror("Error", "No recipients found in the CSV file.")
                return
            
            # Check if email column exists
            email_col = self.email_column.get()
            if email_col not in self.emails_to_send[0]:
                messagebox.showerror("Error", f"Column '{email_col}' not found in CSV. Available columns: {', '.join(self.emails_to_send[0].keys())}")
                return
            
            # Start sending in a separate thread
            self.is_sending = True
            self.send_button.config(state=tk.DISABLED)
            self.progress['maximum'] = len(self.emails_to_send)
            self.progress['value'] = 0
            
            thread = threading.Thread(target=self.send_emails)
            thread.daemon = True
            thread.start()
        except Exception as e:
            messagebox.showerror("Error", f"Error preparing to send: {str(e)}")
    
    def send_emails(self):
        """Send emails one by one with delay"""
        success_count = 0
        delay_seconds = int(self.delay.get())
        email_col = self.email_column.get()
        
        for i, recipient in enumerate(self.emails_to_send):
            if not self.is_sending:
                break
            
            # Get recipient email
            recipient_email = recipient.get(email_col)
            if not recipient_email:
                self.status_var.set(f"Missing email for recipient {i+1}, skipping")
                continue
            
            # Update UI
            self.status_var.set(f"Sending to {recipient_email} ({i+1}/{len(self.emails_to_send)})")
            self.root.update_idletasks()
            
            # Personalize message
            try:
                personalized_message = self.personalize_message(self.template_content, recipient)
                
                # Send email
                if self.send_email(
                    self.sender_email.get(),
                    self.sender_password.get(),
                    recipient_email,
                    self.email_subject.get(),
                    personalized_message
                ):
                    success_count += 1
                
                # Update progress
                self.progress['value'] = i + 1
                self.root.update_idletasks()
                
                # Wait before sending next email
                time.sleep(delay_seconds)
            except Exception as e:
                self.status_var.set(f"Error with {recipient_email}: {str(e)}")
        
        # Update UI when done
        self.is_sending = False
        self.send_button.config(state=tk.NORMAL)
        self.status_var.set(f"Completed: Successfully sent {success_count} of {len(self.emails_to_send)} emails.")
        messagebox.showinfo("Complete", f"Email campaign completed.\nSuccessfully sent {success_count} of {len(self.emails_to_send)} emails.")

if __name__ == "__main__":
    root = tk.Tk()
    app = SimpleEmailSender(root)
    root.mainloop()