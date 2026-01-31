import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import threading
import time
from tkinter import font
import webbrowser
import tempfile
import json


class EmailSenderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("μLearn Email Sender - Professional Email Marketing Tool")
        self.root.geometry("1200x800")
        self.root.configure(bg='#f0f0f0')
        
        # Variables
        self.csv_file_path = tk.StringVar()
        self.smtp_server = tk.StringVar(value="email-smtp.ap-south-1.amazonaws.com")
        self.smtp_port = tk.StringVar(value="587")
        self.username = tk.StringVar()
        self.password = tk.StringVar()
        self.sender_email = tk.StringVar()
        self.reply_to_email = tk.StringVar(value="info@mulearn.org")
        self.subject = tk.StringVar()
        self.csv_data = None
        self.current_preview_index = 0
        self.attachments = []
        
        self.setup_styles()
        self.create_widgets()
        self.load_config()
        
    def setup_styles(self):
        """Setup custom styles for the application"""
        style = ttk.Style()
        
        # Configure custom styles
        style.configure('Title.TLabel', font=('Arial', 16, 'bold'), background='#f0f0f0')
        style.configure('Heading.TLabel', font=('Arial', 12, 'bold'), background='#f0f0f0')
        style.configure('Custom.TButton', font=('Arial', 10))
        style.configure('Success.TLabel', foreground='green', font=('Arial', 10, 'bold'))
        style.configure('Error.TLabel', foreground='red', font=('Arial', 10, 'bold'))
        
    def create_widgets(self):
        """Create all GUI widgets"""
        # Create main container with tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Tab 1: SMTP Configuration
        self.create_smtp_tab()
        
        # Tab 2: CSV Upload & Preview
        self.create_csv_tab()
        
        # Tab 3: Email Composition
        self.create_compose_tab()
        
        # Tab 4: Preview & Send
        self.create_preview_tab()
        
    def create_smtp_tab(self):
        """Create SMTP configuration tab"""
        smtp_frame = ttk.Frame(self.notebook)
        self.notebook.add(smtp_frame, text="📧 SMTP Configuration")
        
        # Title
        title_label = ttk.Label(smtp_frame, text="SMTP Server Configuration", style='Title.TLabel')
        title_label.pack(pady=20)
        
        # Main configuration frame
        config_frame = ttk.LabelFrame(smtp_frame, text="Email Server Settings", padding=20)
        config_frame.pack(fill='x', padx=20, pady=10)
        
        # SMTP Server
        ttk.Label(config_frame, text="SMTP Server:", style='Heading.TLabel').grid(row=0, column=0, sticky='w', pady=5)
        smtp_entry = ttk.Entry(config_frame, textvariable=self.smtp_server, width=40)
        smtp_entry.grid(row=0, column=1, padx=10, pady=5, sticky='ew')
        
        # SMTP Port
        ttk.Label(config_frame, text="SMTP Port:", style='Heading.TLabel').grid(row=1, column=0, sticky='w', pady=5)
        port_entry = ttk.Entry(config_frame, textvariable=self.smtp_port, width=40)
        port_entry.grid(row=1, column=1, padx=10, pady=5, sticky='ew')
        
        # Username
        ttk.Label(config_frame, text="Username:", style='Heading.TLabel').grid(row=2, column=0, sticky='w', pady=5)
        username_entry = ttk.Entry(config_frame, textvariable=self.username, width=40)
        username_entry.grid(row=2, column=1, padx=10, pady=5, sticky='ew')
        
        # Password
        ttk.Label(config_frame, text="Password:", style='Heading.TLabel').grid(row=3, column=0, sticky='w', pady=5)
        password_entry = ttk.Entry(config_frame, textvariable=self.password, show="*", width=40)
        password_entry.grid(row=3, column=1, padx=10, pady=5, sticky='ew')
        
        # Sender Email
        ttk.Label(config_frame, text="Sender Email:", style='Heading.TLabel').grid(row=4, column=0, sticky='w', pady=5)
        sender_entry = ttk.Entry(config_frame, textvariable=self.sender_email, width=40)
        sender_entry.grid(row=4, column=1, padx=10, pady=5, sticky='ew')
        
        # Reply-To Email
        ttk.Label(config_frame, text="Reply-To Email:", style='Heading.TLabel').grid(row=5, column=0, sticky='w', pady=5)
        reply_to_entry = ttk.Entry(config_frame, textvariable=self.reply_to_email, width=40)
        reply_to_entry.grid(row=5, column=1, padx=10, pady=5, sticky='ew')
        
        # Configure grid weights
        config_frame.columnconfigure(1, weight=1)
        
        # Test connection button
        test_btn = ttk.Button(config_frame, text="🔍 Test SMTP Connection", 
                             command=self.test_smtp_connection, style='Custom.TButton')
        test_btn.grid(row=6, column=1, pady=20, sticky='e')
        
        # Connection status label
        self.connection_status = ttk.Label(config_frame, text="", font=('Arial', 10))
        self.connection_status.grid(row=7, column=1, sticky='ew')
        
        # Preset configurations
        presets_frame = ttk.LabelFrame(smtp_frame, text="Quick Presets", padding=15)
        presets_frame.pack(fill='x', padx=20, pady=10)
        
        ttk.Button(presets_frame, text="AWS SES", command=self.load_aws_preset).pack(side='left', padx=5)
        ttk.Button(presets_frame, text="Gmail", command=self.load_gmail_preset).pack(side='left', padx=5)
        ttk.Button(presets_frame, text="Outlook", command=self.load_outlook_preset).pack(side='left', padx=5)
        
        # Save Config Button
        ttk.Button(presets_frame, text="💾 Save Configuration", command=self.save_config).pack(side='right', padx=5)
        
    def create_csv_tab(self):
        """Create CSV upload and data preview tab"""
        csv_frame = ttk.Frame(self.notebook)
        self.notebook.add(csv_frame, text="📊 CSV Data")
        
        # Title
        title_label = ttk.Label(csv_frame, text="CSV File Upload & Data Preview", style='Title.TLabel')
        title_label.pack(pady=20)
        
        # Upload section
        upload_frame = ttk.LabelFrame(csv_frame, text="Upload CSV File", padding=15)
        upload_frame.pack(fill='x', padx=20, pady=10)
        
        # File selection
        file_frame = ttk.Frame(upload_frame)
        file_frame.pack(fill='x')
        
        ttk.Label(file_frame, text="Selected File:", style='Heading.TLabel').pack(side='left')
        file_label = ttk.Label(file_frame, textvariable=self.csv_file_path, foreground='blue')
        file_label.pack(side='left', padx=10)
        
        ttk.Button(file_frame, text="📁 Browse CSV File", 
                  command=self.browse_csv_file).pack(side='right', padx=5)
        ttk.Button(file_frame, text="🔄 Reload", 
                  command=self.load_csv_data).pack(side='right', padx=5)
        
        # CSV Requirements
        req_frame = ttk.Frame(upload_frame)
        req_frame.pack(fill='x', pady=10)
        
        req_label = ttk.Label(req_frame, text="📋 CSV Requirements:", style='Heading.TLabel')
        req_label.pack(anchor='w')
        
        requirements = [
            "• Must contain 'Name' and 'Email' columns",
            "• Email addresses should be valid",
            "• Additional columns can be used for personalization"
        ]
        
        for req in requirements:
            ttk.Label(req_frame, text=req, font=('Arial', 9)).pack(anchor='w', padx=20)
        
        # Data preview section
        preview_frame = ttk.LabelFrame(csv_frame, text="Data Preview", padding=15)
        preview_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        # Treeview for data display
        self.tree = ttk.Treeview(preview_frame)
        self.tree.pack(fill='both', expand=True)
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(preview_frame, orient='vertical', command=self.tree.yview)
        v_scrollbar.pack(side='right', fill='y')
        self.tree.configure(yscrollcommand=v_scrollbar.set)
        
        h_scrollbar = ttk.Scrollbar(preview_frame, orient='horizontal', command=self.tree.xview)
        h_scrollbar.pack(side='bottom', fill='x')
        self.tree.configure(xscrollcommand=h_scrollbar.set)
        
        # Statistics
        self.stats_label = ttk.Label(preview_frame, text="No data loaded", style='Heading.TLabel')
        self.stats_label.pack(pady=10)
        
    def create_compose_tab(self):
        """Create email composition tab"""
        compose_frame = ttk.Frame(self.notebook)
        self.notebook.add(compose_frame, text="✏️ Compose Email")
        
        # Title
        title_label = ttk.Label(compose_frame, text="Email Composition & Formatting", style='Title.TLabel')
        title_label.pack(pady=20)
        
        # Subject section
        subject_frame = ttk.LabelFrame(compose_frame, text="Email Subject", padding=15)
        subject_frame.pack(fill='x', padx=20, pady=10)
        
        ttk.Label(subject_frame, text="Subject Line:", style='Heading.TLabel').pack(anchor='w')
        subject_entry = ttk.Entry(subject_frame, textvariable=self.subject, font=('Arial', 11))
        subject_entry.pack(fill='x', pady=5)
        
        # Formatting toolbar
        toolbar_frame = ttk.LabelFrame(compose_frame, text="Formatting Tools", padding=10)
        toolbar_frame.pack(fill='x', padx=20, pady=5)
        
        # Toolbar buttons
        btn_frame = ttk.Frame(toolbar_frame)
        btn_frame.pack(fill='x')
        
        formatting_buttons = [
            ("B", "Bold", self.make_bold),
            ("I", "Italic", self.make_italic),
            ("U", "Underline", self.make_underline),
            ("🔗", "Link", self.add_link),
            ("📋", "Template", self.load_template),
            ("👁️", "Preview HTML", self.preview_html)
        ]
        
        for text, tooltip, command in formatting_buttons:
            btn = ttk.Button(btn_frame, text=text, command=command, width=8)
            btn.pack(side='left', padx=2)
        
        # Email content area
        content_frame = ttk.LabelFrame(compose_frame, text="Email Content (HTML Supported)", padding=15)
        content_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        # Text editor with scrollbar
        text_frame = ttk.Frame(content_frame)
        text_frame.pack(fill='both', expand=True)
        
        self.email_content = scrolledtext.ScrolledText(text_frame, height=15, font=('Consolas', 10))
        self.email_content.pack(fill='both', expand=True)
        
        # Quick templates
        template_frame = ttk.Frame(content_frame)
        template_frame.pack(fill='x', pady=10)
        
        ttk.Label(template_frame, text="Quick Templates:", style='Heading.TLabel').pack(side='left')
        
        templates = [
            ("Welcome", self.load_welcome_template),
            ("Notification", self.load_notification_template),
            ("Certificate", self.load_certificate_template)
        ]
        
        for name, command in templates:
            ttk.Button(template_frame, text=name, command=command).pack(side='left', padx=5)
        
        for name, command in templates:
            ttk.Button(template_frame, text=name, command=command).pack(side='left', padx=5)
        
        # Variable Insertion Helper
        var_frame = ttk.LabelFrame(content_frame, text="Insert Variable", padding=10)
        var_frame.pack(fill='x', pady=5)
        
        ttk.Label(var_frame, text="Select Field:", style='Heading.TLabel').pack(side='left', padx=5)
        
        self.var_combo = ttk.Combobox(var_frame, state='readonly', width=30)
        self.var_combo.pack(side='left', padx=5)
        self.var_combo.set("Load CSV first...")
        
        ttk.Button(var_frame, text="Insert into Email", command=self.insert_variable).pack(side='left', padx=5)
        
        ttk.Label(var_frame, text="(Automatically replaces {Field} with CSV data for each person)", font=('Arial', 8, 'italic')).pack(side='left', padx=10)

        # Attachment Section
        attachment_frame = ttk.LabelFrame(compose_frame, text="Attachments", padding=15)
        attachment_frame.pack(fill='x', padx=20, pady=10)
        
        # Attachment List
        self.attachment_list = tk.Listbox(attachment_frame, height=4)
        self.attachment_list.pack(side='left', fill='x', expand=True, padx=(0, 10))
        
        # Attachment Buttons
        att_btn_frame = ttk.Frame(attachment_frame)
        att_btn_frame.pack(side='right', fill='y')
        
        ttk.Button(att_btn_frame, text="➕ Add Files", command=self.add_attachment).pack(fill='x', pady=2)
        ttk.Button(att_btn_frame, text="➖ Remove", command=self.remove_attachment).pack(fill='x', pady=2)
        ttk.Button(att_btn_frame, text="🗑️ Clear All", command=self.clear_attachments).pack(fill='x', pady=2)
        
    def create_preview_tab(self):
        """Create email preview and sending tab"""
        preview_frame = ttk.Frame(self.notebook)
        self.notebook.add(preview_frame, text="📤 Preview & Send")
        
        # Title
        title_label = ttk.Label(preview_frame, text="Email Preview & Bulk Sending", style='Title.TLabel')
        title_label.pack(pady=20)
        
        # Preview section
        preview_section = ttk.LabelFrame(preview_frame, text="Email Preview", padding=15)
        preview_section.pack(fill='both', expand=True, padx=20, pady=10)
        
        # Navigation controls
        nav_frame = ttk.Frame(preview_section)
        nav_frame.pack(fill='x', pady=10)
        
        ttk.Button(nav_frame, text="◀ Previous", command=self.prev_preview).pack(side='left')
        
        self.preview_info = ttk.Label(nav_frame, text="No data loaded", style='Heading.TLabel')
        self.preview_info.pack(side='left', padx=20)
        
        ttk.Button(nav_frame, text="Next ▶", command=self.next_preview).pack(side='left')
        ttk.Button(nav_frame, text="🌐 Open in Browser", command=self.open_in_browser).pack(side='right')
        
        # Preview display
        self.preview_display = scrolledtext.ScrolledText(preview_section, height=20, state='disabled')
        self.preview_display.pack(fill='both', expand=True, pady=10)
        
        # Sending section
        send_frame = ttk.LabelFrame(preview_frame, text="Send Emails", padding=15)
        send_frame.pack(fill='x', padx=20, pady=10)
        
        # Send controls
        control_frame = ttk.Frame(send_frame)
        control_frame.pack(fill='x')
        
        self.send_button = ttk.Button(control_frame, text="🚀 Send All Emails", 
                                     command=self.send_all_emails, style='Custom.TButton')
        self.send_button.pack(side='left', padx=5)
        
        ttk.Button(control_frame, text="📧 Send Test Email", 
                  command=self.send_test_email).pack(side='left', padx=5)
        
        ttk.Button(control_frame, text="📮 Custom Test Email", 
                  command=self.send_custom_test_email).pack(side='left', padx=5)
        
        self.stop_button = ttk.Button(control_frame, text="⏹️ Stop Sending", 
                                     command=self.stop_sending, state='disabled')
        self.stop_button.pack(side='left', padx=5)
        
        # Progress section
        progress_frame = ttk.Frame(send_frame)
        progress_frame.pack(fill='x', pady=10)
        
        ttk.Label(progress_frame, text="Progress:", style='Heading.TLabel').pack(anchor='w')
        
        self.progress_bar = ttk.Progressbar(progress_frame, mode='determinate')
        self.progress_bar.pack(fill='x', pady=5)
        
        self.progress_label = ttk.Label(progress_frame, text="Ready to send")
        self.progress_label.pack(anchor='w')
        
        # Log section
        log_frame = ttk.LabelFrame(send_frame, text="Sending Log", padding=10)
        log_frame.pack(fill='x', pady=10)
        
        self.log_display = scrolledtext.ScrolledText(log_frame, height=8, state='disabled')
        self.log_display.pack(fill='both', expand=True)
        
        # Initialize variables
        self.sending_stopped = False
        
    def browse_csv_file(self):
        """Browse and select CSV file"""
        file_path = filedialog.askopenfilename(
            title="Select CSV File",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if file_path:
            self.csv_file_path.set(file_path)
            self.load_csv_data()
    
    def load_csv_data(self):
        """Load and validate CSV data"""
        if not self.csv_file_path.get():
            messagebox.showwarning("Warning", "Please select a CSV file first!")
            return
        
        try:
            self.csv_data = pd.read_csv(self.csv_file_path.get())
            
            # Validate required columns
            required_columns = ['Name', 'Email']
            missing_columns = [col for col in required_columns if col not in self.csv_data.columns]
            
            if missing_columns:
                messagebox.showerror("Error", f"Missing required columns: {', '.join(missing_columns)}")
                return
            
            # Update treeview
            self.update_data_preview()
            
            # Update statistics
            total_rows = len(self.csv_data)
            valid_emails = self.csv_data['Email'].notna().sum()
            self.stats_label.config(text=f"Total records: {total_rows} | Valid emails: {valid_emails}")
            
            # Reset preview index
            self.current_preview_index = 0
            self.update_email_preview()
            
            self.log_message(f"✅ Loaded {total_rows} records from CSV file")
            
            # Update variable combo
            self.var_combo['values'] = list(self.csv_data.columns)
            if len(self.csv_data.columns) > 0:
                self.var_combo.current(0)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load CSV file: {str(e)}")
            self.log_message(f"❌ Error loading CSV: {str(e)}")
    
    def update_data_preview(self):
        """Update the data preview treeview"""
        # Clear existing data
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        if self.csv_data is None:
            return
        
        # Configure columns
        columns = list(self.csv_data.columns)
        self.tree['columns'] = columns
        self.tree['show'] = 'headings'
        
        # Configure column headings and widths
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120, minwidth=80)
        
        # Insert data (limit to first 100 rows for performance)
        for index, row in self.csv_data.head(100).iterrows():
            values = [str(row[col]) if pd.notna(row[col]) else '' for col in columns]
            self.tree.insert('', 'end', values=values)
    
    def test_smtp_connection(self):
        """Test SMTP connection"""
        if not all([self.smtp_server.get(), self.smtp_port.get(), self.username.get(), self.password.get()]):
            messagebox.showwarning("Warning", "Please fill in all SMTP configuration fields!")
            return
        
        def test_connection():
            try:
                self.connection_status.config(text="🔄 Testing connection...", foreground='blue')
                self.root.update()
                
                server = smtplib.SMTP(self.smtp_server.get(), int(self.smtp_port.get()))
                server.starttls()
                server.login(self.username.get(), self.password.get())
                server.quit()
                
                self.connection_status.config(text="✅ Connection successful!", foreground='green')
                self.log_message("✅ SMTP connection test successful")
                
            except Exception as e:
                self.connection_status.config(text=f"❌ Connection failed: {str(e)}", foreground='red')
                self.log_message(f"❌ SMTP connection failed: {str(e)}")
        
        # Run test in separate thread
        threading.Thread(target=test_connection, daemon=True).start()
    
    def load_aws_preset(self):
        """Load AWS SES preset configuration"""
        self.smtp_server.set("email-smtp.ap-south-1.amazonaws.com")
        self.smtp_port.set("587")
        self.log_message("📝 Loaded AWS SES preset configuration")
    
    def load_gmail_preset(self):
        """Load Gmail preset configuration"""
        self.smtp_server.set("smtp.gmail.com")
        self.smtp_port.set("587")
        self.log_message("📝 Loaded Gmail preset configuration")
    
    def load_outlook_preset(self):
        """Load Outlook preset configuration"""
        self.smtp_server.set("smtp-mail.outlook.com")
        self.smtp_port.set("587")
        self.log_message("📝 Loaded Outlook preset configuration")
    
    def make_bold(self):
        """Insert bold formatting"""
        self.insert_formatting("<strong>", "</strong>")
    
    def make_italic(self):
        """Insert italic formatting"""
        self.insert_formatting("<em>", "</em>")
    
    def make_underline(self):
        """Insert underline formatting"""
        self.insert_formatting("<u>", "</u>")
    
    def insert_formatting(self, start_tag, end_tag):
        """Insert HTML formatting tags"""
        try:
            selection = self.email_content.selection_get()
            self.email_content.delete("sel.first", "sel.last")
            self.email_content.insert("insert", f"{start_tag}{selection}{end_tag}")
        except tk.TclError:
            # No selection, just insert tags
            self.email_content.insert("insert", f"{start_tag}{end_tag}")
            # Move cursor between tags
            current_pos = self.email_content.index("insert")
            new_pos = f"{current_pos.split('.')[0]}.{int(current_pos.split('.')[1]) - len(end_tag)}"
            self.email_content.mark_set("insert", new_pos)
    
    def add_link(self):
        """Add a hyperlink"""
        url = tk.simpledialog.askstring("Add Link", "Enter URL:")
        text = tk.simpledialog.askstring("Add Link", "Enter link text:")
        
        if url and text:
            link_html = f'<a href="{url}" style="color: #007bff; text-decoration: none;">{text}</a>'
            self.email_content.insert("insert", link_html)
    
    def load_template(self):
        """Load a predefined template"""
        template_window = tk.Toplevel(self.root)
        template_window.title("Select Template")
        template_window.geometry("400x300")
        
        # Template list
        templates = {
            "Welcome Email": self.get_welcome_template(),
            "Event Notification": self.get_notification_template(),
            "Certificate Email": self.get_certificate_template()
        }
        
        listbox = tk.Listbox(template_window)
        listbox.pack(fill='both', expand=True, padx=10, pady=10)
        
        for template_name in templates.keys():
            listbox.insert('end', template_name)
        
        def load_selected():
            selection = listbox.curselection()
            if selection:
                template_name = listbox.get(selection[0])
                self.email_content.delete(1.0, 'end')
                self.email_content.insert(1.0, templates[template_name])
                template_window.destroy()
        
        ttk.Button(template_window, text="Load Template", command=load_selected).pack(pady=10)
    
    def preview_html(self):
        """Preview HTML in browser"""
        self.open_in_browser()
    
    def load_welcome_template(self):
        """Load welcome email template"""
        self.email_content.delete(1.0, 'end')
        self.email_content.insert(1.0, self.get_welcome_template())
    
    def load_notification_template(self):
        """Load notification template"""
        self.email_content.delete(1.0, 'end')
        self.email_content.insert(1.0, self.get_notification_template())
    
    def load_certificate_template(self):
        """Load certificate template"""
        self.email_content.delete(1.0, 'end')
        self.email_content.insert(1.0, self.get_certificate_template())
    
    def get_welcome_template(self):
        """Return welcome email template"""
        return '''<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; background-color: #f8f9fa;">
    <div style="background-color: white; padding: 30px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1);">
        <h2 style="color: #2c3e50; margin-bottom: 20px;">Welcome <span style="color: #e74c3c;">{Name}</span>! 🎉</h2>
        
        <p style="color: #495057; line-height: 1.6;">We're <strong>thrilled</strong> to have you join our community!</p>
        
        <div style="background-color: #d4edda; border-left: 4px solid #28a745; padding: 15px; margin: 20px 0;">
            <p style="margin: 0; color: #155724;"><strong>✅ Your account is now active!</strong></p>
        </div>
        
        <p style="color: #495057; line-height: 1.6;">Start exploring our platform and discover all the amazing features we have prepared for you.</p>
        
        <div style="text-align: center; margin: 25px 0;">
            <a href="#" style="display: inline-block; background-color: #007bff; color: white; padding: 12px 25px; text-decoration: none; border-radius: 5px; font-weight: bold;">🚀 Get Started</a>
        </div>
        
        <p style="color: #6c757d; font-size: 14px; margin-top: 30px; border-top: 1px solid #dee2e6; padding-top: 20px;">
            Best regards,<br><strong>The Team</strong>
        </p>
    </div>
</div>'''
    
    def get_notification_template(self):
        """Return notification template"""
        return '''<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; background-color: #f8f9fa;">
    <div style="background-color: white; padding: 25px; border-radius: 8px; border-left: 5px solid #ffc107;">
        <h3 style="color: #856404; margin-top: 0;">📢 Important Update for {Name}</h3>
        
        <p style="color: #495057; line-height: 1.6;">We have an <em>important update</em> to share with you regarding your recent activity.</p>
        
        <div style="background-color: #fff3cd; border: 1px solid #ffeaa7; border-radius: 5px; padding: 15px; margin: 15px 0;">
            <p style="margin: 0; color: #856404;"><strong>Action Required:</strong> Please review the changes and update your settings accordingly.</p>
        </div>
        
        <p style="color: #495057; line-height: 1.6;">If you have any questions, don't hesitate to contact our support team.</p>
        
        <p style="color: #6c757d; font-size: 14px; margin-top: 20px;">Best regards,<br><strong>The Team</strong></p>
    </div>
</div>'''
    
    def get_certificate_template(self):
        """Return certificate template"""
        return '''<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; background-color: #f0f0f0;">
    <div style="background-color: white; padding: 30px; border-radius: 10px; box-shadow: 0 2px 10px rgba(0,0,0,0.1);">
        <p>Dear <strong>{Name}</strong>,</p>
        
        <p>We hope this message finds you well. We are pleased to extend our <strong>heartfelt congratulations</strong> to you for your active participation in the recent <em>Grand Tech Adventure: CodeStorm Hackathon</em>.</p>

        <p>Feel free to share this accomplishment with your peers, colleagues, and networks.</p>
        
        <p>We believe that your experience in the <strong>GTA: CodeStorm Hackathon</strong> has not only enriched your own skill set but has also contributed to the larger goal of fostering innovation and positive change.</p>

        <p>Once again, <strong>congratulations</strong> on your achievement, and we hope to see you in our future endeavors.</p>
        
        <p>If you have any questions or require further information, please don't hesitate to reach out.<br>Please find the attached certificate below.</p>

        <p><a href="https://drive.google.com/drive/folders/1IgMy7vXK2vy0bbJjOtmZcuZp3ywYdTIo?usp=sharing" style="color: #007bff; text-decoration: none;"><strong>📂 Access Your Certificate</strong></a></p>

        <p>Best regards,<br><strong>Team µLearn</strong></p>
    </div>
</div>'''
    
    def update_email_preview(self):
        """Update email preview with current data"""
        if self.csv_data is None or len(self.csv_data) == 0:
            self.preview_info.config(text="No data loaded")
            self.preview_display.config(state='normal')
            self.preview_display.delete(1.0, 'end')
            self.preview_display.insert(1.0, "No data to preview")
            self.preview_display.config(state='disabled')
            return
        
        # Get current row
        if self.current_preview_index >= len(self.csv_data):
            self.current_preview_index = 0
        
        row = self.csv_data.iloc[self.current_preview_index]
        
        # Update info
        self.preview_info.config(text=f"Preview {self.current_preview_index + 1} of {len(self.csv_data)}: {row['Name']} ({row['Email']})")
        
        # Generate preview content
        content = self.email_content.get(1.0, 'end-1c')
        
        # Replace variables
        for column in self.csv_data.columns:
            content = content.replace(f"{{{column}}}", str(row[column]) if pd.notna(row[column]) else "")
        
        # Update preview display
        self.preview_display.config(state='normal')
        self.preview_display.delete(1.0, 'end')
        self.preview_display.insert(1.0, content)
        self.preview_display.config(state='disabled')
    
    def prev_preview(self):
        """Show previous email preview"""
        if self.csv_data is not None and len(self.csv_data) > 0:
            self.current_preview_index = (self.current_preview_index - 1) % len(self.csv_data)
            self.update_email_preview()
    
    def next_preview(self):
        """Show next email preview"""
        if self.csv_data is not None and len(self.csv_data) > 0:
            self.current_preview_index = (self.current_preview_index + 1) % len(self.csv_data)
            self.update_email_preview()
    
    def open_in_browser(self):
        """Open current preview in browser"""
        if self.csv_data is None or len(self.csv_data) == 0:
            messagebox.showwarning("Warning", "No data to preview!")
            return
        
        # Get current content
        content = self.preview_display.get(1.0, 'end-1c')
        
        # Create temporary HTML file
        with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False) as f:
            f.write(f"""
            <!DOCTYPE html>
            <html>
            <head>
                <title>Email Preview</title>
                <meta charset="utf-8">
            </head>
            <body>
                {content}
            </body>
            </html>
            """)
            temp_path = f.name
        
        # Open in browser
        webbrowser.open(f'file://{temp_path}')
    
    def send_test_email(self):
        """Send a test email to the first recipient"""
        if not self.validate_send_requirements():
            return
        
        if len(self.csv_data) == 0:
            messagebox.showwarning("Warning", "No recipients in CSV file!")
            return
        
        # Get first recipient
        first_row = self.csv_data.iloc[0]
        
        def send_test():
            try:
                self.log_message(f"📧 Sending test email to {first_row['Email']}...")
                self.send_single_email(first_row['Email'], first_row['Name'], first_row.to_dict())
                self.log_message("✅ Test email sent successfully!")
                messagebox.showinfo("Success", f"Test email sent to {first_row['Email']}")
            except Exception as e:
                self.log_message(f"❌ Test email failed: {str(e)}")
                messagebox.showerror("Error", f"Test email failed: {str(e)}")
        
        threading.Thread(target=send_test, daemon=True).start()
    
    def send_custom_test_email(self):
        """Send a test email to a custom email address"""
        # Check basic requirements first (except CSV data)
        if not all([self.smtp_server.get(), self.smtp_port.get(), self.username.get(), 
                   self.password.get(), self.sender_email.get(), self.reply_to_email.get()]):
            messagebox.showwarning("Warning", "Please complete SMTP configuration!")
            return
        
        if not self.subject.get().strip():
            messagebox.showwarning("Warning", "Please enter email subject!")
            return
        
        if not self.email_content.get(1.0, 'end-1c').strip():
            messagebox.showwarning("Warning", "Please enter email content!")
            return
        
        # Create custom test email dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Send Custom Test Email")
        dialog.geometry("400x250")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center the dialog
        dialog.geometry("+%d+%d" % (self.root.winfo_rootx() + 50, self.root.winfo_rooty() + 50))
        
        # Create form
        main_frame = ttk.Frame(dialog, padding=20)
        main_frame.pack(fill='both', expand=True)
        
        ttk.Label(main_frame, text="Send test email to custom recipient", style='Heading.TLabel').pack(pady=(0, 15))
        
        # Email input
        ttk.Label(main_frame, text="Email Address:").pack(anchor='w')
        email_var = tk.StringVar()
        email_entry = ttk.Entry(main_frame, textvariable=email_var, width=40)
        email_entry.pack(fill='x', pady=(5, 10))
        email_entry.focus()
        
        # Name input (optional)
        ttk.Label(main_frame, text="Name (optional):").pack(anchor='w')
        name_var = tk.StringVar()
        name_entry = ttk.Entry(main_frame, textvariable=name_var, width=40)
        name_entry.pack(fill='x', pady=(5, 15))
        
        # Info text
        info_text = "Note: If no name is provided, '{Name}' placeholders will be replaced with the email address."
        ttk.Label(main_frame, text=info_text, font=('Arial', 8), wraplength=350).pack(pady=(0, 15))
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill='x')
        
        def send_custom():
            email = email_var.get().strip()
            if not email:
                messagebox.showwarning("Warning", "Please enter an email address!")
                return
            
            # Basic email validation
            if '@' not in email or '.' not in email:
                messagebox.showwarning("Warning", "Please enter a valid email address!")
                return
            
            name = name_var.get().strip() or email
            
            dialog.destroy()
            
            def send_custom_thread():
                try:
                    self.log_message(f"📮 Sending custom test email to {email}...")
                    
                    # Create a mock row data with the custom email and name
                    custom_row_data = {
                        'Email': email,
                        'Name': name
                    }
                    
                    # Add any additional columns from CSV if available
                    if self.csv_data is not None and len(self.csv_data) > 0:
                        # Use first row as template for any additional fields
                        first_row = self.csv_data.iloc[0]
                        for col in self.csv_data.columns:
                            if col not in custom_row_data:
                                custom_row_data[col] = first_row[col] if pd.notna(first_row[col]) else ""
                    
                    self.send_single_email(email, name, custom_row_data)
                    self.log_message("✅ Custom test email sent successfully!")
                    messagebox.showinfo("Success", f"Test email sent to {email}")
                except Exception as e:
                    self.log_message(f"❌ Custom test email failed: {str(e)}")
                    messagebox.showerror("Error", f"Test email failed: {str(e)}")
            
            threading.Thread(target=send_custom_thread, daemon=True).start()
        
        ttk.Button(button_frame, text="Send Test Email", command=send_custom).pack(side='right', padx=(5, 0))
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side='right')
        
        # Bind Enter key to send
        dialog.bind('<Return>', lambda e: send_custom())
        dialog.bind('<Escape>', lambda e: dialog.destroy())
    
    def send_all_emails(self):
        """Send emails to all recipients"""
        if not self.validate_send_requirements():
            return
        
        if len(self.csv_data) == 0:
            messagebox.showwarning("Warning", "No recipients in CSV file!")
            return
        
        # Confirm sending
        result = messagebox.askyesno("Confirm", f"Send emails to {len(self.csv_data)} recipients?")
        if not result:
            return
        
        # Start sending
        self.sending_stopped = False
        self.send_button.config(state='disabled')
        self.stop_button.config(state='normal')
        
        def send_all():
            try:
                total = len(self.csv_data)
                sent = 0
                failed = 0
                
                self.progress_bar['maximum'] = total
                
                for index, row in self.csv_data.iterrows():
                    if self.sending_stopped:
                        self.log_message("⏹️ Sending stopped by user")
                        break
                    
                    try:
                        email = row['Email']
                        name = row['Name']
                        
                        if pd.isna(email) or email.strip() == '':
                            self.log_message(f"⚠️ Skipping {name}: No email address")
                            continue
                        
                        self.log_message(f"📧 Sending to {name} ({email})...")
                        
                        self.send_single_email(email, name, row.to_dict())
                        sent += 1
                        
                        self.log_message(f"✅ Sent to {name}")
                        
                        # Update progress
                        self.progress_bar['value'] = sent + failed
                        self.progress_label.config(text=f"Sent: {sent} | Failed: {failed} | Remaining: {total - sent - failed}")
                        self.root.update()
                        
                        # Small delay to avoid overwhelming the server
                        time.sleep(2)
                        
                    except Exception as e:
                        failed += 1
                        self.log_message(f"❌ Failed to send to {row['Name']}: {str(e)}")
                        
                        # Update progress
                        self.progress_bar['value'] = sent + failed
                        self.progress_label.config(text=f"Sent: {sent} | Failed: {failed} | Remaining: {total - sent - failed}")
                        self.root.update()
                
                # Final summary
                self.log_message(f"🏁 Sending complete! Sent: {sent}, Failed: {failed}")
                self.progress_label.config(text=f"Complete! Sent: {sent} | Failed: {failed}")
                
                if not self.sending_stopped:
                    messagebox.showinfo("Complete", f"Email sending complete!\nSent: {sent}\nFailed: {failed}")
                
            except Exception as e:
                self.log_message(f"❌ Critical error: {str(e)}")
                messagebox.showerror("Error", f"Critical error: {str(e)}")
            
            finally:
                self.send_button.config(state='normal')
                self.stop_button.config(state='disabled')
        
        threading.Thread(target=send_all, daemon=True).start()
    
    def stop_sending(self):
        """Stop the email sending process"""
        self.sending_stopped = True
        self.stop_button.config(state='disabled')
    
    def send_single_email(self, to_email, name, row_data):
        """Send a single email"""
        # Create email content
        content = self.email_content.get(1.0, 'end-1c')
        
        # Replace variables
        for column, value in row_data.items():
            content = content.replace(f"{{{column}}}", str(value) if pd.notna(value) else "")
        
        # Create message
        msg = MIMEMultipart()
        msg['From'] = self.sender_email.get()
        msg['To'] = to_email
        msg['Reply-To'] = self.reply_to_email.get()
        msg['Subject'] = self.subject.get().replace("{Name}", name)
        
        # Attach body
        msg.attach(MIMEText(content, 'html'))
        
        # Attach files
        for file_path in self.attachments:
            try:
                filename = os.path.basename(file_path)
                with open(file_path, 'rb') as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
                    msg.attach(part)
            except Exception as e:
                self.log_message(f"⚠️ Failed to attach {filename}: {str(e)}")

        
        # Send email
        server = smtplib.SMTP(self.smtp_server.get(), int(self.smtp_port.get()))
        server.starttls()
        server.login(self.username.get(), self.password.get())
        server.sendmail(self.sender_email.get(), to_email, msg.as_string())
        server.quit()
    
    def validate_send_requirements(self):
        """Validate all requirements for sending emails"""
        # Check SMTP configuration
        if not all([self.smtp_server.get(), self.smtp_port.get(), self.username.get(), 
                   self.password.get(), self.sender_email.get(), self.reply_to_email.get()]):
            messagebox.showwarning("Warning", "Please complete SMTP configuration!")
            return False
        
        # Check subject
        if not self.subject.get().strip():
            messagebox.showwarning("Warning", "Please enter email subject!")
            return False
        
        # Check content
        if not self.email_content.get(1.0, 'end-1c').strip():
            messagebox.showwarning("Warning", "Please enter email content!")
            return False
        
        # Check CSV data
        if self.csv_data is None:
            messagebox.showwarning("Warning", "Please load CSV data!")
            return False
        
        return True
    
    def log_message(self, message):
        """Add message to log display"""
        self.log_display.config(state='normal')
        self.log_display.insert('end', f"{time.strftime('%H:%M:%S')} - {message}\n")
        self.log_display.see('end')
        self.log_display.config(state='disabled')
        self.root.update()

    def add_attachment(self):
        """Add files to attachment list"""
        filenames = filedialog.askopenfilenames(title="Select Attachments")
        if filenames:
            for filename in filenames:
                if filename not in self.attachments:
                    self.attachments.append(filename)
                    self.attachment_list.insert('end', os.path.basename(filename))

    def remove_attachment(self):
        """Remove selected attachment"""
        selection = self.attachment_list.curselection()
        if selection:
            index = selection[0]
            self.attachment_list.delete(index)
            self.attachments.pop(index)

    def clear_attachments(self):
        """Clear all attachments"""
        self.attachment_list.delete(0, 'end')
        self.attachments.clear()

    def save_config(self):
        """Save SMTP configuration to file"""
        config = {
            "smtp_server": self.smtp_server.get(),
            "smtp_port": self.smtp_port.get(),
            "username": self.username.get(),
            # Security: Do not save password
            "sender_email": self.sender_email.get(),
            "reply_to_email": self.reply_to_email.get(),
            "subject": self.subject.get()
        }
        
        try:
            with open('config.json', 'w') as f:
                json.dump(config, f)
            messagebox.showinfo("Success", "Configuration saved successfully!\n(Password was not saved for security)")
            self.log_message("💾 Configuration saved (excluding password)")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save config: {str(e)}")

    def load_config(self):
        """Load SMTP configuration from file"""
        if os.path.exists('config.json'):
            try:
                with open('config.json', 'r') as f:
                    config = json.load(f)
                
                self.smtp_server.set(config.get("smtp_server", ""))
                self.smtp_port.set(config.get("smtp_port", ""))
                self.username.set(config.get("username", ""))
                # Security: Password is not loaded
                self.sender_email.set(config.get("sender_email", ""))
                self.reply_to_email.set(config.get("reply_to_email", ""))
                
                # Only load subject if saved
                if "subject" in config:
                    self.subject.set(config["subject"])
                    
                self.log_message("📂 Configuration loaded")
            except Exception as e:
                self.log_message(f"⚠️ Failed to load config: {str(e)}")

    def insert_variable(self):
        """Insert selected variable into email content"""
        selection = self.var_combo.get()
        if selection and selection != "Load CSV first...":
            # Insert at cursor position
            self.email_content.insert("insert", f"{{{selection}}}")
            self.email_content.focus()



# Import missing module
try:
    import tkinter.simpledialog
except ImportError:
    pass

def main():
    root = tk.Tk()
    app = EmailSenderGUI(root)
    
    # Set initial values
    app.subject.set("Welcome to our platform!")
    app.sender_email.set("events@mulearn.in")
    app.reply_to_email.set("info@mulearn.org")
    app.load_welcome_template()
    
    root.mainloop()

if __name__ == "__main__":
    main()
