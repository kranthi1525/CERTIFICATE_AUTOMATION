import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser, simpledialog
from PIL import Image, ImageDraw, ImageFont, ImageTk
import pandas as pd
import os
import json
import base64
from datetime import datetime
import requests
import webbrowser
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from PyPDF2 import PdfReader, PdfWriter
import io

class CertificateGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("Advanced Dynamic Certificate Generator with Email")
        self.root.geometry("1500x950")
        
        # Variables
        self.template_image = None
        self.template_path = None
        self.csv_path = None
        self.df_current = None
        self.preview_row_var = tk.IntVar(value=1)
        self.canvas_scale = 1.0
        self.show_axis = tk.BooleanVar(value=True)
        self.show_crosshair = tk.BooleanVar(value=True)
        self.send_email = tk.BooleanVar(value=False)
        # Global feature flag (enabled by default)
        self.enable_claude_haiku_4_5 = tk.BooleanVar(value=True)
        
        # Text fields configuration
        self.text_fields = []
        self.current_field_index = 0
        self.preview_image = None
        
        # Email settings
        self.apps_script_url = ""
        self.email_column = ""
        
        # Default field types
        self.field_types = [
            "Name", "Roll Number", "Branch", "Course", "Date", 
            "Grade", "Score", "Department", "Year", "Email", "Custom"
        ]

        # NEW: Verification feature state
        self.enable_verification = tk.BooleanVar(value=False)
        self.uid_column_var = tk.StringVar()
        self.verification_position = None
        self.setting_verification_position = False
        self.verification_font_size = tk.IntVar(value=14)
        
        self.setup_ui()
    
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Left panel for controls
        left_panel = ttk.Frame(main_frame, width=400)
        left_panel.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        left_panel.pack_propagate(False)
        
        # Right panel for preview
        right_panel = ttk.Frame(main_frame)
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Setup panels
        self.setup_controls(left_panel)
        self.setup_preview(right_panel)
        
        # Available CSV columns
        self.csv_columns = []
    
    def setup_controls(self, parent):
        # Create scrollable frame for controls
        canvas = tk.Canvas(parent)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Title
        title_label = ttk.Label(scrollable_frame, text="Advanced Certificate Generator", 
                               font=("Arial", 14, "bold"))
        title_label.pack(pady=(0, 15))

        # Global settings
        self.setup_global_settings(scrollable_frame)
        
        # Template selection
        self.setup_template_section(scrollable_frame)
        
        # CSV selection
        self.setup_csv_section(scrollable_frame)
        
        # Text fields configuration
        self.setup_fields_section(scrollable_frame)

        # NEW: Verification (Optional) — unnumbered so existing numbering stays same
        self.setup_verification_section(scrollable_frame)
        
        # Email configuration
        self.setup_email_section(scrollable_frame)
        
        # Generation controls
        self.setup_generation_section(scrollable_frame)
        
        # Mouse wheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind("<MouseWheel>", _on_mousewheel)
    
    def setup_template_section(self, parent):
        template_frame = ttk.LabelFrame(parent, text="1. Select Template", padding=10)
        template_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(template_frame, text="Browse Template Image", 
                  command=self.browse_template).pack(fill=tk.X)
        
        self.template_label = ttk.Label(template_frame, text="No template selected", 
                                       foreground="gray")
        self.template_label.pack(pady=(5, 0))
        
        # Display options
        options_frame = ttk.Frame(template_frame)
        options_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Checkbutton(options_frame, text="Show Axis", variable=self.show_axis,
                       command=self.update_display).pack(side=tk.LEFT)
        
        ttk.Checkbutton(options_frame, text="Show Crosshair", variable=self.show_crosshair,
                       command=self.update_display).pack(side=tk.LEFT, padx=(10, 0))

    # NEW: Separate section for verification controls (unnumbered)
    def setup_verification_section(self, parent):
        vf = ttk.LabelFrame(parent, text="Verification (Optional)", padding=10)
        vf.pack(fill=tk.X, pady=(0, 10))

        ttk.Checkbutton(
            vf, text="Enable verification link",
            variable=self.enable_verification,
            command=self.on_toggle_verification
        ).pack(anchor=tk.W)

        row1 = ttk.Frame(vf); row1.pack(fill=tk.X, pady=(6,0))
        ttk.Label(row1, text="UID Column:").pack(side=tk.LEFT)
        self.uid_combo = ttk.Combobox(row1, textvariable=self.uid_column_var, state="readonly")
        self.uid_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(6,0))

        ttk.Button(
            vf, text="Select Verification Position",
            command=self.begin_set_verification_position
        ).pack(fill=tk.X, pady=(8,0))

        # Font size for verification ID
        row2 = ttk.Frame(vf); row2.pack(fill=tk.X, pady=(6,0))
        ttk.Label(row2, text="Verification Font Size:").pack(side=tk.LEFT)
        self.verification_size_spin = tk.Spinbox(row2, from_=8, to=72, width=5, textvariable=self.verification_font_size)
        self.verification_size_spin.pack(side=tk.LEFT, padx=(6,0))

        info = ttk.Label(
            vf,
            text="Adds 'Verification ID: <UID>' and an active link to the PDF:\n"
                 "https://avishkaar.co/s3_virtual/verify.php?uid=<UID>",
            foreground="gray"
        )
        info.pack(anchor=tk.W, pady=(6,0))

    def on_toggle_verification(self):
        # Populate UID dropdown if CSV is loaded
        if self.enable_verification.get() and self.csv_columns:
            self.uid_combo['values'] = self.csv_columns

        # Refresh preview so Verification ID overlay shows/hides immediately
        self.update_display()

    def begin_set_verification_position(self):
        if not self.enable_verification.get():
            messagebox.showwarning("Verification", "Please enable verification first.")
            return
        if not self.uid_column_var.get():
            messagebox.showwarning("Verification", "Please choose the UID column first.")
            return
        self.setting_verification_position = True
        messagebox.showinfo("Verification", "Click on the certificate preview to set the Verification ID position.")

    def setup_email_section(self, parent):
        email_frame = ttk.LabelFrame(parent, text="4. Email Configuration (Optional)", padding=10)
        email_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Enable email checkbox
        ttk.Checkbutton(email_frame, text="Send certificates via email", 
                       variable=self.send_email,
                       command=self.toggle_email_settings).pack(anchor=tk.W)
        
        # Email settings frame
        self.email_settings_frame = ttk.Frame(email_frame)
        self.email_settings_frame.pack(fill=tk.X, pady=(10, 0))
        
        # Apps Script URL
        url_frame = ttk.Frame(self.email_settings_frame)
        url_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(url_frame, text="Google Apps Script URL:").pack(anchor=tk.W)
        self.url_entry = ttk.Entry(url_frame, font=("Arial", 9))
        self.url_entry.pack(fill=tk.X, pady=(2, 0))
        
        # Email column selection
        email_col_frame = ttk.Frame(self.email_settings_frame)
        email_col_frame.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Label(email_col_frame, text="Email Column:").pack(anchor=tk.W)
        self.email_column_var = tk.StringVar()
        self.email_combo = ttk.Combobox(email_col_frame, textvariable=self.email_column_var,
                                       state="readonly")
        self.email_combo.pack(fill=tk.X, pady=(2, 0))
        
        # Email template settings
        template_frame = ttk.Frame(self.email_settings_frame)
        template_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Label(template_frame, text="Email Subject:").pack(anchor=tk.W)
        self.subject_entry = ttk.Entry(template_frame)
        self.subject_entry.pack(fill=tk.X, pady=(2, 5))
        self.subject_entry.insert(0, "Your Certificate")
        
        ttk.Label(template_frame, text="Email Message:").pack(anchor=tk.W)
        self.message_text = tk.Text(template_frame, height=4, wrap=tk.WORD)
        self.message_text.pack(fill=tk.X, pady=(2, 0))
        self.message_text.insert(tk.END, "Dear {Name},\n\nPlease find your certificate attached.\n\nBest regards,\nCertificate Team")
        
        # Setup Apps Script button
        ttk.Button(self.email_settings_frame, text="Setup Google Apps Script", 
                  command=self.show_apps_script_setup).pack(pady=(10, 0))
        
        # Initially hide email settings
        self.toggle_email_settings()

    def setup_global_settings(self, parent):
        settings_frame = ttk.LabelFrame(parent, text="General Settings", padding=8)
        settings_frame.pack(fill=tk.X, pady=(0, 10))

        # Enable Claude Haiku 4.5 for all clients
        ttk.Checkbutton(settings_frame, text='Enable Claude Haiku 4.5 for all clients',
                        variable=self.enable_claude_haiku_4_5).pack(anchor=tk.W)
    
    def toggle_email_settings(self):
        if self.send_email.get():
            self.email_settings_frame.pack(fill=tk.X, pady=(10, 0))
            # Update email column options if CSV is loaded
            if self.csv_columns:
                self.email_combo['values'] = self.csv_columns
                # Auto-select email column
                for col in self.csv_columns:
                    if 'email' in col.lower() or 'mail' in col.lower():
                        self.email_column_var.set(col)
                        break
        else:
            self.email_settings_frame.pack_forget()
    
    def show_apps_script_setup(self):
        setup_window = tk.Toplevel(self.root)
        setup_window.title("Google Apps Script Setup")
        setup_window.geometry("800x600")
        setup_window.transient(self.root)
        setup_window.grab_set()
        
        # Create scrollable text widget
        main_frame = ttk.Frame(setup_window)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        title_label = ttk.Label(main_frame, text="Google Apps Script Setup Guide", 
                               font=("Arial", 14, "bold"))
        title_label.pack(pady=(0, 10))
        
        text_widget = tk.Text(main_frame, wrap=tk.WORD, font=("Arial", 10))
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        setup_text = """STEP 1: Create Google Apps Script
1. Go to https://script.google.com
2. Click "New Project"
3. Replace the default code with the provided script below
4. Save the project with a meaningful name

STEP 2: Apps Script Code
Copy and paste this code into your Google Apps Script editor:

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const { to, subject, message, attachmentData, attachmentName } = data;
    
    // Decode base64 attachment
    const blob = Utilities.newBlob(
      Utilities.base64Decode(attachmentData), 
      'application/pdf', 
      attachmentName
    );
    
    // Send email with attachment
    GmailApp.sendEmail(to, subject, message, {
      attachments: [blob]
    });
    
    return ContentService
      .createTextOutput(JSON.stringify({success: true, message: 'Email sent successfully'}))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({success: false, error: error.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

STEP 3: Deploy the Script
1. Click "Deploy" → "New Deployment"
2. Click the gear icon next to "Type" and select "Web app"
3. Set "Execute as" to "Me"
4. Set "Who has access" to "Anyone" (this is required for external requests)
5. Click "Deploy"
6. Copy the deployment URL and paste it in the URL field below

STEP 4: Authorize Permissions
1. You'll be prompted to authorize permissions
2. Click "Review permissions"
3. Choose your Google account
4. Click "Advanced" → "Go to [your project name] (unsafe)"
5. Click "Allow"

IMPORTANT NOTES:
- The Apps Script will send emails from your Gmail account
- Make sure you have sufficient Gmail sending limits
- Test with a small batch first
- The script requires Gmail API permissions to send emails

STEP 5: Test Setup
Use the "Test Email Setup" button below to verify your configuration works correctly.
"""
        
        text_widget.insert(tk.END, setup_text)
        text_widget.config(state=tk.DISABLED)
        
        text_widget.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Test button
        button_frame = ttk.Frame(setup_window)
        button_frame.pack(fill=tk.X, padx=20, pady=(10, 20))
        
        ttk.Button(button_frame, text="Test Email Setup", 
                  command=self.test_email_setup).pack(side=tk.LEFT)
        
        ttk.Button(button_frame, text="Close", 
                  command=setup_window.destroy).pack(side=tk.RIGHT)
    
    def test_email_setup(self):
        url = self.url_entry.get().strip()
        if not url:
            messagebox.showerror("Error", "Please enter the Google Apps Script URL first")
            return
        
        test_email = simpledialog.askstring("Test Email", "Enter test email address:")
        if not test_email:
            return
        
        try:
            # Create a simple test PDF (just text)
            from PIL import Image, ImageDraw
            test_img = Image.new('RGB', (400, 300), 'white')
            draw = ImageDraw.Draw(test_img)
            draw.text((50, 150), "TEST CERTIFICATE", fill='black')
            
            # Convert to PDF bytes
            import io
            pdf_buffer = io.BytesIO()
            test_img.save(pdf_buffer, format='PDF')
            pdf_data = pdf_buffer.getvalue()
            
            # Encode to base64
            attachment_data = base64.b64encode(pdf_data).decode('utf-8')
            
            # Prepare email data
            email_data = {
                'to': test_email,
                'subject': 'Test Certificate Email',
                'message': 'This is a test email from the Certificate Generator.',
                'attachmentData': attachment_data,
                'attachmentName': 'test_certificate.pdf'
            }
            
            # Send request
            response = requests.post(url, json=email_data, timeout=30)
            
            if response.status_code == 200:
                result = response.json()
                if result.get('success'):
                    messagebox.showinfo("Success", "Test email sent successfully!")
                else:
                    messagebox.showerror("Error", f"Email sending failed: {result.get('error', 'Unknown error')}")
            else:
                messagebox.showerror("Error", f"HTTP Error {response.status_code}: {response.text}")
                
        except Exception as e:
            messagebox.showerror("Error", f"Test failed: {str(e)}")
    
    def setup_fields_section(self, parent):
        fields_frame = ttk.LabelFrame(parent, text="3. Configure Text Fields", padding=10)
        fields_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Add field button
        add_frame = ttk.Frame(fields_frame)
        add_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(add_frame, text="+ Add Text Field", 
                  command=self.add_text_field).pack(side=tk.LEFT)
        
        ttk.Button(add_frame, text="Clear All Fields", 
                  command=self.clear_all_fields).pack(side=tk.RIGHT)
        
        # Fields list frame
        self.fields_list_frame = ttk.Frame(fields_frame)
        self.fields_list_frame.pack(fill=tk.X)
        
        # Instructions
        instructions = ttk.Label(fields_frame, 
                               text="Load CSV first, then add fields and position them on template",
                               foreground="gray", font=("Arial", 9))
        instructions.pack(pady=(10, 0))
    
    def setup_csv_section(self, parent):
        csv_frame = ttk.LabelFrame(parent, text="2. Select Data CSV", padding=10)
        csv_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(csv_frame, text="Browse CSV File", 
                  command=self.browse_csv).pack(fill=tk.X)
        
        self.csv_label = ttk.Label(csv_frame, text="No CSV selected", 
                                  foreground="gray")
        self.csv_label.pack(pady=(5, 0))
        
        # Preview row selector
        preview_frame = ttk.Frame(csv_frame)
        preview_frame.pack(fill=tk.X, pady=(6, 0))
        ttk.Label(preview_frame, text="Preview Row:").pack(side=tk.LEFT)
        self.preview_row_spin = tk.Spinbox(preview_frame, from_=1, to=1, width=6, textvariable=self.preview_row_var, command=self.on_preview_row_change)
        self.preview_row_spin.pack(side=tk.LEFT, padx=(6,0))
        
        # CSV columns preview
        self.columns_frame = ttk.Frame(csv_frame)
        self.columns_frame.pack(fill=tk.X, pady=(10, 0))
    
    def setup_generation_section(self, parent):
        generate_frame = ttk.LabelFrame(parent, text="5. Generate Certificates", padding=10)
        generate_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.generate_button = ttk.Button(generate_frame, text="Generate All Certificates", 
                                         command=self.generate_certificates, state=tk.DISABLED)
        self.generate_button.pack(fill=tk.X)
        
        # Progress bar
        self.progress = ttk.Progressbar(generate_frame, mode='determinate')
        self.progress.pack(fill=tk.X, pady=(10, 0))
        
        # Status label
        self.status_label = ttk.Label(generate_frame, text="Ready", foreground="blue")
        self.status_label.pack(pady=(5, 0))
    
    def setup_preview(self, parent):
        preview_frame = ttk.LabelFrame(parent, text="Preview & Positioning", padding=10)
        preview_frame.pack(fill=tk.BOTH, expand=True)
        
        # Coordinates display
        coords_frame = ttk.Frame(preview_frame)
        coords_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.coords_label = ttk.Label(coords_frame, text="Mouse Position: (0, 0)", 
                                     font=("Courier", 10))
        self.coords_label.pack(side=tk.LEFT)
        
        self.current_field_label = ttk.Label(coords_frame, text="No field selected", 
                                           foreground="gray")
        self.current_field_label.pack(side=tk.RIGHT)
        
        # Canvas for image preview
        canvas_frame = ttk.Frame(preview_frame)
        canvas_frame.pack(fill=tk.BOTH, expand=True)
        
        self.canvas = tk.Canvas(canvas_frame, bg="white", cursor="crosshair")
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        h_scrollbar = ttk.Scrollbar(canvas_frame, orient="horizontal", command=self.canvas.xview)
        
        self.canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Pack scrollbars and canvas
        v_scrollbar.pack(side="right", fill="y")
        h_scrollbar.pack(side="bottom", fill="x")
        self.canvas.pack(side="left", fill="both", expand=True)
        
        # Bind events
        self.canvas.bind("<Button-1>", self.on_canvas_click)
        self.canvas.bind("<Motion>", self.on_mouse_move)
        self.canvas.bind("<MouseWheel>", self.on_mouse_wheel)
    
    def add_text_field(self):
        if not self.csv_columns:
            messagebox.showwarning("CSV Required", 
                                 "Please load a CSV file first to enable column selection for text fields.")
            return
            
        field_id = len(self.text_fields)
        field_data = {
            'id': field_id,
            'type': 'Name',
            'csv_column': '',
            'font_path': None,
            'font_size': 50,
            'font_color': '#000000',
            'position': None,
            'sample_text': 'SAMPLE TEXT'
        }
        
        self.text_fields.append(field_data)
        self.create_field_widget(field_data)
        self.current_field_index = field_id
        self.update_current_field_display()
        self.check_generate_ready()
    
    def create_field_widget(self, field_data):
        field_frame = ttk.LabelFrame(self.fields_list_frame, 
                                   text=f"Field {field_data['id'] + 1}: {field_data['type']}", 
                                   padding=5)
        field_frame.pack(fill=tk.X, pady=(0, 5))
        
        # Store reference to frame
        field_data['frame'] = field_frame
        
        # Field type selection
        type_frame = ttk.Frame(field_frame)
        type_frame.pack(fill=tk.X)
        
        ttk.Label(type_frame, text="Type:").pack(side=tk.LEFT)
        
        type_var = tk.StringVar(value=field_data['type'])
        field_data['type_var'] = type_var
        
        type_combo = ttk.Combobox(type_frame, textvariable=type_var, 
                                 values=self.field_types, width=12, state="readonly")
        type_combo.pack(side=tk.LEFT, padx=(5, 0))
        type_combo.bind('<<ComboboxSelected>>', 
                       lambda e, fid=field_data['id']: self.on_field_type_change(fid))
        
        # CSV column selection
        csv_frame = ttk.Frame(field_frame)
        csv_frame.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Label(csv_frame, text="CSV Column:").pack(side=tk.LEFT)
        
        csv_var = tk.StringVar(value=field_data['csv_column'])
        field_data['csv_var'] = csv_var
        
        csv_combo = ttk.Combobox(csv_frame, textvariable=csv_var, width=15, 
                                values=self.csv_columns, state="readonly")
        csv_combo.pack(side=tk.LEFT, padx=(5, 0))
        csv_combo.bind('<<ComboboxSelected>>', 
                      lambda e, fid=field_data['id']: self.on_csv_column_change(fid))
        field_data['csv_combo'] = csv_combo
        
        # Auto-select matching column if CSV is loaded
        if self.csv_columns:
            self.auto_select_column(field_data)
        
        # Font settings
        font_frame = ttk.Frame(field_frame)
        font_frame.pack(fill=tk.X, pady=(5, 0))
        
        # Font size
        ttk.Label(font_frame, text="Size:").pack(side=tk.LEFT)
        size_var = tk.StringVar(value=str(field_data['font_size']))
        field_data['size_var'] = size_var
        
        size_spin = ttk.Spinbox(font_frame, from_=10, to=300, width=6,
                               textvariable=size_var, 
                               command=lambda fid=field_data['id']: self.on_field_change(fid))
        size_spin.pack(side=tk.LEFT, padx=(2, 10))
        
        # Font color
        ttk.Label(font_frame, text="Color:").pack(side=tk.LEFT)
        color_button = tk.Button(font_frame, text="  ", bg=field_data['font_color'],
                               width=2, command=lambda fid=field_data['id']: self.choose_field_color(fid))
        color_button.pack(side=tk.LEFT, padx=(2, 10))
        field_data['color_button'] = color_button
        
        # Control buttons
        btn_frame = ttk.Frame(field_frame)
        btn_frame.pack(fill=tk.X, pady=(5, 0))
        
        select_btn = ttk.Button(btn_frame, text="Select for Positioning", 
                               command=lambda fid=field_data['id']: self.select_field(fid))
        select_btn.pack(side=tk.LEFT)
        
        font_btn = ttk.Button(btn_frame, text="Font", 
                             command=lambda fid=field_data['id']: self.browse_field_font(fid))
        font_btn.pack(side=tk.LEFT, padx=(5, 0))
        
        delete_btn = ttk.Button(btn_frame, text="Delete", 
                               command=lambda fid=field_data['id']: self.delete_field(fid))
        delete_btn.pack(side=tk.RIGHT)
        
        # Link button
        link_btn = ttk.Button(btn_frame, text="Set Link", 
                     command=lambda fid=field_data['id']: self.set_field_link(fid))
        link_btn.pack(side=tk.LEFT, padx=(5,0))
        
        # Position display
        pos_label = ttk.Label(field_frame, text="Position: Not set", foreground="gray")
        pos_label.pack(pady=(5, 0))
        field_data['pos_label'] = pos_label
        # Link display
        link_label = ttk.Label(field_frame, text="Link: None", foreground="gray")
        link_label.pack(pady=(2, 0))
        field_data['link_label'] = link_label
        # Store link URL (static)
        field_data.setdefault('link_url', None)
    
    def select_field(self, field_id):
        self.current_field_index = field_id
        self.update_current_field_display()

    def set_field_link(self, field_id):
        """Prompt user to set a static URL for the field. Display text comes from CSV column (field value)."""
        field = self.text_fields[field_id]
        # Ask for static URL (no placeholders)
        url = simpledialog.askstring("Set Link URL", "Enter the URL to open when clicking this field:")
        if url is None:
            return
        url = url.strip()
        if url == "":
            # Clear link
            field['link_url'] = None
            field['link_label'].config(text="Link: None", foreground="gray")
            self.update_display()
            return

        # Store the static URL
        field['link_url'] = url
        field['link_label'].config(text=f"Link: {url}", foreground="blue")
        self.update_display()
    
    def update_current_field_display(self):
        if self.current_field_index < len(self.text_fields):
            field = self.text_fields[self.current_field_index]
            self.current_field_label.config(
                text=f"Selected: Field {field['id'] + 1} ({field['type']})",
                foreground="blue"
            )
        else:
            self.current_field_label.config(text="No field selected", foreground="gray")
    
    def on_field_type_change(self, field_id):
        field = self.text_fields[field_id]
        new_type = field['type_var'].get()
        field['type'] = new_type
        
        # Update frame title
        field['frame'].config(text=f"Field {field_id + 1}: {new_type}")
        
        # Set default sample text based on type
        sample_texts = {
            'Name': 'JOHN DOE',
            'Roll Number': '2023001',
            'Branch': 'COMPUTER SCIENCE',
            'Course': 'B.TECH',
            'Date': '2024-01-15',
            'Grade': 'A+',
            'Score': '95%',
            'Department': 'CSE',
            'Year': '2024',
            'Email': 'john.doe@example.com'
        }
        field['sample_text'] = sample_texts.get(new_type, 'SAMPLE TEXT')
        
        self.update_display()
    
    def auto_select_column(self, field):
        """Auto-select the best matching CSV column for a field type"""
        if not self.csv_columns:
            return
        
        field_type = field['type'].lower()
        
        # Define mapping of field types to possible column names
        column_mappings = {
            'name': ['name', 'full_name', 'participant_name', 'student_name', 'full name', 'participant full name'],
            'roll number': ['roll_number', 'roll_no', 'student_id', 'id', 'roll number', 'roll no'],
            'branch': ['branch', 'department', 'stream', 'course'],
            'course': ['course', 'program', 'degree'],
            'date': ['date', 'issue_date', 'completion_date', 'issue date'],
            'grade': ['grade', 'result', 'class'],
            'score': ['score', 'marks', 'percentage', 'points'],
            'department': ['department', 'dept', 'branch'],
            'year': ['year', 'academic_year', 'batch', 'academic year'],
            'email': ['email', 'email_address', 'mail', 'e-mail', 'email address']
        }
        
        # Find best match
        possible_names = column_mappings.get(field_type, [])
        
        for possible_name in possible_names:
            for column in self.csv_columns:
                if possible_name.lower() in column.lower():
                    field['csv_column'] = column
                    field['csv_var'].set(column)
                    return
    
    def on_csv_column_change(self, field_id):
        """Handle CSV column selection change"""
        field = self.text_fields[field_id]
        field['csv_column'] = field['csv_var'].get()
        self.check_generate_ready()
    
    def on_field_change(self, field_id):
        field = self.text_fields[field_id]
        try:
            field['font_size'] = int(field['size_var'].get())
        except ValueError:
            field['font_size'] = 50
        
        self.update_display()

    def on_preview_row_change(self):
        """Handle preview row spinbox change - update display to show new row values"""
        self.update_display()
    
    def choose_field_color(self, field_id):
        field = self.text_fields[field_id]
        color = colorchooser.askcolor(color=field['font_color'])[1]
        if color:
            field['font_color'] = color
            field['color_button'].config(bg=color)
            self.update_display()
    
    def browse_field_font(self, field_id):
        field = self.text_fields[field_id]
        file_path = filedialog.askopenfilename(
            title="Select Font File",
            filetypes=[("Font files", "*.ttf *.otf")]
        )
        if file_path:
            field['font_path'] = file_path
            self.update_display()
    
    def delete_field(self, field_id):
        # Remove field data
        self.text_fields = [f for f in self.text_fields if f['id'] != field_id]
        
        # Destroy widget
        for field in self.text_fields:
            if field['id'] == field_id:
                field['frame'].destroy()
                break
        
        # Refresh fields list
        self.refresh_fields_display()
        self.update_display()
        self.check_generate_ready()
    
    def clear_all_fields(self):
        for field in self.text_fields:
            field['frame'].destroy()
        self.text_fields = []
        self.current_field_index = 0
        self.update_current_field_display()
        self.update_display()
        self.check_generate_ready()
    
    def refresh_fields_display(self):
        # Destroy all field widgets
        for widget in self.fields_list_frame.winfo_children():
            widget.destroy()
        
        # Recreate widgets
        for field in self.text_fields:
            self.create_field_widget(field)
    
    def browse_template(self):
        file_path = filedialog.askopenfilename(
            title="Select Certificate Template",
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp *.gif *.tiff")]
        )
        
        if file_path:
            try:
                self.template_image = Image.open(file_path)
                self.template_path = file_path
                self.template_label.config(text=os.path.basename(file_path), foreground="black")
                self.display_template()
                self.check_generate_ready()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load template: {str(e)}")
    
    def browse_csv(self):
        file_path = filedialog.askopenfilename(
            title="Select Data CSV File",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")]
        )
        
        if file_path:
            try:
                if file_path.endswith('.csv'):
                    df = pd.read_csv(file_path)
                else:
                    df = pd.read_excel(file_path)
                
                self.csv_path = file_path
                self.csv_label.config(text=f"{os.path.basename(file_path)} ({len(df)} rows)", 
                                     foreground="black")
                
                # Update CSV columns list
                self.csv_columns = list(df.columns)
                # Keep DataFrame in memory for preview and link insertion
                self.df_current = df
                # Update preview spinbox max
                try:
                    total = len(df)
                    self.preview_row_var.set(1)
                    self.preview_row_spin.config(to=total)
                except Exception:
                    pass
                
                # Update CSV column options for all existing fields
                for field in self.text_fields:
                    field['csv_combo']['values'] = self.csv_columns
                    # Auto-select matching column if available
                    self.auto_select_column(field)
                
                # Update email column options
                if self.send_email.get():
                    self.email_combo['values'] = self.csv_columns
                    # Auto-select email column
                    for col in self.csv_columns:
                        if 'email' in col.lower() or 'mail' in col.lower():
                            self.email_column_var.set(col)
                            break

                # NEW: populate UID dropdown when verification is enabled
                if hasattr(self, 'uid_combo') and self.enable_verification.get():
                    self.uid_combo['values'] = self.csv_columns
                
                # Show available columns
                self.show_csv_columns(self.csv_columns)
                self.check_generate_ready()
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load CSV: {str(e)}")
    
    def show_csv_columns(self, columns):
        # Clear previous columns display
        for widget in self.columns_frame.winfo_children():
            widget.destroy()
        
        ttk.Label(self.columns_frame, text="Available columns:", 
                 font=("Arial", 9, "bold")).pack(anchor=tk.W)
        
        # Show columns in a scrollable text widget
        text_widget = tk.Text(self.columns_frame, height=4, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(self.columns_frame, orient="vertical", command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        text_widget.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        text_widget.insert(tk.END, ", ".join(columns))
        text_widget.config(state=tk.DISABLED)
    
    def display_template(self):
        if not self.template_image:
            return
        
        # Get canvas dimensions
        self.canvas.update_idletasks()
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        
        if canvas_width <= 1 or canvas_height <= 1:
            self.root.after(100, self.display_template)
            return
        
        # Calculate scale
        img_width, img_height = self.template_image.size
        scale_x = canvas_width / img_width
        scale_y = canvas_height / img_height
        self.canvas_scale = min(scale_x, scale_y, 1.0)
        
        # Update canvas scroll region
        display_width = int(img_width * self.canvas_scale)
        display_height = int(img_height * self.canvas_scale)
        self.canvas.configure(scrollregion=(0, 0, display_width, display_height))
        
        self.update_display()
    
    def update_display(self):
        if not self.template_image:
            return
        
        # Create display image
        img_width, img_height = self.template_image.size
        display_width = int(img_width * self.canvas_scale)
        display_height = int(img_height * self.canvas_scale)
        
        # Create preview with all text fields
        preview = self.template_image.copy()
        draw = ImageDraw.Draw(preview)
        
        # Draw all positioned text fields
        for field in self.text_fields:
            if field['position']:
                self.draw_field_on_preview(draw, field)

        # NEW: draw verification ID overlay if enabled and positioned
        if self.enable_verification.get() and self.verification_position and self.df_current is not None:
            uid_col = self.uid_column_var.get()
            if uid_col in self.df_current.columns and len(self.df_current) > 0:
                try:
                    pr = int(self.preview_row_var.get())
                    pr = max(1, min(pr, len(self.df_current)))
                except Exception:
                    pr = 1
                uid_val = str(self.df_current.iloc[pr-1][uid_col])
                if uid_val and uid_val.lower() != "nan":
                    # Draw at original coords (centered baseline like other fields)
                    text = f"Verification ID: {uid_val}"
                    try:
                        vsize = max(8, int(self.verification_font_size.get()))
                        try:
                            font = ImageFont.truetype("arial.ttf", vsize)
                        except Exception:
                            font = ImageFont.load_default()
                        bbox = draw.textbbox((0, 0), text, font=font)
                        tw = bbox[2] - bbox[0]
                        th = bbox[3] - bbox[1]
                    except Exception:
                        tw, th = draw.textsize(text, font=font)
                    x, y = self.verification_position
                    x_draw = x - tw // 2
                    y_draw = y - th // 2
                    draw.text((x_draw, y_draw), text, fill="#0000EE", font=font)
        
        # Resize for display
        display_preview = preview.resize((display_width, display_height), Image.Resampling.LANCZOS)
        
        # Convert to PhotoImage
        self.photo = ImageTk.PhotoImage(display_preview)
        
        # Clear canvas and draw image
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor=tk.NW, image=self.photo, tags="template")
        
        # Draw axis and guides if enabled
        if self.show_axis.get():
            self.draw_axis()
        
        # Draw field markers
        self.draw_field_markers()

        # Ensure link items get pointer cursor
        for i, field in enumerate(self.text_fields):
            if field.get('position') and field.get('link_url'):
                tag_name = f"link_{field['id']}"
                try:
                    self.canvas.tag_bind(tag_name, "<Enter>", lambda e: self.canvas.config(cursor="hand2"))
                    self.canvas.tag_bind(tag_name, "<Leave>", lambda e: self.canvas.config(cursor="crosshair"))
                except Exception:
                    pass
    
    def draw_field_on_preview(self, draw, field):
        # Load font
        try:
            if field['font_path']:
                font = ImageFont.truetype(field['font_path'], field['font_size'])
            else:
                font = ImageFont.load_default()
        except:
            font = ImageFont.load_default()
        
        # Get text size
        text = field['sample_text']
        try:
            bbox = draw.textbbox((0, 0), text, font=font)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
        except AttributeError:
            text_width, text_height = draw.textsize(text, font=font)
        
        # Calculate position (center text at clicked position)
        x = field['position'][0] - text_width // 2
        y = field['position'][1] - text_height // 2
        
        # Draw text
        draw.text((x, y), text, fill=field['font_color'], font=font)
    
    def draw_axis(self):
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        
        # Draw grid lines every 50 pixels (in display coordinates)
        grid_spacing = 50
        
        # Vertical lines
        for x in range(0, int(self.template_image.size[0] * self.canvas_scale), grid_spacing):
            self.canvas.create_line(x, 0, x, canvas_height, fill="lightgray", tags="axis")
            if x % (grid_spacing * 2) == 0:  # Labels every 100 pixels
                orig_x = int(x / self.canvas_scale)
                self.canvas.create_text(x + 2, 10, text=str(orig_x), anchor="nw", 
                                      fill="gray", font=("Arial", 8), tags="axis")
        
        # Horizontal lines
        for y in range(0, int(self.template_image.size[1] * self.canvas_scale), grid_spacing):
            self.canvas.create_line(0, y, canvas_width, y, fill="lightgray", tags="axis")
            if y % (grid_spacing * 2) == 0:  # Labels every 100 pixels
                orig_y = int(y / self.canvas_scale)
                self.canvas.create_text(10, y + 2, text=str(orig_y), anchor="nw", 
                                      fill="gray", font=("Arial", 8), tags="axis")
    
    def draw_field_markers(self):
        for i, field in enumerate(self.text_fields):
            if field['position']:
                # Convert to display coordinates
                display_x = field['position'][0] * self.canvas_scale
                display_y = field['position'][1] * self.canvas_scale
                
                # Draw marker
                color = "red" if i == self.current_field_index else "blue"
                self.canvas.create_oval(display_x - 5, display_y - 5, 
                                      display_x + 5, display_y + 5,
                                      fill=color, outline="white", width=2, tags="marker")
                
                # Draw field number
                self.canvas.create_text(display_x + 10, display_y - 10, 
                                      text=f"F{field['id'] + 1}", 
                                      fill=color, font=("Arial", 10, "bold"), tags="marker")
                
                # If field has link, show field value as clickable link
                if field.get('link_url') and self.df_current is not None:
                    col = field['csv_column']
                    try:
                        if col in self.df_current.columns and len(self.df_current) > 0:
                            # Use selected preview row (1-based)
                            try:
                                pr = int(self.preview_row_var.get())
                                pr = max(1, min(pr, len(self.df_current)))
                            except Exception:
                                pr = 1
                            link_display_text = str(self.df_current.iloc[pr-1][col])
                        else:
                            link_display_text = field.get('sample_text', 'Link')
                    except Exception:
                        link_display_text = field.get('sample_text', 'Link')
                    
                    tag_name = f"link_{field['id']}"
                    self.canvas.create_text(display_x + 10, display_y + 10,
                                           text=link_display_text, fill="blue", font=("Arial", 10, "underline"),
                                           tags=("link", tag_name))
                    
                    # Clicking opens the static URL
                    def _on_click(event, url=field['link_url']):
                        webbrowser.open(url)
                    
                    self.canvas.tag_bind(tag_name, "<Button-1>", _on_click)
    
    def on_canvas_click(self, event):
        if not self.template_image:
            return

        # NEW: If we are in verification position mode, set it and return
        if self.enable_verification.get() and self.setting_verification_position:
            orig_x = int(event.x / self.canvas_scale)
            orig_y = int(event.y / self.canvas_scale)
            self.verification_position = (orig_x, orig_y)
            self.setting_verification_position = False
            messagebox.showinfo("Verification", f"Verification ID position set at ({orig_x}, {orig_y})")
            self.update_display()
            return
        
        if self.current_field_index >= len(self.text_fields):
            return
        
        # Convert display coordinates to original image coordinates
        orig_x = int(event.x / self.canvas_scale)
        orig_y = int(event.y / self.canvas_scale)
        
        # Check bounds
        if (0 <= orig_x <= self.template_image.size[0] and 
            0 <= orig_y <= self.template_image.size[1]):
            
            field = self.text_fields[self.current_field_index]
            field['position'] = (orig_x, orig_y)
            
            # Update position label
            field['pos_label'].config(text=f"Position: ({orig_x}, {orig_y})", 
                                    foreground="green")
            
            self.update_display()
            self.check_generate_ready()
    
    def on_mouse_move(self, event):
        if not self.template_image:
            return
        
        # Convert to original coordinates
        orig_x = int(event.x / self.canvas_scale)
        orig_y = int(event.y / self.canvas_scale)
        
        self.coords_label.config(text=f"Mouse Position: ({orig_x}, {orig_y})")
        
        # Draw crosshair if enabled
        if self.show_crosshair.get():
            self.canvas.delete("crosshair")
            canvas_width = self.canvas.winfo_width()
            canvas_height = self.canvas.winfo_height()
            
            # Vertical line
            self.canvas.create_line(event.x, 0, event.x, canvas_height, 
                                  fill="red", width=1, tags="crosshair")
            # Horizontal line
            self.canvas.create_line(0, event.y, canvas_width, event.y, 
                                  fill="red", width=1, tags="crosshair")
    
    def on_mouse_wheel(self, event):
        # Zoom functionality
        if event.state & 0x4:  # Ctrl key pressed
            zoom_factor = 1.1 if event.delta > 0 else 0.9
            self.canvas_scale *= zoom_factor
            self.canvas_scale = max(0.1, min(3.0, self.canvas_scale))  # Limit zoom
            self.update_display()
        else:
            # Normal scrolling
            self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    def check_generate_ready(self):
        ready = (self.template_image and 
                self.csv_path and 
                len(self.text_fields) > 0 and
                all(field['position'] for field in self.text_fields) and
                all(field['csv_column'] for field in self.text_fields))
        
        # If email is enabled, check email configuration
        if self.send_email.get():
            email_ready = (self.url_entry.get().strip() and 
                          self.email_column_var.get())
            ready = ready and email_ready
        
        self.generate_button.config(state=tk.NORMAL if ready else tk.DISABLED)
    
    def ask_output_folder(self):
        """Ask user for output folder name"""
        folder_name = simpledialog.askstring(
            "Output Folder", 
            "Enter folder name for certificates:",
            initialvalue=f"certificates_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        )
        return folder_name
    
    def _add_pdf_links(self, input_pdf, output_pdf, links_to_add, image_size):
        """Add clickable hyperlinks to PDF using ReportLab overlay.
        
        links_to_add: list of dicts with 'position', 'text', 'url'
        image_size: (width, height) of the certificate image
        """
        try:
            # Create overlay PDF with clickable rectangles
            packet = io.BytesIO()
            c = canvas.Canvas(packet, pagesize=(image_size[0], image_size[1]))
            
            # Add invisible clickable rectangles for each link
            for link_info in links_to_add:
                x, y = link_info['position']
                url = link_info['url']
                # Determine clickable area based on font size (if provided)
                fsize = int(link_info.get('font_size', 12))
                rect_width = max(80, fsize * 6)
                rect_height = max(12, int(fsize * 1.6))

                # ReportLab origin is bottom-left while image coords origin is top-left.
                # Convert Y coordinate accordingly.
                page_w, page_h = image_size[0], image_size[1]
                # link_info position is (x,y) in image coords (top-left origin)
                # compute rect in PDF coordinates (bottom-left origin)
                x1 = x - rect_width // 2
                x2 = x + rect_width // 2
                # top of rect in image coords = y - rect_height//2
                y_top = y - rect_height // 2
                y_bottom = y + rect_height // 2
                # convert to PDF coords
                py1 = page_h - y_bottom
                py2 = page_h - y_top
                rect = (float(x1), float(py1), float(x2), float(py2))

                # Add the clickable link rectangle
                c.linkURL(url, rect, relative=0, thickness=0)
            
            c.showPage()
            c.save()
            packet.seek(0)
            
            # Read the original PDF
            reader = PdfReader(input_pdf)
            writer = PdfWriter()
            
            # Get first page and merge with overlay
            if len(reader.pages) > 0:
                first_page = reader.pages[0]
                overlay_pdf = PdfReader(packet)
                overlay_page = overlay_pdf.pages[0]
                first_page.merge_page(overlay_page)
            
            # Add all pages to writer
            for page in reader.pages:
                writer.add_page(page)
            
            # Write output PDF
            with open(output_pdf, 'wb') as f:
                writer.write(f)
            
            # Clean up temp file
            if os.path.exists(input_pdf):
                os.remove(input_pdf)
                
        except Exception as e:
            print(f"Error adding PDF links: {e}")
            # Fallback: just copy input to output
            if os.path.exists(input_pdf):
                os.rename(input_pdf, output_pdf)
    
    def send_email_with_certificate(self, email, subject, message, pdf_path, recipient_name=""):
        """Send email with certificate attachment using Google Apps Script"""
        try:
            url = self.url_entry.get().strip()
            if not url:
                return False, "Apps Script URL not configured"

            # Read PDF file and encode to base64
            with open(pdf_path, 'rb') as f:
                pdf_data = f.read()

            attachment_data = base64.b64encode(pdf_data).decode('utf-8')
            attachment_name = os.path.basename(pdf_path)

            # Replace placeholders in message
            personalized_message = message.replace("{Name}", recipient_name)

            # Prepare email data
            email_data = {
                'to': email,
                'subject': subject,
                'message': personalized_message,
                'attachmentData': attachment_data,
                'attachmentName': attachment_name
            }

            # Send request with timeout
            response = requests.post(url, json=email_data, timeout=30)

            # Save full response for diagnostics
            try:
                resp_text = response.text
            except Exception:
                resp_text = '<no response body>'

            log_path = os.path.join(os.getcwd(), 'last_apps_script_response.txt')
            try:
                with open(log_path, 'w', encoding='utf-8') as lf:
                    lf.write(f'Status: {response.status_code}\n')
                    lf.write('Headers:\n')
                    for k, v in response.headers.items():
                        lf.write(f'{k}: {v}\n')
                    lf.write('\nBody:\n')
                    lf.write(resp_text)
            except Exception:
                # ignore logging errors
                log_path = None

            if response.status_code == 200:
                # Try to parse JSON, fallback to raw text
                try:
                    result = response.json()
                except Exception:
                    result = None

                if isinstance(result, dict) and result.get('success'):
                    return True, 'Email sent successfully'
                else:
                    err_msg = 'Apps Script returned error'
                    if result and isinstance(result, dict):
                        err_msg = result.get('error', err_msg)
                    else:
                        # Use raw text if JSON not returned
                        err_msg = resp_text[:1000]

                    if log_path:
                        messagebox.showerror('Email Error', f"Failed to send email. Server response saved to:\n{log_path}")
                    else:
                        messagebox.showerror('Email Error', f"Failed to send email. Server response:\n{err_msg}")

                    return False, err_msg
            else:
                err_msg = f'HTTP Error {response.status_code}'
                if log_path:
                    messagebox.showerror('Email HTTP Error', f"{err_msg}. Full response saved to:\n{log_path}")
                else:
                    messagebox.showerror('Email HTTP Error', err_msg)
                return False, f"{err_msg}: {resp_text}"

        except Exception as e:
            return False, str(e)
    
    def generate_certificates(self):
        if not self.text_fields:
            messagebox.showerror("Error", "Please add at least one text field")
            return
        
        # Ask for output folder name
        output_folder = self.ask_output_folder()
        if not output_folder:
            return
        
        try:
            # Load CSV data
            if self.csv_path.endswith('.csv'):
                df = pd.read_csv(self.csv_path)
            else:
                df = pd.read_excel(self.csv_path)
            
            # Validate CSV columns
            missing_columns = []
            for field in self.text_fields:
                if field['csv_column'] not in df.columns:
                    missing_columns.append(field['csv_column'])
            
            # Check email column if email is enabled
            if self.send_email.get():
                email_column = self.email_column_var.get()
                if email_column not in df.columns:
                    missing_columns.append(f"{email_column} (email)")
            
            if missing_columns:
                messagebox.showerror("Error", 
                                   f"Missing CSV columns: {', '.join(missing_columns)}")
                return
            
            # Create output folder
            os.makedirs(output_folder, exist_ok=True)
            
            # Setup progress
            self.progress.config(maximum=len(df))
            self.progress.config(value=0)
            
            generated_files = []
            email_results = {"sent": 0, "failed": 0, "errors": []}
            
            # Get email settings if enabled
            if self.send_email.get():
                email_column = self.email_column_var.get()
                email_subject = self.subject_entry.get()
                email_message = self.message_text.get("1.0", tk.END).strip()
            
            # Generate certificates
            for index, row in df.iterrows():
                # Create certificate
                certificate = self.template_image.copy()
                draw = ImageDraw.Draw(certificate)
                
                # Draw all text fields
                name_parts = []
                recipient_name = ""
                
                for field in self.text_fields:
                    # Get data from CSV
                    field_value = str(row[field['csv_column']]).strip()
                    if pd.isna(row[field['csv_column']]) or not field_value:
                        field_value = "N/A"
                    
                    # Store name for filename and email personalization
                    if field['type'].lower() == 'name':
                        name_parts.append(field_value)
                        recipient_name = field_value
                    
                    # Load font
                    try:
                        if field['font_path']:
                            font = ImageFont.truetype(field['font_path'], field['font_size'])
                        else:
                            font = ImageFont.load_default()
                    except:
                        font = ImageFont.load_default()
                    
                    # Calculate text position (center at clicked position)
                    try:
                        bbox = draw.textbbox((0, 0), field_value, font=font)
                        text_width = bbox[2] - bbox[0]
                        text_height = bbox[3] - bbox[1]
                    except AttributeError:
                        text_width, text_height = draw.textsize(field_value, font=font)
                    
                    text_x = field['position'][0] - text_width // 2
                    text_y = field['position'][1] - text_height // 2
                    
                    # Draw text
                    draw.text((text_x, text_y), field_value, 
                             fill=field['font_color'], font=font)

                # If any fields have links, render them as visible text on certificate (blue)
                for link_field in self.text_fields:
                    if link_field.get('link_url') and link_field.get('position'):
                        try:
                            ffont = ImageFont.truetype(link_field['font_path'], max(12, int(link_field.get('font_size', 12) * 0.6))) if link_field.get('font_path') else ImageFont.load_default()
                        except:
                            ffont = ImageFont.load_default()
                        
                        # Display text is the field value from CSV
                        if link_field['csv_column'] in df.columns:
                            try:
                                link_display_text = str(row[link_field['csv_column']])
                            except Exception:
                                link_display_text = ''
                        else:
                            link_display_text = ''

                        if link_display_text:
                            try:
                                bbox = draw.textbbox((0,0), link_display_text, font=ffont)
                                tw = bbox[2]-bbox[0]
                                th = bbox[3]-bbox[1]
                            except AttributeError:
                                tw, th = draw.textsize(link_display_text, font=ffont)
                            lx = link_field['position'][0] - tw // 2
                            ly = link_field['position'][1] + int(th * 0.8)
                            # Draw in blue to indicate it's a link
                            draw.text((lx, ly), link_display_text, fill="#0000EE", font=ffont)

                # NEW: Draw verification text and add clickable link to PDF
                added_verification_link = False
                if self.enable_verification.get() and self.verification_position:
                    uid_col = self.uid_column_var.get()
                    if uid_col in df.columns:
                        uid_val = str(row[uid_col]).strip()
                        if uid_val and uid_val.lower() != "nan":
                            # Draw visible text in blue using configurable font size
                            vtext = f"Verification ID: {uid_val}"
                            vsize = max(8, int(self.verification_font_size.get()))
                            try:
                                # Prefer Arial if available, otherwise fallback to default
                                try:
                                    fontv = ImageFont.truetype("arial.ttf", vsize)
                                except Exception:
                                    fontv = ImageFont.load_default()

                                bbox = draw.textbbox((0,0), vtext, font=fontv)
                                tw = bbox[2]-bbox[0]
                                th = bbox[3]-bbox[1]
                            except Exception:
                                tw, th = draw.textsize(vtext, font=fontv)
                            vx = self.verification_position[0] - tw // 2
                            vy = self.verification_position[1] - th // 2
                            draw.text((vx, vy), vtext, fill="#0000EE", font=fontv)
                            added_verification_link = True

                # Create filename
                if name_parts:
                    filename_base = "_".join(name_parts)
                else:
                    # Use first column value if no name field
                    filename_base = str(row.iloc[0])
                
                # Sanitize filename
                sanitized_name = "".join(c for c in filename_base if c.isalnum() or c in (" ", "_")).replace(" ", "_")
                sanitized_name = sanitized_name[:50]  # Limit length
                
                # Save as PDF
                pdf_filename = f"{sanitized_name}_{index+1}.pdf"
                pdf_path = os.path.join(output_folder, pdf_filename)
                
                # Save certificate image as temporary PDF
                temp_pdf = pdf_path.replace(".pdf", "_temp.pdf")
                certificate.convert("RGB").save(temp_pdf)
                
                # Now add clickable hyperlinks using ReportLab
                try:
                    # Collect all links to add
                    links_to_add = []
                    
                    # Check field links
                    for link_field in self.text_fields:
                        if link_field.get('link_url') and link_field.get('position'):
                            if link_field['csv_column'] in df.columns:
                                try:
                                    link_display_text = str(row[link_field['csv_column']])
                                except Exception:
                                    link_display_text = ''
                            else:
                                link_display_text = ''
                            
                            if link_display_text:
                                # Store link info: (position, text, url)
                                links_to_add.append({
                                    'position': link_field['position'],
                                    'text': link_display_text,
                                    'url': link_field['link_url'],
                                    'font_size': int(link_field.get('font_size', 12))
                                })
                    
                    # Check verification link
                    verify_url = None
                    if self.enable_verification.get() and self.verification_position:
                        uid_col = self.uid_column_var.get()
                        if uid_col in df.columns:
                            uid_val = str(row[uid_col]).strip()
                            if uid_val and uid_val.lower() != "nan":
                                verify_url = f"https://avishkaar.co/s3_virtual/verify.php?uid={uid_val}"
                                links_to_add.append({
                                    'position': self.verification_position,
                                    'text': f"Verification ID: {uid_val}",
                                    'url': verify_url,
                                    'font_size': int(self.verification_font_size.get())
                                })
                    
                    # If there are links to add, use ReportLab to overlay them
                    if links_to_add:
                        self._add_pdf_links(temp_pdf, pdf_path, links_to_add, certificate.size)
                    else:
                        # No links, just copy temp to final
                        os.rename(temp_pdf, pdf_path)
                    
                except Exception as e:
                    print(f"Warning: Failed to add hyperlinks: {e}")
                    # Fall back to temp file as final
                    if os.path.exists(temp_pdf):
                        os.rename(temp_pdf, pdf_path)

                generated_files.append(pdf_path)
                
                # Send email if enabled (UNCHANGED)
                if self.send_email.get():
                    recipient_email = str(row[email_column]).strip()
                    if recipient_email and '@' in recipient_email:
                        success, error_msg = self.send_email_with_certificate(
                            recipient_email, email_subject, email_message, 
                            pdf_path, recipient_name
                        )
                        
                        if success:
                            email_results["sent"] += 1
                        else:
                            email_results["failed"] += 1
                            email_results["errors"].append(f"{recipient_name} ({recipient_email}): {error_msg}")
                    else:
                        email_results["failed"] += 1
                        email_results["errors"].append(f"{recipient_name}: Invalid email address")
                
                # Update progress
                self.progress.config(value=index + 1)
                status_text = f"Processing: {sanitized_name}"
                if self.send_email.get():
                    status_text += f" | Emails sent: {email_results['sent']}"
                self.status_label.config(text=status_text)
                self.root.update()
            
            # Show completion message
            success_msg = f"Successfully generated {len(generated_files)} certificates in folder: {output_folder}"
            
            if self.send_email.get():
                success_msg += f"\n\nEmail Results:\n• Sent: {email_results['sent']}\n• Failed: {email_results['failed']}"
                
                if email_results["errors"]:
                    # Show first few errors
                    error_preview = "\n".join(email_results["errors"][:5])
                    if len(email_results["errors"]) > 5:
                        error_preview += f"\n... and {len(email_results['errors']) - 5} more errors"
                    success_msg += f"\n\nFirst few email errors:\n{error_preview}"
            
            self.status_label.config(text="Completed!")
            messagebox.showinfo("Success", success_msg)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate certificates: {str(e)}")
        finally:
            self.progress.config(value=0)

def main():
    root = tk.Tk()
    app = CertificateGenerator(root)
    
    # Handle window resize
    def on_window_resize(event):
        if event.widget == root:
            root.after(100, app.display_template)
    
    root.bind('<Configure>', on_window_resize)
    
    # Bind mouse wheel to root for better scrolling
    def bind_mousewheel(event):
        app.canvas.bind_all("<MouseWheel>", app.on_mouse_wheel)
    
    def unbind_mousewheel(event):
        app.canvas.unbind_all("<MouseWheel>")
    
    app.canvas.bind('<Enter>', bind_mousewheel)
    app.canvas.bind('<Leave>', unbind_mousewheel)
    
    root.mainloop()

if __name__ == "__main__":
    main()
