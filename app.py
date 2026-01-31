
import streamlit as st
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import json
import time
from streamlit_quill import st_quill

# Page Configuration
st.set_page_config(
    page_title="µLearn Mailer",
    page_icon="📨",
    layout="wide"
)

# Custom Styling
st.markdown("""
<style>
    /* MuLearn Branding Colors */
    :root {
        --primary-color: #FC4F4F; /* MuLearn Red/Orange */
        --text-color: #333333;
        --bg-color: #FFFFFF;
    }
    
    .main-header {
        font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
        font-weight: 700;
        font-size: 3rem;
        color: #333333;
        text-align: center;
        margin-bottom: 0.5rem;
    }
    
    .sub-header {
        font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
        font-weight: 400;
        font-size: 1.2rem;
        color: #666666;
        text-align: center;
        margin-bottom: 3rem;
    }
    
    /* Minimalist Button Styling */
    .stButton > button {
        background-color: white;
        color: #FC4F4F;
        border: 1px solid #FC4F4F;
        border-radius: 5px;
        transition: all 0.3s ease;
    }
    
    .stButton > button:hover {
        background-color: #FC4F4F;
        color: white;
        border: 1px solid #FC4F4F;
    }
    
    /* Clean inputs */
    .stTextInput > div > div > input {
        border-radius: 5px;
    }
    
    /* Mobile Responsiveness */
    @media (max-width: 640px) {
        .main-header {
            font-size: 2rem;
        }
        .sub-header {
            font-size: 1rem;
            margin-bottom: 1.5rem;
        }
        /* Make columns stack better on mobile if needed */
        div[data-testid="column"] {
            width: 100% !important;
            flex: 1 1 auto;
            min-width: 100%;
        }
    }
    
    /* Card-like container style */
    .css-card {
        padding: 1rem;
        border-radius: 0.5rem;
        background: #F9F9F9;
        margin-bottom: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# Helper Functions
def save_config(config):
    """Save configuration to JSON file (excluding password)"""
    # Create a copy to avoid modifying the original dict in session state
    safe_config = config.copy()
    if 'password' in safe_config:
        del safe_config['password']
    
    with open('config.json', 'w') as f:
        json.dump(safe_config, f)
    return True

def load_config():
    """Load configuration from JSON file"""
    if os.path.exists('config.json'):
        with open('config.json', 'r') as f:
            return json.load(f)
    return {}

def load_templates():
    """Load templates from JSON file"""
    if os.path.exists('templates.json'):
        try:
            with open('templates.json', 'r') as f:
                return json.load(f)
        except:
            return {}
    return {}

def save_templates(templates):
    """Save templates to JSON file"""
    with open('templates.json', 'w') as f:
        json.dump(templates, f, indent=4)


def send_email(smtp_settings, recipient_email, subject, body_html, attachments=None):
    """Send a single email via SMTP"""
    msg = MIMEMultipart()
    msg['From'] = smtp_settings['sender_email']
    msg['To'] = recipient_email
    msg['Subject'] = subject
    if smtp_settings.get('reply_to'):
        msg['Reply-To'] = smtp_settings['reply_to']

    msg.attach(MIMEText(body_html, 'html'))

    if attachments:
        for file in attachments:
            try:
                # In Streamlit, file is an UploadedFile object
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(file.getvalue())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename="{file.name}"')
                msg.attach(part)
            except Exception as e:
                return False, f"Attachment error: {str(e)}"

    try:
        server = smtplib.SMTP(smtp_settings['server'], int(smtp_settings['port']))
        server.starttls()
        server.login(smtp_settings['username'], smtp_settings['password'])
        server.sendmail(smtp_settings['sender_email'], recipient_email, msg.as_string())
        server.quit()
        return True, "Sent successfully"
    except Exception as e:
        return False, str(e)

# --- SIDEBAR: CONFIGURATION ---
with st.sidebar:
    st.header("⚙️ SMTP Configuration")
    
    # Load Config Button
    if st.button("📂 Load Saved Config"):
        loaded_conf = load_config()
        if loaded_conf:
            st.session_state.update(loaded_conf)
            st.success("Configuration loaded! (Enter password below)")
    
    smtp_server = st.text_input("SMTP Server", value=st.session_state.get('smtp_server', 'email-smtp.ap-south-1.amazonaws.com'))
    smtp_port = st.text_input("SMTP Port", value=st.session_state.get('smtp_port', '587'))
    username = st.text_input("Username", value=st.session_state.get('username', ''))
    password = st.text_input("Password", type="password", value=st.session_state.get('password', ''))
    sender_email = st.text_input("Sender Email", value=st.session_state.get('sender_email', ''))
    reply_to = st.text_input("Reply-To Email", value=st.session_state.get('reply_to_email', 'info@mulearn.org'))
    
    if st.button("💾 Save Config (Safe)"):
        conf_to_save = {
            'smtp_server': smtp_server,
            'smtp_port': smtp_port,
            'username': username,
            'sender_email': sender_email,
            'reply_to_email': reply_to
        }
        save_config(conf_to_save)
        st.success("Settings saved (Password not saved)")

    st.divider()
    st.info("💡 **MuMailer Web**\n\nA modern tool for bulk email marketing.")

# --- MAIN PAGE ---
st.markdown('<h1 class="main-header">µLearn Mailer</h1>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Professional Email Marketing Tool</p>', unsafe_allow_html=True)


tab1, tab2, tab3 = st.tabs(["1️⃣ Upload & Data", "2️⃣ Compose & Attachments", "3️⃣ Preview & Send"])

# --- TAB 1: DATA ---
with tab1:
    st.subheader("📊 Upload Recipient Data")
    uploaded_file = st.file_uploader("Upload CSV File (Req: Name, Email columns)", type=['csv'])
    
    if uploaded_file is not None:
        try:
            df = pd.read_csv(uploaded_file)
            st.session_state['csv_data'] = df
            st.success(f"Loaded {len(df)} recipients successfully!")
            
            with st.expander("👀 View Data Preview"):
                st.dataframe(df.head())
            
            st.divider()
            st.subheader("🛠️ Map Columns")
            
            # Smart detection of columns
            all_cols = list(df.columns)
            email_default_idx = 0
            name_default_idx = 0
            
            # Try to find 'email' and 'name' in columns (case-insensitive)
            for i, col in enumerate(all_cols):
                if 'email' in col.lower():
                    email_default_idx = i
                if 'name' in col.lower():
                    name_default_idx = i
            
            # Column Selectors
            col1, col2 = st.columns(2)
            with col1:
                email_col = st.selectbox("Select Email Column", all_cols, index=email_default_idx)
                st.session_state['email_col'] = email_col
            with col2:
                name_col = st.selectbox("Select Name Column", all_cols, index=name_default_idx)
                st.session_state['name_col'] = name_col
            
            st.info(f"Using **{email_col}** for emails and **{name_col}** for names.")
                
        except Exception as e:
            st.error(f"Error reading CSV: {e}")
    else:
        st.info("👆 Please upload a CSV file to begin.")

# --- TAB 2: COMPOSE ---
with tab2:
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.subheader("✏️ Compose Email")
        
        # Template Manager
        templates = load_templates()
        template_names = ["Select a Template..."] + list(templates.keys())
        
        # Template Selection
        selected_template = st.selectbox("📂 Load Template", template_names)
        
        # Load Template Logic
        if selected_template != "Select a Template..." and selected_template in templates:
            # We use a session state flag to check if we just loaded to avoid overwriting user edits immediately on rerun
            # But for simplicity, we just check if it changed from last run or use a button to "Apply"
            if st.button("Apply Template"):
                t_data = templates[selected_template]
                st.session_state.email_subject = t_data.get('subject', '')
                st.session_state.email_body = t_data.get('body', '')
                # Force Quill update
                if 'quill_version' not in st.session_state:
                    st.session_state.quill_version = 0
                st.session_state.quill_version += 1
                st.rerun()

        email_subject = st.text_input("Subject Line", value=st.session_state.get('email_subject', ''))
            
        # Initialize session state for body if not exists
        if 'email_body' not in st.session_state:
            st.session_state.email_body = ""
            
        # Quill Editor
        if 'quill_version' not in st.session_state:
            st.session_state.quill_version = 0

        # Toolbar Configuration
        toolbar_config = [
            [{"header": [1, 2, 3, False]}],
            ["bold", "italic", "underline", "strike"],
            [{"color": []}, {"background": []}],
            [{"list": "ordered"}, {"list": "bullet"}],
            ["link", "image"],
            ["clean"]
        ]

        email_content = st_quill(
            value=st.session_state.email_body, 
            placeholder="Write your email here...",
            html=True,
            toolbar=toolbar_config,
            key=f"quill_editor_{st.session_state.quill_version}"
        )
        
        # Sync back to session state so other tabs see it
        st.session_state.email_body = email_content
        
        st.divider()
        st.subheader("💾 Template Storage")
        
        # Template Actions
        col_t1, col_t2 = st.columns([2, 1])
        
        with col_t1:
             new_template_name = st.text_input("New Template Name", placeholder="e.g., Monthly Newsletter")
             if st.button("Save Current as Template"):
                if new_template_name:
                    templates[new_template_name] = {
                        "subject": email_subject,
                        "body": st.session_state.get('email_body', '')
                    }
                    save_templates(templates)
                    st.success(f"✅ Template '{new_template_name}' saved!")
                    time.sleep(1)
                    st.rerun()
                else:
                    st.error("Please enter a name for the template.")
                    
        with col_t2:
            if selected_template != "Select a Template..." and selected_template in templates:
                st.write(f"selected: **{selected_template}**")
                if st.button("🗑️ Delete Selected"):
                    del templates[selected_template]
                    save_templates(templates)
                    st.warning(f"Template '{selected_template}' deleted.")
                    time.sleep(1)
                    st.rerun()
            else:
                st.info("Select a template above to delete it.")
        
    with col2:
        st.subheader("📎 Attachments")
        uploaded_attachments = st.file_uploader("Add Files", accept_multiple_files=True)
        
        st.subheader("🧩 Variables")
        if 'csv_data' in st.session_state:
            st.info("💡 **Tip:** Click a variable below to append it to your email body. The editor will reload to reflect changes.")
            
            # Helper to append variable
            def append_var(v):
                if 'email_body' not in st.session_state:
                    st.session_state.email_body = ""
                # Append variable
                st.session_state.email_body += f" {{{v}}}"
                # Force reload of Quill
                if 'quill_version' not in st.session_state:
                    st.session_state.quill_version = 0
                st.session_state.quill_version += 1

            # Create a grid of buttons
            cols = st.columns(3)
            for i, col_name in enumerate(st.session_state['csv_data'].columns):
                with cols[i % 3]:
                    st.button(f"{{{col_name}}}", key=f"btn_{col_name}", on_click=append_var, args=(col_name,))
        else:
            st.warning("Upload CSV to see variables")


# --- TAB 3: PREVIEW & SEND ---
with tab3:
    st.subheader("📤 Preview & Dispatch")

    # --- Quick Test Section (Always Visible) ---
    with st.expander("🧪 Quick Test (No CSV required)", expanded=True):
        col_test1, col_test2 = st.columns(2)
        with col_test1:
            q_test_email = st.text_input("Recipient Email", placeholder="test@example.com")
        with col_test2:
            q_test_name = st.text_input("Recipient Name (for {Name} variable)", placeholder="Test User")
            
        if st.button("🚀 Send Quick Test"):
             if not password:
                st.error("Please enter SMTP Password in Sidebar first!")
             elif not q_test_email:
                st.error("Please enter a recipient email.")
             else:
                # Mock a row for variables
                # We use the manually entered name for {Name}
                # And empty strings for any other variables found in the body
                
                # Get current body
                current_body = st.session_state.get('email_body', '')
                current_subject = st.session_state.get('email_subject', '')
                
                # Replace Name
                test_body = current_body.replace("{Name}", q_test_name)
                test_subject = current_subject.replace("{Name}", q_test_name)
                
                # Replace other variables with placeholders
                # (Simple regex or loop could work if we knew all vars, but standard replace is safe enough)
                
                smtp_settings = {
                    'server': smtp_server, 'port': smtp_port,
                    'username': username, 'password': password,
                    'sender_email': sender_email, 'reply_to': reply_to
                }
                
                with st.spinner("Sending..."):
                    success, msg = send_email(smtp_settings, q_test_email, test_subject, test_body, uploaded_attachments)
                
                if success:
                    st.success(f"✅ Test email sent to {q_test_email}!")
                else:
                    st.error(f"❌ Failed: {msg}")

    st.divider()

    if 'csv_data' in st.session_state:
        df = st.session_state['csv_data']
        
        # Previewer
        preview_index = st.number_input("Preview Row Index", min_value=0, max_value=len(df)-1, value=0, step=1)
        
        if 0 <= preview_index < len(df):
            row = df.iloc[preview_index]
            
            # Get mapped columns
            email_col = st.session_state.get('email_col', 'Email')
            name_col = st.session_state.get('name_col', 'Name')
            
            # Replace Variables
            preview_subject = email_subject.replace("{Name}", str(row.get(name_col, '')))
            
            # Get current body from state
            current_body = st.session_state.get('email_body', '')
            preview_body = current_body
            for col in df.columns:
                preview_body = preview_body.replace(f"{{{col}}}", str(row[col]))
            
            st.markdown("### 👁️ Email Preview")
            st.markdown(f"**To:** {row.get(email_col, 'Unknown')}")
            st.markdown(f"**Subject:** {preview_subject}")
            
            # HTML Preview container
            # Wrap in a white container to simulate actual email appearance and ensure readability in Dark Mode
            styled_preview = f"""
            <div style="background-color: white; color: black; padding: 20px; border-radius: 10px; font-family: Helvetica, Arial, sans-serif; min-height: 380px;">
                {preview_body}
            </div>
            """
            st.components.v1.html(styled_preview, height=400, scrolling=True)
            
            st.divider()
            
            # Sending Actions
            col_send1, col_send2 = st.columns(2)
            
            with col_send1:
                st.markdown("### 🧪 Test Send")
                test_email = st.text_input("Test Email Address", value=row.get(email_col, ''))
                if st.button("🚀 Send Test Email"):
                    if not password:
                        st.error("Please enter SMTP Password in Sidebar!")
                    else:
                        smtp_settings = {
                            'server': smtp_server, 'port': smtp_port,
                            'username': username, 'password': password,
                            'sender_email': sender_email, 'reply_to': reply_to
                        }
                        success, msg = send_email(smtp_settings, test_email, preview_subject, preview_body, uploaded_attachments)
                        if success:
                            st.success(f"✅ Test email sent to {test_email}")
                        else:
                            st.error(f"❌ Failed: {msg}")

            with col_send2:
                st.markdown("### 🌍 Bulk Send")
                st.write(f"Ready to send to **{len(df)} recipients**.")
                
                if st.button("🔥 Start Bulk Sending"):
                    if not password:
                        st.error("Please enter SMTP Password in Sidebar!")
                    else:
                        smtp_settings = {
                            'server': smtp_server, 'port': smtp_port,
                            'username': username, 'password': password,
                            'sender_email': sender_email, 'reply_to': reply_to
                        }
                        
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        results = []
                        
                        for i, r in df.iterrows():
                            # Get correct email and name
                            target_email = r.get(email_col)
                            target_name = r.get(name_col, '')
                            
                            # Personalize
                            p_curr_body = st.session_state.get('email_body', '')
                            p_curr_sub = email_subject.replace("{Name}", str(target_name))
                            for col in df.columns:
                                p_curr_body = p_curr_body.replace(f"{{{col}}}", str(r[col]))
                            
                            status_text.text(f"Sending to {target_email} ({i+1}/{len(df)})...")
                            
                            success, msg = send_email(smtp_settings, target_email, p_curr_sub, p_curr_body, uploaded_attachments)
                            results.append({"Email": target_email, "Status": "Sent" if success else "Failed", "Error": msg})
                            
                            progress_bar.progress((i + 1) / len(df))
                            time.sleep(1) # Rate limit safety
                        
                        status_text.text("✅ Bulk sending finished!")
                        st.success("Campaign Completed!")
                        st.dataframe(pd.DataFrame(results))
                        
    else:
        st.info("👆 To use **Bulk Sending**, please upload a CSV file in Tab 1.")
