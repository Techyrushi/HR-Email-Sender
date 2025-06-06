import streamlit as st
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def send_emails(sender_email, sender_password, subject, template, df):
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(sender_email, sender_password)

    sent_count = 0
    failed_emails = []

    for index, row in df.iterrows():
        recipient_email = row[2]  # 3rd column 
        st.write(f"Sending to: {recipient_email}")

        msg = MIMEMultipart('alternative')
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject

        msg.attach(MIMEText(template, 'html'))

        try:
            server.sendmail(sender_email, recipient_email, msg.as_string())
            sent_count += 1
        except Exception as e:
            failed_emails.append(recipient_email)
            st.write(f"❌ Error sending to {recipient_email}: {e}")

    server.quit()
    return sent_count, failed_emails

def main():
    st.set_page_config(page_title="HR Email Sender", page_icon="📧")
    st.title("📧 HR Email Sender Automation")

    st.write("""
    Upload an Excel file with HR email addresses in the 3rd column.
    Customize your email below and click "Send Emails".
    """)

    uploaded_file = st.file_uploader("Upload Excel file (.xlsx)", type=["xlsx"])
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        # Show only a preview of the first 10 rows
        st.write("Preview (first 10 rows):", df.head(490))
        # Display the total number of records
        st.write(f"Total records to process: {len(df)}")
        
        email_subject = st.text_input("Email Subject", "Exciting Collaboration Opportunity")
        email_template = st.text_area("Email Body (HTML allowed)", """
        <p>Hello,</p>
        <p>We’d like to connect with you regarding exciting collaboration opportunities.</p>
        <p>Best regards,<br>Your Name</p>
        """, height=200)

        sender_email = st.text_input("Sender Email (e.g. you@gmail.com)")
        sender_password = st.text_input("Sender App Password", type="password")

        if st.button("🚀 Send Emails"):
            if not sender_email or not sender_password:
                st.warning("Please enter your sender email and app password.")
            else:
                # Limit to first 490 emails
                df_to_send = df.head(490)
                st.write(f"Sending emails to first 490 recipients out of {len(df)} total records.")
                
                with st.spinner("Sending emails..."):
                    sent_count, failed_emails = send_emails(
                        sender_email, sender_password, email_subject, email_template, df_to_send
                    )
                    st.success(f"✅ Sent {sent_count} emails successfully!")
                    if failed_emails:
                        st.warning("⚠️ Some emails failed:")
                        st.write(failed_emails)

if __name__ == "__main__":
    main()
