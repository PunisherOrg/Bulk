import pandas as pd
import streamlit as st
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def send_email(subject, body, cc_address, smtp_server, smtp_port, sender_email, sender_password):
    try:
        # Create the email object
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = 'himanshu825406@gmail.com'
        msg['Subject'] = subject
        msg['Cc'] = cc_address  # Add the CC recipient

        # Add the body to the email
        msg.attach(MIMEText(body, 'plain'))

        # Create the SMTP session for sending the email
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()  # Upgrade the connection to a secure encrypted SSL/TLS connection
            server.login(sender_email, sender_password)  # Login to the SMTP server
            text = msg.as_string()
            server.sendmail(sender_email, [msg['To'], cc_address], text)

        return f"Email sent to {msg['To']} with CC to {cc_address} and subject '{subject}'"
    except Exception as e:
        return f"Failed to send email: {str(e)}"

st.title('Bulk SF-Case Update through Excel')
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

# Email configuration inputs
st.write("Enter SMTP details for sending emails:")
smtp_server = st.text_input("SMTP Server (e.g., smtp.gmail.com for Gmail)", "smtp.gmail.com")
smtp_port = st.number_input("SMTP Port", min_value=1, max_value=65535, value=587)
sender_email = st.text_input("Your Email Address")
sender_password = st.text_input("Your Email Password", type="password")

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # Display the content of the DataFrame
    st.write("Data from Excel file:")
    st.write(df)

    if st.button('Send Emails'):
        if not sender_email or not sender_password:
            st.error("Please provide email and password to send emails.")
        else:
            results = []
            for _, row in df.iterrows():
                owner_mail = row['OwnerEmail'].strip()  # Remove any extra spaces
                update = row['Update']
                subject = row['Subject']
                body = update
                cc_address = owner_mail
                result = send_email(subject, body, cc_address, smtp_server, smtp_port, sender_email, sender_password)
                results.append(result)
            
            # Display the results
            st.write("Email sending results:")
            for result in results:
                st.write(result)
