import pandas as pd
import win32com.client
import streamlit as st
def send_email(subject, body, cc_address):
    ol = win32com.client.Dispatch('Outlook.Application')
    to_address=f'himanshu825406@gmail.com'
    olmailitem = 0x0
    newmail = ol.CreateItem(olmailitem)
    newmail.Subject = subject
    newmail.To = to_address
    newmail.CC = cc_address  # Set the Owner mail in CC
    newmail.Body = body
    newmail.Send()
    return f"Email sent to {to_address} with CC to {cc_address} and subject '{subject}'"

st.title('Bulk SF-Case Update though Excel')
uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # Display the content of the DataFrame
    st.write("Data from Excel file:")
    st.write(df)

    if st.button('Send Emails'):
        results = []
        for _, row in df.iterrows():
            owner_mail = row['OwnerEmail'].strip()  # Remove any extra spaces
            update = row['Update']
            subject = row['Subject']
            body = update
            cc_address = owner_mail
            result = send_email(subject, body, cc_address)
            results.append(result)
        # Display the results
        st.write("Email sending results:")
        for result in results:
            st.write(result)