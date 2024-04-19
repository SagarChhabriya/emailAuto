from email.message import EmailMessage
import ssl
import smtplib
import os
from pathlib import Path
from dotenv import load_dotenv
import pandas as pd
import datetime
import streamlit as st
from email.utils import formataddr # instead of your email  address, you can use a name 

global uploaded_file

# ---------------Load the environemnt Variables------------------------#
current_dir =   Path(__file__).resolve().parent if '__file__' in locals() else Path.cwd()
envars = current_dir / '.env'
load_dotenv(envars)

# --------------Read environment variables-----------------------------#
sender_email = os.getenv("EMAIL")
password_email = os.getenv("PASS")


# -----------------------mailData( ) method----------------------------#


# Get the Email Object
def mailData(email_reciever,subject,body):
    em = EmailMessage()    

    # Add data to the email object
    em['From']      = formataddr(("Sagar Chhabriya.", sender_email))
    em['To']        = email_reciever
    em['Subject']   = subject
    em.set_content(body)

    # Pack Context
    context = ssl.create_default_context()


    # Send mail using Simple-Mail-Transfer-Protocol
    with smtplib.SMTP_SSL('smtp.gmail.com',465,context=context) as smtp:
        smtp.login(sender_email,password_email)
        smtp.sendmail(sender_email,email_reciever,em.as_string())
        
        # Print time
        now = datetime.datetime.now()
        sentTime = f"{now.strftime('%a %b %d')} {now.strftime('%I:%M %p')} {now.year}"
        st.write(f"Email sent to: {email_reciever}" ,f": {sentTime}")

# --------------------mailData( ) body complete

# ---------------------sendMail( ) body start
def sendMail():

    # read emails sheet
    df = pd.read_excel(uploaded_file)
    
    # get value of each cell
    for x in range(df['email'].size):
        email   = df['email'].values[x]
        subject = df['subject'].values[x]
        body    = df['body'].values[x]
        # insert mail data to be sent
        mailData(email_reciever=email, subject=subject,body=body)
# --------------------sendMail() body complete

# UI
st.title("Email Automater")
uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Read Excel file into a Pandas DataFrame
        df = pd.read_excel(uploaded_file)

        # --- Display the DataFrame ---
        st.subheader("Excel File Data")
        st.dataframe(df) 

        # --- Extracting and Using Data ---
        st.subheader("Select Specific Columns")
        columns = df.columns.tolist()
        selected_columns = st.multiselect("Choose columns", columns)

        if selected_columns:
            st.write("Data from selected columns:")
            st.dataframe(df[selected_columns])

    except Exception as e:
        st.error(f"Error reading Excel file: {e}")

try:
    if st.button('Send Emails'):
        sendMail()
except:
    st.error("Please select an excel sheet first.")


# Project Complete: RADHE RADHE 🙏🏼    
