#=========================================== RPA =======================================================================

import pandas as pd
import sqlalchemy as sal
import pyautogui as pg
import os
import pyodbc
import streamlit as st


#=========================================== SQL CONNECTION ============================================================

# SQL DB information
username = ''
password = ''
port = '1433'
server_name = 'ZAIN-PC'
database = 'MISC'
DSN = 'ZAIN-PC-NHS-ENGLAND' # You must create a DSN (using "ODBC Data Sources" on the control panel) to create a relevant connection
                # In this case 'ZAIN-PC' links to the MISC database within the 'ZAIN-PC' server

# Use the following code to connect to the SQL engine - Please note! this is for a native SQL client
#engine = sal.create_engine(f'mssql+pyodbc://{username}:{password}@{server_name}/{database}?driver=SQL+Server+Native+Client+11.0')

# With a non-native client please use the code below
#engine = sal.create_engine(f'postgresql+psycopg2://{username}:{password}@{server_name}:{port}/{database}')
engine = sal.create_engine(f'mssql+pyodbc://{username}:{password}@{DSN}')

conn = engine.connect()

#=========================================== SQL QUERY WRITING =========================================================

# executing a query with the engine
#result = engine.execute('select * from [dbo].[Jon_Jones]')

# reading a SQL query using pandas - using the engine
sql_query = pd.read_sql_query("""SELECT TOP 1000 * FROM [NHS ENGLAND].[dbo].[Recosted_2021_v2]""", engine)
# saving SQL table in a pandas data frame
df = pd.DataFrame(sql_query)
print(df)
df.to_csv(r'C:\Users\Zain\PycharmProjects\RoboticProcessAutomation\df.csv')
filepath = r'C:\Users\Zain\PycharmProjects\RoboticProcessAutomation\df.csv'
# uploading a df to SQL - first read the CSV
#df = pd.read_csv(r'C:\Users\Zain\PycharmProjects\RoboticProcessAutomation\df.csv')
# create a new table called 'df' and append data frame values to this table if it already exists
#df.to_sql('df', con=engine, if_exists='append',index=False,chunksize=1000)

# To close the connection
conn.close()

#===================================== SENDING EMAILS ==================================================================

# The below code uses the default SMTP module within PYTHON to send emails - but YAGMAIL is a better module
#--------------------------------------------------------------------------------------
# import smtplib, ssl
#
# port = 465  # For SSL
# email = "zaineisa.work@gmail.com"
# password = ''
#
# # Create a secure SSL context
# context = ssl.create_default_context()
#
# with smtplib.SMTP_SSL("smtp.gmail.com", port, context=context) as server:
#     server.login(email, password)
#
#     sender_email = 'zaineisa.work@gmail.com'
#     receiver_email = 'zaineisa.work@gmail.com'
#     message = 'THIS IS A PYTHON TEST'
#
#     server.sendmail(sender_email, receiver_email, message)
#--------------------------------------------------------------------------------------

# YAGMAIL is a library that is better at sending emails
import yagmail
email_password = ''
receiver = "zaineisa.work@gmail.com"
body = f"YAGMAIL test email being sent with the dataset from {database}"
filename = filepath

yag = yagmail.SMTP("zaineisa.work@gmail.com",email_password)
yag.send(
        to=receiver,
        subject="Yagmail test with attachment",
        contents=body,
        attachments=filename,
    )

# ========================================= STREAMLIT ==================================================================

st.subheader(f'Summary of Data from SQL Database "{server_name}"')
st.markdown(f'Data shown for the database {database} seen below')
st.dataframe(df)