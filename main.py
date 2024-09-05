import pandas as pd
from datetime import datetime, timedelta
import requests
import json
import pytz
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment
from dotenv import load_dotenv
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

load_dotenv()

# ... (rest of the imports and function definitions remain the same)

# Set your Azure AD app details
client_id = os.getenv("CLIENT_ID")
client_secret = os.getenv("CLIENT_SECRET")
tenant_id = os.getenv("TENANT_ID")

# Get the list of userPrincipalNames from the .env file
user_principal_names = os.getenv("USER_PRINCIPAL_NAMES", "").split(",")
manager_email = os.getenv("MANAGER_EMAIL")

# Get date range (last 24 hours)
end_date = datetime.now(pytz.utc)
start_date = end_date - timedelta(hours=24)
from_date = start_date.isoformat()
to_date = end_date.isoformat()

print(f"Querying for calls from {from_date} to {to_date}")

# Get the token and call logs
token = get_token(client_id, client_secret, tenant_id)
if not token:
    print("Failed to get token. Exiting.")
    exit(1)

all_call_logs = get_call_logs(token, from_date, to_date)

if not all_call_logs:
    print("No call data found for the specified date range.")
    exit(0)

# Process the call logs
df = pd.DataFrame(all_call_logs['value'])
filtered_df = df[df['userPrincipalName'].isin(user_principal_names)]

# Format start and end dates as strings
start_date_str = start_date.strftime('%Y%m%d_%H%M')
end_date_str = end_date.strftime('%Y%m%d_%H%M')

# Generate the consolidated report for the manager
consolidated_excel_filename = f'consolidated_calls_report_{start_date_str}_{end_date_str}.xlsx'
generate_excel_report(filtered_df, consolidated_excel_filename)

# Initialize a list to store users without call data
users_without_data = []

# Generate and send individual reports to each user
for user_principal_name in user_principal_names:
    # Filter the call logs for the current user
    user_df = df[df['userPrincipalName'] == user_principal_name]
    
    if user_df.empty:
        print(f"No call data found for user: {user_principal_name}")
        users_without_data.append(user_principal_name)
        continue
    
    # Generate the individual report for the user
    user_excel_filename = f'calls_report_{user_principal_name}_{start_date_str}_{end_date_str}.xlsx'
    generate_excel_report(user_df, user_excel_filename)
    
    # Send email to the user with their individual report
    subject = f"Your Call Data Report ({start_date.strftime('%B %d, %Y %H:%M')} - {end_date.strftime('%B %d, %Y %H:%M')})"
    body = f"Dear {user_df['userDisplayName'].iloc[0]},<br><br>Please find attached your call data report for the period {start_date.strftime('%B %d, %Y %H:%M')} to {end_date.strftime('%B %d, %Y %H:%M')}."
    send_email(subject, body, user_principal_name, [user_excel_filename])
    print(f"Individual report sent to {user_principal_name}")

# Update the email body for the manager
body = f"Dear Manager,<br><br>Please find attached the consolidated call data report for the period {start_date.strftime('%B %d, %Y %H:%M')} to {end_date.strftime('%B %d, %Y %H:%M')}."

if users_without_data:
    body += "<br><br>The following users had no call data during this period:<br>"
    body += "<br>".join(users_without_data)
else:
    body += "<br><br>All users had call data during this period."

# Send the consolidated report to the manager
subject = f"Consolidated Call Data Report ({start_date.strftime('%B %d, %Y %H:%M')} - {end_date.strftime('%B %d, %Y %H:%M')})"
send_email(subject, body, manager_email, [consolidated_excel_filename])
print(f"Consolidated report sent to the manager: {manager_email}")
