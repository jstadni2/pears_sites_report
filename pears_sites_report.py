import os
import pandas as pd

import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders

# Calculate the path to the root directory of this script
ROOT_DIR = os.path.realpath(os.path.join(os.path.dirname(__file__), '.'))

sites = pd.read_excel(ROOT_DIR + "/example_inputs/Site_Export.xlsx", sheet_name='Site Data')
users = pd.read_excel(ROOT_DIR + "/example_inputs/User_Export.xlsx", sheet_name='User Data')


# Sites Report


prev_month = (pd.to_datetime("today") - pd.DateOffset(months=1))
report_month_int = prev_month.month
report_year_int = prev_month.year

sites['created'] = pd.to_datetime(sites['created'])
sites['created_month'] = sites['created'].dt.month
sites['created_year'] = sites['created'].dt.year

sites = sites.loc[
    (sites['created_month'] == report_month_int) & (sites['created_year'] == report_year_int),
    ['site_id',
     'site_name',
     'created_by',
     'created_by_email',
     'created',
     'address',
     'city', 'city__county',
     'zip_code',
     'setting']
]

sites = pd.merge(sites, users[['full_name', 'program_area']], how='left', left_on='created_by',
                 right_on='full_name').drop(columns=['full_name'])
sites.insert(5, 'program_area', sites.pop('program_area'))

sites['created'] = sites['created'].dt.strftime('%m-%d-%Y')


# Export the Sites Report as an Excel file


sites_report_filename = 'PEARS Sites Report ' + prev_month.strftime('%Y-%m') + '.xlsx'
out_path = ROOT_DIR + "/example_outputs"
sites_report_path = out_path + '/' + sites_report_filename


# Export a list of dataframes as an Excel workbook
# file: string for the name or path of the file
# sheet_names: list of strings for the name of each sheet
# dfs: list of dataframes for the report
def write_report(file, sheet_names, dfs):
    report_dict = dict(zip(sheet_names, dfs))
    writer = pd.ExcelWriter(file, engine='xlsxwriter')
    # Loop through dict of dataframes
    for sheet_name, df in report_dict.items():
        # Send df to writer
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        # Pull worksheet object
        worksheet = writer.sheets[sheet_name]
        # Loop through all columns
        for idx, col in enumerate(df):
            series = df[col]
            max_len = max((
                # Len of the largest item
                series.astype(str).map(len).max(),
                # Len of column name/header
                len(str(series.name))
            )) + 1  # adding a little extra space
            # Set column width
            worksheet.set_column(idx, idx, max_len)
            worksheet.autofilter(0, 0, 0, len(df.columns) - 1)
    writer.close()


write_report(sites_report_path, ['PEARS Sites Report'], [sites])


# Email Sites Report


# Set the following variables with the appropriate credentials and recipients
admin_username = 'your_username@domain.com'
admin_password = 'your_password'
admin_send_from = 'your_username@domain.com'
report_cc = 'list@domain.com, of_recipients@domain.com'


# Send an email with or without a xlsx attachment
# send_from: string for the sender's email address
# send_to: string for the recipient's email address
# Cc: string of comma-separated cc addresses
# subject: string for the email subject line
# html: string for the email body
# username: string for the username to authenticate with
# password: string for the password to authenticate with
# isTls: boolean, True to put the SMTP connection in Transport Layer Security mode (default: True)
# wb: boolean, whether an Excel file should be attached to this email (default: False)
# file_path: string for the xlsx attachment's filepath (default: '')
# filename: string for the xlsx attachments filename (default: '')
def send_mail(send_from,
              send_to,
              cc,
              subject,
              html,
              username,
              password,
              is_tls=True,
              wb=False,
              file_path='',
              filename=''):
    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = send_to
    msg['Cc'] = cc
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject
    msg.attach(MIMEText(html, 'html'))

    if wb:
        fp = open(file_path, 'rb')
        part = MIMEBase('application', 'vnd.ms-excel')
        part.set_payload(fp.read())
        fp.close()
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment', filename=filename)
        msg.attach(part)

    smtp = smtplib.SMTP('smtp.office365.com', 587)
    if is_tls:
        smtp.starttls()
    try:
        smtp.login(username, password)
        smtp.sendmail(send_from, send_to.split(',') + msg['Cc'].split(','), msg.as_string())
    except smtplib.SMTPAuthenticationError:
        print("Authentication failed. Make sure to provide a valid username and password.")
    smtp.quit()


report_recipients = 'recipient@domain.com'
report_subject = 'PEARS Sites Report ' + prev_month.strftime('%Y-%m')

report_html = """<html>
  <head></head>
<body>
            <p>
            Hello,<br><br>

            Here is the PEARS Sites Report for the previous month.

            SITES ADMIN-
            Would you mind verifying that the correct setting is selected and duplicate sites are merged? 

            If you have any questions, please reply to this email and I will respond at my earliest opportunity.<br>

            <br>Best Regards,<br>       
            <br> <b> FCS Evaluation Team </b> <br>
            <a href = "mailto: your_username@domain.com ">your_username@domain.com </a><br>
            </p>
  </body>
</html>
"""

send_mail(send_from=admin_send_from,
          send_to=report_recipients,
          cc=report_cc,
          subject=report_subject,
          html=report_html,
          username=admin_username,
          password=admin_password,
          is_tls=True,
          wb=True,
          file_path=sites_report_path,
          filename=sites_report_filename)


# Email Unauthorized Site Creators


# List of PEARS users authorized to create sites
site_creators = ['names', 'of', 'PEARS', 'users']

# Create list of staff to notify
staff_list = sites.loc[~sites['created_by'].isin(site_creators)
                       & sites['created_by_email'].str.contains('|'.join(['illinois.edu', 'uic.edu']), na=False),
                       ['created_by', 'created_by_email']].drop_duplicates(keep='first').values.tolist()

notification_subject = "Friendly REMINDER: Adding new sites to PEARS " + prev_month.strftime('%Y-%m')

notification_cc = 'list@domain.com, of_recipients@domain.com'

notification_html = """
<html>
  <head></head>
<body>
            <p>
            Hello {0},<br>

            <br>You are receiving this email as our records show you have added a new site to the PEARS database within
            the last month. This a friendly reminder that new site additions to PEARS are conducted centrally on campus 
            for all Extension program areas. Requests for new sites in PEARS must be sent to 
            <a href = "mailto: sites_admin@@domain.com">sites_admin@@domain.com </a> for entry.
            We do this to keep our database clean, accurate, and free of accidental duplicates.<br>

            <br>We ask that field staff not add new sites on their own. A member of the state Evaluation Team is trained
            in the process of adding new sites so that they are usable for staff across all  Extension program areas.
            If the individual in receipt of your request has questions they will reach out to you for clarification.<br>
            
            <br>Please reply to this email if you have any questions or think you have received this message
            in error.<br>

            <br>Thanks and have a great day!<br>       
            <br> <b> FCS Evaluation Team </b> <br>
            <a href = "mailto: your_username@domain.com ">your_username@domain.com </a><br>
  </body>
</html>
"""

failed_recipients = []

for x in staff_list:
    staff_name = x[0]
    notification_send_to = x[1]
    user_html = notification_html.format(staff_name)
    # Try to send the email, otherwise add the recipient's email address to failed_recipients
    try:
        send_mail(send_from=admin_send_from,
                  send_to=notification_send_to,
                  cc=notification_cc,
                  subject=notification_subject,
                  html=user_html,
                  username=admin_username,
                  password=admin_password,
                  is_tls=True)
    except smtplib.SMTPException:
        failed_recipients.append(x)

# Notify admin of any failed attempts to send an email
# Else, print success notification to console
if failed_recipients:
    fail_html = """The following recipients failed to receive an email:<br>
    {}    
    """
    new_string = '<br>'.join(map(str, failed_recipients))
    new_html = html.format(new_string)
    fail_subject = 'PEARS Sites Report  ' + prev_month.strftime('%b-%Y') + ' Failure Notice'
    send_mail(send_from=admin_send_from,
              send_to='your_username@domain.com',
              cc='',
              subject=fail_subject,
              html=fail_html,
              username=admin_username,
              password=admin_password,
              is_tls=True)
else:
    print("Unauthorized site creation notifications sent successfully.")
