from google.oauth2 import service_account
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
import base64
import csv

# giving the permissions needed
SCOPES = ['https://www.googleapis.com/auth/gmail.send']

# authentication to use the account 
flow = InstalledAppFlow.from_client_secrets_file(
    'client_secret_754206838234-b6eakou7temfjv57mdav2jdu9hgijth2.apps.googleusercontent.com.json', SCOPES)
creds = flow.run_local_server(port=0)

# making the gmail service
service = build('gmail', 'v1', credentials=creds)

# making the email itself
def write_email(to_email, subject, template_file, attachment_path, inline_image_path, recipient_name, startup_name, member, title):
    message = MIMEMultipart('related')  # 'related' to include inline images
    message['to'] = to_email
    message['subject'] = subject

    # reading and formatting the HTML template
    with open(template_file, 'r') as file:
        body_text = file.read().format(
            recipient_name=recipient_name,
            startup_name=startup_name,
            member=member,
            title=title
        )
    
    # attaching the HTML body
    message.attach(MIMEText(body_text, 'html'))

    # attaching the inline image
    with open(inline_image_path, 'rb') as img_file:
        msg_image = MIMEImage(img_file.read())
        msg_image.add_header('Content-ID', '<logo_image>')
        msg_image.add_header('Content-Disposition', 'inline', filename="logo.png")
        message.attach(msg_image)

    # attaching the pdf
    with open(attachment_path, 'rb') as file:
        attachment = MIMEApplication(file.read(), Name="Winternship_Employer_Package.pdf")
        attachment['Content-Disposition'] = 'attachment; filename="Winternship_Employer_Package.pdf"'
        message.attach(attachment)

    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
    return {'raw': raw_message}

# fcn for sending email
def send_email(service, to_email, subject, template_file, attachment_path, inline_image_path, recipient_name, startup_name, member, title):
    message = write_email(to_email, subject, template_file, attachment_path, inline_image_path, recipient_name, startup_name, member, title)
    try:
        sent_message = service.users().messages().send(userId='me', body=message).execute()
        print(f'Email sent to {to_email}: Message ID - {sent_message["id"]}')
    except Exception as e:
        print(f'An error occurred while sending email to {to_email}: {e}')

# sending emails
with open('contacts.csv', newline='') as csvfile:
    reader = csv.reader(csvfile)
    for row in reader: 
        email = row[0]  # email address
        startup_name = row[1]  # startup name
        recipient_name = row[2]  # recipient name
        member = row[3]  # member name
        title = row[4]  # title

        send_email(service,
           to_email=email,
           subject="Winternship partnership opportunity",
           template_file="email.html",  # HTML template with placeholders
           attachment_path='2024_25_Winternship_Package.pdf',  # PDF attachment
           inline_image_path='StartupLogo.png',  # Path to the inline logo image
           recipient_name=recipient_name,
           startup_name=startup_name,
           member=member,
           title=title)
