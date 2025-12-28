import smtplib
from email.message import EmailMessage
import ssl

def send_email(email, password, to_email, subject, content):

    # create the email message

    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = email
    msg['To'] = to_email
    msg.set_content(content)

    # smtp server configuration
    SMTP_SERVER = "smtps.uhserver.com"
    SMTP_PORT = 465

    try:
        # create a secure SSL context
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, context=context, timeout=30) as server:

            # login
            server.login(email, password)

            # send email
            server.send_message(msg)

        print("Email sent successfully!")
    except Exception as e:
        print(f"Error: {e}")
