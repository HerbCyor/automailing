import smtplib
from email.message import EmailMessage
from email.mime.base import MIMEBase
from email.utils import make_msgid
import mimetypes
from email import encoders
import os.path

'''
    Python script to handle automatic e-mail sending with in message hmtl (for signature attachment or other purposes)
    and support for sending attached files.

    Next step would be to create a better handling for connection failures.

'''

LOGIN_NAME = 'youremail@yourdomain.com'
LOGIN_PASSWORD = 'yourpassword' #it's better to use .env or pickle here DON'T SAVE YOUR PASSWORD IN PLAIN TEXT
SMTP_SERVER = 'outlook.office365.com' #standard outlook server
EMAIL_SENDER = 'youremail@yourdoamin.com'
SMTP_PORT = 587 #standard port

def send_email(email_recipient, email_subject, email_text, attachment_location = ''):
    global LOGIN_PASSWORD
    global LOGIN_NAME
    global SMTP_SERVER
    global EMAIL_SENDER
    global SMTP_PORT
    
    #creating a message
    msg = EmailMessage()
    #Email Headers
    msg['From'] = EMAIL_SENDER
    msg['To'] = email_recipient
    msg['Subject'] = email_subject
    
    # set the plain text body
    msg.set_content('this is a plain text')
    # now create a Content-ID for the image
    image_cid = make_msgid(domain='xyz.com')
    # if `domain` argument isn't provided, it will 
    # use your computer's name
    
    # set an alternative html body.
    msg.add_alternative("""\
    <html>
        <body>
            <p>{email_message}<br>
            Sample text.html
            </p>
            <img src="cid:{image_cid}" width="820" height="260">
        </body>
    </html>
    """.format(email_message = email_text, image_cid=image_cid[1:-1]), subtype='html')
    # image_cid looks like <long.random.number@xyz.com>
    # to use it as the img src, we don't need `<` or `>`
    # so we use [1:-1] to strip them off


    # now open the image and attach it to the email
    with open('signature.jpg', 'rb') as img:

        # know the Content-Type of the image
        maintype, subtype = mimetypes.guess_type(img.name)[0].split('/')

        # attach it
        msg.get_payload()[1].add_related(img.read(), 
                                            maintype=maintype, 
                                            subtype=subtype, 
                                            cid=image_cid)



        #handling attachments
    if attachment_location != '':
        filename = os.path.basename(attachment_location)
        attachment = open(attachment_location, "rb")
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
        msg.attach(part)

    try:
        #server connections
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.ehlo()
        server.starttls()
        server.login(LOGIN_NAME, LOGIN_PASSWORD)
        text = msg.as_string()
        server.sendmail(EMAIL_SENDER, email_recipient, text)
        print('email sent to '+ email_recipient)
        server.quit()
    except:
        
        print(f"SMPT server connection error: email to {email_recipient} failed")
        
    return True