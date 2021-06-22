import pandas as pd
import smtplib

'''
Change these to your credentials and name
'''
your_name = "Traideas Technology"
your_email = "walter.y.m.g@gmail.com"
your_password = "password"

# If you are using something other than gmail
# then change the 'smtp.gmail.com' and 465 in the line below
server = smtplib.SMTP('smtp.gmail.com:587')

server.ehlo()
server.starttls()
server.ehlo()
server.login(your_email, your_password)

# Read the file
email_list = pd.read_excel("~/Downloads/Email list.xlsx", engine='openpyxl')

# Get all the Names, Email Addreses, Subjects and Messages
all_emails = email_list['Email']
all_names = email_list['Name']



# Loop through the emails
counter = 0
for idx in range(len(all_emails)):
    counter = counter + 1
    # Get each records name, email, subject and message
    name = all_names[idx]
    email = all_emails[idx]
    subject = "Do you need any custom software"
    message = "Hi this is test"

    # Create the email to send
    full_email = ("From: {0} <{1}>\n"
                  "To: {2} <{3}>\n"
                  "Subject: {4}\n\n"
                  "{5}"
                  .format(your_name, your_email, name, email, subject, message))

    # In the email field, you can add multiple other emails if you want
    # all of them to receive the same text
    try:
        server.sendmail(your_email, [email], full_email)
        print(str(counter) + ' Email to {} successfully sent!\n\n'.format(email))
    except Exception as e:
        print(str(counter) + ' Email to {} could not be sent :( because {}\n\n'.format(email, str(e)))

# Close the smtp server
server.close()











