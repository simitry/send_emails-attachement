import numpy as np
import pandas as pd
from email.mime.multipart import MIMEMultipart  
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from time import sleep
import ssl
import smtplib

processed_emails = []

excel_file = 'rh_email.xlsx'
data = pd.read_excel(excel_file)

con = data.loc[data["VILLE"] == "MARRAKECH", ["E-mail", "PERSONNE"]]
con = con.dropna()
print(con)

for index, row in con.iterrows():
    emails = row['E-mail']
    name = row['PERSONNE']
    email_list = [receiver.strip() for receiver in emails.split()]

    for receiver in email_list:
        processed_emails.append({'Name': name, 'E-mail': receiver})

processed_email_df = pd.DataFrame(processed_emails)
processed_email_df.to_excel('processed_emails.xlsx', index=False)
print('All the email addresses processed are stored in "processed_emails.xlsx"')
sleep(5)

for i in range(len(processed_emails)):
    print(processed_emails[i])

myemail = 'enter email here'
mypassword = 'password'

subject = 'Demande de stage'

for i, row in processed_email_df.iterrows():
    name = row['Name']
    email_receiver = row['E-mail']
    
    em = MIMEMultipart()  
    em['From'] = myemail
    em['To'] = email_receiver
    em['Subject'] = subject
    
    body = f"""
    Bonjour {name},
    
    Veuillez trouver ci-joint mon CV et demande de stage.
    
    Cordialement.
    """
    
    em.attach(MIMEText(body, 'plain'))  
    
    # Attach the first file
    attachment1 = r'attachment path here'
    with open(attachment1, 'rb') as f:
        attachment_data1 = f.read()

    attachment_part1 = MIMEApplication(attachment_data1, Name='cv.pdf')
    attachment_part1['Content-Disposition'] = 'attachment; filename="name.txt"'
    em.attach(attachment_part1)

    # Attach the second file
    attachment2 = r'attachement path here'
    with open(attachment2, 'rb') as f:
        attachment_data2 = f.read()

    attachment_part2 = MIMEApplication(attachment_data2, Name='name.txt')
    attachment_part2['Content-Disposition'] = 'attachment; filename="name.txt"'
    em.attach(attachment_part2)

    context = ssl.create_default_context()
    
    with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
        smtp.login(myemail, mypassword)
        smtp.sendmail(myemail, email_receiver, em.as_string())
