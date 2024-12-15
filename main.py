import datetime
from email.header import decode_header
from email.message import EmailMessage
import os
import smtplib
from imaplib import IMAP4_SSL
import email

# fetch encrypted mail username and password of school's mail
sch_user: str = os.environ.get('SCH_USER')
sch_pass: str = os.environ.get('SCH_PASS')

# which sender we are looking for and from which date we are looking
sender: str = "saman"
date: datetime.datetime = datetime.datetime.now()
four_days_ago: datetime.datetime = date - datetime.timedelta(days=4)

# what imap and smtp server we are using
imap_server: str = 'imap.gmail.com'
smtp_server: str = 'smtp.gmail.com'

# function for sending a mail with attachment to the relevant people
def send_mail(file_data, file_name: str, classroom: str, week_num: int, student: bool):

    # create a mail
    msg = EmailMessage()
    msg['Subject'] = f'Ukeplannr. {week_num} for klasse {classroom}'
    msg['From'] = sch_user
    msg['To'] = 'undisclosed-recipients:;'

    msg.set_content(f"""\
Vedlagt ukeplannr. {week_num}

Mvh
Salahaddin skoleadministrasjon
    """)
    msg.add_attachment(file_data, maintype='application', subtype='pdf', filename=file_name)

    # getting recipents
    bcc_addresses = []

    # check if we need to retrieve from student mails or teacher mails
    if student:
        mail_file = open('mails/students.txt', 'r')
    else:
        mail_file = open('mails/teachers.txt', 'r')

    # parsing through the mails, sending to mail from corresponding classroom
    mail_list = mail_file.read().split('\n')
    for mail in mail_list:
        cr = mail.split()[0]
        ml = mail.split()[1]
        if cr.lower() == classroom.lower():
            bcc_addresses.append(ml)

    # converting bcc addressess to be usable
    bcc_addresses = ', '.join(bcc_addresses)

    # sending the mail
    with smtplib.SMTP_SSL(smtp_server, 465) as smtp:
        smtp.login(sch_user, sch_pass)
        smtp.send_message(msg, to_addrs=[bcc_addresses])


# mail server i am connecting to is outlook/hotmail
imap: IMAP4_SSL = IMAP4_SSL(imap_server)

# log in to the mail
imap.login(sch_user, sch_pass)

# looking for mails from sender within the last 5 days in the inbox category
imap.select("Inbox")
formatted = four_days_ago.strftime('%d-%b-%Y')
status, tot_msgs = imap.search(None, f'FROM "{sender}" SINCE {formatted}')

# if it finds a mail that matches the criteria
if status == 'OK':
    print("scanning mail...\n")
    # the latest mail

    # we should be matching with excactly 2 mails
    tot_mails = len(tot_msgs[0].split())
    if tot_mails != 2:
        print(f"Wrong number of mails found: {tot_mails}, should be 2.")
        # closing and logging out for safe measure
        imap.close()
        imap.logout()
        os._exit(1)

    for msg in tot_msgs[0].split():

        # fetch each mail as bytes
        _, data = imap.fetch(msg, "(RFC822)")   
        # turn into message class
        msg = email.message_from_bytes(data[0][1])
        subject = msg.get('Subject')
        decoded_subject = decode_header(subject)

        # see if its weekly schedule for teachers or studfents
        sub_list = decoded_subject[0][0].split()
        student = False
        elev = sub_list[1].lower()
        if elev == 'elev':
            student = True

        # get the week number
        week_num = str(sub_list[-1]).split('.')[1]

        # iterates through all the parts of the mail
        for part in msg.walk():

            # looks for all the weekly schedules 
            file_name = part.get_filename()

            if file_name:

                # retrieve which class the pdf is for
                classroom = file_name.split()[-1].split('.')[0]

                # send the schedule to the relevant people
                file_data = part.get_payload(decode=True)
                print(f'sending {file_name}...')
                send_mail(file_data, file_name, classroom, week_num, student)
                print(f'{file_name} sendt successfully!\n')


else:
    print("Did not find the e-mails")

# closing and logging out for safe measure
imap.close()
imap.logout()
