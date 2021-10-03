import smtplib
import ssl

from email.utils import formatdate
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase

import os
import datetime as dt
from string import Template

import icalendar
import pytz


class sendInvite:
    def __init__(self, filename_contacts, filename_message, filename_senderdetails, filename_time_and_date):
        self.filename_contacts = filename_contacts
        self.filename_message = filename_message
        self.filename_senderdetails = filename_senderdetails
        self.filename_time_and_date = filename_time_and_date

    def get_contacts(self):
        filename = self.filename_contacts
        print(filename)
        names = []
        emails = []
        with open(filename, mode='r', encoding='utf-8') as contacts_file:
            for a_contact in contacts_file:
                names.append(a_contact.split()[0])
                emails.append(a_contact.split()[1])
        return names, emails

    def read_template(self):
        filename = self.filename_message
        with open(filename, 'r', encoding='utf-8') as template_file:
            template_file_content = template_file.read()
        return Template(template_file_content)

    def getSenderEmailPwd(self):
        filename = self.filename_senderdetails
        with open(filename, mode='r', encoding='utf-8') as sender_file:
            sender_details = sender_file.read()
        return sender_details.split()[0], sender_details.split()[1], sender_details.split()[2]

    def get_timeAndDate(self):
        filename = self.filename_time_and_date
        dict_timedate = {}
        with open(filename, mode='r', encoding='utf-8') as timedate_file:
            for lineread in timedate_file:
                dict_timedate[lineread.split()[0]] = int(lineread.split()[1])
        return dict_timedate

    def generateMeetingCall(self, sender_email, receiver_email, filename_message):
        timedate = self.get_timeAndDate()
        start_hour = timedate['start_hour']
        start_minute = timedate['start_minute']
        year = timedate['year']
        month = timedate['month']
        day = timedate['day']
        start_hour = 7
        start_minute = 30
        reminder_hours = 1
        subj = "Minus-Zero Meeting Invitation"
        description = filename_message
        location = "Jalandhar"

        # Timezone to use for our dates - change as needed
        tz = pytz.timezone("Europe/London")
        start_date = dt.date(year, month, day)
        start = tz.localize(dt.datetime.combine(
            start_date, dt.time(start_hour, start_minute, 0)))
        # Build the event itself
        cal = icalendar.Calendar()
        cal.add('prodid', '-//My calendar application//')
        cal.add('version', '2.0')
        cal.add('method', "REQUEST")
        event = icalendar.Event()
        event.add('attendee', receiver_email)
        event.add('organizer', sender_email)
        event.add('status', "confirmed")
        event.add('category', "Event")
        event.add('summary', subj)
        event.add('description', description)
        event.add('location', location)
        event.add('dtstart', start)
        event.add('dtend', tz.localize(dt.datetime.combine(
            start_date, dt.time(start_hour + 1, start_minute, 0))))
        event.add('dtstamp', tz.localize(
            dt.datetime.combine(start_date, dt.time(6, 0, 0))))
        # event['uid'] = self.get_unique_id() # Generate some unique ID
        event.add('priority', 5)
        event.add('sequence', 1)
        event.add('created', tz.localize(dt.datetime.now()))

        # Add a reminder
        alarm = icalendar.Alarm()
        alarm.add("action", "DISPLAY")
        alarm.add('description', "Reminder")
        # The only way to convince Outlook to do it correctly
        alarm.add("TRIGGER;RELATED=START", "-PT{0}H".format(reminder_hours))
        event.add_component(alarm)
        cal.add_component(event)

        # Build the email message and attach the event to it
        msg = MIMEMultipart("alternative")

        msg["Subject"] = subj
        msg["From"] = sender_email
        msg["To"] = receiver_email
        msg["Content-class"] = "urn:content-classes:calendarmessage"

        msg.attach(MIMEText(description))

        filename = "invite.ics"
        part = MIMEBase('text', "calendar", method="REQUEST", name=filename)
        part.set_payload(cal.to_ical())
        # email.Encoders.encode_base64(part)
        part.add_header('Content-Description', filename)
        part.add_header("Content-class", "urn:content-classes:calendarmessage")
        part.add_header("Filename", filename)
        part.add_header("Path", filename)
        msg.attach(part)
        return msg

    def sendEmail(self, msg, sender_password):
        smtp_server = "smtp.gmail.com"
        port = 587  # For starttls
        # sender_email = msg["From"]
        # receiver_email = msg["To"]
        sender_email = "soyames@gmail.com"
        receiver_email = "yao.sossou@edu.fh-joanneum.at"
        # input("Type your password and press enter: ")
        sender_password = "7566124839"
        password = sender_password

        # Create a secure SSL context
        context = ssl.create_default_context()

        # Try to log in to server and send email
        try:

            server = smtplib.SMTP(smtp_server, port)
            server.ehlo()  # Can be omitted
            server.starttls(context=context)  # Secure the connection
            server.ehlo()  # Can be omitted
            server.login(sender_email, password)
            # TODO: Send email here
            server.sendmail(sender_email, receiver_email, msg.as_string())
            #server.sendmail(msg["From"], msg["To"], msg.as_string())

            print("email sent!")
            '''
            with smtplib.SMTP(smtp_server, port) as server:
                server.ehlo()  # Can be omitted
                server.starttls(context=context)
                server.ehlo()  # Can be omitted
                server.login(sender_email, password)
                server.sendmail(sender_email, receiver_email, message)
                print("email sent!")
             '''
        except Exception as e:
            # Print any error messages to stdout
            print(e)
        finally:
            server.quit()

    def validation(self, sender_email, names, emails, message_template, sender_password):
        self.sender_email = sender_email
        self.names = names
        self.emails = emails
        self.message_template = message_template
        for name, email in zip(self.names, self.emails):
            print(name, ':', email)
            s = self.message_template.substitute(PERSON_NAME=name)
            msg = self.generateMeetingCall(self.sender_email, email, s)
            self.sendEmail(msg, sender_password)
            # print(s)


THIS_FOLDER = os.path.dirname(os.path.abspath(__file__))

file_contacts = os.path.join(THIS_FOLDER, 'myContacts.txt')
file_message = os.path.join(THIS_FOLDER, 'message.txt')
file_sender_details = os.path.join(THIS_FOLDER, 'sender_details.txt')
file_time_and_date = os.path.join(THIS_FOLDER, 'meeting_time_and_date.txt')


inviter = sendInvite(file_contacts, file_message,
                     file_sender_details, file_time_and_date)
names, emails = inviter.get_contacts()  # read contacts
message_template = inviter.read_template()

sender_name, sender_email, sender_password = inviter.getSenderEmailPwd()
print(sender_email, sender_password)
sender_email = "soyames@gmail.com"
receiver_email = "yao.sossou@edu.fh-joanneum.at"
msg = inviter.generateMeetingCall(sender_email, receiver_email, file_message)
inviter.validation(sender_email, names, emails,
                   message_template, sender_password)
inviter.sendEmail(msg)

del inviter

print("Sucess!")
