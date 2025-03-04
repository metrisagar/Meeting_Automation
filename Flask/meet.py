from flask import Flask, request, render_template
import win32com.client as win32
import pandas as pd
import pythoncom
from datetime import datetime, timedelta
import pytz
import time

app = Flask(__name__)

def create_meeting(subject, body, start_time, end_time, location, required_attendees, sender_email, retries=5, delay=1):
    pythoncom.CoInitialize()  # Initialize COM library
    attempt = 0
    while attempt < retries:
        try:
            outlook = win32.Dispatch('outlook.application')
            meeting = outlook.CreateItem(1)  # 1: olAppointmentItem

            meeting.Subject = subject
            meeting.Body = body
            meeting.Start = start_time
            meeting.End = end_time
            meeting.Location = location
            meeting.MeetingStatus = 1  # 1: olMeeting

            for attendee in required_attendees:
                meeting.Recipients.Add(attendee)

            # Set the sender email
            account = None
            for acc in outlook.Session.Accounts:
                if acc.SmtpAddress == sender_email:
                    account = acc
                    break

            if account:
                meeting.SendUsingAccount = account

            meeting.Save()
            meeting.Send()
            pythoncom.CoUninitialize()  # Uninitialize COM library
            print("Meeting created successfully!")
            break
        except pythoncom.com_error as e:
            print(f"Attempt {attempt + 1} failed: {e}")
            attempt += 1
            time.sleep(delay)
    else:
        print("Failed to create meeting after several attempts.")

def get_meeting_details_from_excel(ritm_number, excel_file, default_attendees, sender_name, meeting_type):
    df = pd.read_excel(excel_file)
    row = df[df['RITMNumber'] == ritm_number].iloc[0]

    if meeting_type == "application":
        subject = f"{row['RITMNumber']}_{row['PanarolaID']}_{row['Appname']} Application Walkthrough"
        body = f"\n\nPlease join the meeting to discuss {row['RITMNumber']}_{row['PanarolaID']}_{row['Appname']} Application Walkthrough.\n\nBelow is the agenda of the meeting:\n• To understand application workflow/functionality\n• Application architecture overview.\n• What type of application [Custom/COTS/SAAS/etc.] and application hosted environment [Haleon/3rd party]\n• User roles available in the application and respective functionalities.\n• Application access?\n Is provided URL is a stable environment, with latest code drop and prod like configuration.\n• When the AST tested, code is planned to move into production?\n• Is there any major release planned, if yes when?\n\nBest Regards,\n{sender_name}"
    else:
        subject = f"{row['RITMNumber']}_{row['PanarolaID']}_{row['Appname']} Report Walkthrough"
        body = f"Hi {row['Requester']},\n\nScheduling this meeting for {row['RITMNumber']}_{row['PanarolaID']}_{row['Appname']} Report Walkthrough.\n\nNote- Request you to Extend this invite to other Vendors or Whomsoever it may concern.\n\nBest Regards,\n{sender_name}"

    required_attendees = [row['Requester']] + default_attendees
    return subject, body, required_attendees

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/create_meeting', methods=['POST'])
def create_meeting_route():
    ritm_number = request.form['ritm_number']
    start_time = request.form['start_time']
    sender_name = request.form['sender_name']
    meeting_type = request.form['meeting_type']
    sender_email = "chappsec@gskconsumer.com"
    excel_file = 'meeting_details.xlsx'  # Replace with the path to your Excel file
    default_attendees = ["sagar.x.metri@haleon.com"]  # Add your default attendees here

    subject, body, required_attendees = get_meeting_details_from_excel(ritm_number, excel_file, default_attendees, sender_name, meeting_type)
    start_time = datetime.strptime(start_time, "%Y-%m-%dT%H:%M")
    start_time = pytz.timezone('Asia/Kolkata').localize(start_time)
    end_time = start_time + timedelta(minutes=30)
    location = "Microsoft Teams Meeting"

    create_meeting(subject, body, start_time, end_time, location, required_attendees, sender_email)
    return "Meeting created successfully!"

if __name__ == '__main__':
    app.run(debug=True)