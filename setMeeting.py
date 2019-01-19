import win32com.client
import datetime
OUTLOOK_FORMAT = '%m-%d-%Y %H:%M'
outlook_date   = lambda dt: dt.strftime(OUTLOOK_FORMAT)
reminder = 15.5
oOutlook = win32com.client.Dispatch("Outlook.Application")
appt = oOutlook.CreateItem(1) # 1 - olAppointmentItem
#appt.Start = '2012-03-8 17:00'
appt.Start = outlook_date(datetime.datetime.now() - datetime.timedelta(0,reminder*60))
appt.Subject = 'Follow Up Meeting'
appt.Duration = 20
appt.Location = 'Office - Room 132A'
appt.MeetingStatus = 1 # 1 - olMeeting; Changing the appointment to meeting
#only after changing the meeting status recipients can be added
appt.Recipients.Add("saurabh.kumar@shell.com")
#appt.Recipients.Add("aaaaaa@shell.com")
appt.ReminderMinutesBeforeStart = 15
appt.Save()
appt.Send()
print "Done"