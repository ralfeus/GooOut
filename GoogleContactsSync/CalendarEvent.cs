using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Google.GData.Calendar;
using Microsoft.Office.Interop.Outlook;
using Google.GData.Extensions;

namespace R.GoogleOutlookSync
{
    class CalendarEvent
    {
        internal EventSchedule Schedule { get; set; }
        internal string Subject { get; set; }
        internal string Body { get; set; }
        internal bool ReminderSet { get; set; }
        internal int ReminderMinutes { get; set; }
        internal List<Attendee> Attendees { get; set; }
        internal string Location { get; set; }

        internal CalendarEvent(EventEntry googleItem)
        {
            this.Subject = googleItem.Title.Text;
            this.Body = googleItem.Content.Content;
            this.ReminderSet = googleItem.Reminder != null;
            this.ReminderMinutes = this.ReminderSet ? googleItem.Reminder.Minutes : 0;
            this.Attendees = new List<Attendee>(googleItem.Participants.Count - 1);
            foreach (Who participant in googleItem.Participants)
                if (participant.Rel != "http://schemas.google.com/g/2005#event.organizer")
                    this.Attendees.Add(new Attendee(participant.Email, participant.Attendee_Type, participant.Attendee_Status));
            this.Location = googleItem.Locations[0].ValueString;
            this.Schedule = new EventSchedule(googleItem);
        }

        internal CalendarEvent(AppointmentItem outlookItem)
        {
            this.Subject = outlookItem.Subject;
            this.Body = outlookItem.Body;
            this.ReminderSet = outlookItem.ReminderSet;
            this.ReminderMinutes = outlookItem.ReminderMinutesBeforeStart;
            this.Attendees = new List<Attendee>(outlookItem.Recipients.Count);
            foreach (Recipient recipient in outlookItem.Recipients)
                this.Attendees.Add(new Attendee(recipient.Address, recipient.Type, recipient.MeetingResponseStatus));
            this.Location = outlookItem.Location;
            this.Schedule = new EventSchedule(outlookItem);
        }

        public override bool Equals(object obj)
        {
            if (object.ReferenceEquals(this, obj))
                return true;
            CalendarEvent target = (CalendarEvent)obj;
            return
                this.Attendees.SequenceEqual(target.Attendees, new AttendeeComparer()) &&
                this.Body == target.Body &&
                this.Location == target.Location &&
                this.ReminderSet == target.ReminderSet &&
                this.ReminderMinutes == target.ReminderMinutes &&
                this.Subject == target.Subject &&
                this.Schedule.Equals(target.Schedule);
        }

        public override int GetHashCode()
        {
            return
                this.Attendees.GetHashCode() ^
                this.Body.GetHashCode() ^
                this.Location.GetHashCode() ^
                this.ReminderMinutes.GetHashCode() ^
                this.ReminderSet.GetHashCode() ^
                this.Schedule.GetHashCode() ^
                this.Subject.GetHashCode();
        }

        internal void ToGoogle(EventEntry googleItem)
        {
            this.Schedule.ToGoogle(googleItem);

            if (this.ReminderSet)
            {
                googleItem.Reminder = new Google.GData.Extensions.Reminder();
                googleItem.Reminder.Minutes = this.ReminderMinutes;
            }
        }

        internal EventEntry ToGoogle()
        {
            EventEntry googleItem = new EventEntry(this.Subject, this.Body, this.Location);
            this.ToGoogle(googleItem);
            return googleItem;
        }
    }
}