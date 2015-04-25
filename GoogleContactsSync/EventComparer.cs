using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Google.GData.Calendar;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace R.GoogleOutlookSync
{
    class EventComparer
    {
        internal static bool Equals(EventEntry googleItem, Outlook.AppointmentItem outlookItem)
        {
            return
                googleItem.Title.Text == outlookItem.Subject &&
                googleItem.Content.Content == outlookItem.Body &&
                AttendeeComparer.Equals(googleItem.Participants, outlookItem.Recipients) &&
                LocationIsEqual(googleItem, outlookItem) &&
                ReminderIsEqual(googleItem, outlookItem);
        }

        internal static bool Equals(CalendarEvent x, CalendarEvent y)
        {
            var a = x.Attendees.SequenceEqual(y.Attendees, new AttendeeComparer());
            return
                StringIsEqual(x.Subject, y.Subject) &&
                StringIsEqual(x.Body, y.Body) &&
                x.Attendees.SequenceEqual(y.Attendees, new AttendeeComparer()) &&
                StringIsEqual(x.Location, y.Location) &&
                ReminderIsEqual(x, y);
        }

        private static bool StringIsEqual(string x, string y)
        {
            return
                (String.IsNullOrEmpty(x) && string.IsNullOrEmpty(y)) ||
                (x == y);
        }

        [FieldComparer(Field.Location)]
        private static bool LocationIsEqual(EventEntry googleItem, Outlook.AppointmentItem outlookItem)
        {
            return
                (String.IsNullOrEmpty(googleItem.Locations[0].ValueString) && String.IsNullOrEmpty(outlookItem.Location)) ||
                (googleItem.Locations[0].ValueString == outlookItem.Location);
        }

        [FieldComparer(Field.Reminder)]
        private static bool ReminderIsEqual(EventEntry googleItem, Outlook.AppointmentItem outlookItem)
        {
            if ((googleItem.Reminder == null) && !outlookItem.ReminderSet)
                return true;
            if (googleItem.Reminder != null)
                return googleItem.Reminder.Minutes == outlookItem.ReminderMinutesBeforeStart;
            return false;
        }

        private static bool ReminderIsEqual(CalendarEvent x, CalendarEvent y)
        {
            return
                (x.ReminderSet == y.ReminderSet) &&
                (!x.ReminderSet || x.ReminderMinutes == y.ReminderMinutes);
        }
    }
}
