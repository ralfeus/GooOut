using Google.Apis.Calendar.v3.Data;
using System;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace R.GoogleOutlookSync
{
    class EventComparer
    {
        internal static bool Equals(Event googleItem, Outlook.AppointmentItem outlookItem)
        {
            return
                googleItem.Summary == outlookItem.Subject &&
                googleItem.Description == outlookItem.Body &&
                AttendeeComparer.Equals(googleItem.Attendees, outlookItem.Recipients) &&
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
        private static bool LocationIsEqual(Event googleItem, Outlook.AppointmentItem outlookItem)
        {
            return
                (String.IsNullOrEmpty(googleItem.Location) && String.IsNullOrEmpty(outlookItem.Location)) ||
                (googleItem.Location == outlookItem.Location);
        }

        [FieldComparer(Field.Reminder)]
        private static bool ReminderIsEqual(Event googleItem, Outlook.AppointmentItem outlookItem)
        {
            if ((googleItem.Reminders == null) && !outlookItem.ReminderSet)
                return true;
            if (googleItem.Reminders != null)
                return googleItem.Reminders.Overrides.First().Minutes == outlookItem.ReminderMinutesBeforeStart;
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
