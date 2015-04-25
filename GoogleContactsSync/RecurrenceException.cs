using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;
using Google.GData.Calendar;

namespace R.GoogleOutlookSync
{
    class RecurrenceException
    {
        private EventRecurrence _parentRecurrence;

        internal bool Deleted { get; set; }
        internal CalendarEvent ModifiedEvent { get; set; }
        internal DateTime OriginalDate { get; set; }

        internal RecurrenceException(EventRecurrence parentRecurrence, DateTime originalDate)
        {
            this._parentRecurrence = parentRecurrence;
            this.OriginalDate = originalDate;
        }

        internal RecurrenceException(EventEntry googleException, EventRecurrence parentRecurrence = null)
        {
            this._parentRecurrence = parentRecurrence != null ? parentRecurrence : null;
            this.OriginalDate = googleException.OriginalEvent.OriginalStartTime.StartTime;
            this.Deleted = googleException.Status.Value == EventEntry.EventStatus.CANCELED_VALUE;
            if (!this.Deleted)
                this.ModifiedEvent = new CalendarEvent(googleException);
        }

        internal RecurrenceException(Outlook.Exception outlookException, EventRecurrence parentRecurrence = null)
        {
            this._parentRecurrence = parentRecurrence != null ? parentRecurrence : new EventRecurrence(outlookException.Parent as Outlook.RecurrencePattern);
            this.OriginalDate = outlookException.OriginalDate;
            this.Deleted = outlookException.Deleted;
            if (!this.Deleted)
                this.ModifiedEvent = new CalendarEvent(outlookException.AppointmentItem);
        }
    }
}
