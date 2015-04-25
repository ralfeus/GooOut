using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Google.GData.Calendar;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;

namespace R.GoogleOutlookSync
{
    class RecurrenceExceptionComparer : IEqualityComparer<RecurrenceException>
    {
        #region IEqualityComparer<RecurrenceException> Members

        bool IEqualityComparer<RecurrenceException>.Equals(RecurrenceException x, RecurrenceException y)
        {
            if (object.ReferenceEquals(x, y))
                return true;
            if (x.Deleted)
                return
                    y.Deleted &&
                    x.OriginalDate == y.OriginalDate;
            else
                return
                    !y.Deleted &&
                    x.OriginalDate == y.OriginalDate &&
                    EventComparer.Equals(x.ModifiedEvent, y.ModifiedEvent);
        }

        int IEqualityComparer<RecurrenceException>.GetHashCode(RecurrenceException obj)
        {
            return
                obj.Deleted.GetHashCode() ^
                obj.ModifiedEvent.GetHashCode() ^
                obj.OriginalDate.GetHashCode();
        }

        #endregion

        internal static bool Equals(EventEntry googleException, Outlook.Exception outlookException)
        {
            bool googleExceptionDeleted = googleException.Status.Value == EventEntry.EventStatus.CANCELED_VALUE;
            bool exceptionEventEqual;
            if (outlookException.Deleted)
                exceptionEventEqual = true;
            else
                exceptionEventEqual = 
                    googleException.Times[0].StartTime == outlookException.AppointmentItem.Start &&
                    googleException.Times[0].EndTime == outlookException.AppointmentItem.End &&
                    googleException.Times[0].AllDay == outlookException.AppointmentItem.AllDayEvent &&
                    EventComparer.Equals(googleException, outlookException.AppointmentItem);
            return
                googleException.OriginalEvent.OriginalStartTime.StartTime.Date == outlookException.OriginalDate.Date &&
                googleExceptionDeleted == outlookException.Deleted &&
                exceptionEventEqual;
        }

        internal static bool Equals(List<EventEntry> googleExceptions, Outlook.Exceptions outlookExceptions)
        {
            var tmpGoogleExceptions = new List<EventEntry>(googleExceptions);
            foreach (Outlook.Exception outlookException in outlookExceptions)
            {
                try
                {
                    var found = false;
                    foreach (var googleException in new List<EventEntry>(tmpGoogleExceptions))
                    {
                        if (Equals(googleException, outlookException))
                        {
                            found = true;
                            tmpGoogleExceptions.Remove(googleException);
                            break;
                        }
                    }
                    if (!found)
                        return false;
                }
                finally
                {
                    Marshal.ReleaseComObject(outlookException);
                }
            }                    
            return tmpGoogleExceptions.Count == 0;
        }

        internal static bool Contains(List<EventEntry> googleExceptions, Outlook.Exception outlookException)
        {
            foreach (var googleException in googleExceptions)
                if (Equals(googleException, outlookException))
                    return true;
            return false;
        }

        internal static bool Contains(Outlook.Exceptions outlookExceptions, EventEntry googleException)
        {
            foreach (Outlook.Exception outlookException in outlookExceptions)
                try
                {
                    if (Equals(googleException, outlookException))
                        return true;
                }
                finally
                {
                    Marshal.ReleaseComObject(outlookException);
                }
            return false;
        }
    }
}
