using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;
using Google.Apis.Calendar.v3.Data;

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

        internal static bool Equals(Event googleException, Outlook.Exception outlookException)
        {
            bool googleExceptionDeleted = googleException.Status == "cancelled";
            bool exceptionEventEqual;
            if (outlookException.Deleted)
                exceptionEventEqual = true;
            else
                exceptionEventEqual = outlookException.AppointmentItem.AllDayEvent ?
                    (
                        googleException.Start.Date == outlookException.AppointmentItem.Start.ToString("yyyy-MM-dd") &&
                        googleException.End.Date == outlookException.AppointmentItem.End.ToString("yyyy-MM-dd")
                    ) :
                    (
                        googleException.Start.DateTime.Value == outlookException.AppointmentItem.Start &&
                        googleException.End.DateTime.Value == outlookException.AppointmentItem.End
                    );
                    EventComparer.Equals(googleException, outlookException.AppointmentItem);
            return
                googleException.OriginalStartTime.DateTime == outlookException.OriginalDate.Date &&
                googleExceptionDeleted == outlookException.Deleted &&
                exceptionEventEqual;
        }

        internal static bool Equals(List<Event> googleExceptions, Outlook.Exceptions outlookExceptions)
        {
            var tmpGoogleExceptions = new List<Event>(googleExceptions);
            foreach (Outlook.Exception outlookException in outlookExceptions)
            {
                try
                {
                    var found = false;
                    foreach (var googleException in new List<Event>(tmpGoogleExceptions))
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

        internal static bool Contains(List<Event> googleExceptions, Outlook.Exception outlookException)
        {
            foreach (var googleException in googleExceptions)
                if (Equals(googleException, outlookException))
                    return true;
            return false;
        }

        internal static bool Contains(Outlook.Exceptions outlookExceptions, Event googleException)
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
