using Google.GData.Calendar;
using Outlook = Microsoft.Office.Interop.Outlook;
using System;
using Google.GData.Extensions;
using System.Linq;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Google.GData.Client;

namespace R.GoogleOutlookSync
{
    internal partial class CalendarSynchronizer
    {
        /// <summary>
        /// Compares AllDayEvent attribute for items
        /// </summary>
        /// <param name="googleItem"></param>
        /// <param name="outlookItem"></param>
        /// <returns></returns>
        //[FieldComparer(Field.AllDayEvent)]
        private bool CompareAllDayEvent(EventEntry googleItem, Outlook.AppointmentItem outlookItem)
        {
            return googleItem.Times[0].AllDay == outlookItem.AllDayEvent;
        }

        //private bool CompareRecurrence(EventEntry googleItem, Outlook.AppointmentItem outlookItem)
        //{
        //    EventRecurrence googleRec = null;
        //    EventRecurrence outlookRec = null;
        //    try 
        //    {
        //        if (googleItem.Recurrence != null)
        //            googleRec = new EventRecurrence(googleItem.Recurrence); 
        //    } 
        //    catch (ArgumentNullException) { }
        //    try 
        //    { 
        //        if (outlookItem.IsRecurring)
        //            outlookRec = new EventRecurrence(outlookItem.GetRecurrencePattern()); 
        //    }
        //    catch (ArgumentNullException) { }
        //    return googleRec == outlookRec;
        //}

        /// Attachments are not accessible yet by Google API
        //[FieldComparer(Field.Attachments)]
        //private bool CompareAttachments(EventEntry googleItem, Outlook.AppointmentItem outlookItem)
        //{
        //    return true;
        //}

        [FieldComparer(Field.Attendees)]
        private bool CompareAttendees(EventEntry googleItem, Outlook.AppointmentItem outlookItem)
        {
            return AttendeeComparer.Equals(googleItem.Participants, outlookItem.Recipients); // ((vbMAPI_AppointmentItem)vbMAPI_Init.NewOutlookWrapper(outlookItem)).Recipients);
        }

        [FieldComparer(Field.Description)]
        private bool CompareDescription(EventEntry googleItem, Outlook.AppointmentItem outlookItem)
        {
            //var outlookItemBody = ((vbMAPI_AppointmentItem)vbMAPI_Init.NewOutlookWrapper(outlookItem)).Body.Value;
            return
                (String.IsNullOrEmpty(googleItem.Content.Content) && String.IsNullOrEmpty(outlookItem.Body)) ||
                (googleItem.Content.Content == outlookItem.Body);
        }

        [FieldComparer(Field.Location)]
        private bool CompareLocation(EventEntry googleItem, Outlook.AppointmentItem outlookItem)
        {
            return
                (String.IsNullOrEmpty(googleItem.Locations[0].ValueString) && String.IsNullOrEmpty(outlookItem.Location)) ||
                (googleItem.Locations[0].ValueString == outlookItem.Location);
        }

        //[FieldComparer(Field.Reminder)]
        private bool CompareReminder(EventEntry googleItem, Outlook.AppointmentItem outlookItem)
        {
            if ((googleItem.Reminder == null) && !outlookItem.ReminderSet)
                return true;
            if (googleItem.Reminder != null)
                return googleItem.Reminder.Minutes == outlookItem.ReminderMinutesBeforeStart;
            return false;
        }

        [FieldComparer(Field.ShowTimeAs)]
        private bool CompareShowTimeAs(EventEntry googleItem, Outlook.AppointmentItem outlookItem)
        {
            return googleItem.EventTransparency.Value == ConvertTo.Google(outlookItem.BusyStatus).Value;
        }
        
        [FieldComparer(Field.Subject)]
        private bool CompareSubjects(EventEntry googleItem, Outlook.AppointmentItem outlookItem)
        {
            return googleItem.Title.Text == outlookItem.Subject;
        }

        [FieldComparer(Field.Time)]
        private bool CompareTime(EventEntry googleItem, Outlook.AppointmentItem outlookItem)
        {
            List<EventEntry> googleExceptions = null;
            if (this._googleExceptions.ContainsKey(googleItem.EventId))
                googleExceptions = this._googleExceptions[googleItem.EventId];
            
            return
                new EventSchedule(googleItem, googleExceptions).Equals(new EventSchedule(outlookItem)) &&
                this.CompareReminder(googleItem, outlookItem);
        }

        [FieldComparer(Field.Visibility)]
        private bool CompareVisibility(EventEntry googleItem, Outlook.AppointmentItem outlookItem)
        {
            return googleItem.EventVisibility.Value == ConvertTo.Google(outlookItem.Sensitivity).Value;
        }

        //[FieldGetter(Field.Subject)]
        //private string GetSubject(object item)
        //{
        //    if (item is EventEntry)
        //        return ((EventEntry)item).Title.Text;
        //    else if (item is Outlook.AppointmentItem)
        //        return ((Outlook.AppointmentItem)item).Subject;
        //    else
        //        throw new ArgumentException(Properties.Settings.Default.ErrorMessage_InvalidItem, "item");
        //}

        //[FieldSetter(Field.AllDayEvent)]
        //private void SetAllDayEvent(EventEntry googleItem, Outlook.AppointmentItem outlookItem, Targets target)
        //{
        //    if (target == Targets.Google)
        //    {
        //        if (googleItem.Times.Count == 0)
        //            googleItem.Times.Add(new When());
        //        googleItem.Times[0].AllDay = outlookItem.AllDayEvent;
        //    }
        //    else
        //        outlookItem.AllDayEvent = googleItem.Times[0].AllDay;
        //}

        /// Attachments are not accessible yet by Google API
        //[FieldSetter(Field.Attachments)]
        //private void SetAttachments(EventEntry googleItem, Outlook.AppointmentItem outlookItem, Target target)
        //{
        //}

        [FieldSetter(Field.Attendees)]
        private void SetAttendees(EventEntry googleItem, Outlook.AppointmentItem outlookItem, Target target)
        {
            if (target == Target.Google)
            {
                foreach (Outlook.Recipient outlookRecipient in outlookItem.Recipients)
                {
                    try 
                    {
                        /// Organizator is omited
                        if (((Outlook.OlMeetingRecipientType)outlookRecipient.Type == Outlook.OlMeetingRecipientType.olOrganizer) ||
                            (outlookRecipient.Address == null) ||
                            !Utilities.SMTPAddressPattern.IsMatch(outlookRecipient.Address))
                            continue;
                        var googleRecipient = googleItem.Participants.FirstOrDefault(recipient => recipient.Email == outlookRecipient.Address);
                        if (googleRecipient == null)
                        {
                            googleRecipient = new Who();
                            googleRecipient.Email = outlookRecipient.Address;
                            googleItem.Participants.Add(googleRecipient);
                        }
                        else if (AttendeeComparer.Equals(googleRecipient, outlookRecipient))
                        {
                            continue;
                        }
                        googleRecipient.Rel = ConvertTo.Google((Outlook.OlMeetingRecipientType)outlookRecipient.Type);
                        googleRecipient.Attendee_Status = ConvertTo.Google(outlookRecipient.MeetingResponseStatus);
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(outlookRecipient);
                    }
                }
            }
            else
            {
                foreach (var googleRecipient in googleItem.Participants)
                {
                    /// Organizator and not valid SMTP addresses are omited
                    if ((googleRecipient.Rel == "http://schemas.google.com/g/2005#event.organizer") ||
                        !Utilities.SMTPAddressPattern.IsMatch(googleRecipient.Email))
                        continue;
                    var matchIsFound = false;
                    foreach (Outlook.Recipient outlookRecipient in outlookItem.Recipients)
                    {
                        if (googleRecipient.Email == outlookRecipient.Address)
                        {
                            matchIsFound = true;
                            Marshal.ReleaseComObject(outlookRecipient);
                            break;
                        }
                        Marshal.ReleaseComObject(outlookRecipient);
                    }
                    if (!matchIsFound)
                    {
                        Outlook.Recipient outlookRecipient = outlookItem.Recipients.Add(googleRecipient.Email);
                        outlookRecipient.Type = (int)ConvertTo.Outlook(googleRecipient.Rel);
                        Marshal.ReleaseComObject(outlookRecipient);
                    }
                }
            }
        }

        [FieldSetter(Field.Description)]
        private void SetDescription(EventEntry googleItem, Outlook.AppointmentItem outlookItem, Target target)
        {
            if (target == Target.Google)
                googleItem.Content.Content = outlookItem.Body;
            else
                outlookItem.Body = googleItem.Content.Content;
        }

        [FieldSetter(Field.Location)]
        private void SetLocation(EventEntry googleItem, Outlook.AppointmentItem outlookItem, Target target)
        {
            if (target == Target.Google)
            {
                if (googleItem.Locations.Count == 0)
                    googleItem.Locations.Add(new Where("location", outlookItem.Location, outlookItem.Location));
                else
                    googleItem.Locations[0].ValueString = outlookItem.Location;
            }
            else
                outlookItem.Location = googleItem.Locations[0].ValueString;
        }

        /// <summary>
        /// Sets recurrence in direction Google -> Outlook
        /// </summary>
        /// <param name="googleItem">Source Google event</param>
        /// <param name="outlookItem">Target Outlook event</param>
        private void SetRecurrence(EventEntry googleItem, Outlook.AppointmentItem outlookItem)
        {
            if (googleItem.Recurrence != null)
            {
                Logger.Log(googleItem.Recurrence.Value, EventType.Debug);
                new EventSchedule(googleItem).GetOutlookRecurrence(outlookItem);
                this.SetRecurrenceExceptions(googleItem, outlookItem);
            }
            else
                outlookItem.ClearRecurrencePattern();
        }
        
        /// <summary>
        /// Sets recurrence in direction Outlook -> Google
        /// </summary>
        /// <param name="googleItem">Source Outlook event</param>
        /// <param name="outlookItem">Target Google event</param>
        private void SetRecurrence(Outlook.AppointmentItem outlookItem, EventEntry googleItem)
        {
            if (outlookItem.IsRecurring)
            {
                /// Times property doesn't contain anything for recurrent events
                googleItem.Times.Clear();
                var outlookItemSchedule = new EventSchedule(outlookItem);
                googleItem.Recurrence = outlookItemSchedule.GetGoogleRecurrence();
                this.SetRecurrenceExceptions(outlookItem, googleItem);
            }
            else
                googleItem.Recurrence = null;
        }

        //[FieldSetter(Field.Reminder)]
        private void SetReminder(EventEntry googleItem, Outlook.AppointmentItem outlookItem, Target target)
        {
            if (target == Target.Google)
                if (outlookItem.ReminderSet)
                {
                    var reminder = new Reminder();
                    reminder.Minutes = outlookItem.ReminderMinutesBeforeStart;
                    reminder.Method = Reminder.ReminderMethod.all;
                    if (googleItem.Reminder == null)
                        try
                        {
                            googleItem.Reminders.Add(reminder);
                        }
                        catch (NullReferenceException)
                        {
                            throw new EventScheduleIsNotSetException(Properties.Resources.Error_EventScheduleIsNotSet);
                        }
                }
                else
                    googleItem.Reminder = null;
            else
                if (googleItem.Reminder == null)
                    outlookItem.ReminderSet = false;
                else
                {
                    outlookItem.ReminderSet = true;
                    outlookItem.ReminderMinutesBeforeStart = googleItem.Reminder.Minutes;
                }
        }

        /// <summary>
        /// Sets recurrence exceptions in direction Google -> Outlook
        /// </summary>
        /// <param name="googleItem">Source Google event</param>
        /// <param name="outlookItem">Target Outlook event</param>
        private void SetRecurrenceExceptions(EventEntry googleItem, Outlook.AppointmentItem outlookItem)
        {
            if (!this._googleExceptions.ContainsKey(googleItem.EventId))
                return;
            outlookItem.Save();
            Outlook.RecurrencePattern outlookRecurrence = outlookItem.GetRecurrencePattern();
            try
            {
                foreach (EventEntry googleException in this._googleExceptions[googleItem.EventId])
                {
                    /// If the exception is already in Outlook event this one is omited
                    //if (RecurrenceExceptionComparer.Contains(outlookRecurrence.Exceptions, googleException))
                    //    continue;
                    /// Get occurence of the recurrence and modify it. Thus new exception is created
                    Outlook.AppointmentItem outlookExceptionItem = outlookRecurrence.GetOccurrence(googleException.OriginalEvent.OriginalStartTime.StartTime);
                    try
                    {
                        if (googleException.Status.Value == EventEntry.EventStatus.CANCELED_VALUE)
                        {
                            outlookExceptionItem.Delete();
                        }
                        else
                        {
                            this.SetSubject(googleException, outlookExceptionItem, Target.Outlook);
                            this.SetDescription(googleException, outlookExceptionItem, Target.Outlook);
                            this.SetLocation(googleException, outlookExceptionItem, Target.Outlook);
                            this.SetTime(googleException, outlookExceptionItem, Target.Outlook);
                            outlookExceptionItem.Save();
                        }
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(outlookExceptionItem);
                    }
                }
            }
            finally
            {
                Marshal.ReleaseComObject(outlookRecurrence);
            }
        }

        /// <summary>
        /// Sets recurrence exceptions in direction Outlook -> Google
        /// </summary>
        /// <param name="googleItem">Source Outlook event</param>
        /// <param name="outlookItem">Target Google event</param>
        private void SetRecurrenceExceptions(Outlook.AppointmentItem outlookItem, EventEntry googleItem)
        {
            Outlook.RecurrencePattern outlookRecurrence = outlookItem.GetRecurrencePattern();
            try
            {
                foreach (Outlook.Exception outlookException in outlookRecurrence.Exceptions)
                {
                    try
                    {
                        /// If the exception is already in Google event this one is omited
                        if (this._googleExceptions.ContainsKey(googleItem.EventId) &&
                            RecurrenceExceptionComparer.Contains(this._googleExceptions[googleItem.EventId], outlookException))
                            continue;
                        //googleItem.RecurrenceException = new Google.GData.Extensions.RecurrenceException();
                        var googleException = new EventEntry();
                        googleException.OriginalEvent = new OriginalEvent();
                        googleException.OriginalEvent.IdOriginal = googleItem.EventId;
                        googleException.OriginalEvent.Href = googleItem.SelfUri.Content;
                        googleException.OriginalEvent.OriginalStartTime = new When();
                        googleException.OriginalEvent.OriginalStartTime.AllDay = outlookItem.AllDayEvent;
                        //if (googleException.Times.Count == 0)
                        //    googleException.Times.Add(
                        //        new When(
                        //            outlookException.ModifiedEvent.Schedule.StartTime,
                        //            outlookException.ModifiedEvent.Schedule.EndTime,
                        //            outlookException.ModifiedEvent.Schedule.AllDayEvent));
                        if (outlookException.Deleted)
                        {
                            /// If Outlook exception is deletion it contains only date of exception
                            /// Otherwise it contains exact time.
                            /// But Google exception always requires time of original occurence. Therefore it's set depending on exception type
                            googleException.OriginalEvent.OriginalStartTime.StartTime = outlookException.OriginalDate;
                            if (!googleException.OriginalEvent.OriginalStartTime.AllDay)
                                googleException.OriginalEvent.OriginalStartTime.StartTime += outlookItem.Start.TimeOfDay;
                            /// Specify the exception is deletion
                            googleException.Status = EventEntry.EventStatus.CANCELED;
                        }
                        else
                        {
                            googleException.OriginalEvent.OriginalStartTime.StartTime = outlookException.OriginalDate;
                            googleException.Status = EventEntry.EventStatus.CONFIRMED;
                            this.SetSubject(googleException, outlookException.AppointmentItem, Target.Google);
                            this.SetDescription(googleException, outlookException.AppointmentItem, Target.Google);
                            this.SetLocation(googleException, outlookException.AppointmentItem, Target.Google);
                            this.SetTime(googleException, outlookException.AppointmentItem, Target.Google);
                        }
                        //googleException = this._googleConnection.Insert(new Uri(this._calendarFeedUrl), googleException);
                        /// Add exception to update batch
                        googleException.BatchData = new GDataBatchEntryData(GDataBatchOperationType.insert);
                        this._googleBatchUpdateFeed.Entries.Add(googleException);
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(outlookException);
                    }
                }
            }
            finally
            {
                Marshal.ReleaseComObject(outlookRecurrence);
            }
        }

        [FieldSetter(Field.ShowTimeAs)]
        private void SetShowTimeAs(EventEntry googleItem, Outlook.AppointmentItem outlookItem, Target target)
        {
            if (target == Target.Google)
                googleItem.EventTransparency = ConvertTo.Google(outlookItem.BusyStatus);
            else
                outlookItem.BusyStatus = ConvertTo.Outlook(googleItem.EventTransparency);
        }

        [FieldSetter(Field.Time)]
        private void SetTime(EventEntry googleItem, Outlook.AppointmentItem outlookItem, Target target)
        {
            /// in order to prevent situation when start is later then end
            /// we should process both times simultaniously and check what time should be changed first
            if (target == Target.Google)
                if (!(outlookItem.IsRecurring && (outlookItem.RecurrenceState == Outlook.OlRecurrenceState.olApptMaster)))
                {
                    //var schedule = new EventSchedule(outlookItem);
                    if (googleItem.Times.Count == 0)
                        googleItem.Times.Add(new When());
                    if (outlookItem.Start >= googleItem.Times[0].EndTime)
                    {
                        googleItem.Times[0].EndTime = outlookItem.End;
                        googleItem.Times[0].StartTime = outlookItem.Start;
                    }
                    else
                    {
                        googleItem.Times[0].StartTime = outlookItem.Start;
                        googleItem.Times[0].EndTime = outlookItem.End;
                    }
                    googleItem.Times[0].AllDay = outlookItem.AllDayEvent; // schedule.AllDayEvent;
                }
                else
                    this.SetRecurrence(outlookItem, googleItem);
            else 
                if (googleItem.Recurrence == null)
                {
                    var schedule = new EventSchedule(googleItem);
                    outlookItem.Start = schedule.StartTime;
                    outlookItem.End = schedule.EndTime;
                    outlookItem.AllDayEvent = schedule.AllDayEvent;
                }
                else
                    this.SetRecurrence(googleItem, outlookItem);

            this.SetReminder(googleItem, outlookItem, target);
        }

        [FieldSetter(Field.Subject)]
        private void SetSubject(EventEntry googleItem, Outlook.AppointmentItem outlookItem, Target target)
        {
            if (target == Target.Google)
            {
                googleItem.Title.Text = outlookItem.Subject;
            }
            else
            {
                outlookItem.Subject = googleItem.Title.Text;
            }
        }

        [FieldSetter(GoogleOutlookSync.Field.Visibility)]
        private void SetVisibility(EventEntry googleItem, Outlook.AppointmentItem outlookItem, Target target)
        {
            if (target == Target.Google)
                googleItem.EventVisibility = ConvertTo.Google(outlookItem.Sensitivity);
            else 
                outlookItem.Sensitivity = ConvertTo.Outlook(googleItem.EventVisibility);
        }
    }
}