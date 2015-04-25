using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Outlook;
using Google.GData.Extensions;
using Google.GData.Calendar;

namespace R.GoogleOutlookSync
{
    internal static class ConvertTo
    {
        /// <summary>
        /// Convert Outlook recipient type to Google
        /// </summary>
        /// <param name="recipientType">Outlook's recipient type</param>
        /// <returns></returns>
        internal static string Google(OlMeetingRecipientType recipientType)
        {
            switch (recipientType)
            {
                case OlMeetingRecipientType.olOptional:
                    return "http://schemas.google.com/g/2005#event.optional";
                case OlMeetingRecipientType.olOrganizer:
                    return "http://schemas.google.com/g/2005#event.organizer";
                case OlMeetingRecipientType.olRequired:
                    return "http://schemas.google.com/g/2005#event.required";
                case OlMeetingRecipientType.olResource:
                default:
                    return "http://schemas.google.com/g/2005#event.attendee";
            }
        }

        /// <summary>
        /// Convert invitee response status from Outlook to Google
        /// </summary>
        /// <param name="outlookStatus"></param>
        /// <returns></returns>
        internal static Who.AttendeeStatus Google(OlResponseStatus outlookStatus)
        {
            var status = new Who.AttendeeStatus();
            switch (outlookStatus)
            {
                case OlResponseStatus.olResponseAccepted:
                    status.Value = Who.AttendeeStatus.EVENT_ACCEPTED;
                    break;
                case OlResponseStatus.olResponseDeclined:
                    status.Value = Who.AttendeeStatus.EVENT_DECLINED;
                    break;
                case OlResponseStatus.olResponseTentative:
                    status.Value = Who.AttendeeStatus.EVENT_TENTATIVE;
                    break;
                default:
                    status.Value = Who.AttendeeStatus.EVENT_INVITED;
                    break;
            }
            return status;
        }

        /// <summary>
        /// Convert Google recipient type to Outlook
        /// </summary>
        /// <param name="googleRel">Google's recipient's relation</param>
        /// <returns></returns>
        /// Probably it will be necessary to replace string with Attendee_Type type 
        internal static OlMeetingRecipientType Outlook(string rel)
        {
            switch (rel)
            {
                case "http://schemas.google.com/g/2005#event.optional":
                    return OlMeetingRecipientType.olOptional;
                case "http://schemas.google.com/g/2005#event.organizer":
                    return OlMeetingRecipientType.olOrganizer;
                case "http://schemas.google.com/g/2005#event.required":
                default:
                    return OlMeetingRecipientType.olRequired;
            }
        }

        /// <summary>
        /// Convert Google invitee response status to Outlook's one
        /// </summary>
        /// <param name="googleStatus"></param>
        /// <returns></returns>
        internal static OlResponseStatus Outlook(Who.AttendeeStatus googleStatus)
        {
            switch (googleStatus.Value)
            {
                case Who.AttendeeStatus.EVENT_ACCEPTED:
                    return OlResponseStatus.olResponseAccepted;
                case Who.AttendeeStatus.EVENT_DECLINED:
                    return OlResponseStatus.olResponseDeclined;
                case Who.AttendeeStatus.EVENT_TENTATIVE:
                    return OlResponseStatus.olResponseTentative;
                default:
                    return OlResponseStatus.olResponseNotResponded;
            }
        }

        /// <summary>
        /// Convert Outlook free/busy status to Google's one.
        /// In fact Google has just two statuses: free and busy.
        /// Therefore OlBusyStatus.olFree will be converted to Google's free
        /// and rest will be converted to Google's busy
        /// </summary>
        /// <param name="outlookBusyStatus"></param>
        /// <returns></returns>
        internal static EventEntry.Transparency Google(OlBusyStatus outlookBusyStatus)
        {
            switch (outlookBusyStatus)
            {
                case OlBusyStatus.olFree:
                    return EventEntry.Transparency.TRANSPARENT;
                default:
                    return EventEntry.Transparency.OPAQUE;
            }
        }

        /// <summary>
        /// Convert Google free/busy appearance to Outlook's one
        /// </summary>
        /// <param name="googleBusyStatus"></param>
        /// <returns></returns>
        internal static OlBusyStatus Outlook(EventEntry.Transparency googleBusyStatus)
        {
            switch (googleBusyStatus.Value)
            {
                case EventEntry.Transparency.OPAQUE_VALUE:
                    return OlBusyStatus.olBusy;
                default:
                    return OlBusyStatus.olFree;
            }
        }

        ///// <summary>
        ///// Convert any calendar event to Google's EventEntry
        ///// </summary>
        ///// <param name="calendarEvent"></param>
        ///// <returns></returns>
        //internal static EventEntry Google(CalendarEvent calendarEvent)
        //{
        //    return calendarEvent.ToGoogle();
        //}

        ///// <summary>
        ///// Convert any calendar event to Outlook's AppointmentItem
        ///// </summary>
        ///// <param name="calendarEvent"></param>
        ///// <returns></returns>
        //internal static AppointmentItem Outlook(CalendarEvent calendarEvent)
        //{
        //    return calendarEvent.ToOutlook();
        //}

        /// <summary>
        /// Converts Outlook item privacy/sensitivity/visibility setting to Google one
        /// </summary>
        /// <param name="olSensitivity"></param>
        /// <returns></returns>
        internal static EventEntry.Visibility Google(OlSensitivity outlookVisibility)
        {
            switch (outlookVisibility)
            {
                case OlSensitivity.olNormal:
                    return EventEntry.Visibility.DEFAULT;
                case OlSensitivity.olPrivate:
                    return EventEntry.Visibility.PRIVATE;
                case OlSensitivity.olConfidential:
                    return EventEntry.Visibility.CONFIDENTIAL;
                default:
                    return EventEntry.Visibility.DEFAULT;
            }
        }

        /// <summary>
        /// Converts Google item privacy setting to Outlook's one
        /// </summary>
        /// <param name="googleVisibility"></param>
        /// <returns></returns>
        internal static OlSensitivity Outlook(EventEntry.Visibility googleVisibility)
        {
            switch (googleVisibility.Value)
            {
                case EventEntry.Visibility.CONFIDENTIAL_VALUE:
                    return OlSensitivity.olConfidential;
                case EventEntry.Visibility.PRIVATE_VALUE:
                    return OlSensitivity.olPrivate;
                case EventEntry.Visibility.PUBLIC_VALUE:
                default:
                    return OlSensitivity.olNormal;
            }
        }
    }
}
