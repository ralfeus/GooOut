using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Google.GData.Calendar;
using Google.GData.Extensions;
using Microsoft.Office.Interop.Outlook;

namespace R.GoogleOutlookSync
{
    class Attendee
    {
        internal string Email { get; set; }
        internal bool Required { get; set; }
        internal OlResponseStatus Status { get; set; }

        internal Attendee(string email, int type, OlResponseStatus status)
        {
            this.Email = email;
            this.Required = (OlMeetingRecipientType)type == OlMeetingRecipientType.olRequired;
            this.Status = status;
        }

        public Attendee(string email, Who.AttendeeType type, Who.AttendeeStatus status)
        {
            this.Email = email;
            this.Required = type.Value == Who.AttendeeType.EVENT_REQUIRED;
            switch (status.Value)
            {
                case Who.AttendeeStatus.EVENT_ACCEPTED:
                    this.Status = OlResponseStatus.olResponseAccepted;
                    break;
                case Who.AttendeeStatus.EVENT_DECLINED:
                    this.Status = OlResponseStatus.olResponseDeclined;
                    break;
                case Who.AttendeeStatus.EVENT_TENTATIVE:
                    this.Status = OlResponseStatus.olResponseTentative;
                    break;
                default:
                    this.Status = OlResponseStatus.olResponseNotResponded;
                    break;
            }
        }
    }
}
