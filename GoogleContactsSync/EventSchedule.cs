using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Google.GData.Calendar;
using Microsoft.Office.Interop.Outlook;
using System.Text.RegularExpressions;
using Google.GData.Extensions;
using System.Runtime.InteropServices;

namespace R.GoogleOutlookSync
{
    class EventSchedule
    {
        internal bool AllDayEvent { get; set; }
        internal DateTime EndTime { get; set; }
        internal EventRecurrence RecurrencePattern { get; set; }
        internal DateTime StartTime { get; set; }
        internal string TimeZone { get; set; }

        internal EventSchedule(EventEntry googleItem, List<EventEntry> recurrenceExceptions = null)
        {
            if (googleItem.Recurrence == null)
            {
                this.StartTime = googleItem.Times[0].StartTime;
                this.EndTime = googleItem.Times[0].EndTime;
                this.AllDayEvent = googleItem.Times[0].AllDay;
            }
            else
            {

                var startTimeMatch = Regex.Match(googleItem.Recurrence.Value, "DTSTART(;TZID=.+)?(;VALUE=DATE)?:(\\d{4})(\\d{2})(\\d{2})(T(\\d{2})(\\d{2})(\\d{2})Z?)?\r\n");
                var endTimeMatch = Regex.Match(googleItem.Recurrence.Value, "DTEND(;TZID=.+)?(;VALUE=DATE)?:(\\d{4})(\\d{2})(\\d{2})(T(\\d{2})(\\d{2})(\\d{2})Z?)?\r\n");
                this.AllDayEvent = startTimeMatch.Groups[2].Value == ";VALUE=DATE";
                if (this.AllDayEvent)
                {
                    this.StartTime = new DateTime(
                        Convert.ToInt32(startTimeMatch.Groups[3].Value),
                        Convert.ToInt32(startTimeMatch.Groups[4].Value),
                        Convert.ToInt32(startTimeMatch.Groups[5].Value));
                    this.EndTime = new DateTime(
                        Convert.ToInt32(endTimeMatch.Groups[3].Value),
                        Convert.ToInt32(endTimeMatch.Groups[4].Value),
                        Convert.ToInt32(endTimeMatch.Groups[5].Value));
                }
                else
                {
                    this.StartTime = new DateTime(
                        Convert.ToInt32(startTimeMatch.Groups[3].Value),
                        Convert.ToInt32(startTimeMatch.Groups[4].Value),
                        Convert.ToInt32(startTimeMatch.Groups[5].Value),
                        Convert.ToInt32(startTimeMatch.Groups[7].Value),
                        Convert.ToInt32(startTimeMatch.Groups[8].Value),
                        Convert.ToInt32(startTimeMatch.Groups[9].Value));
                    this.EndTime = new DateTime(
                        Convert.ToInt32(endTimeMatch.Groups[3].Value),
                        Convert.ToInt32(endTimeMatch.Groups[4].Value),
                        Convert.ToInt32(endTimeMatch.Groups[5].Value),
                        Convert.ToInt32(endTimeMatch.Groups[7].Value),
                        Convert.ToInt32(endTimeMatch.Groups[8].Value),
                        Convert.ToInt32(endTimeMatch.Groups[9].Value));
                    this.TimeZone = startTimeMatch.Groups[1].Value.Substring(6);
                }
                this.RecurrencePattern = new EventRecurrence(googleItem.Recurrence, recurrenceExceptions);
            }
        }

        internal EventSchedule(AppointmentItem outlookItem)
        {
            this.StartTime = outlookItem.Start;
            this.EndTime = outlookItem.End;
            this.AllDayEvent = outlookItem.AllDayEvent;
            this.TimeZone = outlookItem.StartTimeZone.ID;
            if (outlookItem.IsRecurring && (outlookItem.RecurrenceState == OlRecurrenceState.olApptMaster))
                this.RecurrencePattern = new EventRecurrence(outlookItem.GetRecurrencePattern());
        }

        public override bool Equals(object obj)
        {
            var target = (EventSchedule)obj;
            return
                (this.StartTime == target.StartTime) &&
                (this.EndTime == target.EndTime) &&
                (this.AllDayEvent == target.AllDayEvent) &&
                (this.RecurrencePattern == target.RecurrencePattern);
        }

        public override int GetHashCode()
        {
            return this.StartTime.GetHashCode() ^ this.EndTime.GetHashCode();
        }

        /// <summary>
        /// Generates Google recurrence pattern
        /// Because Google recurrence pattern contains scheduling information (like start time and end time) 
        /// it's better to delegate such work to EventSchedule object
        /// </summary>
        internal Recurrence GetGoogleRecurrence()
        {
            string result = "";
            if (this.AllDayEvent)
            {
                result = String.Format("DTSTART;VALUE=DATE:{0}\r\n", this.StartTime.ToString("yyyyMMdd"));
                result += String.Format("DTEND;VALUE=DATE:{0}\r\n", this.EndTime.ToString("yyyyMMdd"));
            }
            else
            {
                result = String.Format("DTSTART;TZID={0}:{1}\r\n", TimeZones.GetTZ(this.TimeZone), this.StartTime.ToString("yyyyMMddTHHmmss"));
                result += String.Format("DTEND;TZID={0}:{1}\r\n", TimeZones.GetTZ(this.TimeZone), this.EndTime.ToString("yyyyMMddTHHmmss"));
            }
            /// RFC 2445 tells VTIMEZONE must be set if event start is defined by DTSTART and DTEND elements
            //var timeZone = String.Format("VTIMEZONE={0};", "");

            /// It's necessary to define week interval in such way in order to add week interval
            /// (if any) to BYDAY element
            var instance = this.RecurrencePattern.WeekInterval != 0 ? this.RecurrencePattern.WeekInterval.ToString() : "";
            var byDay = this.RecurrencePattern.DayOfWeekMask != 0 ? String.Format("BYDAY={0};", Regex.Replace(this.RecurrencePattern.DayOfWeekMask.ToString(), "((\\w{2})\\w+)", instance + "$2").ToUpper().Replace(" ", "")) : "";
            var byMonthDay = this.RecurrencePattern.DayOfMonth != 0 ? String.Format("BYMONTHDAY={0};", this.RecurrencePattern.DayOfMonth) : "";
            var count = this.RecurrencePattern.Count != 0 ? String.Format("COUNT={0};", this.RecurrencePattern.Count) : "";
            var frequency = String.Format("FREQ={0};", this.RecurrencePattern.Frequency.ToString().ToUpper());
            var interval = this.RecurrencePattern.Interval > 1 ? String.Format("INTERVAL={0};", this.RecurrencePattern.Interval) : "";
            var month = this.RecurrencePattern.Frequency == RecurrenceFrequency.Yearly ? String.Format("BYMONTH={0};", this.RecurrencePattern.Month) : "";
            var until = this.RecurrencePattern.EndMethod == EndBy.Date ? String.Format("UNTIL={0}", this.RecurrencePattern.EndDate.ToString("yyyyMMddTHHmmssZ")) : "";

            var rule = String.Format("RRULE:{0}{1}{2}{3}{4}{5}{6}", frequency, interval, count, month, byDay, byMonthDay, until);
            result += rule.Substring(0, rule.Length - 1);

            var recurrence = new Recurrence();
            recurrence.Value = result;
            return recurrence;
        }

        private Google.GData.Extensions.RecurrenceException GetGoogleRecurrenceExceptions()
        {
            throw new NotImplementedException();
        }

        internal void GetOutlookRecurrence(AppointmentItem outlookItem)
        {
            RecurrencePattern outlookRec = outlookItem.GetRecurrencePattern();
            if (!this.AllDayEvent)
            {
                /// Set time of the even taking consideration difference of time zones. 
                /// This makes sense only in case the event isn't all day one
                var googleTimeZone = TimeZoneInfo.FindSystemTimeZoneById(TimeZones.GetWindowsTimeZone(this.TimeZone));
                var outlookTimeZone = TimeZoneInfo.FindSystemTimeZoneById(outlookItem.StartTimeZone.ID);
                outlookRec.StartTime = TimeZoneInfo.ConvertTime(this.StartTime, googleTimeZone, outlookTimeZone);
            }
 
            outlookRec.Duration = (int)(this.EndTime - this.StartTime).TotalMinutes;
            outlookRec.RecurrenceType = (OlRecurrenceType)this.RecurrencePattern.Frequency;
            if ((this.RecurrencePattern.Frequency == RecurrenceFrequency.Weekly) || (this.RecurrencePattern.Frequency == RecurrenceFrequency.MonthlyNth))
                outlookRec.DayOfWeekMask = (OlDaysOfWeek)this.RecurrencePattern.DayOfWeekMask;
            if (this.RecurrencePattern.Interval > 1)
                if ((outlookRec.RecurrenceType != OlRecurrenceType.olRecursYearly) || (Properties.Settings.Default.OutlookVersion >= 14))
                    outlookRec.Interval = this.RecurrencePattern.Interval;
            outlookRec.PatternStartDate = this.RecurrencePattern.StartDate;
            if (this.RecurrencePattern.EndMethod == EndBy.Date)
                outlookRec.PatternEndDate = this.RecurrencePattern.EndDate;
            if (this.RecurrencePattern.EndMethod == EndBy.OccurencesCount)
                outlookRec.Occurrences = this.RecurrencePattern.Count;
            if ((this.RecurrencePattern.Frequency == RecurrenceFrequency.Monthly) || (this.RecurrencePattern.Frequency == RecurrenceFrequency.Yearly))
                outlookRec.DayOfMonth = this.RecurrencePattern.DayOfMonth;
            /// As it was found out RecurrencePattern.Instance is valid for olRecursMonthNth only. 
            /// Not sure yet whether Google provides such frequency model. If not - it will be just omited
            //outlookRec.Instance = this.WeekInterval;
            if (this.RecurrencePattern.Frequency == RecurrenceFrequency.Yearly)
                outlookRec.MonthOfYear = this.RecurrencePattern.Month;
            //outlookRec.NoEndDate = this.EndMethod == EndBy.NoEnd;

        }

        internal void ToGoogle(EventEntry googleItem)
        {
            if (this.RecurrencePattern != null)
            {
                googleItem.Recurrence = this.GetGoogleRecurrence();
                googleItem.RecurrenceException = this.GetGoogleRecurrenceExceptions();
            }
            else
            {
                if (googleItem.Times.Count != 0)
                    googleItem.Times.Clear();
                googleItem.Times.Add(new When(this.StartTime, this.EndTime, this.AllDayEvent));
            }
        }

        internal void ToOutlook(AppointmentItem outlookItem)
        {
            
        }
    }
}
