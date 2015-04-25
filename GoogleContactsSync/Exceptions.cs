using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace R.GoogleOutlookSync
{
    class CannotDefineSynchronizationTargetException : Exception
    {
    }

    class EventScheduleIsNotSetException : Exception 
    {
        internal EventScheduleIsNotSetException(string message)
            : base(message)
        { }
    }

    class NoDataToSyncronizeSpecifiedException : Exception
    {
        internal NoDataToSyncronizeSpecifiedException(string message)
            : base(message)
        { }
    }

    class UnsynchronizableItemException : Exception
    {
        internal Target ItemType { get; set; }

        internal UnsynchronizableItemException(Target target, string message)
            : base(message)
        {
            this.ItemType = target;
        }
    }

    class IncompatibleRecurrencePatternException : UnsynchronizableItemException
    {
        internal IncompatibleRecurrencePatternException(Target target, string message)
            : base(target, message)
        { }
    }

    class GoogleConnectionException : Exception
    {
        public GoogleConnectionException(Exception lastError)
            : base(Properties.Resources.Error_GoogleConnectionFailure, lastError)
        { }
    }

    class OutlookConnectionException : Exception 
    {
        public OutlookConnectionException(Exception lastError)
            : base(Properties.Resources.Error_OutlookConnectionFailure, lastError)
        { }
        public OutlookConnectionException(string message)
            : base(message)
        { }
    }

    class ProfileNotConfiguredException : OutlookConnectionException
    {
        public ProfileNotConfiguredException(string profile)
            : base(String.Format(Properties.Resources.Error_ProfileNotConfigured, profile))
        { }
    }
}
