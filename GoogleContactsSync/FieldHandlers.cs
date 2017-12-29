using System;
using System.Reflection;
using Google.Apis.Calendar.v3.Data;

namespace R.GoogleOutlookSync
{
    internal delegate bool ComparerDelegate(Event googleItem, object outlookItem);
    internal delegate void SetterDelegate(Event googleItem, object outlookItem, Target target);

    //TODO: Replace MethodInfo with functional types (I believe it should work faster than reflection)
    internal class FieldHandlers
    {
        internal MethodInfo Comparer { get; private set; }
        //internal MethodInfo Getter { get; private set; }
        internal MethodInfo Setter { get; private set; }
        internal Func<Event, object, bool> comparer { get; private set; }
        internal Action<Event, object, Target> setter { get; private set; }

        internal FieldHandlers(MethodInfo comparer/*, MethodInfo getter*/, MethodInfo setter)
        {
            this.Comparer = comparer;
            //this.Getter = getter;
            this.Setter = setter;
        }

        internal FieldHandlers(Func<Event, object, bool> comparer, Action<Event, object, Target> setter)
        {
            this.comparer = comparer;
            this.setter = setter;
        }
    }
}
