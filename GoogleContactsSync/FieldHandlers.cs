using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using Google.GData.Client;

namespace R.GoogleOutlookSync
{
    internal delegate bool ComparerDelegate(AtomEntry googleItem, object outlookItem);
    internal delegate void SetterDelegate(AtomEntry googleItem, object outlookItem, Target target);

    //TODO: Replace MethodInfo with functional types (I believe it should work faster than reflection)
    internal class FieldHandlers
    {
        internal MethodInfo Comparer { get; private set; }
        //internal MethodInfo Getter { get; private set; }
        internal MethodInfo Setter { get; private set; }
        internal Func<AtomEntry, object, bool> comparer { get; private set; }
        internal Action<AtomEntry, object, Target> setter { get; private set; }

        internal FieldHandlers(MethodInfo comparer/*, MethodInfo getter*/, MethodInfo setter)
        {
            this.Comparer = comparer;
            //this.Getter = getter;
            this.Setter = setter;
        }

        internal FieldHandlers(Func<AtomEntry, object, bool> comparer, Action<AtomEntry, object, Target> setter)
        {
            this.comparer = comparer;
            this.setter = setter;
        }
    }
}
