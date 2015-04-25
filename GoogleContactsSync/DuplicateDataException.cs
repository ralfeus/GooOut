using System;
using System.Collections.Generic;
using System.Text;

namespace R.GoogleOutlookSync
{
 
    [global::System.Serializable]
    public class DuplicateDataException : Exception
    {
        //
        // For guidelines regarding the creation of new exception types, see
        //    http://msdn.microsoft.com/library/default.asp?url=/library/en-us/cpgenref/html/cpconerrorraisinghandlingguidelines.asp
        // and
        //    http://msdn.microsoft.com/library/default.asp?url=/library/en-us/dncscol/html/csharp07192001.asp
        //

        public DuplicateDataException() : base() { }
        public DuplicateDataException(string message) : base(message) { }
        public DuplicateDataException(string message, Exception inner) : base(message, inner) { }
        protected DuplicateDataException(
          System.Runtime.Serialization.SerializationInfo info,
          System.Runtime.Serialization.StreamingContext context)
            : base(info, context) { }
    }
}
