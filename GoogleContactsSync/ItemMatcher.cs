using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Google.GData.Client;

namespace R.GoogleOutlookSync
{
    internal class ItemMatcher 
    {
        internal AtomEntry Google { get; set; }
        internal object Outlook { get; set; }
        internal SyncAction SyncAction { get; set; }

        internal ItemMatcher(AtomEntry googleItem, object outlookItem)
        {
            this.Google = googleItem;
            this.Outlook = outlookItem;
            this.SyncAction = new SyncAction();
        }
    }
}
