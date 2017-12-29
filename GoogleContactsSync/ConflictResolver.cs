using System;

namespace R.GoogleOutlookSync
{
    class ConflictResolver //: IConflictResolver
    {
        private ConflictResolverForm _form;

        public ConflictResolver()
        {
            _form = new ConflictResolverForm();
        }
        public ConflictResolver(ConflictResolverForm form)
        {
            _form = form;
        }



        #region IConflictResolver Members

        //public ConflictResolution Resolve(Microsoft.Office.Interop.Outlook.ContactItem outlookContact, Contact googleContact)
        //{
        //    string name = googleContact.Title;
        //    if (string.IsNullOrEmpty(name))
        //        name = googleContact.Name.FullName;
        //    if (string.IsNullOrEmpty(name) && googleContact.Organizations.Count > 0)
        //        name = googleContact.Organizations[0].Name;
        //    if (string.IsNullOrEmpty(name) && googleContact.Emails.Count > 0)
        //        name = googleContact.Emails[0].Address;

        //    _form.messageLabel.Text =
        //        "Both the outlook contact and the google contact \"" + name +
        //        "\" have been changed. Choose which you would like to keep.";

        //    switch (_form.ShowDialog())
        //    {
        //        case System.Windows.Forms.DialogResult.Ignore:
        //            // skip
        //            return ConflictResolution.Skip;
        //        case System.Windows.Forms.DialogResult.Cancel:
        //            // cancel
        //            return ConflictResolution.Cancel;
        //        case System.Windows.Forms.DialogResult.No:
        //            // google wins
        //            return ConflictResolution.GoogleWins;
        //        case System.Windows.Forms.DialogResult.Yes:
        //            // outlook wins
        //            return ConflictResolution.OutlookWins;
        //        default:
        //            throw new Exception();
        //    }
        //}

        //public ConflictResolution Resolve(Microsoft.Office.Interop.Outlook.NoteItem outlookNote, Document googleNote)
        //{
        //    string name = googleNote.Title;
            
        //    _form.messageLabel.Text =
        //        "Both the outlook note and the google note \"" + name +
        //        "\" have been changed. Choose which you would like to keep.";

        //    switch (_form.ShowDialog())
        //    {
        //        case System.Windows.Forms.DialogResult.Ignore:
        //            // skip
        //            return ConflictResolution.Skip;
        //        case System.Windows.Forms.DialogResult.Cancel:
        //            // cancel
        //            return ConflictResolution.Cancel;
        //        case System.Windows.Forms.DialogResult.No:
        //            // google wins
        //            return ConflictResolution.GoogleWins;
        //        case System.Windows.Forms.DialogResult.Yes:
        //            // outlook wins
        //            return ConflictResolution.OutlookWins;
        //        default:
        //            throw new Exception();
        //    }
        //}

        #endregion
    }
}
