using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Google.GData.Client;
using Google.GData.Extensions;
using Microsoft.Office.Interop.Outlook;
using System.Net;

namespace R.GoogleOutlookSync
{
    internal static class GoogleUtilities
    {
        internal static string GetItemID(AtomEntry item)
        {
            return item.Id.AbsoluteUri;
        }

        internal static DateTime GetLastModificationTime(AtomEntry item)
        {
            return item.Updated;
        }

        internal static string GetOutlookID(AtomEntry item)
        {
            var result =
                ((ExtendedProperty)item.ExtensionElements.FirstOrDefault(
                    element =>
                        element.XmlName == "extendedProperty" &&
                        ((ExtendedProperty)element).Name == Properties.Settings.Default.ExtendedPropertyName_OutlookIDInGoogleItem)
                );
            if (result == null)
                return null;
            else
                return result.Value;
        }

        internal static void SetOutlookID(AtomEntry item, string outlookID)
        {
            item.ExtensionElements.Add(new ExtendedProperty(outlookID, Properties.Settings.Default.ExtendedPropertyName_OutlookIDInGoogleItem));
        }

        internal static void RemoveOutlookID(AtomEntry googleItem)
        {
            var outlookIDs =
                (from outlookIDProperty in googleItem.ExtensionElements
                where
                    outlookIDProperty is ExtendedProperty &&
                    ((ExtendedProperty)outlookIDProperty).Name == Properties.Settings.Default.ExtendedPropertyName_OutlookIDInGoogleItem
                select outlookIDProperty).ToList<IExtensionElementFactory>();
            foreach (var outlookID in outlookIDs)
                googleItem.ExtensionElements.Remove(outlookID);
        }

        public static T TryDo<T>(Func<T> function)
        {
            System.Exception lastError = null;
            var attemptsAmount = Properties.Settings.Default.AttemptsAmount;
            do
            {
                try
                {
                    return function();
                }
                catch (WebException exc)
                {
                    Logger.Log("During Google operation an error has occured. Error details:\r\n" + ErrorHandler.BuildExceptionDescription(exc), EventType.Debug);
                    --attemptsAmount;
                    lastError = exc;
                }
            } while (attemptsAmount > 0);
            throw new GoogleConnectionException(lastError);
        }

        public static void TryDo(System.Action function)
        {
            TryDo<object>(() => { function(); return null; });
        }
    }
}
