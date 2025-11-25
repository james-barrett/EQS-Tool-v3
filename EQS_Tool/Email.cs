using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace EQS_Tool
{
    internal class Email
    {
        public static void Send(string addresses, string filePath, string fileName, string subject)
        {
            try
            {
                var outlookApp = new Outlook.Application();
                var mailItem = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);

                // Split the addresses string into an array of email addresses
                var addressList = addresses.Split(',').Select(address => address.Trim());

                foreach (var address in addressList)
                {
                    var recipient = mailItem.Recipients.Add(address);
                    recipient.Resolve();
                    Marshal.ReleaseComObject(recipient); // Release the recipient object after each addition
                }

                mailItem.Subject = subject;
                mailItem.Body = "Please find attached.";

                int position = (int)mailItem.Body.Length + 1;
                int attachmentType = (int)Outlook.OlAttachmentType.olByValue;

                var attachment = mailItem.Attachments.Add(filePath, attachmentType, position, fileName);

                mailItem.Save();
                mailItem.Send();

                // Release COM objects
                Marshal.ReleaseComObject(attachment);
                Marshal.ReleaseComObject(mailItem);
                Marshal.ReleaseComObject(outlookApp);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception caught: {ex.Message}");
            }
        }
    }
}
