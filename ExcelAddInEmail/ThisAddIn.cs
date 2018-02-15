using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Outlook;

namespace SendEmailAddIn
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            SendEmailtoContacts();
        }

        private void SendEmailtoContacts()
        {
            string subjectEmail = "Meeting has been rescheduled.";
            string bodyEmail = "Meeting is one hour later.";
            Microsoft.Office.Tools.Outlook.MAPIFolder sentContacts = (Microsoft.Office.Tools.Outlook.MAPIFolder)
                this.Application.ActiveExplorer().Session.GetDefaultFolder
                (Microsoft.Office.Tools.Outlook.OlDefaultFolders.olFolderContacts);
            foreach (Microsoft.Office.Tools.Outlook.ContactItem contact in sentContacts.Items)
            {
                if (contact.Email1Address.Contains("example.com"))
                {
                    this.CreateEmailItem(subjectEmail, contact
                        .Email1Address, bodyEmail);
                }
            }
        }

        private void CreateEmailItem(string subjectEmail,
               string toEmail, string bodyEmail)
        {
            Microsoft.Office.Tools.Outlook.MailItem eMail = (Microsoft.Office.Tools.Outlook.MailItem)
                this.Application.CreateItem(Microsoft.Office.Tools.Outlook.OlItemType.olMailItem);
            eMail.Subject = subjectEmail;
            eMail.To = toEmail;
            eMail.Body = bodyEmail;
            eMail.Importance = Microsoft.Office.Tools.Outlook.OlImportance.olImportanceLow;
            ((Microsoft.Office.Tools.Outlook._MailItem)eMail).Send();
        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
