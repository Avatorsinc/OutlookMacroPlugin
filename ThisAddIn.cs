using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;

namespace OutlookMacroPlugin
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Add your startup code here
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Add your shutdown code here
        }

        public void MoveSelectedToTreatedSD()
        {
            ChangeSelectedSubject("Servicedesk - Shared Mailbox");
            MoveSelectedMailsTo("1 - Treated mail", "Servicedesk - Shared Mailbox");
        }

        public void Sign()
        {
            ChangeSelectedSubject("Servicedesk - Shared Mailbox");
        }

        public void ServiceNow()
        {
            ChangeSelectedSubject("Servicedesk - Shared Mailbox");
            MoveSelectedMailsTo("1.2 - Sent to ServiceNow", "Servicedesk - Shared Mailbox");
        }

        public void IBM()
        {
            ChangeSelectedSubject("Servicedesk - Shared Mailbox");
            MoveSelectedMailsTo("3 - Feedback 3rd. parties\\IBM", "Servicedesk - Shared Mailbox");
        }

        public void EG()
        {
            ChangeSelectedSubject("Servicedesk - Shared Mailbox");
            MoveSelectedMailsTo("3 - Feedback 3rd. parties\\EG", "Servicedesk - Shared Mailbox");
        }

        public void NCR()
        {
            ChangeSelectedSubject("Servicedesk - Shared Mailbox");
            MoveSelectedMailsTo("3 - Feedback 3rd. parties\\NCR", "Servicedesk - Shared Mailbox");
        }

        public void Wincor()
        {
            ChangeSelectedSubject("Servicedesk - Shared Mailbox");
            MoveSelectedMailsTo("3 - Feedback 3rd. parties\\Wincor", "Servicedesk - Shared Mailbox");
        }

        public void ND()
        {
            ChangeSelectedSubject("Servicedesk - Shared Mailbox");
            MoveSelectedMailsTo("3 - Feedback 3rd. parties\\NetDesign", "Servicedesk - Shared Mailbox");
        }

        public void HPE()
        {
            ChangeSelectedSubject("Servicedesk - Shared Mailbox");
            MoveSelectedMailsTo("5.1 - HPE Incident Reports", "Servicedesk - Shared Mailbox");
        }

        public void SIMA()
        {
            ChangeSelectedSubject("Servicedesk - Shared Mailbox");
            MoveSelectedMailsTo("5 - SIMA Reports", "Servicedesk - Shared Mailbox");
        }

        public void MIM()
        {
            ChangeSelectedSubject("Servicedesk - Shared Mailbox");
            MoveSelectedMailsTo("8 - MIM", "Servicedesk - Shared Mailbox");
        }

        public void Lexmark()
        {
            ChangeSelectedSubject("Servicedesk - Shared Mailbox");
            MoveSelectedMailsTo("3 - Feedback 3rd. parties\\Lexmark", "Servicedesk - Shared Mailbox");
        }

        public void Atea()
        {
            ChangeSelectedSubject("Servicedesk - Shared Mailbox");
            MoveSelectedMailsTo("3 - Feedback 3rd. parties\\Atea", "Servicedesk - Shared Mailbox");
        }

        public void Ricoh()
        {
            ChangeSelectedSubject("Servicedesk - Shared Mailbox");
            MoveSelectedMailsTo("3 - Feedback 3rd. parties\\Ricoh", "Servicedesk - Shared Mailbox");
        }

        private void ChangeSelectedSubject(string sharedMailbox)
        {
            Application outlookApp = Application;
            NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");
            Recipient recipient = outlookNamespace.CreateRecipient(sharedMailbox);
            recipient.Resolve();

            if (recipient.Resolved)
            {
                var currentUser = outlookNamespace.CurrentUser.Name;
                var initials = GetInitials(currentUser);

                Explorer explorer = outlookApp.ActiveExplorer();
                Selection selection = explorer.Selection;

                for (int i = 1; i <= selection.Count; i++)
                {
                    MailItem mailItem = selection[i] as MailItem;
                    if (mailItem != null)
                    {
                        string[] subjectParts = mailItem.Subject.Split(' ');
                        if (subjectParts[0] != initials)
                        {
                            mailItem.Subject = initials + " " + mailItem.Subject;
                            mailItem.UnRead = false;
                            mailItem.Save();
                        }
                    }
                }
            }
        }

        private void MoveSelectedMailsTo(string destinationFolder, string sharedMailbox)
        {
            Application outlookApp = Application;
            NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");
            Recipient recipient = outlookNamespace.CreateRecipient(sharedMailbox);
            recipient.Resolve();

            if (recipient.Resolved)
            {
                Explorer explorer = outlookApp.ActiveExplorer();
                Selection selection = explorer.Selection;
                MAPIFolder destination = GetFolder(sharedMailbox + "\\" + destinationFolder);

                for (int i = 1; i <= selection.Count; i++)
                {
                    MailItem mailItem = selection[i] as MailItem;
                    if (mailItem != null && destination != null)
                    {
                        mailItem.Move(destination);
                    }
                }
            }
        }

        private MAPIFolder GetFolder(string folderPath)
        {
            Application outlookApp = Application;
            NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");

            string[] folderNames = folderPath.Split(new string[] { "\\" }, StringSplitOptions.None);
            MAPIFolder folder = outlookNamespace.Folders[folderNames[0]] as MAPIFolder;

            if (folder != null)
            {
                for (int i = 1; i < folderNames.Length; i++)
                {
                    Folders subFolders = folder.Folders;
                    folder = subFolders[folderNames[i]] as MAPIFolder;
                    if (folder == null) break;
                }
            }

            return folder;
        }

        private string GetInitials(string fullName)
        {
            var nameParts = fullName.Split(' ');
            string initials = string.Empty;

            if (nameParts.Length > 1)
            {
                initials = nameParts[0].Substring(0, 2) + nameParts[1].Substring(0, 1);
            }
            else
            {
                initials = nameParts[0].Substring(0, 3);
            }

            return initials.ToUpper();
        }

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
    }

}
