using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Diagnostics;
using System.Windows.Forms;
using System.IO;

namespace WebTechOutlkInMsg
{
    public partial class ThisAddIn
    {

        public static string VER = "1.5.3.7";
        public static string LOG_FILE_PATH = @"C:\WebTechOutlkInMsg\log.txt";


        private int proccedCounter = 0;

        Outlook.NameSpace outlookNameSpace;
        Outlook.MAPIFolder inbox;
        Outlook.Items items;
        Outlook.Inspectors inspectors;
        Outlook.Explorer activeExplorer;

        private String subjectFilter = String.Empty;
        private String bodyFilter = String.Empty;
        private String url2open = String.Empty;
        private List<String> incomingItems = new List<String>();
        private List<String> incomingItemsPoped = new List<String>();



        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                // MessageBox.Show($"Version {VER}", "WebTech", MessageBoxButtons.OK, MessageBoxIcon.Information);
                outlookNameSpace = this.Application.GetNamespace("MAPI");
                inbox = outlookNameSpace.GetDefaultFolder(
                        Microsoft.Office.Interop.Outlook.
                        OlDefaultFolders.olFolderInbox);

                // 4 incoming event
                items = inbox.Items;
                items.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(items_ItemAdd);

                // 4 open mail event
                inspectors = this.Application.Inspectors;
                inspectors.NewInspector += new Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);

                // 4 selection event
                this.activeExplorer = this.Application.ActiveExplorer();
                this.activeExplorer.SelectionChange += new Outlook.ExplorerEvents_10_SelectionChangeEventHandler(Explorer_SelectionChange);

                this.subjectFilter = readLineFromFile(@"C:\WebTechOutlkInMsg\subject.txt");
                this.bodyFilter = readLineFromFile(@"C:\WebTechOutlkInMsg\body.txt");
                this.url2open = readLineFromFile(@"C:\WebTechOutlkInMsg\url.txt");

                logit($"subject filter: {this.subjectFilter}");
                logit($"body filter: {this.bodyFilter}");
                logit($"URL to open: {this.url2open}");
                logit($"VER : {VER}");
                logit($"Application.Version {this.Application.Version}");
                logit($"Application.Name {this.Application.Name}");

                if (File.Exists(LOG_FILE_PATH))
                {
                    File.Delete(LOG_FILE_PATH);
                    logit($"LOG File deleted");
                }
            }
            catch (Exception ee)
            {
                logit($"Exception WebTechOutlkInMsg: {ee.Message}");
                logit($"Source: {ee.Source}");
                logit($"StackTrace: {ee.StackTrace}");

                if (ee.InnerException != null && ee.InnerException.StackTrace != null)
                    logit($"InnerException.StackTrace: {ee.InnerException.StackTrace}");

                throw;
            }

        }

        private string readLineFromFile(string fileName)
        {
            try
            {
                System.IO.StreamReader file = new System.IO.StreamReader(fileName);
                return file.ReadLine();

            }
            catch (Exception e)
            {
                EventLog.WriteEntry("Application", $"Error reading from file '{fileName}': {e.StackTrace}");
                return String.Empty;
            }
        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        private void items_ItemAdd(object Item)
        {
            Boolean processed = false;

            Outlook.MailItem mail = (Outlook.MailItem)Item;
            if (Item != null)
            {
                if (mail.MessageClass == "IPM.Note")
                {
                    if (!string.IsNullOrEmpty(this.subjectFilter) &&
                        mail.Subject.Contains(this.subjectFilter))
                    {
                        logit($"Recieved message subject!!");

                        mail.Subject = "Processed by WebTechOutlkInMsg - subject: " + mail.Subject;
                        mail.Save();
                        this.proccedCounter++;
                        processed = true;

                    }
                    else if (!string.IsNullOrEmpty(this.bodyFilter) && mail.Body.Contains(this.bodyFilter))
                    {
                        logit($"Recieved message body!! ");
                        mail.Subject = "Processed by WebTechOutlkInMsg - body: " + mail.Subject;
                        mail.Save();
                        this.proccedCounter++;
                        processed = true;
                    }
                    if (processed)
                    {
                        logit($"Adde {mail.EntryID} mail id - processed so far {this.proccedCounter}  mails");
                        this.incomingItems.Add(mail.EntryID);
                    }
                }
            }
        }

        private void Explorer_SelectionChange()
        {
            logit($"Explorer_SelectionChange event");
            if (this.Application.ActiveExplorer().Selection.Count == 1)
            {
                Outlook.MailItem item = this.Application.ActiveExplorer().Selection[1] as Outlook.MailItem;
                this.ProcessMail(item);
            } else
                logit($"Explorer_SelectionChange event - Selection.Count > 1 => {this.Application.ActiveExplorer().Selection.Count}");
        }

        void Inspectors_NewInspector(Microsoft.Office.Interop.Outlook.Inspector Inspector)
        {
            if (Inspector.CurrentItem is Outlook.MailItem)
            {
                Outlook.MailItem mail = Inspector.CurrentItem as Outlook.MailItem;
                this.ProcessMail(mail);
            }
            else
                logit($"Inspectors_NewInspector event - Inspector.CurrentItem is NOT Outlook.MailItem");
        }

        private void ProcessMail(Outlook.MailItem item)
        {
            logit($"ProcessMail event");

            if (item != null)
            {
                if (this.incomingItems.Contains(item.EntryID))
                {
                    if (!this.incomingItemsPoped.Contains(item.EntryID))
                    {
                        logit($"Processed item '{item.EntryID}' opened.");
                        // MessageBox.Show($"Processed item '{item.EntryID}' opened.", "WebTech", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        this.incomingItemsPoped.Add(item.EntryID);
                        runApp(item.EntryID);
                    }
                    {
                        logit($"ProcessMail-Item already process -> '{item.EntryID}'.");
                    }
                } else
                {
                    logit($"ProcessMail event - this.incomingItems NOT Contains {item.EntryID}");
                }
            }
            else
            {
                logit($"ProcessMail item is null");
            }
        }

        private void runApp(string id)
        {
            string url = this.url2open.Replace("{KEY}", id);
            logit($"Openning url:  {url}");

            using (Process myProcess = new Process())
            {
                myProcess.StartInfo.UseShellExecute = false;
                myProcess.StartInfo.FileName = @"powershell.exe";
                myProcess.StartInfo.Arguments = @"C:\WebTechOutlkInMsg\openBrowser.ps1 " + id;
                myProcess.StartInfo.CreateNoWindow = true;
                myProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                myProcess.Start();
            }


            System.Diagnostics.Process.Start(url);
            if (true) return;

            ProcessStartInfo start = new ProcessStartInfo();
            // Enter in the command line arguments, everything you would enter after the executable name itself
            start.Arguments = id;
            // Enter the executable to run, including the complete path
            start.FileName = "";
            // Do you want to show a console window?
            start.WindowStyle = ProcessWindowStyle.Hidden;
            start.CreateNoWindow = true;
            int exitCode;


            // Run the external process & wait for it to finish
            using (Process proc = Process.Start(start))
            {
                proc.WaitForExit();

                // Retrieve the app's exit code
                exitCode = proc.ExitCode;
            }
        }

        private void logit(string msg)
        {
            EventLog.WriteEntry("Application", $"WebTechOutlkInMsg: {msg}");
            Debug.WriteLine(msg);

            using (StreamWriter sw = File.AppendText(LOG_FILE_PATH))
            {
                sw.WriteLine($"{DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss")}:{msg}");
            }
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
