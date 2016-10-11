using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace DeleteFromFolder
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            ((Outlook.ApplicationEvents_11_Event)Application).Quit += new Outlook.ApplicationEvents_11_QuitEventHandler(ThisAddIn_Quit);
            ((Outlook.ApplicationEvents_11_Event)Application).OptionsPagesAdd += new Outlook.ApplicationEvents_11_OptionsPagesAddEventHandler(ThisAddIn_OptionPagesAdd);

            DeleteFromFolder.Properties.Settings.Default.Upgrade(); //this could solve the addin from forgetting its saved state
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
            // additional: http://stackoverflow.com/questions/6677349/no-application-quit-event-in-outlook
        }
        void ThisAddIn_Quit()
        {
            Outlook.Stores stores = Application.Session.Stores;
            
            foreach(Outlook.Store s in stores)
            {
                Outlook.MAPIFolder mapi = s.GetRootFolder();

                foreach (Outlook.Folder f in mapi.Folders)
                {
                    //get saved list of folders to empty on application exit
                    string[] checkedList = DeleteFromFolder.Properties.Settings.Default.CheckedItems.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                    if (checkedList.Any(item => item.Contains(f.Name)))
                    {
                        int iCount = f.Items.Count;

                        for (int i = iCount; i > 0; i--)
                            f.Items.Remove(i);
                    }
                }
            }
        }
        void ThisAddIn_OptionPagesAdd(Outlook.PropertyPages Pages)
        {
            //add options page listing all folders that we can select to empty
            Pages.Add(new UCOptions(Application.Session.Stores), "Delete From Folder");
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
