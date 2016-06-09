using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Act.UI;
using Act.UI.Core;
using System.Windows.Forms;
using Act.UI.Email;

namespace EmailTestPlugin
{
    public class EmailTestPlugin : IPlugin
    {
        //This was an attempt to do mail merges in Act from multiple or selected email fields on multiple contacts at once.
        ActApplication application;
        public string TEST_EMAIL_MAIL_MERGE_URN
        {
            get { return "act-ui://com.act/application/menu/contact/testMailMerge"; }
            set { }
        }
        public string TEST_EMAIL_MAIL_MERGE_TEXT
        {
            get { return "Mail Merge Testing"; }
            set { }
        }

        void IPlugin.OnLoad(ActApplication application)
        {
            this.application = application;
            //TODO Plug in event handlers here
            application.ViewLoaded += new ViewEventHandler(Application_ViewLoaded);
        }

        private void Application_ViewLoaded(object sender, ViewEventArgs vea)
        {
            if (application.CurrentViewName == "Act.UI.IContactDetailView")
                if (!menuItemExists(TEST_EMAIL_MAIL_MERGE_URN, application))
                    addMenuItem(TEST_EMAIL_MAIL_MERGE_URN, TEST_EMAIL_MAIL_MERGE_TEXT, new Act.UI.CommandHandler(testMailMergeCommand), application, 0, true, null);
        }

        private void testMailMergeCommand(string command)
        {
            Act.UI.Correspondence.EmailMailMergeInfo emmi = new Act.UI.Correspondence.EmailMailMergeInfo();
            //TODO Email Field Name List is the list of physical field names that have the Email data type - might be usable for the mail merge
            var results = emmi.EmailFieldNameList(application);
            //foreach (string s in results)
            //{
            //    MessageBox.Show(s);
            //}
            //MessageBox.Show(application.ApplicationState.CurrentContact.FullName);
            //Act.UI.Correspondence.UICorrespondenceManager uicm = application.UICorrespondenceManager;
            //uicm.WriteEmail(application.ApplicationState.CurrentContact, results);
            //uicm.WriteEmail(application.ApplicationState.CurrentContact, "TBL_CONTACT.BUSINESS_EMAIL");

            //Act.UI.Email.UIEmailManager uiem = application.UIEmailManager;
            //TODO We can write individual emails, but what about mail merges?
            //uiem.CreateEmailDraft(application.ApplicationState.CurrentContact, results, "Hello, and a Blarg Honk to you, too.", false);
            
            //TODO It would seem the only way we're going to be sending emails to the same contact but to its different email addresses is via .net mail, which precludes the use of Act's mail merge and history recording engines
            //TODO Can we do Microsoft Mail Merges using Act fields but not the Act mail merge engine?

        }

        void IPlugin.OnUnLoad()
        {
            if (application != null)
            {
                //TODO Unplug event handlers here
            }
        }

        #region "ACT UI Helpers"
        /* Shamelessly borrowed from Jim Durkin at Durkin Computing */
        /* Add a menu item under TOOLS */
        private void addMenuItem(string urn, string menuText, Act.UI.CommandHandler handler, Act.UI.ActApplication actApplication, int position, bool hasSeperator, System.Drawing.Icon ico)
        {
            // Check if the connected menu exists
            if (actApplication.Explorer.CommandBarCollection["Connected Menus"] == null)
            {
                return;
            }
            // Add a menu item that delegates back to us
            if (menuItemExists(urn, actApplication) == true)
            {
                removeMenuItem(urn, actApplication);
            }
            try
            {
                CommandBarControl parentMenu = actApplication.Explorer.CommandBarCollection["Connected Menus"].ControlCollection[getParentControlURN(urn)];
                CommandBarButton newMenu = new CommandBarButton(menuText, menuText, null, urn, null, null);
                actApplication.RegisterCommand(urn, handler, RegisterType.Shell);

                // Are we displaying an icon?
                if (ico != null)
                {
                    newMenu.Icon = ico;
                    newMenu.DisplayStyle = CommandBarControl.ItemDisplayStyle.ImageAndText;
                }
                else
                    newMenu.DisplayStyle = CommandBarControl.ItemDisplayStyle.TextOnly;

                // If this is the first add-on under this plugin, add the seperator
                if (parentMenu == null)
                {
                    newMenu.HasSeparator = true;
                }
                else
                    newMenu.HasSeparator = hasSeperator;

                if (position == 0)
                    parentMenu.AddSubItem(newMenu);
                else
                    parentMenu.AddSubItem(newMenu, position);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }
        private string getParentControlURN(string urn)
        {
            try
            {
                return urn.Substring(0, urn.LastIndexOf("/"));
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            return "";
        }
        private void removeMenuItem(string urn, ActApplication actApplication)
        {
            try
            {
                if (actApplication.Explorer.CommandBarCollection["Connected Menus"] == null)
                    return;
                actApplication.RevokeCommand(urn);
                CommandBarControl removeMenu = actApplication.Explorer.CommandBarCollection["Connected Menus"].ControlCollection[urn];
                if (!(removeMenu == null))
                {
                    actApplication.Explorer.CommandBarCollection["Connected Menus"].ControlCollection[getParentControlURN(urn)].RemoveSubItem(removeMenu);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }
        private bool menuItemExists(string urn, ActApplication actApplication)
        {
            try
            {
                if (actApplication.Explorer.CommandBarCollection["Connected Menus"] == null)
                    return !(actApplication.Explorer.CommandBarCollection["Connected Menus"].ControlCollection[urn] == null);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            return false;
        }
        #endregion
    }
}