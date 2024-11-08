using Gone_Phishing.Properties;
using Microsoft.Office.Core;
using Microsoft.Win32;
using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Gone_Phishing
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon1() { }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            // Only return the custom UI for the main Outlook window
            if (ribbonID == "Microsoft.Outlook.Explorer")
            {
                return GetResourceText("Gone_Phishing.Ribbon1.xml");
            }
            // Return empty string for email windows and other views
            return string.Empty;
        }

        #endregion

        #region Ribbon Callbacks

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnButtonClick_Phish(Office.IRibbonControl control)
        {
            ForwardSelectedEmail();
        }

        public Bitmap ButtonImage_Phish(Office.IRibbonControl control)
        {
            return Resources.phish;
        }

        public string ReadFromRegistry(string keyPath, string valueName)
        {
            try
            {
                using (RegistryKey key = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Default).OpenSubKey(keyPath))
                {
                    if (key != null)
                    {
                        object value = key.GetValue(valueName);
                        if (value != null)
                        {
                            return value.ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error reading from registry: {ex.Message}");
            }

            return null;
        }

        public string SendTo()
        {
            string registryKeyPath = null;
            if (File.Exists(@"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"))
            {
                registryKeyPath = @"Software\Microsoft\Office\Outlook\Addins\GonePhishing";
            }
            else if (File.Exists(@"C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE"))
            {
                registryKeyPath = @"Software\WOW6432Node\Microsoft\Office\Outlook\Addins\GonePhishing";
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Outlook is installed in a weird path, and this probably won't work.", "Incorrect", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }
            return ReadFromRegistry(registryKeyPath, "ReportTo");
        }

        public string Prefix()
        {
            string registryKeyPath = null;
            if (File.Exists(@"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"))
            {
                registryKeyPath = @"Software\Microsoft\Office\Outlook\Addins\GonePhishing";
            }
            else if (File.Exists(@"C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE"))
            {
                registryKeyPath = @"Software\WOW6432Node\Microsoft\Office\Outlook\Addins\GonePhishing";
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Outlook is installed in a weird path, and this probably won't work.", "Incorrect", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);

            }
            return ReadFromRegistry(registryKeyPath, "Prefix");
        }

        public void ForwardSelectedEmail()
        {
            Outlook.Explorer explorer = Globals.ThisAddIn.Application.ActiveExplorer();

            if (explorer.Selection.Count == 0)
            {
                System.Windows.Forms.MessageBox.Show("Please select an email to forward.", "No Email Selected", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }
            else if (explorer.Selection.Count == 1 && explorer.Selection[1] is Outlook.MailItem)
            {
                Outlook.MailItem selectedMail = explorer.Selection[1] as Outlook.MailItem;
                DialogResult result = MessageBox.Show($"Do you want to forward the email:\n'{selectedMail.Subject}'\nto {SendTo()} and move it to Deleted Items?", "Gone Phishing", MessageBoxButtons.YesNo);

                if (result == DialogResult.Yes)
                {
                    try
                    {
                        Outlook.MailItem newMail = Globals.ThisAddIn.Application.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;

                        newMail.Subject = Prefix() + selectedMail.Subject;
                        newMail.To = SendTo();

                        string tempFile = System.IO.Path.GetTempFileName();
                        selectedMail.SaveAs(tempFile, Outlook.OlSaveAsType.olMSG);
                        newMail.Attachments.Add(tempFile, Outlook.OlAttachmentType.olEmbeddeditem, 1, selectedMail.Subject);
                        newMail.Send();

                        System.IO.File.Delete(tempFile);
                        Outlook.MAPIFolder deletedItems = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems);
                        selectedMail.Move(deletedItems);

                    }
                    catch (Exception ex)
                    {
                        System.Windows.Forms.MessageBox.Show($"{ex.Message}", "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    }
                }
            }
            else if (explorer.Selection.Count > 1)
            {
                System.Windows.Forms.MessageBox.Show("Please only forward one email", "Too Many Emails Selected", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
