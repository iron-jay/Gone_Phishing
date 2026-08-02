using System;
using System.IO;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Gone_Phishing
{
    /// <summary>
    /// Forwards the selected message to the configured reporting address and files the original
    /// in Deleted Items.
    ///
    /// Every Outlook object below is a short-lived local, which is what keeps Outlook from being
    /// pinned in memory. Marshal.ReleaseComObject is deliberately not used: the runtime hands out
    /// one wrapper per COM identity for process-wide singletons like Application, ActiveExplorer
    /// and Session, so releasing one breaks it for Outlook itself.
    /// </summary>
    internal static class PhishingReporter
    {
        internal static void ReportSelectedEmail(Outlook.Application application)
        {
            if (application == null)
            {
                ShowError("Gone Phishing did not start correctly. Please restart Outlook.", "Not Connected");
                return;
            }

            try
            {
                Outlook.Explorer explorer = application.ActiveExplorer();
                if (explorer == null)
                {
                    return;
                }

                Outlook.Selection selection = explorer.Selection;

                if (selection.Count == 0)
                {
                    ShowWarning("Please select an email to forward.", "No Email Selected");
                    return;
                }

                if (selection.Count > 1)
                {
                    ShowWarning("Please only forward one email", "Too Many Emails Selected");
                    return;
                }

                Outlook.MailItem selectedMail = selection[1] as Outlook.MailItem;
                if (selectedMail == null)
                {
                    ShowWarning("Only email messages can be reported. Please select a single email and try again.", "Unsupported Item");
                    return;
                }

                string reportTo = AddInSettings.Read("ReportTo");
                if (string.IsNullOrEmpty(reportTo))
                {
                    ShowError("Gone Phishing has not been configured with a reporting address, so nothing was sent.\n\nPlease contact your IT administrator.", "Not Configured");
                    return;
                }

                string prefix = AddInSettings.Read("Prefix") ?? string.Empty;

                DialogResult result = MessageBox.Show(
                    $"Do you want to forward the email:\n'{selectedMail.Subject}'\nto {reportTo} and move it to Deleted Items?",
                    "Gone Phishing",
                    MessageBoxButtons.YesNo);

                if (result == DialogResult.Yes)
                {
                    ReportMail(application, selectedMail, reportTo, prefix);
                }
            }
            catch (Exception ex)
            {
                // Nothing may escape into Outlook's COM call - an exception crossing that
                // boundary takes the whole process down.
                ShowError(ex.Message, "Error");
            }
        }

        private static void ReportMail(Outlook.Application application, Outlook.MailItem selectedMail, string reportTo, string prefix)
        {
            // A .msg extension keeps the attachment unambiguous, and a GUID name avoids the
            // 65535-file limit that Path.GetTempFileName() runs into.
            string tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".msg");

            try
            {
                selectedMail.SaveAs(tempFile, Outlook.OlSaveAsType.olMSG);

                Outlook.MailItem newMail = application.CreateItem(Outlook.OlItemType.olMailItem) as Outlook.MailItem;
                newMail.Subject = prefix + selectedMail.Subject;
                newMail.To = reportTo;
                newMail.Attachments.Add(tempFile, Outlook.OlAttachmentType.olEmbeddeditem, 1, selectedMail.Subject);
                newMail.Send();

                Outlook.MAPIFolder deletedItems = application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems);
                selectedMail.Move(deletedItems);
            }
            finally
            {
                TryDeleteFile(tempFile);
            }
        }

        private static void ShowWarning(string message, string caption)
        {
            MessageBox.Show(message, caption, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private static void ShowError(string message, string caption)
        {
            MessageBox.Show(message, caption, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private static void TryDeleteFile(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch (Exception)
            {
                // A leftover temp file is not worth interrupting the user for.
            }
        }
    }
}
