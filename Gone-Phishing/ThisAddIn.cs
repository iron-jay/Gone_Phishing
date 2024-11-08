using Outlook = Microsoft.Office.Interop.Outlook;

namespace Gone_Phishing
{
    public partial class ThisAddIn
    {
        private Outlook.Application application;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            application = this.Application;
            ((Outlook.ApplicationEvents_11_Event)application).ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(ItemSend);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {

        }

        private void ItemSend(object item, ref bool cancel)
        {
            Outlook.MailItem mailItem = item as Outlook.MailItem;
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1();
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
