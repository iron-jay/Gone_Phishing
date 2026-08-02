using Gone_Phishing.Properties;
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
    /// <summary>
    /// The add-in itself. Outlook creates this class by CLSID, connects it through
    /// <see cref="IDTExtensibility2"/>, and asks it for ribbon markup through
    /// <see cref="Office.IRibbonExtensibility"/>.
    /// </summary>
    // AutoDual is the documented arrangement for a C# COM add-in with ribbon callbacks: Office
    // resolves onAction and getImage by name, and the dual class interface is what carries the
    // public members for that lookup. ClassInterfaceType.None leaves the callbacks unbound, which
    // shows up as a ribbon that appears with a blank button that does nothing.
    [ComVisible(true)]
    [Guid(ClassId)]
    [ProgId(ProgIdentifier)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class Connect : IDTExtensibility2, Office.IRibbonExtensibility
    {
        internal const string ClassId = "2c6cf44d-af64-4b2c-aa5d-59a64f09e8ef";
        internal const string ProgIdentifier = "GonePhishing.Connect";

        private const string FriendlyName = "Gone Phishing";
        private const string Description = "Report phishing and spam emails to your security team.";

        private Outlook.Application application;
        private Office.IRibbonUI ribbon;

        #region IDTExtensibility2

        // Every method in this region is called by Outlook across a COM boundary. An exception
        // that escapes one of them terminates Outlook rather than surfacing anywhere useful, so
        // each catches its own.

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                application = Application as Outlook.Application;
            }
            catch (Exception)
            {
                application = null;
            }
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            // Dropping the references is all that is wanted here. Do NOT call
            // Marshal.ReleaseComObject on the Application: the runtime shares one wrapper per COM
            // identity across the whole process, so releasing it severs the wrapper Outlook and
            // every other add-in are still using, and the CLR dies with an ExecutionEngineException.
            ribbon = null;
            application = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }

        public void OnStartupComplete(ref Array custom) { }

        public void OnBeginShutdown(ref Array custom) { }

        #endregion

        #region IRibbonExtensibility

        public string GetCustomUI(string RibbonID)
        {
            try
            {
                // Only return the custom UI for the main Outlook window
                if (RibbonID == "Microsoft.Outlook.Explorer")
                {
                    return GetResourceText("Gone_Phishing.Ribbon.xml") ?? string.Empty;
                }
            }
            catch (Exception)
            {
                // No ribbon is better than no Outlook.
            }

            // Return empty string for email windows and other views
            return string.Empty;
        }

        #endregion

        #region Ribbon Callbacks
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            ribbon = ribbonUI;
        }

        public void OnButtonClick_Phish(Office.IRibbonControl control)
        {
            try
            {
                PhishingReporter.ReportSelectedEmail(application);
            }
            catch (Exception)
            {
                // ReportSelectedEmail already reports its own failures; this only stops anything
                // unexpected from reaching Outlook.
            }
        }

        public stdole.IPictureDisp ButtonImage_Phish(Office.IRibbonControl control)
        {
            try
            {
                // The ribbon expects an OLE picture. Unlike VSTO, a plain COM add-in has to do
                // this conversion itself.
                return PictureConverter.ToPictureDisp(Resources.phish);
            }
            catch (Exception)
            {
                // The button renders without an icon rather than taking Outlook down.
                return null;
            }
        }

        #endregion

        #region COM Registration

        // These run under regasm and write the keys Outlook reads to discover the add-in. See
        // readme.md for the exact commands.

        [ComRegisterFunction]
        public static void Register(Type type)
        {
            // Earlier builds registered under a key name that did not match the ProgID, which
            // Outlook rejects. Clear it so the two registrations cannot fight.
            try
            {
                Registry.LocalMachine.DeleteSubKeyTree(AddInSettings.LegacyRegistryKeyPath, false);
            }
            catch (Exception)
            {
                // Nothing to clean up, or no rights to do so - neither blocks registration.
            }

            using (RegistryKey key = Registry.LocalMachine.CreateSubKey(AddInSettings.RegistryKeyPath))
            {
                if (key == null)
                {
                    return;
                }

                key.SetValue("FriendlyName", FriendlyName, RegistryValueKind.String);
                key.SetValue("Description", Description, RegistryValueKind.String);
                key.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);

                // Left behind by the previous VSTO build. If it survives, Outlook tries to load
                // the add-in through the VSTO runtime instead of the CLSID and fails.
                if (key.GetValue("Manifest") != null)
                {
                    key.DeleteValue("Manifest", false);
                }
            }
        }

        [ComUnregisterFunction]
        public static void Unregister(Type type)
        {
            using (RegistryKey key = Registry.LocalMachine.OpenSubKey(AddInSettings.RegistryKeyPath, true))
            {
                if (key == null)
                {
                    return;
                }

                foreach (string valueName in new[] { "FriendlyName", "Description", "LoadBehavior" })
                {
                    key.DeleteValue(valueName, false);
                }

                // Anything else under this key was put there by someone else - configuration
                // lives under Software\Policies. Only tidy up the key if nothing is left.
                if (key.ValueCount > 0 || key.SubKeyCount > 0)
                {
                    return;
                }
            }

            Registry.LocalMachine.DeleteSubKey(AddInSettings.RegistryKeyPath, false);
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

        /// <summary>
        /// AxHost exposes the only public managed path from a Bitmap to an OLE picture, and its
        /// constructor is protected, so a throwaway subclass is the accepted way to reach it.
        /// </summary>
        private sealed class PictureConverter : AxHost
        {
            private PictureConverter() : base(string.Empty) { }

            internal static stdole.IPictureDisp ToPictureDisp(Image image)
            {
                return (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
            }
        }

        #endregion
    }
}







