using Microsoft.Win32;
using System;

namespace Gone_Phishing
{
    /// <summary>
    /// Configuration for the add-in. Group Policy overrides the defaults written by the
    /// installer. Every location read is one only an administrator can write, so a user cannot
    /// quietly redirect their own phishing reports.
    /// </summary>
    internal static class AddInSettings
    {
        /// <summary>
        /// Where the ADMX template writes. Keys under Software\Policies are ACL'd against
        /// standard users and removed when the policy stops applying.
        /// </summary>
        internal const string PolicyKeyPath = @"Software\Policies\iron-jay\GonePhishing";

        /// <summary>
        /// Where the installer writes the defaults supplied on the msiexec command line. Under
        /// HKLM, so it is administrator-writable only; policy takes precedence over it.
        /// </summary>
        internal const string InstallKeyPath = @"Software\iron-jay\GonePhishing";

        /// <summary>
        /// Where the add-in registers itself with Outlook. Settings are deliberately not read
        /// from here - it holds the COM registration only.
        ///
        /// The leaf name MUST equal the ProgID: Outlook takes the key name as the ProgID it
        /// looks up, and reports "not a valid Office Add-in" if the two disagree.
        /// </summary>
        internal const string RegistryKeyPath = @"Software\Microsoft\Office\Outlook\Addins\GonePhishing.Connect";

        /// <summary>
        /// The pre-1.0 key name, which did not match the ProgID. Removed on install.
        /// </summary>
        internal const string LegacyRegistryKeyPath = @"Software\Microsoft\Office\Outlook\Addins\GonePhishing";

        /// <summary>
        /// Search order, highest precedence first: computer policy, then user policy, then the
        /// default the installer wrote. Policy therefore supersedes the installed value without
        /// having to change it, so a fleet can be retargeted without reinstalling.
        /// </summary>
        private static readonly Location[] SearchOrder =
        {
            new Location(RegistryHive.LocalMachine, PolicyKeyPath),
            new Location(RegistryHive.CurrentUser, PolicyKeyPath),
            new Location(RegistryHive.LocalMachine, InstallKeyPath)
        };

        /// <summary>
        /// Reads a configuration value from the first location that defines it. Both registry
        /// views are probed so the same code works regardless of Office bitness.
        /// </summary>
        internal static string Read(string valueName)
        {
            RegistryView[] views = { RegistryView.Registry64, RegistryView.Registry32 };

            foreach (Location location in SearchOrder)
            {
                foreach (RegistryView view in views)
                {
                    try
                    {
                        using (RegistryKey baseKey = RegistryKey.OpenBaseKey(location.Hive, view))
                        using (RegistryKey key = baseKey.OpenSubKey(location.Path))
                        {
                            string value = key?.GetValue(valueName) as string;
                            if (!string.IsNullOrWhiteSpace(value))
                            {
                                return value.Trim();
                            }
                        }
                    }
                    catch (Exception)
                    {
                        // An inaccessible hive or view is not fatal - try the next location.
                    }
                }
            }

            return null;
        }

        private struct Location
        {
            internal readonly RegistryHive Hive;
            internal readonly string Path;

            internal Location(RegistryHive hive, string path)
            {
                Hive = hive;
                Path = path;
            }
        }
    }
}
