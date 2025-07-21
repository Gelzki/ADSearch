using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;
using System.Net.NetworkInformation;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Security.Principal;
using System.Runtime.InteropServices;

namespace ADSearch {
    class OutputFormatting {

        enum OUTPUT_TYPE { 
            SUCCESS = '+',
            VERBOSE = '*',
            ERROR = '!'
        };

        public static void PrintADProperties(DirectoryEntry directoryEntry) {
            // Added below to list all attributes without fail - this is for OU walking
            string border = new string('-', 100);

            string name = "<no cn>";
            if (directoryEntry.Properties.Contains("cn") && directoryEntry.Properties["cn"].Value != null)
                name = directoryEntry.Properties["cn"].Value.ToString();
            else if (directoryEntry.Properties.Contains("ou") && directoryEntry.Properties["ou"].Value != null)
                name = directoryEntry.Properties["ou"].Value.ToString();
            else if (directoryEntry.Properties.Contains("name") && directoryEntry.Properties["name"].Value != null)
                name = directoryEntry.Properties["name"].Value.ToString();

            PrintVerbose(string.Format("     |-> {0,-30} | {1}", $"NAME ({name})", "VALUE"));
            PrintVerbose(border);

            foreach (var prop in directoryEntry.Properties.PropertyNames)
            {
                try
                {
                    var propName = prop.ToString();
                    var propVals = directoryEntry.Properties[propName];

                    string valueStr;

                    // For resolving comobjects and guids
                    if (propVals.Count > 1)
                    {
                        var values = propVals.Cast<object>().Select(val => ConvertAdValue(val, propName));
                        valueStr = $"[{string.Join(", ", values)}]";
                    }
                    else
                    {
                        valueStr = ConvertAdValue(propVals.Value, propName);
                    }
                    
                    /* Original if logic
                    if (propVals.Count > 1)
                    {
                        //var values = propVals.Cast<object>().Select(val => val?.ToString() ?? "<null>");
                        valueStr = $"[{string.Join(", ", values)}]";
                    }
                    else
                    {
                        valueStr = ConvertAdValue(propVals.Value);
                    }
                    */

                    PrintSuccess(string.Format("     |-> {0,-30} | {1}", propName, valueStr));
                }
                catch (Exception ex)
                {
                    PrintError(string.Format("     |-> {0,-30} | <error: {1}>", prop.ToString(), ex.Message));
                }
            }

            PrintVerbose(border);

            /* original PrintADProperties function
            string border = new String('-', 100);
            string cn = directoryEntry.Properties["cn"].Value.ToString();
            PrintVerbose(String.Format("     |-> {0,-30} | {1}", String.Format("NAME ({0})", cn), "VALUE"));
            PrintVerbose(border);
            foreach (var prop in directoryEntry.Properties.PropertyNames) {
                PrintSuccess(String.Format("     |-> {0,-30} | {1}", prop.ToString(), directoryEntry.Properties[prop.ToString()].Value));
            }
            PrintVerbose(border);
            */
        }

        public static int GetFormatLenSpecifier(string[] keys) {
            keys.Select((text, index) => new { Index = index, Text = text, Length = text.Length })
                .OrderByDescending(arr => arr.Length)
                .ToList();

            return keys.Max(arr => arr.Length);
        }

        public static void PrintJson(object obj) {
            Console.WriteLine(JsonConvert.SerializeObject(obj, Formatting.Indented));
        }

        public static void PrintSuccess(string msg, int indentation = 0) {
            Print(OUTPUT_TYPE.SUCCESS, msg, indentation);
        }

        public static void PrintVerbose(string msg, int indentation = 0) {
            Print(OUTPUT_TYPE.VERBOSE, msg, indentation);
        }

        public static void PrintError(string msg, int indentation = 0) {
            Print(OUTPUT_TYPE.ERROR, msg, indentation);
        }

        private static void Print(OUTPUT_TYPE msgType, string msg, int indentation = 0) {
            if (indentation != 0) {
                string tabs = new String('\t', indentation);
                Console.WriteLine("{0}[{1}] {2}", tabs, (char)msgType, msg);
            } else {
                Console.WriteLine("[{0}] {1}", (char)msgType, msg);
            }
        }
        // Add private helper for decoding System COMObject and GUIDs
        private static string ConvertAdValue(object val)
        {
            if (val == null)
                return "<null>";

            if (val is byte[] byteArray)
            {
                // Check for GUID
                if (byteArray.Length == 16)
                    return new Guid(byteArray).ToString();

                // Try interpreting as SID
                try
                {
                    var sid = new SecurityIdentifier(byteArray, 0);
                    return sid.Value;
                }
                catch
                {
                    return BitConverter.ToString(byteArray).Replace("-", "");
                }
            }

            if (val is DateTime dt)
                return dt.ToString("dd/MM/yyyy h:mm:ss tt");

            if (System.Runtime.InteropServices.Marshal.IsComObject(val))
                return "<COMObject>";

            return val.ToString();
        }

        // Added to decode ComObject
        // Interface for LargeInteger COM object
        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsDual)]
        [Guid("9068270B-0939-11D1-8BE1-00C04FD8D503")]
        private interface IADsLargeInteger
        {
            int HighPart { get; }
            int LowPart { get; }
        }

        private static DateTime? ConvertLargeIntegerToDateTime(object comObj)
        {
            if (comObj == null)
                return null;

            try
            {
                var largeInt = (IADsLargeInteger)comObj;
                long fileTime = ((long)largeInt.HighPart << 32) + (uint)largeInt.LowPart;

                if (fileTime == 0) return null;

                return DateTime.FromFileTimeUtc(fileTime);
            }
            catch
            {
                return null;
            }
        }

        private static string ConvertSecurityDescriptor(object sdObject)
        {
            try
            {
                var sdBytes = sdObject as byte[];
                if (sdBytes != null)
                {
                    var rawSd = new System.Security.AccessControl.RawSecurityDescriptor(sdBytes, 0);
                    return rawSd.GetSddlForm(System.Security.AccessControl.AccessControlSections.All);
                }
            }
            catch { }
            return "<SecurityDescriptor>";
        }

        private static string ConvertAdValue(object val, string propName = null)
        {
            if (val == null)
                return "<null>";

            // Try to decode LargeInteger COM objects (for timestamps)
            if (System.Runtime.InteropServices.Marshal.IsComObject(val))
            {
                var dt = ConvertLargeIntegerToDateTime(val);
                if (dt.HasValue)
                    return dt.Value.ToString("dd/MM/yyyy h:mm:ss tt");

                // Special case for nTSecurityDescriptor if passed as COM object? Usually byte[]
                return "<COMObject>";
            }

            // Decode Security Descriptor (byte array)
            if (propName == "nTSecurityDescriptor")
            {
                string sddl = ConvertSecurityDescriptor(val);
                if (!string.IsNullOrEmpty(sddl))
                    return sddl;
            }

            // Decode byte arrays (like objectGUID)
            if (val is byte[] bytes)
            {
                // Example: GUID formatting
                if (propName == "objectGUID" && bytes.Length == 16)
                    return new Guid(bytes).ToString();

                return BitConverter.ToString(bytes).Replace("-", "");
            }

            // Default fallback
            return val.ToString();
        }

    }
}
