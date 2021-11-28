using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Diagnostics;

namespace Quadz
{
    public partial class Program
    {
        static string _path = @"C:\Quadz";
        static string _master = _path + @"\_master";
        static string _inventory = _path + @"\Inventory";
        static string _reports = _path + @"\Reports";

        static string _version = "1.0.0";
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            while (true)
            {
                Console.Clear();
                Console.WriteLine($"Hello QUADZ Las Vegas - version [{_version}]");
                Console.WriteLine();
                Console.WriteLine("Select an option:");
                Console.WriteLine();
                Console.WriteLine("1. Inventory");
                Console.WriteLine("2. Monthly Reports");
                Console.WriteLine();
                Console.WriteLine("99. Ext");

                try
                {
                    var opt = Convert.ToInt32(Console.ReadLine());
                    if (opt == 1)
                        InventoryMenu();
                    else if (opt == 2)
                        ReportsMenu();
                    else if (opt == 99)
                        break;
                }
                catch (Exception e)
                {
                    Console.WriteLine();
                    Console.WriteLine(e.Message);
                    Console.ReadKey();
                }
            }
        }

        public static string GetExcelLocation()
        {
            var dir = "";
            RegistryKey key = Registry.LocalMachine;
            RegistryKey excelKey = key.OpenSubKey(@"SOFTWARE\MicroSoft\Office");
            if (excelKey != null)
            {
                foreach (string valuename in excelKey.GetSubKeyNames())
                {
                    int version = 9;
                    double currentVersion = 0;
                    if (Double.TryParse(valuename, out currentVersion) && currentVersion >= version)
                    {
                        RegistryKey rootdir = excelKey.OpenSubKey(currentVersion + @".0\Excel\InstallRoot");
                        if (rootdir != null)
                        {
                            dir = rootdir.GetValue(rootdir.GetValueNames()[0]).ToString();
                            break;
                        }
                    }
                }
            }

            return dir;
        }

        public static void OpenExcel(string dir, string file)
        {
            Process ExternalProcess = new Process();
            ExternalProcess.StartInfo.FileName = dir + @"\Excel.exe";
            ExternalProcess.StartInfo.Arguments = "\"" + file + "\"";
            ExternalProcess.StartInfo.WindowStyle = ProcessWindowStyle.Maximized;
            ExternalProcess.Start();
            ExternalProcess.WaitForExit();
        }
    }
}
