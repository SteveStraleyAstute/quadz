using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace Quadz
{
    public partial class Program
    {
        public static void ReportsMenu()
        {
            while (true)
            {
                Console.Clear();

                Console.WriteLine();
                Console.WriteLine("Reports Menu");
                Console.WriteLine();
                Console.WriteLine("1. Create Report File");
                Console.WriteLine("2. Open Report File");
                Console.WriteLine();
                Console.WriteLine("99. Return to Previous Menu");
                Console.WriteLine();
                try
                {
                    var opt = Convert.ToInt32(Console.ReadLine());
                    if (opt == 1) { ReportCreateMenu(); }
                    else if (opt == 2) { ReportOpenMenu(); }
                    if (opt == 99)
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
        public static void ReportCreateMenu()
        {
            // Show top 12 files in INVENTORY folder
            var files = Directory.GetFiles(_inventory).OrderByDescending(t => t).Take(12).ToList();

            while (true)
            {
                Console.Clear();
                Console.WriteLine();
                Console.WriteLine("Create Report Menu - Select File");
                Console.WriteLine();

                var ndx = 1;
                foreach (var file in files)
                {
                    var temp = new FileInfo(file);
                    var baseTemp = temp.Name.Replace(temp.Extension, "");
                    var splitTemp = baseTemp.Split("-");

                    var mnth = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(Convert.ToInt32(splitTemp[1]));

                    Console.WriteLine($"{ndx++}. {splitTemp[0]} - {mnth}");
                }

                Console.WriteLine();
                Console.WriteLine("99. Return to Previous Menu");

                try
                {
                    var opt = Convert.ToInt32(Console.ReadLine());
                    if (opt == 99)
                        break;
                    else 
                    {
                        var temp = new FileInfo(files[opt-1]);
                        var baseTemp = temp.Name.Replace(temp.Extension, "");
                        var splitTemp = baseTemp.Split("-");

                        var mnth = Convert.ToInt32(splitTemp[1]);
                        var year = Convert.ToInt32(splitTemp[0]);

                        CreateReport(year, mnth);
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine();
                    Console.WriteLine(e.Message);
                    Console.ReadKey();
                }
            }
        }

        public static void ReportOpenMenu()
        {
            // Show top 12 files in REPORT folder
            var files = Directory.GetFiles(_reports).OrderByDescending(t => t).Take(18).ToList();

            while (true)
            {
                Console.Clear();
                Console.WriteLine();
                Console.WriteLine("Open Report Menu - Select File");
                Console.WriteLine();

                var ndx = 1;
                foreach (var file in files)
                {
                    var temp = new FileInfo(file);

                    Console.WriteLine($"{ndx++}. {temp.Name}");
                }

                Console.WriteLine();
                Console.WriteLine("99. Return to Previous Menu");

                try
                {
                    var opt = Convert.ToInt32(Console.ReadLine());
                    if (opt == 99)
                        break;
                    else 
                    {
                        if (opt > 0 || opt <= files.Count())
                        {
                            ndx = opt - 1;
                            var selectedFile = files[ndx];

                            try
                            {
                                var dir = GetExcelLocation();

                                if (!string.IsNullOrEmpty(dir))
                                {
                                    OpenExcel(dir, selectedFile);
                                }
                                else
                                {
                                    Console.WriteLine();
                                    Console.WriteLine($"Please go to EXCEL and open {selectedFile}");
                                    Console.WriteLine();
                                    Console.WriteLine("... press ENTER to continue...");
                                    Console.ReadLine();

                                }
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine(e.Message);
                                Console.Write("-- Any Key to Continue --");
                                Console.ReadKey();
                            }

                        }
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine();
                    Console.WriteLine(e.Message);
                    Console.ReadKey();
                }

            }

        }
    }
}
