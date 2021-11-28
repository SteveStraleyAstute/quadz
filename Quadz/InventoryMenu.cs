using System;
using System.Linq;
using System.IO;
using System.Globalization;
using System.Diagnostics;
using Microsoft.Win32;

namespace Quadz
{
    public partial class Program
    {
        public static void InventoryMenu()
        {
            while (true)
            {
                Console.Clear();

                Console.WriteLine();
                Console.WriteLine("Invetory Menu");
                Console.WriteLine();
                Console.WriteLine("1. Create Inventory File");
                Console.WriteLine("2. Open Existing Inventory File");
                Console.WriteLine();
                Console.WriteLine("99. Return to Previous Menu");

                try
                {
                    var opt = Convert.ToInt32(Console.ReadLine());
                    if (opt == 1) { InventoryCreateMenu(); }
                    else if (opt == 2) { InventoryOpenMenu(); }
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
    
        public static void InventoryCreateMenu()
        {
            int whatYear = DateTime.Now.Year;

            while (true)
            {
                Console.Clear();

                Console.WriteLine();
                Console.WriteLine("Create Inventory Menu - Select Month");
                Console.WriteLine();
                Console.WriteLine("1. January");
                Console.WriteLine("2. February");
                Console.WriteLine("3. March");
                Console.WriteLine("4. April");
                Console.WriteLine("5. May");
                Console.WriteLine("6. June");
                Console.WriteLine("7. July");
                Console.WriteLine("8. August");
                Console.WriteLine("9. September");
                Console.WriteLine("10. October");
                Console.WriteLine("11. November");
                Console.WriteLine("12. December");
                Console.WriteLine();
                Console.WriteLine("99. Return to Previous Menu");

                try
                {
                    var opt = Convert.ToInt32(Console.ReadLine());
                    if (opt == 12) 
                    {
                        Console.Clear();
                        Console.WriteLine("For what year????");
                        Console.WriteLine();
                        Console.WriteLine($"1. {whatYear - 1}");
                        Console.WriteLine($"2. {whatYear}");
                        Console.WriteLine();
                        opt = Convert.ToInt32(Console.ReadLine());
                        if (opt == 1)
                            CreateInventory(whatYear - 1, 12);
                        else if (opt == 2)
                            CreateInventory(whatYear, 12);

                        break;
                    }
                    else if (opt == 99)
                        break;
                    else 
                    {
                        CreateInventory(whatYear, opt);
                        break;
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

        public static void InventoryOpenMenu()
        {
            while (true)
            {
                Console.Clear();

                var files = Directory.GetFiles(_inventory).OrderByDescending(t => t).Take(10).ToList();

                Console.WriteLine();
                Console.WriteLine("Open Invetory Menu - Select File");
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
                                    Console.Clear();
                                    Console.WriteLine();
                                    Console.WriteLine($"Cannot find Excel.  Please go to EXCEL and open {selectedFile}");
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
