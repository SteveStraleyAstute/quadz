using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace Quadz
{
    public partial class Program
    {
        public static void CreateInventory(int year, int month)
        {
            bool canContinue = false;

            // set month to 2 digits
            var theMonth = month.ToString().PadLeft(2, '0');

            var outFile = $"{year}-{theMonth}";
            var masterFile = $@"{_master}\quadzmaster.xlsx";
            var daysInMonth = DateTime.DaysInMonth(year, month);
            FileInfo inventoryFile = new FileInfo($@"{_inventory}\{outFile}.xlsx");

            List<string> category = new List<string>() { "Hard", "Beer/Wine", "Misc" };
            var categoryCount = 0;

            // check in inventory file exists or not
            if (File.Exists(inventoryFile.FullName))
            {

                Console.Clear();
                Console.WriteLine($"The inventory file [{inventoryFile.FullName}] already exists.   Overwrite??");
                Console.WriteLine();
                Console.WriteLine("1. Continue");
                Console.WriteLine("9. Do NOT continue");
                Console.WriteLine();
                Console.WriteLine("Enter Selection....");
                var opt = Convert.ToInt32(Console.ReadLine());

                if (opt == 1)
                    canContinue = true;
            }
            else
                canContinue = true;

            if (canContinue)
            {
                if (File.Exists(masterFile))
                {
                    FileInfo mastFile = new FileInfo(masterFile);
                    using (ExcelPackage master = new ExcelPackage(mastFile))
                    {
                        //get the first worksheet in the workbook
                        ExcelWorksheet masterWs = master.Workbook.Worksheets[2];
                        int colCount = masterWs.Dimension.End.Column;  //get Column Count
                        int rowCount = masterWs.Dimension.End.Row;     //get row count

                        // create inventory file

                        using (ExcelPackage inventory = new ExcelPackage())
                        {
                            ExcelWorksheet ws = inventory.Workbook.Worksheets.Add($"Inventory-{DateTime.Now.Year}-{DateTime.Now.Month}");
                            ws.Cells[1, 1].Value = "Product";
                            ws.Cells[1, 2].Value = "Code";
                            ws.Row(1).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            ws.Row(1).Style.Fill.BackgroundColor.SetColor(Color.Teal);
                            ws.Row(1).Style.Font.Color.SetColor(Color.White);
                            ws.View.ShowGridLines = true;
                            ws.PrinterSettings.ShowGridLines = true;
                            ws.PrinterSettings.BottomMargin = .1M;
                            ws.PrinterSettings.HeaderMargin = .1M;
                            ws.PrinterSettings.LeftMargin = .1M;
                            ws.PrinterSettings.RightMargin = .1M;
                            ws.PrinterSettings.ShowHeaders = false;
                            ws.PrinterSettings.Orientation = eOrientation.Landscape;

                            ws.PrinterSettings.FitToPage = true;
                            ws.PrinterSettings.FitToWidth = 1;
                            ws.PrinterSettings.FitToHeight = 0;

                            for (int i = 1; i <= daysInMonth; i++)
                            {
                                ws.Cells[1, 2 + i].Value = i.ToString();
                                ws.Cells[1, 2 + i].Style.Numberformat.Format = "#";
                                ws.Cells[1, 2 + i].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                                ws.Column(2 + i).Width = 5;
                            }

                            ws.Column(2 + daysInMonth).PageBreak = true;

                            // pull inventory from master
                            ws.Cells[2, 1].Value = category[categoryCount++];
                            ws.Cells[2, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            ws.Row(2).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            ws.Row(2).Style.Fill.BackgroundColor.SetColor(Color.LightGray);


                            for (int row = 2; row <= rowCount; row++)
                            {
                                if (masterWs.Cells[row, 1].Value != null)
                                {
                                    ws.Cells[1 + row, 1].Value = masterWs.Cells[row, 1].Value.ToString();
                                    ws.Cells[1 + row, 2].Value = masterWs.Cells[row, 4].Value.ToString();
                                }
                                else
                                {
                                    ws.Cells[1 + row, 1].Value = category[categoryCount++];
                                    ws.Cells[1 + row, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                    ws.Row(1 + row).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    ws.Row(1 + row).Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                                }
                            }

                            ws.Column(1).AutoFit();
                            ws.Column(2).Hidden = true;

                            ws.View.FreezePanes(2, 1);

                            inventory.SaveAs(inventoryFile);

                            Console.Clear();
                            Console.WriteLine();
                            Console.WriteLine($"Inventory file [{inventoryFile}] has been created.   You may open it");
                            Console.WriteLine();
                            Console.WriteLine("Press ENTER key to continue");
                            
                            Console.ReadLine();

                            var dir = GetExcelLocation();
                            if (!string.IsNullOrEmpty(dir))
                            {
                                OpenExcel(dir, inventoryFile.FullName);
                            }
                            else
                            {
                                Console.Clear();
                                Console.WriteLine();
                                Console.WriteLine($"Cannot find Excel.  Please go to EXCEL and open {inventoryFile.FullName}");
                                Console.WriteLine();
                                Console.WriteLine("... press ENTER to continue...");
                                Console.ReadLine();
                            }



                        }
                    }
                }
                else
                    Console.WriteLine("The QUADZMASTER.XLSX file is missing");
            }
        }
    }
}
