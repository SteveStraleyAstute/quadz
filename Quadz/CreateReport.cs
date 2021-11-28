using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;

namespace Quadz
{
    public partial class Program
    {
        static void CreateReport(int yearReport, int monthReport)
        {

            var outFilePrefix = $"{yearReport}-{monthReport.ToString().PadLeft(2, '0')}";
            var reportDate = new DateTime(yearReport, monthReport, 1);

            FileInfo master = new FileInfo($@"{_master}\quadzmaster.xlsx");
            FileInfo invent = new FileInfo($@"{_inventory}\{outFilePrefix}.xlsx");

            try
            {
                if (master.Exists)
                {
                    using (ExcelPackage masterPkg = new ExcelPackage(master))
                    {
                        using (ExcelPackage inventoryPkg = new ExcelPackage(invent))
                        {
                            if (inventoryPkg.Workbook.Worksheets.Count == 0)
                            {
                                Console.WriteLine();
                                Console.WriteLine($"Error in reading inventory file [{invent.FullName}]... not enough worksheets.  Contact support");
                                Console.WriteLine();
                                Console.WriteLine("Press ENTER key to continue...");
                                Console.ReadLine();
                            }
                            else
                            {
                                ExcelWorksheet inventoryWs = inventoryPkg.Workbook.Worksheets[0];

                                ExcelWorksheet masterInventoryWs = masterPkg.Workbook.Worksheets["Inventory"];
                                ExcelWorksheet masterTypeWs = masterPkg.Workbook.Worksheets["Type"];
                                ExcelWorksheet masterCategoryWs = masterPkg.Workbook.Worksheets["Category"];
                                // figure out how many reports based on weeksheet tabs
                                var masterTabs = masterPkg.Workbook.Worksheets.Count - 3;  // First 3 are reservered
                                Console.WriteLine();
                                for (int x = 0; x < masterTabs; x++)
                                {
                                    ExcelWorksheet masterWs = masterPkg.Workbook.Worksheets[masterTabs + x];
                                    var reportOut = $@"{_reports}\{outFilePrefix}-{masterWs.Name}.xlsx";
                                    var rowCount = masterWs.Dimension.End.Row;

                                    FileInfo reportFI = new FileInfo(reportOut);

                                    if (reportFI.Exists)
                                        reportFI.Delete();

                                    using (ExcelPackage reportPkg = new ExcelPackage(reportFI))
                                    {
                                        Console.WriteLine($"Creating: [{reportFI.FullName}]");
                                        #region Header of Report
                                        ExcelWorksheet ws = reportPkg.Workbook.Worksheets.Add("Report");
                                        ws.Cells[1, 1].Value = masterWs.Name;
                                        ws.Cells[1, 1].Style.Font.Bold = true;
                                        ws.Column(1).Width = 5;

                                        ws.Cells[1, 4].Value = $"Date: {reportDate.ToString("MMMM", CultureInfo.InvariantCulture)} {yearReport}";
                                        ws.Cells[2, 4].Value = masterWs.Cells[1, 1].Value;

                                        ws.Cells[4, 2].Value = "Product";
                                        ws.Cells[4, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                        ws.Cells[4, 2].Style.Font.UnderLine = true;
                                        ws.Cells[4, 4].Value = "Week 1";
                                        ws.Cells[4, 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                        ws.Cells[4, 4].Style.Font.UnderLine = true;
                                        ws.Cells[4, 5].Value = "Week 2";
                                        ws.Cells[4, 5].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                        ws.Cells[4, 5].Style.Font.UnderLine = true;
                                        ws.Cells[4, 6].Value = "Week 3";
                                        ws.Cells[4, 6].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                        ws.Cells[4, 6].Style.Font.UnderLine = true;
                                        ws.Cells[4, 7].Value = "Week 4";
                                        ws.Cells[4, 7].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                        ws.Cells[4, 7].Style.Font.UnderLine = true;
                                        ws.Cells[4, 8].Value = "Week 5";
                                        ws.Cells[4, 8].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                        ws.Cells[4, 8].Style.Font.UnderLine = true;
                                        ws.Cells[4, 10].Value = "Total";
                                        ws.Cells[4, 10].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                        ws.Cells[4, 10].Style.Font.UnderLine = true;
                                        #endregion

                                        #region Key Value Pair of Category items for running totals
                                        Dictionary<string, int> CategoryCounts = new Dictionary<string, int>();
                                        var mc = masterCategoryWs.Dimension.End.Row;
                                        for (int i = 1; i <= mc; i++)
                                        {
                                            if (masterCategoryWs.Cells[i, 2].Value != null)
                                            {
                                                string categoryCode = masterCategoryWs.Cells[i, 2].Value.ToString();
                                                CategoryCounts.Add(categoryCode, 0);
                                            }
                                        }
                                        #endregion

                                        #region Inventory
                                        for (int y = 2; y <= rowCount; y++)
                                        {
                                            string productCode = masterWs.Cells[y, 1].Value.ToString();  // get product code

                                            var invRow = FindCellValue(inventoryWs, productCode);

                                            var product = inventoryWs.Cells[invRow, 1].Value;

                                            var productCategory = FindProductCategory(masterInventoryWs, productCode);
                                            var categoryColor = FindCategoryColor(masterCategoryWs, productCategory);

                                            if (invRow > 0)
                                            {
                                                var week1 = GetColumnTotal(inventoryWs, invRow, 1);
                                                var week2 = GetColumnTotal(inventoryWs, invRow, 8);
                                                var week3 = GetColumnTotal(inventoryWs, invRow, 15);
                                                var week4 = GetColumnTotal(inventoryWs, invRow, 22);
                                                var week5 = GetColumnTotal(inventoryWs, invRow, 29);
                                                var total = week1 + week2 + week3 + week4 + week5;

                                                // Find product category in KVP and add total
                                                if (!string.IsNullOrEmpty(productCategory))
                                                    CategoryCounts[productCategory] += total;

                                                #region Generate Columns
                                                ws.Cells[3 + y, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                                ws.Cells[3 + y, 1].Style.Fill.BackgroundColor.SetColor(categoryColor);

                                                ws.Cells[3 + y, 2].Value = product;
                                                ws.Cells[3 + y, 4].Value = week1;
                                                ws.Cells[3 + y, 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                                ws.Cells[3 + y, 5].Value = week2;
                                                ws.Cells[3 + y, 5].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                                ws.Cells[3 + y, 6].Value = week3;
                                                ws.Cells[3 + y, 6].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                                ws.Cells[3 + y, 7].Value = week4;
                                                ws.Cells[3 + y, 7].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                                ws.Cells[3 + y, 8].Value = week5;
                                                ws.Cells[3 + y, 8].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                                ws.Cells[3 + y, 10].Value = total;
                                                ws.Cells[3 + y, 10].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                                #endregion
                                            }
                                            else
                                                Console.WriteLine($"Unable to find product [{productCode}] in inventory file");
                                        }
                                        #endregion

                                        var startTotalRow = rowCount + 6;

                                        ws.Cells[startTotalRow++, 2].Value = "Totals";
                                        // Now totals
                                        foreach (var item in CategoryCounts)
                                        {
                                            // get category name and color for category count
                                            var categoryColor = FindCategoryColor(masterCategoryWs, item.Key);
                                            var categoryName = FindCategoryName(masterCategoryWs, item.Key);

                                            ws.Cells[startTotalRow, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                            ws.Cells[startTotalRow, 1].Style.Fill.BackgroundColor.SetColor(categoryColor);
                                            ws.Cells[startTotalRow, 2].Value = categoryName;
                                            ws.Cells[startTotalRow++, 3].Value = item.Value;
                                        }

                                        ws.Column(9).Width = 5;
                                        ws.Column(2).AutoFit();

                                        ws.PrinterSettings.FitToPage = true;
                                        ws.PrinterSettings.FitToWidth = 1;
                                        ws.PrinterSettings.FitToHeight = 0;

                                        reportPkg.SaveAs(reportFI);
                                    }
                                }
                                Console.WriteLine();
                                Console.WriteLine($"Report files have been created in REPORTS folder");
                                Console.WriteLine();
                                Console.WriteLine("Press ENTER key to continue...");
                                Console.ReadLine();

                            }
                        }
                    }
                }
                else
                    Console.WriteLine("The QUADZMASTER.XLSX file is missing");
            }
            catch (Exception e)
            {
                var asd = e;
                throw e;
            }
        }
        static int FindCellValue(ExcelWorksheet ws, string code)
        {
            var rowCount = ws.Dimension.End.Row;
            int theRow = -1;

            for (int x = 3; x <= rowCount; x++)
            {
                if (ws.Cells[x, 2].Value != null)
                {
                    if (ws.Cells[x, 2].Value.ToString() == code)
                    {
                        theRow = x;
                        break;
                    }
                }
            }

            return theRow;
        }
        static string FindProductCategory(ExcelWorksheet ws, string code)
        {
            var rowCount = ws.Dimension.End.Row;
            var category = "";
            for (int x = 2; x <= rowCount; x++)
            {
                if (ws.Cells[x, 4].Value != null)
                {
                    if (ws.Cells[x, 4].Value.ToString() == code)
                    {
                        if (ws.Cells[x, 3].Value != null)
                        {
                            category = ws.Cells[x, 3].Value.ToString();
                            break;
                        }
                        else
                            break;
                    }
                }
            }
            return category;
        }
        static Color FindCategoryColor(ExcelWorksheet ws, string category)
        {
            var rowCount = ws.Dimension.End.Row;
            var color = Color.White;
            for (int x = 1; x <= rowCount; x++)
            {
                if (ws.Cells[x, 2].Value != null)
                {
                    if (ws.Cells[x, 2].Value.ToString() == category)
                    {
                        var c = ws.Cells[x, 2].Style.Fill.BackgroundColor;
                        string oldColor = EPPLookupColorFixed(c);
                        color = ColorTranslator.FromHtml(oldColor);
                        if (color.Name == "Black")
                        {
                            color = Color.White;
                        }
                        break;
                    }
                }
            }

            return color;
        }
        static string FindCategoryName(ExcelWorksheet ws, string category)
        {
            var rowCount = ws.Dimension.End.Row;
            var name = "";
            for (int x = 1; x <= rowCount; x++)
            {
                if (ws.Cells[x, 2].Value != null)
                {
                    if (ws.Cells[x, 2].Value.ToString() == category)
                    {
                        name = ws.Cells[x, 1].Value.ToString();
                        break;
                    }
                }
            }

            return name;
        }
        static List<string> FindAllCodesWithCategories(ExcelWorksheet ws, string catCode)
        {
            List<string> retValue = new List<string>();
            var rowCount = ws.Dimension.End.Row;
            for (int i = 2; i <= rowCount; i++)
            {
                if (ws.Cells[i, 3].Value != null)
                {
                    if (ws.Cells[i, 3].Value.ToString() == catCode)
                    {
                        retValue.Add(ws.Cells[i, 4].Value.ToString());
                    }
                }
            }
            return retValue;
        }
        static string EPPLookupColorFixed(ExcelColor sourceColor)
        {
            var lookupColor = sourceColor.LookupColor();
            const int maxLookup = 63;
            bool isFromTable = (0 <= sourceColor.Indexed) && (maxLookup > sourceColor.Indexed);
            bool isFromRGB = (null != sourceColor.Rgb && 0 < sourceColor.Rgb.Length);
            if (isFromTable || isFromRGB)
                return lookupColor;

            // Ok, we know we entered the else block in EPP - the one 
            // that doesn't quite behave as expected.

            string shortString = "0000";
            switch (lookupColor.Length)
            {
                case 6:
                    // Of the form #FF000
                    shortString = lookupColor.Substring(3, 1).PadLeft(4, '0');
                    break;
                case 9:
                    // Of the form #FFAAAAAA
                    shortString = lookupColor.Substring(3, 2).PadLeft(4, '0');
                    break;
                case 12:
                    // Of the form #FF200200200
                    shortString = lookupColor.Substring(3, 3).PadLeft(4, '0');
                    break;
            }
            var actualValue = short.Parse(shortString, System.Globalization.NumberStyles.HexNumber);
            var percent = ((double)actualValue) / 0x200d;
            var byteValue = (byte)Math.Round(percent * 0xFF, 0);
            var byteText = byteValue.ToString("X");
            byteText = byteText.Length == 2 ? byteText : byteText.PadLeft(2, '0');
            return $"{lookupColor.Substring(0, 3)}{byteText}{byteText}{byteText}";
        }
        static int GetColumnTotal(ExcelWorksheet ws, int row, int start)
        {
            int total = 0;
            start += 2;
            var end = start + 6;
            //Console.WriteLine($"{start} - {end}");
            for (int x = start; x <= end; x++)
            {
                if (ws.Cells[row, x].Value != null)
                {
                    total += Convert.ToInt32(ws.Cells[row, x].Value);
                }
            }

            return total;
        }
    }
}
