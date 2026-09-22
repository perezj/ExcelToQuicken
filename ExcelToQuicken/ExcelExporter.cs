using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Text;


namespace BanortePDFToExcel
{
    internal class ExcelExporter
    {
        public static void Export(List<Account> accounts, string outputPath)
        {
            using var workbook = new XLWorkbook();
            

            for (int i = 0; i < accounts.Count; i++)
            {
                var account = accounts[i];

                var ws = workbook.Worksheets.Add($"{account.Name}");


                // Headers
                ws.Cell(1, 1).Value = "Date";
                ws.Cell(1, 2).Value = "Description";
                ws.Cell(1, 3).Value = "Debit";
                ws.Cell(1, 4).Value = "Credit";
                ws.Cell(1, 5).Value = "Balance";
                ws.Cell(1, 6).Value = "Memo";
                ws.Cell(1, 7).Value = "Category";
                ws.Cell(1, 8).Value = "Tag";

                int row = 2;

                foreach (var tx in account.Transactions)
                {
                    ws.Cell(row, 1).Value = tx.Date;
                    ws.Cell(row, 2).Value = tx.Description;
                    
                    ws.Cell(row, 3).Value = tx.Debit;
                    ws.Cell(row, 3).Style.NumberFormat.Format = "$#,##0.00";
                    ws.Cell(row, 4).Value = tx.Credit;
                    ws.Cell(row, 4).Style.NumberFormat.Format = "$#,##0.00";
                    ws.Cell(row, 5).Value = tx.Balance;
                    ws.Cell(row, 5).Style.NumberFormat.Format = "$#,##0.00";
                    ws.Cell(row, 6).Value = tx.Memo;
                    ws.Cell(row, 7).Value = tx.Category;
                    ws.Cell(row, 8).Value = tx.Tag;
                    row++;
                }

                // Auto size
                var lastRow = row - 1;
                var tableRange = ws.Range(1, 1, lastRow, 8);
                var table = tableRange.CreateTable();

                table.Theme = XLTableTheme.TableStyleMedium2;

                ws.Columns().AdjustToContents();
                // Resize description header so its not too large
                ws.Column(2).Style.Alignment.WrapText = true;
                const double maxDescriptionWidth = 50.0; // adjust to taste (units ≈ character widths)
                if (ws.Column(2).Width > maxDescriptionWidth)
                    ws.Column(2).Width = maxDescriptionWidth;

                if (ws.Column(6).Width > maxDescriptionWidth)
                    ws.Column(6).Width = maxDescriptionWidth;
                // Optionally adjust row heights to fit wrapped text
                ws.Rows().AdjustToContents();
            }

            

            workbook.SaveAs(outputPath);
        }
    }
}
