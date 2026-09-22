using ClosedXML.Excel;
using System.Text;

namespace BanortePDFToExcel.Tests;

public sealed class QifExporterTests
{
    [Fact]
    public void Export_WritesAccountTransactionsCategoriesAndTags()
    {
        var directory = CreateTestDirectory();
        var workbookPath = Path.Combine(directory, "categorized.xlsx");

        try
        {
            CreateWorkbook(workbookPath, includeUncategorized: false);
            var exporter = new QifExporter();

            var discovered = Assert.Single(exporter.DiscoverAccounts(workbookPath));
            Assert.Equal("Personal", discovered.WorksheetName);
            Assert.Equal(2, discovered.TransactionCount);

            var result = exporter.Export(workbookPath,
                new[] { new QifAccountMapping("Personal", "Banorte Personal", "Bank") });

            var file = Assert.Single(result.Files);
            var qif = File.ReadAllText(file.OutputPath, Encoding.Latin1);
            Assert.Equal(2, file.TransactionCount);
            Assert.Contains("!Account", qif);
            Assert.Contains("NBanorte Personal", qif);
            Assert.Contains("!Type:Bank", qif);
            Assert.Contains("D01/05/2026", qif);
            Assert.Contains("T-250.50", qif);
            Assert.Contains("T500.00", qif);
            Assert.Contains("PComisión CFE", qif);
            Assert.Contains("LHome:Bills/Utilities", qif);
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void Export_StopsWhenTransactionsAreUncategorizedByDefault()
    {
        var directory = CreateTestDirectory();
        var workbookPath = Path.Combine(directory, "categorized.xlsx");

        try
        {
            CreateWorkbook(workbookPath, includeUncategorized: true);
            var exporter = new QifExporter();

            var exception = Assert.Throws<QifValidationException>(() => exporter.Export(workbookPath,
                new[] { new QifAccountMapping("Personal", "Banorte Personal", "Bank") }));

            Assert.Contains("do not have a category", exception.Message);
            Assert.Empty(Directory.GetFiles(directory, "*.qif"));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void Export_ShortensLongPayeeAndPreservesOriginalInMemo()
    {
        var directory = CreateTestDirectory();
        var workbookPath = Path.Combine(directory, "categorized.xlsx");
        const string longDescription = "TRASP FONDOS 0000260408 AL R.F.C. PECJ850130T21 , A LA CUENTA: 1171346382 From personal to biz IVA: . 0";

        try
        {
            using (var workbook = new XLWorkbook())
            {
                var sheet = workbook.Worksheets.Add("Personal");
                sheet.Cell(1, 1).Value = "Date";
                sheet.Cell(1, 2).Value = "Description";
                sheet.Cell(1, 3).Value = "Debit";
                sheet.Cell(1, 4).Value = "Category";
                sheet.Cell(2, 1).Value = new DateTime(2026, 4, 8);
                sheet.Cell(2, 2).Value = longDescription;
                sheet.Cell(2, 3).Value = 10000m;
                sheet.Cell(2, 4).Value = "[Banorte Negocio]";
                workbook.SaveAs(workbookPath);
            }

            var result = new QifExporter().Export(
                workbookPath,
                new[] { new QifAccountMapping("Personal", "Banorte Personal", "Bank") });
            var qif = File.ReadAllLines(Assert.Single(result.Files).OutputPath, Encoding.Latin1);
            var payee = Assert.Single(qif.Where(line => line.StartsWith('P')));

            Assert.True(payee.Length - 1 <= 64);
            Assert.Contains("PTRASP FONDOS 0000260408", payee);
            Assert.Contains(qif, line => line.StartsWith("MDescripción original: ", StringComparison.Ordinal));
            Assert.Contains(qif, line => line.Contains(longDescription, StringComparison.Ordinal));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void Export_OmitsEmptyCategoryRecordForUncategorizedTransaction()
    {
        var directory = CreateTestDirectory();
        var workbookPath = Path.Combine(directory, "categorized.xlsx");

        try
        {
            using (var workbook = new XLWorkbook())
            {
                var sheet = workbook.Worksheets.Add("Personal");
                sheet.Cell(1, 1).Value = "Date";
                sheet.Cell(1, 2).Value = "Description";
                sheet.Cell(1, 3).Value = "Debit";
                sheet.Cell(2, 1).Value = new DateTime(2026, 4, 15);
                sheet.Cell(2, 2).Value = "COMPRA ORDEN DE PAGO SPEI 0260415 NO INGRESAD Aileen canto";
                sheet.Cell(2, 3).Value = 1250m;
                workbook.SaveAs(workbookPath);
            }

            var result = new QifExporter().Export(
                workbookPath,
                new[] { new QifAccountMapping("Personal", "Banorte Personal", "Bank") },
                allowUncategorizedTransactions: true);
            var qif = File.ReadAllText(Assert.Single(result.Files).OutputPath, Encoding.Latin1);

            Assert.Contains("PCOMPRA ORDEN DE PAGO SPEI 0260415 NO INGRESAD Aileen canto", qif);
            Assert.DoesNotContain("\nL\n", qif.Replace("\r\n", "\n", StringComparison.Ordinal));
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    private static void CreateWorkbook(string path, bool includeUncategorized)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Personal");
        sheet.Cell(1, 1).Value = "Fecha";
        sheet.Cell(1, 2).Value = "Descripcion";
        sheet.Cell(1, 3).Value = "Cargo";
        sheet.Cell(1, 4).Value = "Abono";
        sheet.Cell(1, 5).Value = "Memo";
        sheet.Cell(1, 6).Value = "Category";
        sheet.Cell(1, 7).Value = "Tag";
        sheet.Cell(2, 1).Value = new DateTime(2026, 1, 5);
        sheet.Cell(2, 2).Value = "Comisión CFE";
        sheet.Cell(2, 3).Value = 250.50m;
        sheet.Cell(2, 5).Value = "January bill";
        sheet.Cell(2, 6).Value = "Home:Bills";
        sheet.Cell(2, 7).Value = "Utilities";
        sheet.Cell(3, 1).Value = new DateTime(2026, 1, 6);
        sheet.Cell(3, 2).Value = "Payroll";
        sheet.Cell(3, 4).Value = 500m;
        if (!includeUncategorized) sheet.Cell(3, 6).Value = "Income:Salary";
        workbook.SaveAs(path);
    }

    private static string CreateTestDirectory()
    {
        var path = Path.Combine(Path.GetTempPath(), "BanortePDFToExcel.Tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(path);
        return path;
    }
}
