using ClosedXML.Excel;

namespace BanortePDFToExcel.Tests;

public sealed class ExcelWorkbookCategorizerTests
{
    [Fact]
    public void Categorize_DetectsNewLayoutAndPreservesEtiqueta()
    {
        var testDirectory = CreateTestDirectory();
        var inputPath = Path.Combine(testDirectory, "input.xlsx");
        var outputPath = Path.Combine(testDirectory, "output.xlsx");

        try
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Cuenta");
                worksheet.Cell(6, 1).Value = "Fecha";
                worksheet.Cell(6, 2).Value = "Descripcion";
                worksheet.Cell(6, 3).Value = "Cargo";
                worksheet.Cell(6, 4).Value = "Abono";
                worksheet.Cell(6, 5).Value = "Saldo";
                worksheet.Cell(6, 6).Value = "Etiqueta";
                worksheet.Cell(7, 2).Value = "TRASPASO, DE LA CUENTA: 1171346382";
                worksheet.Cell(7, 6).Value = "Direccion reparada";
                worksheet.Cell(8, 2).Value = "Unknown merchant";
                workbook.SaveAs(inputPath);
            }

            var categorizer = new ExcelWorkbookCategorizer(new TransactionPostProcessor());
            var result = categorizer.Categorize(inputPath, outputPath);

            using var categorizedWorkbook = new XLWorkbook(outputPath);
            var categorizedSheet = categorizedWorkbook.Worksheet("Cuenta");

            Assert.Equal("Etiqueta", categorizedSheet.Cell(6, 6).GetString());
            Assert.Equal("Direccion reparada", categorizedSheet.Cell(7, 6).GetString());
            Assert.Equal("Category", categorizedSheet.Cell(6, 7).GetString());
            Assert.Equal("[Banorte Negocio]", categorizedSheet.Cell(7, 7).GetString());
            Assert.Equal(string.Empty, categorizedSheet.Cell(8, 7).GetString());
            Assert.Equal(1, result.CategoriesAdded);
            Assert.Equal(1, result.UnmatchedTransactions);
        }
        finally
        {
            Directory.Delete(testDirectory, recursive: true);
        }
    }

    [Fact]
    public void Categorize_DoesNotReplaceAnExistingCategory()
    {
        var testDirectory = CreateTestDirectory();
        var inputPath = Path.Combine(testDirectory, "input.xlsx");
        var outputPath = Path.Combine(testDirectory, "output.xlsx");

        try
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Cuenta");
                worksheet.Cell(1, 1).Value = "Date";
                worksheet.Cell(1, 2).Value = "Description";
                worksheet.Cell(1, 3).Value = "Category";
                worksheet.Cell(2, 2).Value = "CFE";
                worksheet.Cell(2, 3).Value = "My manual category";
                workbook.SaveAs(inputPath);
            }

            var categorizer = new ExcelWorkbookCategorizer(new TransactionPostProcessor());
            var result = categorizer.Categorize(inputPath, outputPath);

            using var categorizedWorkbook = new XLWorkbook(outputPath);
            Assert.Equal("My manual category", categorizedWorkbook.Worksheet("Cuenta").Cell(2, 3).GetString());
            Assert.Equal(0, result.CategoriesAdded);
            Assert.Equal(1, result.ExistingCategories);
        }
        finally
        {
            Directory.Delete(testDirectory, recursive: true);
        }
    }

    [Fact]
    public void Categorize_CreatesMemoColumnForLongSpeiDescription()
    {
        var testDirectory = CreateTestDirectory();
        var inputPath = Path.Combine(testDirectory, "input.xlsx");
        var outputPath = Path.Combine(testDirectory, "output.xlsx");

        try
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Cuenta");
                worksheet.Cell(1, 1).Value = "Date";
                worksheet.Cell(1, 2).Value = "Description";
                worksheet.Cell(1, 3).Value = "Debit";
                worksheet.Cell(1, 4).Value = "Credit";
                worksheet.Cell(2, 1).Value = "04-03-26";
                worksheet.Cell(2, 2).Value = "COMPRA ORDEN DE PAGO SPEI 0260403 =REFERENCIA CTA/CLABE: 4152314364776081 BXI SPEI BCO:012 BENEF:LORENA YOSUHA RAMOS SANCHE (DATO NO VERIF POR ESTA INST), Homeschool teacher CVE RASTREO: 38432P05202604035153553752";
                worksheet.Cell(2, 3).Value = 100m;
                workbook.SaveAs(inputPath);
            }

            new ExcelWorkbookCategorizer(new TransactionPostProcessor()).Categorize(inputPath, outputPath);

            using var categorizedWorkbook = new XLWorkbook(outputPath);
            var categorizedSheet = categorizedWorkbook.Worksheet("Cuenta");
            var memoColumn = Enumerable.Range(1, categorizedSheet.LastColumnUsed()!.ColumnNumber())
                .Single(column => categorizedSheet.Cell(1, column).GetString() == "Memo");

            Assert.True(categorizedSheet.Cell(2, 2).GetString().Length <= 64);
            Assert.Contains("LORENA", categorizedSheet.Cell(2, 2).GetString());
            Assert.Contains("Homeschool teacher", categorizedSheet.Cell(2, memoColumn).GetString());
        }
        finally
        {
            Directory.Delete(testDirectory, recursive: true);
        }
    }

    private static string CreateTestDirectory()
    {
        var path = Path.Combine(Path.GetTempPath(), "BanortePDFToExcel.Tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(path);
        return path;
    }
}
