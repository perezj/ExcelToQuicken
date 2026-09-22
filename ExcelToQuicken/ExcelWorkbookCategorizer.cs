using ClosedXML.Excel;
using System.Globalization;
using System.IO;
using System.Text;

namespace BanortePDFToExcel;

public sealed class ExcelWorkbookCategorizer
{
    private const int HeaderSearchLimit = 20;

    private static readonly HashSet<string> DescriptionHeaders =
        new(StringComparer.OrdinalIgnoreCase) { "DESCRIPCION", "DESCRIPTION", "CONCEPTO" };

    private static readonly HashSet<string> CategoryHeaders =
        new(StringComparer.OrdinalIgnoreCase) { "CATEGORY", "CATEGORIA" };

    private static readonly HashSet<string> MemoHeaders =
        new(StringComparer.OrdinalIgnoreCase) { "MEMO" };

    private readonly TransactionPostProcessor _postProcessor;

    public ExcelWorkbookCategorizer(TransactionPostProcessor postProcessor)
    {
        _postProcessor = postProcessor;
    }

    public WorkbookCategorizationResult Categorize(string inputPath, string? outputPath = null, bool overwriteExistingCategories = false)
    {
        var fullInputPath = Path.GetFullPath(inputPath);
        ValidateInput(fullInputPath);

        var fullOutputPath = Path.GetFullPath(outputPath ?? BuildOutputPath(fullInputPath));
        if (string.Equals(fullInputPath, fullOutputPath, StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException("The output file must be different from the input file.", nameof(outputPath));
        }

        var outputDirectory = Path.GetDirectoryName(fullOutputPath);
        if (!string.IsNullOrWhiteSpace(outputDirectory))
        {
            Directory.CreateDirectory(outputDirectory);
        }

        using var workbook = new XLWorkbook(fullInputPath);
        var worksheetsProcessed = 0;
        var categoriesAdded = 0;
        var existingCategories = 0;
        var unmatchedTransactions = 0;
        var skippedWorksheets = new List<string>();

        foreach (var worksheet in workbook.Worksheets)
        {
            var layout = FindLayout(worksheet);
            if (layout is null)
            {
                skippedWorksheets.Add(worksheet.Name);
                continue;
            }

            worksheetsProcessed++;
            var categoryColumn = layout.CategoryColumn ?? AppendCategoryColumn(worksheet, layout);
            int? memoColumn = layout.MemoColumn;
            var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? layout.HeaderRow;

            for (var row = layout.HeaderRow + 1; row <= lastRow; row++)
            {
                var descriptionCell = worksheet.Cell(row, layout.DescriptionColumn);
                var description = descriptionCell.GetString().Trim();
                if (string.IsNullOrWhiteSpace(description))
                {
                    continue;
                }

                var memo = memoColumn.HasValue
                    ? worksheet.Cell(row, memoColumn.Value).GetString()
                    : string.Empty;
                var formatted = _postProcessor.FormatTransaction(description, memo);
                if (!string.Equals(description, formatted.Description, StringComparison.Ordinal))
                {
                    descriptionCell.Value = formatted.Description;
                }

                if (!string.IsNullOrWhiteSpace(formatted.Memo))
                {
                    memoColumn ??= AppendMemoColumn(worksheet, layout, categoryColumn);
                    worksheet.Cell(row, memoColumn.Value).Value = formatted.Memo;
                }

                var categoryCell = worksheet.Cell(row, categoryColumn);
                var hadExistingCategory = categoryCell.HasFormula || !string.IsNullOrWhiteSpace(categoryCell.GetString());
                if (hadExistingCategory) existingCategories++;
                if (!overwriteExistingCategories && hadExistingCategory)
                {
                    continue;
                }

                var match = _postProcessor.DetermineMatch(formatted.Description, formatted.Memo);
                if (match.Error != null) throw new InvalidDataException(match.Error);
                var category = match.Category;

                if (string.IsNullOrWhiteSpace(category))
                {
                    unmatchedTransactions++;
                    continue;
                }

                categoryCell.Value = category;
                if (!string.IsNullOrWhiteSpace(match.Tag))
                {
                    var tagColumn = FindOrAppendTagColumn(worksheet, layout);
                    worksheet.Cell(row, tagColumn).Value = match.Tag;
                }
                categoriesAdded++;
            }

            if (!layout.CategoryColumn.HasValue)
            {
                worksheet.Column(categoryColumn).Width = 34;
            }
        }

        if (worksheetsProcessed == 0)
        {
            throw new InvalidDataException(
                "No transaction worksheet was found. Expected a header named Descripcion, Description, or Concepto in the first 20 rows.");
        }

        workbook.SaveAs(fullOutputPath);

        return new WorkbookCategorizationResult(
            fullOutputPath,
            worksheetsProcessed,
            categoriesAdded,
            existingCategories,
            unmatchedTransactions,
            skippedWorksheets);
    }

    private static void ValidateInput(string inputPath)
    {
        if (!File.Exists(inputPath))
        {
            throw new FileNotFoundException("The Excel file was not found.", inputPath);
        }

        if (!string.Equals(Path.GetExtension(inputPath), ".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException("Only .xlsx files are supported. Save macro-enabled files as .xlsx before processing.", nameof(inputPath));
        }
    }

    private static string BuildOutputPath(string inputPath)
    {
        var directory = Path.GetDirectoryName(inputPath) ?? Directory.GetCurrentDirectory();
        var baseName = Path.GetFileNameWithoutExtension(inputPath);
        var candidate = Path.Combine(directory, $"{baseName}-categorized.xlsx");

        if (!File.Exists(candidate))
        {
            return candidate;
        }

        return Path.Combine(directory, $"{baseName}-categorized-{DateTime.Now:yyyyMMdd-HHmmss}.xlsx");
    }

    private static WorksheetLayout? FindLayout(IXLWorksheet worksheet)
    {
        var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 0;
        var lastRow = Math.Min(worksheet.LastRowUsed()?.RowNumber() ?? 0, HeaderSearchLimit);

        for (var row = 1; row <= lastRow; row++)
        {
            int? descriptionColumn = null;
            int? categoryColumn = null;
            int? memoColumn = null;

            for (var column = 1; column <= lastColumn; column++)
            {
                var header = NormalizeHeader(worksheet.Cell(row, column).GetString());
                if (DescriptionHeaders.Contains(header))
                {
                    descriptionColumn = column;
                }
                else if (CategoryHeaders.Contains(header))
                {
                    categoryColumn = column;
                }
                else if (MemoHeaders.Contains(header))
                {
                    memoColumn = column;
                }
            }

            if (descriptionColumn.HasValue)
            {
                return new WorksheetLayout(row, descriptionColumn.Value, categoryColumn, memoColumn, lastColumn);
            }
        }

        return null;
    }

    private static int AppendCategoryColumn(IXLWorksheet worksheet, WorksheetLayout layout)
    {
        var categoryColumn = layout.LastUsedColumn + 1;
        var headerCell = worksheet.Cell(layout.HeaderRow, categoryColumn);
        var sourceHeaderCell = worksheet.Cell(layout.HeaderRow, layout.DescriptionColumn);

        headerCell.Value = "Category";
        headerCell.Style = sourceHeaderCell.Style;

        var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? layout.HeaderRow;
        for (var row = layout.HeaderRow + 1; row <= lastRow; row++)
        {
            worksheet.Cell(row, categoryColumn).Style = worksheet.Cell(row, layout.DescriptionColumn).Style;
        }

        return categoryColumn;
    }

    private static int AppendMemoColumn(IXLWorksheet worksheet, WorksheetLayout layout, int categoryColumn)
    {
        var memoColumn = worksheet.LastColumnUsed()?.ColumnNumber() + 1 ?? Math.Max(layout.LastUsedColumn, categoryColumn) + 1;
        var headerCell = worksheet.Cell(layout.HeaderRow, memoColumn);
        var sourceHeaderCell = worksheet.Cell(layout.HeaderRow, layout.DescriptionColumn);

        headerCell.Value = "Memo";
        headerCell.Style = sourceHeaderCell.Style;
        worksheet.Column(memoColumn).Width = 50;

        var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? layout.HeaderRow;
        for (var row = layout.HeaderRow + 1; row <= lastRow; row++)
        {
            worksheet.Cell(row, memoColumn).Style = worksheet.Cell(row, layout.DescriptionColumn).Style;
        }

        return memoColumn;
    }

    private static int FindOrAppendTagColumn(IXLWorksheet worksheet, WorksheetLayout layout)
    {
        var last = worksheet.LastColumnUsed()?.ColumnNumber() ?? layout.LastUsedColumn;
        for (var c = 1; c <= last; c++)
        {
            var header = NormalizeHeader(worksheet.Cell(layout.HeaderRow, c).GetString());
            if (header is "TAG" or "ETIQUETA") return c;
        }
        var column = last + 1;
        worksheet.Cell(layout.HeaderRow, column).Value = "Tag";
        worksheet.Cell(layout.HeaderRow, column).Style = worksheet.Cell(layout.HeaderRow, layout.DescriptionColumn).Style;
        return column;
    }

    private static string NormalizeHeader(string value)
    {
        var decomposed = value.Trim().Normalize(NormalizationForm.FormD);
        var builder = new StringBuilder(decomposed.Length);

        foreach (var character in decomposed)
        {
            if (CharUnicodeInfo.GetUnicodeCategory(character) != UnicodeCategory.NonSpacingMark &&
                !char.IsWhiteSpace(character))
            {
                builder.Append(char.ToUpperInvariant(character));
            }
        }

        return builder.ToString().Normalize(NormalizationForm.FormC);
    }

    private sealed record WorksheetLayout(
        int HeaderRow,
        int DescriptionColumn,
        int? CategoryColumn,
        int? MemoColumn,
        int LastUsedColumn);
}

public sealed record WorkbookCategorizationResult(
    string OutputPath,
    int WorksheetsProcessed,
    int CategoriesAdded,
    int ExistingCategories,
    int UnmatchedTransactions,
    IReadOnlyList<string> SkippedWorksheets);
