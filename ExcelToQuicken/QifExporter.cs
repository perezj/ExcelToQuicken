using ClosedXML.Excel;
using System.Globalization;
using System.IO;
using System.Text;

namespace BanortePDFToExcel;

public static class QifAccountTypes
{
    public static IReadOnlyList<string> All { get; } = new[] { "Bank", "CCard", "Cash" };
}

public sealed record QifAccountMapping(
    string WorksheetName,
    string QuickenAccountName,
    string AccountType,
    bool Include = true);

public sealed record QifAccountDiscovery(string WorksheetName, int TransactionCount);

public sealed record QifFileResult(
    string WorksheetName,
    string AccountName,
    string OutputPath,
    int TransactionCount,
    int UncategorizedTransactions);

public sealed record QifExportResult(IReadOnlyList<QifFileResult> Files)
{
    public int TransactionCount => Files.Sum(file => file.TransactionCount);
    public int UncategorizedTransactions => Files.Sum(file => file.UncategorizedTransactions);
}

public sealed class QifValidationException : Exception
{
    public QifValidationException(string message) : base(message) { }
}

public sealed class QifExporter
{
    private const int HeaderSearchLimit = 20;
    // Conservative import compatibility limit; the workbook retains the longer memo.
    private const int QifMemoLimit = 64;
    private static readonly Encoding QifEncoding = CreateQifEncoding();

    private static Encoding CreateQifEncoding()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        return Encoding.GetEncoding(1252);
    }

    private static readonly HashSet<string> DateHeaders =
        new(StringComparer.OrdinalIgnoreCase) { "DATE", "FECHA" };
    private static readonly HashSet<string> DescriptionHeaders =
        new(StringComparer.OrdinalIgnoreCase) { "DESCRIPTION", "DESCRIPCION", "CONCEPTO", "PAYEE", "BENEFICIARIO" };
    private static readonly HashSet<string> DebitHeaders =
        new(StringComparer.OrdinalIgnoreCase) { "DEBIT", "CARGO", "RETIRO", "WITHDRAWAL" };
    private static readonly HashSet<string> CreditHeaders =
        new(StringComparer.OrdinalIgnoreCase) { "CREDIT", "ABONO", "DEPOSITO", "DEPOSIT" };
    private static readonly HashSet<string> AmountHeaders =
        new(StringComparer.OrdinalIgnoreCase) { "AMOUNT", "IMPORTE", "MONTO" };
    private static readonly HashSet<string> CategoryHeaders =
        new(StringComparer.OrdinalIgnoreCase) { "CATEGORY", "CATEGORIA" };
    private static readonly HashSet<string> MemoHeaders =
        new(StringComparer.OrdinalIgnoreCase) { "MEMO", "NOTA", "NOTAS" };
    private static readonly HashSet<string> TagHeaders =
        new(StringComparer.OrdinalIgnoreCase) { "TAG", "ETIQUETA" };

    private static readonly string[] SupportedDateFormats =
    {
        "MM-dd-yy", "M-d-yy", "MM/dd/yy", "M/d/yy",
        "MM-dd-yyyy", "M-d-yyyy", "MM/dd/yyyy", "M/d/yyyy",
        "yyyy-MM-dd", "yyyy/MM/dd"
    };

    public IReadOnlyList<QifAccountDiscovery> DiscoverAccounts(string workbookPath)
    {
        ValidateWorkbookPath(workbookPath);
        using var workbook = new XLWorkbook(workbookPath);
        var discoveries = new List<QifAccountDiscovery>();

        foreach (var worksheet in workbook.Worksheets)
        {
            var layout = FindLayout(worksheet);
            if (layout is null) continue;
            discoveries.Add(new QifAccountDiscovery(
                worksheet.Name,
                ReadTransactions(worksheet, layout, validateRows: false).Count));
        }

        return discoveries;
    }

    public QifExportResult Export(
        string categorizedWorkbookPath,
        IEnumerable<QifAccountMapping> mappings,
        bool allowUncategorizedTransactions = false)
    {
        ValidateWorkbookPath(categorizedWorkbookPath);
        var selectedMappings = mappings.Where(mapping => mapping.Include).ToList();
        ValidateMappings(selectedMappings);

        using var workbook = new XLWorkbook(categorizedWorkbookPath);
        var pendingFiles = new List<(QifAccountMapping Mapping, List<QifTransaction> Transactions)>();
        var validationErrors = new List<string>();

        foreach (var mapping in selectedMappings)
        {
            var worksheet = workbook.Worksheets.FirstOrDefault(sheet =>
                string.Equals(sheet.Name, mapping.WorksheetName, StringComparison.OrdinalIgnoreCase));
            if (worksheet is null)
            {
                validationErrors.Add($"Worksheet '{mapping.WorksheetName}' was not found.");
                continue;
            }

            var layout = FindLayout(worksheet);
            if (layout is null)
            {
                validationErrors.Add($"Worksheet '{worksheet.Name}' does not contain recognizable transaction headers.");
                continue;
            }

            try
            {
                pendingFiles.Add((mapping, ReadTransactions(worksheet, layout, validateRows: true)));
            }
            catch (QifValidationException ex)
            {
                validationErrors.Add(ex.Message);
            }
        }

        var uncategorized = pendingFiles
            .SelectMany(file => file.Transactions.Select(transaction => (file.Mapping.WorksheetName, transaction)))
            .Where(item => string.IsNullOrWhiteSpace(item.transaction.Category))
            .ToList();

        if (!allowUncategorizedTransactions && uncategorized.Count > 0)
        {
            var examples = uncategorized.Take(5)
                .Select(item => $"{item.WorksheetName}: {item.transaction.Payee}");
            validationErrors.Add(
                $"{uncategorized.Count} transaction(s) do not have a category. Add rules or enable 'Include uncategorized transactions'. " +
                $"Examples: {string.Join("; ", examples)}");
        }

        if (validationErrors.Count > 0)
        {
            throw new QifValidationException(string.Join(Environment.NewLine, validationErrors));
        }

        var directory = Path.GetDirectoryName(Path.GetFullPath(categorizedWorkbookPath))
            ?? Directory.GetCurrentDirectory();
        var workbookName = Path.GetFileNameWithoutExtension(categorizedWorkbookPath);
        var usedPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var output = new List<QifFileResult>();

        foreach (var pending in pendingFiles)
        {
            var path = BuildOutputPath(directory, workbookName, pending.Mapping.QuickenAccountName, usedPaths);
            WriteQif(path, pending.Mapping, pending.Transactions);
            output.Add(new QifFileResult(
                pending.Mapping.WorksheetName,
                pending.Mapping.QuickenAccountName,
                path,
                pending.Transactions.Count,
                pending.Transactions.Count(transaction => string.IsNullOrWhiteSpace(transaction.Category))));
        }

        return new QifExportResult(output);
    }

    private static void ValidateMappings(IReadOnlyList<QifAccountMapping> mappings)
    {
        if (mappings.Count == 0)
        {
            throw new QifValidationException("Select at least one worksheet in the Quicken Export tab.");
        }

        foreach (var mapping in mappings)
        {
            if (string.IsNullOrWhiteSpace(mapping.WorksheetName))
                throw new QifValidationException("A selected worksheet name is empty.");
            if (string.IsNullOrWhiteSpace(mapping.QuickenAccountName))
                throw new QifValidationException($"Enter a Quicken account name for worksheet '{mapping.WorksheetName}'.");
            if (!QifAccountTypes.All.Contains(mapping.AccountType, StringComparer.OrdinalIgnoreCase))
                throw new QifValidationException($"Account type '{mapping.AccountType}' is not supported. Use Bank, CCard, or Cash.");
        }
    }

    private static void WriteQif(string path, QifAccountMapping mapping, IReadOnlyList<QifTransaction> transactions)
    {
        var lines = new List<string>
        {
            $"!Type:{NormalizeAccountType(mapping.AccountType)}"
        };

        foreach (var transaction in transactions)
        {
            var formattedText = TransactionDescriptionFormatter.Format(transaction.Payee, transaction.Memo);
            var amount = transaction.Amount.ToString("0.00", CultureInfo.InvariantCulture);
            lines.Add($"D{transaction.Date:MM/dd/yyyy}");
            lines.Add($"T{amount}");
            lines.Add($"P{CleanControlText(formattedText.Description)}");
            if (!string.IsNullOrWhiteSpace(formattedText.Memo))
                lines.Add($"M{LimitMemo(CleanControlText(formattedText.Memo))}");
            var category = BuildCategory(transaction.Category, transaction.Tag);
            if (!string.IsNullOrWhiteSpace(category))
                lines.Add($"L{category}");
            lines.Add("^");
        }

        File.WriteAllLines(path, lines, QifEncoding);
    }

    private static List<QifTransaction> ReadTransactions(
        IXLWorksheet worksheet,
        WorksheetLayout layout,
        bool validateRows)
    {
        var transactions = new List<QifTransaction>();
        var errors = new List<string>();
        var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? layout.HeaderRow;

        for (var row = layout.HeaderRow + 1; row <= lastRow; row++)
        {
            var payee = worksheet.Cell(row, layout.DescriptionColumn).GetString().Trim();
            if (string.IsNullOrWhiteSpace(payee)) continue;

            if (!TryReadDate(worksheet.Cell(row, layout.DateColumn), out var date))
            {
                if (validateRows) errors.Add($"{worksheet.Name} row {row}: invalid or missing date.");
                continue;
            }

            if (!TryReadAmount(worksheet, row, layout, out var amount))
            {
                if (validateRows) errors.Add($"{worksheet.Name} row {row}: invalid or missing amount.");
                continue;
            }

            transactions.Add(new QifTransaction(
                date,
                amount,
                payee,
                ReadOptional(worksheet, row, layout.MemoColumn),
                ReadOptional(worksheet, row, layout.CategoryColumn),
                ReadOptional(worksheet, row, layout.TagColumn)));
        }

        if (errors.Count > 0)
        {
            var visible = errors.Take(8).ToList();
            if (errors.Count > visible.Count) visible.Add($"...and {errors.Count - visible.Count} more row error(s).");
            throw new QifValidationException(string.Join(Environment.NewLine, visible));
        }

        return transactions;
    }

    private static bool TryReadDate(IXLCell cell, out DateTime date)
    {
        if (cell.TryGetValue<DateTime>(out date)) return true;
        var value = cell.GetString().Trim();
        if (DateTime.TryParseExact(value, SupportedDateFormats, CultureInfo.InvariantCulture,
                DateTimeStyles.AllowWhiteSpaces, out date)) return true;
        return DateTime.TryParse(value, CultureInfo.CurrentCulture, DateTimeStyles.AllowWhiteSpaces, out date);
    }

    private static bool TryReadAmount(IXLWorksheet worksheet, int row, WorksheetLayout layout, out decimal amount)
    {
        if (layout.AmountColumn.HasValue && TryReadDecimal(worksheet.Cell(row, layout.AmountColumn.Value), out amount))
            return true;

        var debit = 0m;
        var credit = 0m;
        var hasDebit = layout.DebitColumn.HasValue &&
                       TryReadDecimal(worksheet.Cell(row, layout.DebitColumn.Value), out debit);
        var hasCredit = layout.CreditColumn.HasValue &&
                        TryReadDecimal(worksheet.Cell(row, layout.CreditColumn.Value), out credit);
        if (!hasDebit && !hasCredit)
        {
            amount = 0;
            return false;
        }

        amount = (hasCredit ? Math.Abs(credit) : 0) - (hasDebit ? Math.Abs(debit) : 0);
        return true;
    }

    private static bool TryReadDecimal(IXLCell cell, out decimal value)
    {
        if (cell.IsEmpty())
        {
            value = 0;
            return false;
        }
        if (cell.TryGetValue<decimal>(out value)) return true;

        var text = cell.GetString().Trim();
        var negative = text.StartsWith('(') && text.EndsWith(')');
        text = text.Trim('(', ')').Replace("$", string.Empty).Replace("MXN", string.Empty, StringComparison.OrdinalIgnoreCase).Trim();
        var styles = NumberStyles.Number | NumberStyles.AllowCurrencySymbol | NumberStyles.AllowLeadingSign;
        if (decimal.TryParse(text, styles, CultureInfo.InvariantCulture, out value) ||
            decimal.TryParse(text, styles, CultureInfo.CurrentCulture, out value))
        {
            if (negative) value = -Math.Abs(value);
            return true;
        }
        return false;
    }

    private static WorksheetLayout? FindLayout(IXLWorksheet worksheet)
    {
        var lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 0;
        var lastRow = Math.Min(worksheet.LastRowUsed()?.RowNumber() ?? 0, HeaderSearchLimit);

        for (var row = 1; row <= lastRow; row++)
        {
            int? date = null, description = null, debit = null, credit = null, amount = null,
                category = null, memo = null, tag = null;

            for (var column = 1; column <= lastColumn; column++)
            {
                var header = NormalizeHeader(worksheet.Cell(row, column).GetString());
                if (DateHeaders.Contains(header)) date = column;
                else if (DescriptionHeaders.Contains(header)) description = column;
                else if (DebitHeaders.Contains(header)) debit = column;
                else if (CreditHeaders.Contains(header)) credit = column;
                else if (AmountHeaders.Contains(header)) amount = column;
                else if (CategoryHeaders.Contains(header)) category = column;
                else if (MemoHeaders.Contains(header)) memo = column;
                else if (TagHeaders.Contains(header)) tag = column;
            }

            if (date.HasValue && description.HasValue && (amount.HasValue || debit.HasValue || credit.HasValue))
                return new WorksheetLayout(row, date.Value, description.Value, debit, credit, amount, category, memo, tag);
        }

        return null;
    }

    private static string BuildOutputPath(
        string directory,
        string workbookName,
        string accountName,
        ISet<string> usedPaths)
    {
        var invalid = Path.GetInvalidFileNameChars().ToHashSet();
        var safeAccountName = new string(accountName.Trim().Select(character => invalid.Contains(character) ? '-' : character).ToArray()).Trim(' ', '.');
        if (string.IsNullOrWhiteSpace(safeAccountName)) safeAccountName = "Quicken-account";

        var basePath = Path.Combine(directory, $"{workbookName}-{safeAccountName}.qif");
        var path = basePath;
        var suffix = 2;
        while (File.Exists(path) || !usedPaths.Add(path))
            path = Path.Combine(directory, $"{workbookName}-{safeAccountName}-{suffix++}.qif");
        usedPaths.Add(path);
        return path;
    }

    private static string BuildCategory(string category, string tag)
    {
        var cleanCategory = CleanControlText(category);
        var cleanTag = CleanControlText(tag);
        if (string.IsNullOrWhiteSpace(cleanTag)) return cleanCategory;
        return $"{cleanCategory}/{cleanTag}";
    }

    private static string ReadOptional(IXLWorksheet worksheet, int row, int? column) =>
        column.HasValue ? worksheet.Cell(row, column.Value).GetString().Trim() : string.Empty;

    private static string NormalizeAccountType(string type) =>
        QifAccountTypes.All.First(candidate => string.Equals(candidate, type, StringComparison.OrdinalIgnoreCase));

    private static string LimitMemo(string value)
    {
        if (value.Length <= QifMemoLimit) return value;
        var shortened = value[..(QifMemoLimit - 3)].TrimEnd();
        var lastSpace = shortened.LastIndexOf(' ');
        if (lastSpace > 0) shortened = shortened[..lastSpace];
        return shortened.TrimEnd() + "...";
    }

    private static string CleanControlText(string value) =>
        string.Join(" ", value.Replace("…", "...").Replace('^', ' ').Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)).Trim();

    private static string NormalizeHeader(string value)
    {
        var decomposed = value.Trim().Normalize(NormalizationForm.FormD);
        var builder = new StringBuilder(decomposed.Length);
        foreach (var character in decomposed)
        {
            if (CharUnicodeInfo.GetUnicodeCategory(character) != UnicodeCategory.NonSpacingMark &&
                !char.IsWhiteSpace(character))
                builder.Append(char.ToUpperInvariant(character));
        }
        return builder.ToString().Normalize(NormalizationForm.FormC);
    }

    private static void ValidateWorkbookPath(string path)
    {
        if (!File.Exists(path)) throw new FileNotFoundException("The Excel file was not found.", path);
        if (!string.Equals(Path.GetExtension(path), ".xlsx", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Only .xlsx files are supported.", nameof(path));
    }

    private sealed record WorksheetLayout(
        int HeaderRow,
        int DateColumn,
        int DescriptionColumn,
        int? DebitColumn,
        int? CreditColumn,
        int? AmountColumn,
        int? CategoryColumn,
        int? MemoColumn,
        int? TagColumn);

    private sealed record QifTransaction(
        DateTime Date,
        decimal Amount,
        string Payee,
        string Memo,
        string Category,
        string Tag);
}
