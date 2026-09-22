namespace BanortePDFToExcel;

public sealed class QifAccountMappingRow
{
    public bool Include { get; set; } = true;
    public required string WorksheetName { get; init; }
    public string QuickenAccountName { get; set; } = string.Empty;
    public string AccountType { get; set; } = "Bank";
    public int TransactionCount { get; init; }
    public IReadOnlyList<string> AccountTypes => QifAccountTypes.All;

    public QifAccountMapping ToMapping() =>
        new(WorksheetName, QuickenAccountName.Trim(), AccountType, Include);
}
