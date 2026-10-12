using System.Text.RegularExpressions;

namespace BanortePDFToExcel;

public sealed record FormattedTransactionText(string Description, string Memo);

public static class TransactionDescriptionFormatter
{
    private const int QuickenPayeeLimit = 64;
    private const int QuickenMemoLimit = 255;
    private const string TruncationMarker = "...";

    private static readonly Regex BeneficiaryRegex = new(
        @"\bBENEF\s*:\s*(?<value>.*?)(?=\s*\(|\s+DATO\b|\s+CVE\s+RASTREO\b|$)",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex ConceptRegex = new(
        @"\bINST\s*\)\s*,?\s*(?<value>.*?)(?=\s+CVE\s+RASTREO\b|$)",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex ReferenceMetadataRegex = new(
        @"\s*=\s*REFERENCIA\b.*$",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);

    public static FormattedTransactionText Format(string? description, string? memo = null)
    {
        var source = Normalize(description);
        var existingMemo = Normalize(memo);
        var details = ExtractDetails(source);

        var operation = ReferenceMetadataRegex.Replace(source, string.Empty).Trim();
        var payee = details.Count == 0
            ? source
            : string.Join(" ", new[] { operation }.Concat(details.Select(detail => detail.Value)))
                .Trim(' ', '-', ',');

        var detailMemo = details.Select(detail => $"{detail.Label}: {detail.Value}").ToList();
        var memoParts = detailMemo
            .Concat(IsSameAsSource(existingMemo, source) ? [] : [existingMemo])
            .Where(part => !string.IsNullOrWhiteSpace(part))
            .Distinct(StringComparer.OrdinalIgnoreCase);

        var formattedMemo = LimitAtWord(string.Join(" | ", memoParts), QuickenMemoLimit);
        return new(LimitAtWord(payee, QuickenPayeeLimit), formattedMemo);
    }

    private static List<(string Label, string Value)> ExtractDetails(string source)
    {
        var details = new List<(string Label, string Value)>();
        AddMatch(details, "BENEF", BeneficiaryRegex.Match(source));
        AddMatch(details, "Concepto", ConceptRegex.Match(source));
        return details;
    }

    private static void AddMatch(List<(string Label, string Value)> details, string label, Match match)
    {
        if (!match.Success) return;

        var value = Normalize(match.Groups["value"].Value).Trim(',', '.', ';');
        if (value.Length > 0) details.Add((label, value));
    }

    private static string Normalize(string? value) =>
        Regex.Replace(value ?? string.Empty, @"\s+", " ").Trim();

    private static bool IsSameAsSource(string memo, string source) =>
        memo.Length > 0 && string.Equals(memo, source, StringComparison.OrdinalIgnoreCase);

    private static string LimitAtWord(string value, int maxLength)
    {
        if (value.Length <= maxLength) return value;

        var cut = value[..(maxLength - TruncationMarker.Length)].TrimEnd();
        var lastSpace = cut.LastIndexOf(' ');
        if (lastSpace > 0) cut = cut[..lastSpace];
        return cut.TrimEnd(' ', '-', ',', ':') + TruncationMarker;
    }
}
