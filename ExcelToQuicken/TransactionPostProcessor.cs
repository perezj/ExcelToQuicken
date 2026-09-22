using System.Text.RegularExpressions;

namespace BanortePDFToExcel;

/// <summary>Applies user-supplied category rules. Description cleanup is intentionally separate.</summary>
public sealed class TransactionPostProcessor
{
    private readonly CategoryRuleEngine _engine;
    public TransactionPostProcessor(IEnumerable<CategoryRule>? rules = null) => _engine = new CategoryRuleEngine(rules ?? new CategoryRuleStore().Load());
    public string DetermineCategory(string description, string memo = "") => _engine.Match(description, memo).Category;
    public CategoryMatchResult DetermineMatch(string description, string memo = "") => _engine.Match(description, memo);

    public FormattedTransactionText FormatTransaction(string description, string memo = "") =>
        TransactionDescriptionFormatter.Format(description, memo);

    public List<Transaction> Process(List<Transaction> transactions)
    {
        foreach (var tx in transactions)
        {
            var formatted = FormatTransaction(tx.Description, tx.Memo);
            tx.Description = formatted.Description;
            tx.Memo = formatted.Memo;
            var match = _engine.Match(tx.Description, tx.Memo);
            if (match.Matched) { tx.Category = match.Category; if (!string.IsNullOrWhiteSpace(match.Tag)) tx.Tag = match.Tag; }
        }
        return transactions;
    }

}
