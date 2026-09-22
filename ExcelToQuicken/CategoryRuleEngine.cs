using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace BanortePDFToExcel;

public sealed record CategoryMatchResult(bool Matched, string Category = "", string Tag = "", string? Error = null);

public sealed class CategoryRuleEngine
{
    private readonly IReadOnlyList<CategoryRule> _rules;
    public CategoryRuleEngine(IEnumerable<CategoryRule> rules) => _rules = rules.Where(r => r.Enabled).OrderBy(r => r.Priority).ThenBy(r => r.Id).ToList();

    public CategoryMatchResult Match(string? description, string? memo)
    {
        CategoryMatchResult? lastMatch = null;
        foreach (var rule in _rules)
        {
            var result = Matches(rule, description ?? "", memo ?? "");
            if (result.Error != null) return result;
            if (result.Matched)
            {
                lastMatch = new(true, rule.Category, rule.Tag);
                if (rule.StopProcessing) return lastMatch;
            }
        }
        return lastMatch ?? new(false);
    }

    public static CategoryMatchResult Test(CategoryRule rule, string description, string memo) => Matches(rule, description, memo);

    private static CategoryMatchResult Matches(CategoryRule rule, string description, string memo)
    {
        var values = rule.MatchField switch { RuleMatchField.Memo => new[] { memo }, RuleMatchField.Both => new[] { description, memo }, _ => new[] { description } };
        try
        {
            foreach (var value in values)
            {
                var left = rule.MatchType == RuleMatchType.RegularExpression ? value : Fold(value);
                var right = rule.MatchType == RuleMatchType.RegularExpression ? rule.MatchText : Fold(rule.MatchText);
                var matched = rule.MatchType switch
                {
                    RuleMatchType.Contains => left.Contains(right, StringComparison.OrdinalIgnoreCase),
                    RuleMatchType.StartsWith => left.StartsWith(right, StringComparison.OrdinalIgnoreCase),
                    RuleMatchType.ExactMatch => string.Equals(left, right, StringComparison.OrdinalIgnoreCase),
                    RuleMatchType.RegularExpression => Regex.IsMatch(left, right, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant),
                    _ => false
                };
                if (matched) return new(true);
            }
            return new(false);
        }
        catch (ArgumentException ex) { return new(false, Error: $"Rule '{rule.Name}' has an invalid regular expression: {ex.Message}"); }
    }

    private static string Fold(string value)
    {
        var form = value.Normalize(NormalizationForm.FormD);
        return new string(form.Where(c => CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark).ToArray()).Normalize(NormalizationForm.FormC);
    }
}
