using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace BanortePDFToExcel;

public enum RuleMatchField { Description, Memo, Both }
public enum RuleMatchType { Contains, StartsWith, ExactMatch, RegularExpression }

public sealed class CategoryRule : INotifyPropertyChanged
{
    private Guid _id = Guid.NewGuid();
    private string _name = "New rule";
    private bool _enabled = true;
    private int _priority;
    private RuleMatchField _matchField = RuleMatchField.Description;
    private RuleMatchType _matchType = RuleMatchType.Contains;
    private string _matchText = "";
    private string _category = "";
    private string _tag = "";
    private string? _payee;
    private bool _stopProcessing = true;

    public Guid Id { get => _id; set => Set(ref _id, value); }
    public string Name { get => _name; set => Set(ref _name, value); }
    public bool Enabled { get => _enabled; set => Set(ref _enabled, value); }
    public int Priority { get => _priority; set => Set(ref _priority, value); }
    public RuleMatchField MatchField { get => _matchField; set => Set(ref _matchField, value); }
    public RuleMatchType MatchType { get => _matchType; set => Set(ref _matchType, value); }
    public string MatchText { get => _matchText; set => Set(ref _matchText, value); }
    public string Category { get => _category; set => Set(ref _category, value); }
    public string Tag { get => _tag; set => Set(ref _tag, value); }
    public string? Payee { get => _payee; set => Set(ref _payee, string.IsNullOrWhiteSpace(value) ? null : value.Trim()); }
    public bool StopProcessing { get => _stopProcessing; set => Set(ref _stopProcessing, value); }

    public CategoryRule Clone() => new()
    {
        Id = Id, Name = Name, Enabled = Enabled, Priority = Priority, MatchField = MatchField,
        MatchType = MatchType, MatchText = MatchText, Category = Category, Tag = Tag, Payee = Payee,
        StopProcessing = StopProcessing
    };

    public event PropertyChangedEventHandler? PropertyChanged;
    private void Set<T>(ref T field, T value, [CallerMemberName] string? name = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value)) return;
        field = value; PropertyChanged?.Invoke(this, new(name));
    }
}

public sealed record CategoryRuleFile(int SchemaVersion, List<CategoryRule> Rules);
