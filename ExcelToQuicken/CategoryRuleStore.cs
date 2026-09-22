using System.Text.Json;
using System.Text.Json.Serialization;
using System.IO;

namespace BanortePDFToExcel;

public sealed class CategoryRuleStore
{
    public const int CurrentSchemaVersion = 1;
    public string DirectoryPath { get; }
    public string FilePath => Path.Combine(DirectoryPath, "rules.json");
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true, Converters = { new JsonStringEnumConverter() } };

    public CategoryRuleStore(string? directoryPath = null)
    {
        DirectoryPath = directoryPath ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Bank2ExcelCategoryApplier");
    }

    public List<CategoryRule> Load()
    {
        Directory.CreateDirectory(DirectoryPath);
        if (!File.Exists(FilePath))
        {
            var defaults = LoadBundledDefaults();
            Save(defaults);
            return defaults;
        }
        var file = JsonSerializer.Deserialize<CategoryRuleFile>(File.ReadAllText(FilePath), JsonOptions)
                   ?? throw new InvalidDataException("The category rules file is empty or invalid.");
        if (file.SchemaVersion != CurrentSchemaVersion) throw new InvalidDataException($"Unsupported category rules schema version {file.SchemaVersion}.");
        Validate(file.Rules);
        return file.Rules;
    }

    public void Save(IEnumerable<CategoryRule> rules)
    {
        var list = rules.Select(r => r.Clone()).ToList(); Validate(list);
        Directory.CreateDirectory(DirectoryPath);
        var temp = FilePath + ".tmp-" + Guid.NewGuid().ToString("N");
        File.WriteAllText(temp, JsonSerializer.Serialize(new CategoryRuleFile(CurrentSchemaVersion, list), JsonOptions));
        File.Move(temp, FilePath, true);
    }

    public static List<CategoryRule> ParseAndValidate(string json)
    {
        var file = JsonSerializer.Deserialize<CategoryRuleFile>(json, JsonOptions) ?? throw new InvalidDataException("Invalid rules JSON.");
        if (file.SchemaVersion != CurrentSchemaVersion) throw new InvalidDataException($"Unsupported category rules schema version {file.SchemaVersion}.");
        Validate(file.Rules); return file.Rules;
    }

    public static void Validate(IEnumerable<CategoryRule> rules)
    {
        foreach (var rule in rules)
        {
            if (string.IsNullOrWhiteSpace(rule.MatchText)) throw new InvalidDataException($"Rule '{rule.Name}' has no match text.");
            if (string.IsNullOrWhiteSpace(rule.Category)) throw new InvalidDataException($"Rule '{rule.Name}' has no category.");
            if (rule.MatchType == RuleMatchType.RegularExpression)
                try { _ = new System.Text.RegularExpressions.Regex(rule.MatchText); } catch (Exception ex) { throw new InvalidDataException($"Rule '{rule.Name}' has an invalid regular expression: {ex.Message}"); }
        }
    }

    private static List<CategoryRule> LoadBundledDefaults()
    {
        var path = Path.Combine(AppContext.BaseDirectory, "default-rules.json");
        if (!File.Exists(path)) return new();
        return ParseAndValidate(File.ReadAllText(path));
    }
}
