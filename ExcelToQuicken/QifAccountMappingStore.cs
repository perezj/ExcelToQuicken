using System.Text.Json;
using System.IO;

namespace BanortePDFToExcel;

public sealed class QifAccountMappingStore
{
    private readonly string _filePath;

    public QifAccountMappingStore(string? directory = null)
    {
        var root = directory ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Bank2ExcelCategoryApplier");
        _filePath = Path.Combine(root, "quicken-account-mappings.json");
    }

    public IReadOnlyDictionary<string, QifAccountMapping> Load()
    {
        if (!File.Exists(_filePath))
            return new Dictionary<string, QifAccountMapping>(StringComparer.OrdinalIgnoreCase);

        try
        {
            var mappings = JsonSerializer.Deserialize<List<QifAccountMapping>>(File.ReadAllText(_filePath)) ?? new();
            return mappings
                .Where(mapping => !string.IsNullOrWhiteSpace(mapping.WorksheetName))
                .GroupBy(mapping => mapping.WorksheetName, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Last(), StringComparer.OrdinalIgnoreCase);
        }
        catch (JsonException)
        {
            return new Dictionary<string, QifAccountMapping>(StringComparer.OrdinalIgnoreCase);
        }
    }

    public void Save(IEnumerable<QifAccountMapping> mappings)
    {
        var directory = Path.GetDirectoryName(_filePath)!;
        Directory.CreateDirectory(directory);
        var json = JsonSerializer.Serialize(mappings.ToList(), new JsonSerializerOptions { WriteIndented = true });
        var temporaryPath = $"{_filePath}.{Guid.NewGuid():N}.tmp";
        File.WriteAllText(temporaryPath, json);
        File.Move(temporaryPath, _filePath, overwrite: true);
    }
}
