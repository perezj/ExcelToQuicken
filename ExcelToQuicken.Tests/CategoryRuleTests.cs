using System.Text.Json;

namespace BanortePDFToExcel.Tests;

public sealed class CategoryRuleTests
{
    [Fact]
    public void Store_RoundTripsVersionedRules()
    {
        var dir = Temp(); try { var store = new CategoryRuleStore(dir); store.Save(new[] { new CategoryRule { Name = "Test", MatchText = "café", Category = "Food" } }); var loaded = store.Load(); Assert.Equal(1, JsonDocument.Parse(File.ReadAllText(store.FilePath)).RootElement.GetProperty("SchemaVersion").GetInt32()); Assert.Equal("café", Assert.Single(loaded).MatchText); } finally { Directory.Delete(dir, true); }
    }
    [Fact]
    public void Engine_SupportsModesCaseAndAccents()
    {
        Assert.True(CategoryRuleEngine.Test(new CategoryRule { MatchText = "cafe", Category = "x" }, "CAFÉ", "").Matched); Assert.True(CategoryRuleEngine.Test(new CategoryRule { MatchType = RuleMatchType.StartsWith, MatchText = "abc", Category = "x" }, "ABC shop", "").Matched); Assert.True(CategoryRuleEngine.Test(new CategoryRule { MatchType = RuleMatchType.ExactMatch, MatchText = "cafe", Category = "x" }, "CAFÉ", "").Matched); Assert.True(CategoryRuleEngine.Test(new CategoryRule { MatchType = RuleMatchType.RegularExpression, MatchText = "^foo\\d+$", Category = "x" }, "FOO42", "").Matched);
    }
    [Fact]
    public void Engine_HonorsPriorityAndDisabledRules() { var rules = new[] { new CategoryRule { Priority = 2, MatchText = "shop", Category = "late" }, new CategoryRule { Priority = 1, MatchText = "shop", Category = "early" }, new CategoryRule { Priority = 0, MatchText = "shop", Category = "disabled", Enabled = false } }; Assert.Equal("early", new CategoryRuleEngine(rules).Match("shop", "").Category); }
    [Fact]
    public void Engine_ReturnsOptionalPayeeReplacement() { var result = new CategoryRuleEngine(new[] { new CategoryRule { MatchText = "merchant", Category = "Food", Payee = "Friendly Merchant" } }).Match("merchant purchase", ""); Assert.Equal("Friendly Merchant", result.Payee); }
    [Fact]
    public void Engine_ReportsInvalidRegexWithoutThrowing() { Assert.NotNull(CategoryRuleEngine.Test(new CategoryRule { MatchType = RuleMatchType.RegularExpression, MatchText = "[", Category = "x" }, "anything", "").Error); }
    [Fact]
    public void Store_RejectsInvalidImports() { Assert.Throws<InvalidDataException>(() => CategoryRuleStore.ParseAndValidate("{\"SchemaVersion\":99,\"Rules\":[]}")); Assert.Throws<InvalidDataException>(() => CategoryRuleStore.ParseAndValidate("{\"SchemaVersion\":1,\"Rules\":[{\"Name\":\"bad\"}]}")); }
    private static string Temp() { var path = Path.Combine(Path.GetTempPath(), "BanortePDFToExcel.Tests", Guid.NewGuid().ToString("N")); Directory.CreateDirectory(path); return path; }
}
