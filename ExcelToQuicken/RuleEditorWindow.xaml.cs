using System.Windows;

namespace BanortePDFToExcel;

public partial class RuleEditorWindow : Window
{
    private readonly CategoryRule _rule;
    public CategoryRule Result => _rule;
    public RuleEditorWindow(CategoryRule rule)
    {
        InitializeComponent(); _rule = rule;
        FieldBox.ItemsSource = Enum.GetValues<RuleMatchField>(); TypeBox.ItemsSource = Enum.GetValues<RuleMatchType>();
        NameBox.Text = rule.Name; PriorityBox.Text = rule.Priority.ToString(); FieldBox.SelectedItem = rule.MatchField; TypeBox.SelectedItem = rule.MatchType; MatchBox.Text = rule.MatchText; CategoryBox.Text = rule.Category; TagBox.Text = rule.Tag; PayeeBox.Text = rule.Payee ?? ""; EnabledBox.IsChecked = rule.Enabled; StopBox.IsChecked = rule.StopProcessing;
    }
    private void Test_Click(object sender, RoutedEventArgs e)
    {
        var test = _rule.Clone(); ReadInto(test); var result = CategoryRuleEngine.Test(test, SampleBox.Text, "");
        MessageBox.Show(this, result.Error ?? (result.Matched ? $"Matches. Category: {test.Category}\nTag: {test.Tag}" : "Does not match."), "Rule test", MessageBoxButton.OK, result.Error == null ? MessageBoxImage.Information : MessageBoxImage.Warning);
    }
    private void Ok_Click(object sender, RoutedEventArgs e) { try { ReadInto(_rule); CategoryRuleStore.Validate(new[] { _rule }); DialogResult = true; } catch (Exception ex) { MessageBox.Show(this, ex.Message, "Invalid rule", MessageBoxButton.OK, MessageBoxImage.Warning); } }
    private void ReadInto(CategoryRule r) { r.Name = NameBox.Text.Trim(); r.Priority = int.TryParse(PriorityBox.Text, out var p) ? p : 0; r.MatchField = (RuleMatchField)FieldBox.SelectedItem!; r.MatchType = (RuleMatchType)TypeBox.SelectedItem!; r.MatchText = MatchBox.Text; r.Category = CategoryBox.Text.Trim(); r.Tag = TagBox.Text.Trim(); r.Payee = PayeeBox.Text; r.Enabled = EnabledBox.IsChecked == true; r.StopProcessing = StopBox.IsChecked != false; }
}
