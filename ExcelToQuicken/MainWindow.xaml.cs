using Microsoft.Win32;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;

namespace BanortePDFToExcel;

public partial class MainWindow : Window
{
    private static readonly Brush DefaultDropBorder = new SolidColorBrush(Color.FromRgb(148, 163, 184));
    private static readonly Brush ActiveDropBorder = new SolidColorBrush(Color.FromRgb(22, 163, 74));
    private readonly CategoryRuleStore _store = new();
    private readonly QifExporter _qifExporter = new();
    private readonly QifAccountMappingStore _qifMappingStore = new();
    private readonly ObservableCollection<CategoryRule> _rules;
    private readonly ObservableCollection<QifAccountMappingRow> _qifMappings = new();
    private readonly ICollectionView _ruleView;
    private string? _selectedFilePath, _lastOutputPath;
    private bool _isConverting, _dirty;
    public MainWindow()
    {
        InitializeComponent();
        _rules = new ObservableCollection<CategoryRule>(_store.Load());
        _ruleView = CollectionViewSource.GetDefaultView(_rules);
        RulesGrid.ItemsSource = _ruleView;
        QifMappingsGrid.ItemsSource = _qifMappings;
        Closing += MainWindow_Closing;
    }
    private void MainWindow_Closing(object? sender, CancelEventArgs e) { if (_dirty && MessageBox.Show(this, "You have unsaved category rule changes. Close without saving?", "Unsaved changes", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes) e.Cancel = true; }
    private void RuleSearchBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e) { var q = RuleSearchBox.Text.Trim(); _ruleView.Filter = o => o is CategoryRule r && (string.IsNullOrWhiteSpace(q) || r.Name.Contains(q, StringComparison.OrdinalIgnoreCase) || r.MatchText.Contains(q, StringComparison.OrdinalIgnoreCase) || r.Category.Contains(q, StringComparison.OrdinalIgnoreCase)); }
    private CategoryRule? Selected => RulesGrid.SelectedItem as CategoryRule;
    private void AddRule_Click(object s, RoutedEventArgs e) { var rule = new CategoryRule { Priority = _rules.Count }; if (Edit(rule)) { _rules.Add(rule); Dirty(); } }
    private void EditRule_Click(object s, RoutedEventArgs e) { if (Selected is { } r && Edit(r)) Dirty(); }
    private bool Edit(CategoryRule rule) { var dialog = new RuleEditorWindow(rule) { Owner = this }; return dialog.ShowDialog() == true; }
    private void DuplicateRule_Click(object s, RoutedEventArgs e) { if (Selected is { } r) { var copy = r.Clone(); copy.Id = Guid.NewGuid(); copy.Name += " (copy)"; _rules.Add(copy); Dirty(); RulesGrid.SelectedItem = copy; } }
    private void DeleteRule_Click(object s, RoutedEventArgs e) { if (Selected is { } r && MessageBox.Show(this, $"Delete '{r.Name}'?", "Confirm delete", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes) { _rules.Remove(r); Dirty(); } }
    private void ToggleRule_Click(object s, RoutedEventArgs e) { if (Selected is { } r) { r.Enabled = !r.Enabled; Dirty(); } }
    private void MoveUp_Click(object s, RoutedEventArgs e) => Move(-1);
    private void MoveDown_Click(object s, RoutedEventArgs e) => Move(1);
    private void Move(int delta) { if (Selected is not { } r) return; var i = _rules.IndexOf(r); var j = i + delta; if (j < 0 || j >= _rules.Count) return; _rules.Move(i, j); for (var n = 0; n < _rules.Count; n++) _rules[n].Priority = n; Dirty(); RulesGrid.SelectedItem = r; }
    private void SaveRules_Click(object s, RoutedEventArgs e) { try { CategoryRuleStore.Validate(_rules); _store.Save(_rules); _dirty = false; SetStatus("Rules saved.", false, true); } catch (Exception ex) { MessageBox.Show(this, ex.Message, "Could not save rules", MessageBoxButton.OK, MessageBoxImage.Warning); } }
    private void CancelRules_Click(object s, RoutedEventArgs e) { if (!_dirty) return; _rules.Clear(); foreach (var r in _store.Load()) _rules.Add(r); _dirty = false; }
    private void ImportRules_Click(object s, RoutedEventArgs e) { var d = new OpenFileDialog { Filter = "JSON rules (*.json)|*.json", CheckFileExists = true }; if (d.ShowDialog(this) != true) return; try { var imported = CategoryRuleStore.ParseAndValidate(File.ReadAllText(d.FileName)); if (MessageBox.Show(this, "Replace the current rules with the imported rules?", "Confirm import", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes) { _rules.Clear(); foreach (var r in imported) _rules.Add(r); Dirty(); } } catch (Exception ex) { MessageBox.Show(this, ex.Message, "Invalid rules file", MessageBoxButton.OK, MessageBoxImage.Warning); } }
    private void ExportRules_Click(object s, RoutedEventArgs e) { var d = new SaveFileDialog { Filter = "JSON rules (*.json)|*.json", FileName = "rules.json" }; if (d.ShowDialog(this) == true) try { File.WriteAllText(d.FileName, JsonSerializer.Serialize(new CategoryRuleFile(CategoryRuleStore.CurrentSchemaVersion, _rules.ToList()), new JsonSerializerOptions { WriteIndented = true, Converters = { new System.Text.Json.Serialization.JsonStringEnumConverter() } })); } catch (Exception ex) { MessageBox.Show(this, ex.Message, "Could not export rules", MessageBoxButton.OK, MessageBoxImage.Error); } }
    private void ResetDefaults_Click(object s, RoutedEventArgs e) { if (MessageBox.Show(this, "Reset all rules to bundled defaults?", "Confirm reset", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes) return; var temp = new CategoryRuleStore(Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"))); var defaults = temp.Load(); _rules.Clear(); foreach (var r in defaults) _rules.Add(r); Dirty(); }
    private void Dirty() { _dirty = true; }
    private void BrowseButton_Click(object s, RoutedEventArgs e) { var d = new OpenFileDialog { Title = "Select a Bank2Excel workbook", Filter = "Excel workbooks (*.xlsx)|*.xlsx", CheckFileExists = true }; if (d.ShowDialog(this) == true) SelectFile(d.FileName); }
    private void DropZone_DragEnter(object s, DragEventArgs e) { e.Effects = !_isConverting && TryGetDroppedWorkbook(e.Data, out _) ? DragDropEffects.Copy : DragDropEffects.None; if (e.Effects == DragDropEffects.Copy) { DropZone.BorderBrush = ActiveDropBorder; DropZone.Background = new SolidColorBrush(Color.FromRgb(240, 253, 244)); } e.Handled = true; }
    private void DropZone_DragLeave(object s, DragEventArgs e) { ResetDropZone(); }
    private void DropZone_Drop(object s, DragEventArgs e) { ResetDropZone(); if (TryGetDroppedWorkbook(e.Data, out var p)) SelectFile(p!); else SetStatus("Drop one .xlsx file here.", true); }
    private async void ConvertButton_Click(object s, RoutedEventArgs e)
    {
        if (_isConverting || _selectedFilePath == null) return;

        QifMappingsGrid.CommitEdit(System.Windows.Controls.DataGridEditingUnit.Cell, true);
        QifMappingsGrid.CommitEdit(System.Windows.Controls.DataGridEditingUnit.Row, true);
        var input = _selectedFilePath;
        var overwriteExistingCategories = OverwriteCategoriesCheckBox.IsChecked == true;
        var createQif = CreateQifCheckBox.IsChecked == true;
        var allowUncategorized = IncludeUncategorizedCheckBox.IsChecked == true;
        var mappings = _qifMappings.Select(row => row.ToMapping()).ToList();

        SetBusy(true);
        SetStatus("Applying category rules...", false);
        OpenOutputButton.Visibility = Visibility.Collapsed;

        try
        {
            var result = await Task.Run(() =>
            {
                var rules = _store.Load();
                return new ExcelWorkbookCategorizer(new TransactionPostProcessor(rules))
                    .Categorize(input, null, overwriteExistingCategories);
            });

            _lastOutputPath = result.OutputPath;
            OpenOutputButton.Visibility = Visibility.Visible;

            if (!createQif)
            {
                SetStatus($"Done. Categorized {result.CategoriesAdded}; already categorized {result.ExistingCategories}; unmatched {result.UnmatchedTransactions}.", false, true);
                return;
            }

            _qifMappingStore.Save(mappings);
            SetStatus("Creating Quicken QIF files...", false);
            try
            {
                var qifResult = await Task.Run(() => _qifExporter.Export(result.OutputPath, mappings, allowUncategorized));
                SetStatus(
                    $"Done. Categorized {result.CategoriesAdded}; created {qifResult.Files.Count} QIF file(s) with {qifResult.TransactionCount} transaction(s).",
                    false,
                    true);
            }
            catch (QifValidationException ex)
            {
                SetStatus("The categorized workbook was saved, but QIF creation needs attention.", true);
                MessageBox.Show(this, ex.Message, "QIF was not created", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        catch (Exception ex)
        {
            _lastOutputPath = null;
            SetStatus(ex.Message, true);
            MessageBox.Show(this, ex.Message, "Could not convert workbook", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            SetBusy(false);
        }
    }
    private void OpenOutputButton_Click(object s, RoutedEventArgs e)
    {
        if (_lastOutputPath == null || !File.Exists(_lastOutputPath)) return;
        var directory = Path.GetDirectoryName(_lastOutputPath);
        if (directory != null) Process.Start(new ProcessStartInfo(directory) { UseShellExecute = true });
    }
    private void SelectFile(string path)
    {
        if (!File.Exists(path) || !string.Equals(Path.GetExtension(path), ".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            SetStatus("Select an existing .xlsx file.", true);
            return;
        }

        try
        {
            var fullPath = Path.GetFullPath(path);
            var discoveries = _qifExporter.DiscoverAccounts(fullPath);
            var savedMappings = _qifMappingStore.Load();
            _qifMappings.Clear();
            foreach (var discovery in discoveries)
            {
                savedMappings.TryGetValue(discovery.WorksheetName, out var saved);
                _qifMappings.Add(new QifAccountMappingRow
                {
                    WorksheetName = discovery.WorksheetName,
                    TransactionCount = discovery.TransactionCount,
                    Include = saved?.Include ?? true,
                    QuickenAccountName = saved?.QuickenAccountName ?? discovery.WorksheetName.Trim(),
                    AccountType = QifAccountTypes.All.Contains(saved?.AccountType ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                        ? saved!.AccountType
                        : "Bank"
                });
            }

            _selectedFilePath = fullPath;
            SelectedFileText.Text = Path.GetFileName(path);
            SelectedFileText.ToolTip = path;
            ConvertButton.IsEnabled = true;
            QifMappingHintText.Text = discoveries.Count == 0
                ? "No QIF-compatible transaction worksheets were detected. You can still create the categorized workbook."
                : $"Detected {discoveries.Count} account worksheet(s). Confirm the Quicken names before converting.";
            SetStatus("Ready to apply categories and create the selected exports.", false);
        }
        catch (Exception ex)
        {
            _selectedFilePath = null;
            _qifMappings.Clear();
            ConvertButton.IsEnabled = false;
            SetStatus(ex.Message, true);
        }
    }
    private void CreateQifCheckBox_Changed(object s, RoutedEventArgs e)
    {
        if (QifOptionsTab != null) QifOptionsTab.IsEnabled = CreateQifCheckBox.IsChecked == true;
    }
    private void SetBusy(bool busy)
    {
        _isConverting = busy;
        BrowseButton.IsEnabled = !busy;
        ConvertButton.IsEnabled = !busy && _selectedFilePath != null;
        QifMappingsGrid.IsEnabled = !busy;
        CreateQifCheckBox.IsEnabled = !busy;
        IncludeUncategorizedCheckBox.IsEnabled = !busy;
        ConversionProgress.Visibility = busy ? Visibility.Visible : Visibility.Collapsed;
        Cursor = busy ? Cursors.Wait : Cursors.Arrow;
    }
    private void SetStatus(string text, bool error, bool success = false) { StatusText.Text = text; StatusText.Foreground = error ? new SolidColorBrush(Color.FromRgb(185, 28, 28)) : success ? new SolidColorBrush(Color.FromRgb(21, 128, 61)) : new SolidColorBrush(Color.FromRgb(71, 85, 105)); }
    private void ResetDropZone() { DropZone.BorderBrush = DefaultDropBorder; DropZone.Background = Brushes.White; }
    private static bool TryGetDroppedWorkbook(IDataObject data, out string? path) { path = null; if (!data.GetDataPresent(DataFormats.FileDrop) || data.GetData(DataFormats.FileDrop) is not string[] files || files.Length != 1 || !File.Exists(files[0]) || !string.Equals(Path.GetExtension(files[0]), ".xlsx", StringComparison.OrdinalIgnoreCase)) return false; path = files[0]; return true; }
}
