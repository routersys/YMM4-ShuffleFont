using System;
using System.Collections.ObjectModel;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Linq;
using System.Collections.Generic;
using System.Windows.Data;

namespace FontShuffle
{
    public partial class FontSettingsWindow : Window
    {
        public ObservableCollection<FontSettingsItem> FontItems { get; set; } = new();
        public ObservableCollection<FontGroup> FontGroups { get; set; } = new();
        public ObservableCollection<FontItem> HiddenFonts { get; set; } = new();

        private ObservableCollection<FontSettingsItem> _workingFontItems = new();
        private ObservableCollection<FontGroup> _workingFontGroups = new();
        private ObservableCollection<FontItem> _workingHiddenFonts = new();

        private Dictionary<string, FontCustomSettings> _originalSettings = new();
        private List<FontGroup> _originalGroups = new();
        private List<FontItem> _originalHiddenFonts = new();

        private FontSettingsItem? _currentSelectedFont;
        private FontGroup? _currentSelectedGroup;
        private bool _isUpdatingUI = false;

        public FontSettingsWindow()
        {
            try
            {
                InitializeComponent();
                DataContext = this;
                FontListBox.ItemsSource = _workingFontItems;
                GroupListBox.ItemsSource = _workingFontGroups;
                HiddenFontListBox.ItemsSource = _workingHiddenFonts;

                Loaded += (s, e) => UpdateHiddenFontsVisibility();

                LogManager.WriteLog("フォント設定ウィンドウが初期化されました");
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "FontSettingsWindow初期化");
            }
        }

        public void SetFonts(IEnumerable<FontItem> fonts, Dictionary<string, FontCustomSettings> customSettings)
        {
            try
            {
                _originalSettings = customSettings?.ToDictionary(kvp => kvp.Key, kvp => kvp.Value.Clone()) ?? new Dictionary<string, FontCustomSettings>();

                _workingFontItems.Clear();
                FontItems.Clear();

                foreach (var font in fonts.OrderBy(f => f.Name))
                {
                    var settingsItem = new FontSettingsItem { Name = font.Name };

                    if (customSettings?.ContainsKey(font.Name) == true)
                    {
                        settingsItem.CustomSettings = customSettings[font.Name].Clone();
                    }
                    else if (FontShuffleEffect.GlobalFontCustomSettings.ContainsKey(font.Name))
                    {
                        settingsItem.CustomSettings = FontShuffleEffect.GlobalFontCustomSettings[font.Name].Clone();
                    }
                    else
                    {
                        settingsItem.CustomSettings = new FontCustomSettings();
                    }

                    _workingFontItems.Add(settingsItem);
                    FontItems.Add(settingsItem);
                }
                LogManager.WriteLog($"フォント設定に{_workingFontItems.Count}個のフォントを設定しました");
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "フォント設定");
            }
        }

        public void SetSingleFont(string fontName, Dictionary<string, FontCustomSettings> customSettings)
        {
            try
            {
                _originalSettings = customSettings?.ToDictionary(kvp => kvp.Key, kvp => kvp.Value.Clone()) ?? new Dictionary<string, FontCustomSettings>();

                _workingFontItems.Clear();
                FontItems.Clear();

                var settingsItem = new FontSettingsItem { Name = fontName };

                if (customSettings?.ContainsKey(fontName) == true)
                {
                    settingsItem.CustomSettings = customSettings[fontName].Clone();
                }
                else if (FontShuffleEffect.GlobalFontCustomSettings.ContainsKey(fontName))
                {
                    settingsItem.CustomSettings = FontShuffleEffect.GlobalFontCustomSettings[fontName].Clone();
                }
                else
                {
                    settingsItem.CustomSettings = new FontCustomSettings();
                }

                _workingFontItems.Add(settingsItem);
                FontItems.Add(settingsItem);
                FontListBox.SelectedIndex = 0;
                LogManager.WriteLog($"単一フォント「{fontName}」の設定を開きました");
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "単一フォント設定");
            }
        }

        public void SetFontGroups(List<FontGroup> fontGroups)
        {
            try
            {
                _originalGroups = fontGroups?.Select(g => new FontGroup
                {
                    Name = g.Name,
                    Type = g.Type,
                    FontNames = new List<string>(g.FontNames ?? new List<string>())
                }).ToList() ?? new List<FontGroup>();

                _workingFontGroups.Clear();
                FontGroups.Clear();

                foreach (var group in fontGroups ?? new List<FontGroup>())
                {
                    var groupCopy = new FontGroup
                    {
                        Name = group.Name,
                        Type = group.Type,
                        FontNames = new List<string>(group.FontNames ?? new List<string>())
                    };
                    _workingFontGroups.Add(groupCopy);
                    FontGroups.Add(groupCopy);
                }
                LogManager.WriteLog($"フォントグループ設定に{_workingFontGroups.Count}個のグループを設定しました");
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "フォントグループ設定");
            }
        }

        public void SetHiddenFonts(List<FontItem> hiddenFonts)
        {
            try
            {
                _originalHiddenFonts = hiddenFonts?.Select(f => new FontItem { Name = f.Name, IsHidden = true }).ToList() ?? new List<FontItem>();

                _workingHiddenFonts.Clear();
                HiddenFonts.Clear();

                if (hiddenFonts != null)
                {
                    var orderedFonts = hiddenFonts.OrderBy(f => f.Name).ToList();
                    foreach (var font in orderedFonts)
                    {
                        var fontCopy = new FontItem { Name = font.Name, IsHidden = true };
                        _workingHiddenFonts.Add(fontCopy);
                        HiddenFonts.Add(fontCopy);
                    }
                }

                UpdateHiddenFontsVisibility();
                LogManager.WriteLog($"非表示フォント設定に{_workingHiddenFonts.Count}個のフォントを設定しました");
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "非表示フォント設定");
            }
        }

        public Dictionary<string, FontCustomSettings> GetUpdatedSettings()
        {
            try
            {
                var result = new Dictionary<string, FontCustomSettings>();
                foreach (var item in _workingFontItems)
                {
                    if (item.CustomSettings.UseCustomSettings)
                    {
                        result[item.Name] = item.CustomSettings.Clone();
                    }
                }
                LogManager.WriteLog($"更新されたフォント設定を取得（{result.Count}個）");
                return result;
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "更新設定取得");
                return new Dictionary<string, FontCustomSettings>();
            }
        }

        public List<FontGroup> GetUpdatedGroups()
        {
            try
            {
                var result = _workingFontGroups.Select(g => new FontGroup
                {
                    Name = g.Name,
                    Type = g.Type,
                    FontNames = new List<string>(g.FontNames ?? new List<string>())
                }).ToList();
                LogManager.WriteLog($"更新されたグループを取得（{result.Count}個）");
                return result;
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "更新グループ取得");
                return new List<FontGroup>();
            }
        }

        public List<FontItem> GetUpdatedHiddenFonts()
        {
            try
            {
                var result = _workingHiddenFonts.Select(f => new FontItem { Name = f.Name, IsHidden = true }).ToList();
                LogManager.WriteLog($"更新された非表示フォントを取得（{result.Count}個）");
                return result;
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "更新非表示フォント取得");
                return new List<FontItem>();
            }
        }

        #region フォント設定タブ

        private void FontListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (FontListBox.SelectedItem is FontSettingsItem selectedFont)
                {
                    _currentSelectedFont = selectedFont;
                    DisplayFontSettings(selectedFont);
                    SettingsScrollViewer.Visibility = Visibility.Visible;
                    NoSelectionText.Visibility = Visibility.Collapsed;
                }
                else
                {
                    _currentSelectedFont = null;
                    SettingsScrollViewer.Visibility = Visibility.Collapsed;
                    NoSelectionText.Visibility = Visibility.Visible;
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "フォント選択変更");
            }
        }

        private void DisplayFontSettings(FontSettingsItem fontItem)
        {
            try
            {
                _isUpdatingUI = true;

                CurrentFontNameText.Text = fontItem.Name;
                try
                {
                    CurrentFontNameText.FontFamily = new FontFamily(fontItem.Name);
                }
                catch
                {
                    CurrentFontNameText.FontFamily = new FontFamily("Yu Gothic UI");
                }

                UseCustomSettingsCheckBox.IsChecked = fontItem.CustomSettings.UseCustomSettings;
                IsProjectSpecificCheckBox.IsChecked = fontItem.CustomSettings.IsProjectSpecific;
                UpdateCustomSettingsVisibility();

                FontSizeSlider.Value = Math.Max(FontSizeSlider.Minimum, Math.Min(FontSizeSlider.Maximum, fontItem.CustomSettings.FontSize));
                FontSizeTextBox.Text = fontItem.CustomSettings.FontSize.ToString("F0");
                UseDynamicSizeCheckBox.IsChecked = fontItem.CustomSettings.UseDynamicSize;

                ColorPreviewBrush.Color = fontItem.CustomSettings.TextColor;
                ColorCodeTextBox.Text = ColorHelper.ColorToHex(fontItem.CustomSettings.TextColor);

                BoldCheckBox.IsChecked = fontItem.CustomSettings.Bold;
                ItalicCheckBox.IsChecked = fontItem.CustomSettings.Italic;

                UpdateFontSizeControls();
                UpdatePreview();

                _isUpdatingUI = false;
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "フォント設定表示");
                _isUpdatingUI = false;
            }
        }

        private void UpdateFontSizeControls()
        {
            if (_currentSelectedFont == null) return;

            try
            {
                bool isDynamic = UseDynamicSizeCheckBox.IsChecked == true;
                FontSizeSlider.IsEnabled = !isDynamic;
                FontSizeTextBox.IsEnabled = !isDynamic;
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "フォントサイズコントロール更新");
            }
        }

        private void UseDynamicSizeCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_currentSelectedFont != null && !_isUpdatingUI)
                {
                    _currentSelectedFont.CustomSettings.UseDynamicSize = UseDynamicSizeCheckBox.IsChecked == true;
                    UpdateFontSizeControls();
                    UpdatePreview();
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "動的サイズチェック変更");
            }
        }

        private void UseCustomSettingsCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_currentSelectedFont != null && !_isUpdatingUI)
                {
                    _currentSelectedFont.CustomSettings.UseCustomSettings = UseCustomSettingsCheckBox.IsChecked == true;
                    UpdateCustomSettingsVisibility();
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "カスタム設定チェック変更");
            }
        }

        private void IsProjectSpecificCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_currentSelectedFont != null && !_isUpdatingUI)
                {
                    _currentSelectedFont.CustomSettings.IsProjectSpecific = IsProjectSpecificCheckBox.IsChecked == true;
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "プロジェクト固有設定チェック変更");
            }
        }


        private void UpdateCustomSettingsVisibility()
        {
            try
            {
                CustomSettingsPanel.Visibility = UseCustomSettingsCheckBox.IsChecked == true
                    ? Visibility.Visible : Visibility.Collapsed;
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "カスタム設定表示更新");
            }
        }

        private void FontSizeSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            try
            {
                if (_currentSelectedFont != null && !_isUpdatingUI)
                {
                    _currentSelectedFont.CustomSettings.FontSize = FontSizeSlider.Value;
                    FontSizeTextBox.Text = FontSizeSlider.Value.ToString("F0");
                    UpdatePreview();
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "フォントサイズスライダー変更");
            }
        }

        private void FontSizeTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                if (_currentSelectedFont != null && !_isUpdatingUI && double.TryParse(FontSizeTextBox.Text, out double value))
                {
                    value = Math.Max(10, Math.Min(1200, value));

                    if (value >= FontSizeSlider.Minimum && value <= FontSizeSlider.Maximum)
                    {
                        FontSizeSlider.Value = value;
                    }

                    _currentSelectedFont.CustomSettings.FontSize = value;
                    UpdatePreview();
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "フォントサイズテキスト変更");
            }
        }

        private void SelectColorButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_currentSelectedFont == null) return;

                var colorDialog = new ColorPickerDialog();
                colorDialog.Owner = this;
                colorDialog.SelectedColor = _currentSelectedFont.CustomSettings.TextColor;

                if (colorDialog.ShowDialog() == true)
                {
                    _currentSelectedFont.CustomSettings.TextColor = colorDialog.SelectedColor;
                    ColorPreviewBrush.Color = colorDialog.SelectedColor;
                    ColorCodeTextBox.Text = ColorHelper.ColorToHex(colorDialog.SelectedColor);
                    UpdatePreview();
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "色選択ダイアログ");
            }
        }

        private void ColorCodeTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                if (_currentSelectedFont != null && !_isUpdatingUI)
                {
                    if (ColorHelper.TryParseHexColor(ColorCodeTextBox.Text, out Color color))
                    {
                        _currentSelectedFont.CustomSettings.TextColor = color;
                        ColorPreviewBrush.Color = color;
                        UpdatePreview();
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "カラーコードテキスト変更");
            }
        }

        private void StyleCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_currentSelectedFont != null && !_isUpdatingUI)
                {
                    _currentSelectedFont.CustomSettings.Bold = BoldCheckBox.IsChecked == true;
                    _currentSelectedFont.CustomSettings.Italic = ItalicCheckBox.IsChecked == true;
                    UpdatePreview();
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "スタイルチェック変更");
            }
        }

        private void UpdatePreview()
        {
            if (_currentSelectedFont == null) return;

            try
            {
                try
                {
                    PreviewTextBlock.FontFamily = new FontFamily(_currentSelectedFont.Name);
                }
                catch
                {
                    PreviewTextBlock.FontFamily = new FontFamily("Yu Gothic UI");
                }

                PreviewTextBlock.Foreground = new SolidColorBrush(_currentSelectedFont.CustomSettings.TextColor);
                PreviewTextBlock.FontWeight = _currentSelectedFont.CustomSettings.Bold ? FontWeights.Bold : FontWeights.Normal;
                PreviewTextBlock.FontStyle = _currentSelectedFont.CustomSettings.Italic ? FontStyles.Italic : FontStyles.Normal;

                var displaySize = Math.Min(_currentSelectedFont.CustomSettings.FontSize * 0.6, 36);
                PreviewTextBlock.FontSize = Math.Max(8, displaySize);
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "プレビュー更新");
            }
        }

        private void ResetButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_currentSelectedFont == null) return;

                if (_originalSettings.ContainsKey(_currentSelectedFont.Name))
                {
                    _currentSelectedFont.CustomSettings = _originalSettings[_currentSelectedFont.Name].Clone();
                }
                else
                {
                    _currentSelectedFont.CustomSettings = new FontCustomSettings();
                }
                DisplayFontSettings(_currentSelectedFont);
                LogManager.WriteLog($"フォント「{_currentSelectedFont.Name}」の設定をリセットしました");
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "設定リセット");
            }
        }

        private void DefaultButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_currentSelectedFont == null) return;

                _currentSelectedFont.CustomSettings = new FontCustomSettings();
                DisplayFontSettings(_currentSelectedFont);
                LogManager.WriteLog($"フォント「{_currentSelectedFont.Name}」の設定をデフォルトに戻しました");
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "デフォルト設定");
            }
        }

        #endregion

        #region グループ管理タブ

        private void GroupListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (GroupListBox.SelectedItem is FontGroup selectedGroup)
                {
                    _currentSelectedGroup = selectedGroup;
                    DisplayGroupSettings(selectedGroup);
                    GroupEditPanel.Visibility = Visibility.Visible;
                    NoGroupSelectionText.Visibility = Visibility.Collapsed;
                }
                else
                {
                    _currentSelectedGroup = null;
                    GroupEditPanel.Visibility = Visibility.Collapsed;
                    NoGroupSelectionText.Visibility = Visibility.Visible;
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "グループ選択変更");
            }
        }

        private void DisplayGroupSettings(FontGroup group)
        {
            try
            {
                _isUpdatingUI = true;
                GroupNameTextBox.Text = group.Name ?? "";
                GroupFontListBox.ItemsSource = group.FontNames ?? new List<string>();
                _isUpdatingUI = false;
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "グループ設定表示");
                _isUpdatingUI = false;
            }
        }

        private void CreateNewGroupButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var inputDialog = new InputDialog();
                inputDialog.Owner = this;
                inputDialog.Title = "新しいグループ";
                inputDialog.Message = "グループ名を入力してください:";
                inputDialog.Value = "新しいグループ";

                if (inputDialog.ShowDialog() == true && !string.IsNullOrWhiteSpace(inputDialog.Value))
                {
                    var newGroup = new FontGroup
                    {
                        Name = inputDialog.Value,
                        Type = FontGroupType.Custom,
                        FontNames = new List<string>()
                    };
                    _workingFontGroups.Add(newGroup);
                    FontGroups.Add(newGroup);
                    GroupListBox.SelectedItem = newGroup;
                    LogManager.WriteLog($"新しいグループ「{inputDialog.Value}」を作成しました");
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "新しいグループ作成");
            }
        }

        private void DeleteGroupButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var groupToDelete = _currentSelectedGroup;
                if (groupToDelete != null)
                {
                    var result = MessageBox.Show($"グループ「{groupToDelete.Name}」を削除しますか？",
                        "グループ削除", MessageBoxButton.YesNo, MessageBoxImage.Question);

                    if (result == MessageBoxResult.Yes)
                    {
                        _workingFontGroups.Remove(groupToDelete);
                        FontGroups.Remove(groupToDelete);
                        LogManager.WriteLog($"グループ「{groupToDelete.Name}」を削除しました");
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "グループ削除");
            }
        }

        private void GroupNameTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                if (_currentSelectedGroup != null && !_isUpdatingUI)
                {
                    _currentSelectedGroup.Name = GroupNameTextBox.Text ?? "";
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "グループ名変更");
            }
        }

        private void SaveGroupButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_currentSelectedGroup != null)
                {
                    MessageBox.Show("グループ設定を保存しました。", "保存完了", MessageBoxButton.OK, MessageBoxImage.Information);
                    LogManager.WriteLog($"グループ「{_currentSelectedGroup.Name}」の設定を保存しました");
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "グループ保存");
            }
        }

        private void RemoveFontFromGroupButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (sender is Button button && button.DataContext is string fontName && _currentSelectedGroup != null)
                {
                    _currentSelectedGroup.FontNames?.Remove(fontName);
                    if (GroupFontListBox.ItemsSource != null)
                    {
                        CollectionViewSource.GetDefaultView(GroupFontListBox.ItemsSource).Refresh();
                    }
                    LogManager.WriteLog($"フォント「{fontName}」をグループから削除しました");
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "グループからフォント削除");
            }
        }

        #endregion

        #region 非表示フォントタブ

        private void RestoreFontButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (sender is Button button && button.DataContext is FontItem font)
                {
                    _workingHiddenFonts.Remove(font);
                    HiddenFonts.Remove(font);
                    UpdateHiddenFontsVisibility();
                    LogManager.WriteLog($"フォント「{font.Name}」を表示に戻しました");
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "フォント表示復元");
            }
        }

        private void RestoreAllFontsButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_workingHiddenFonts.Count > 0)
                {
                    var result = MessageBox.Show($"すべての非表示フォント（{_workingHiddenFonts.Count}個）を表示に戻しますか？",
                        "すべて表示に戻す", MessageBoxButton.YesNo, MessageBoxImage.Question);

                    if (result == MessageBoxResult.Yes)
                    {
                        var count = _workingHiddenFonts.Count;
                        _workingHiddenFonts.Clear();
                        HiddenFonts.Clear();
                        UpdateHiddenFontsVisibility();
                        LogManager.WriteLog($"{count}個のフォントをすべて表示に戻しました");
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "すべてのフォント表示復元");
            }
        }

        private void UpdateHiddenFontsVisibility()
        {
            try
            {
                if (NoHiddenFontsText != null)
                {
                    NoHiddenFontsText.Visibility = _workingHiddenFonts.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "非表示フォント表示更新");
            }
        }

        #endregion

        #region 共通メソッド

        private void OKButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DialogResult = true;
                Close();
                LogManager.WriteLog("フォント設定ウィンドウをOKで閉じました");
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "OK ボタン");
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                RestoreOriginalData();
                DialogResult = false;
                Close();
                LogManager.WriteLog("フォント設定ウィンドウをキャンセルで閉じました");
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "キャンセル ボタン");
            }
        }

        private void RestoreOriginalData()
        {
            try
            {
                _workingFontItems.Clear();
                FontItems.Clear();

                foreach (var originalFont in _originalSettings.Keys.Union(_workingFontItems.Select(f => f.Name)))
                {
                    var settingsItem = new FontSettingsItem { Name = originalFont };
                    if (_originalSettings.ContainsKey(originalFont))
                    {
                        settingsItem.CustomSettings = _originalSettings[originalFont].Clone();
                    }
                    else
                    {
                        settingsItem.CustomSettings = new FontCustomSettings();
                    }
                    _workingFontItems.Add(settingsItem);
                    FontItems.Add(settingsItem);
                }

                _workingFontGroups.Clear();
                FontGroups.Clear();
                foreach (var group in _originalGroups)
                {
                    var groupCopy = new FontGroup
                    {
                        Name = group.Name,
                        Type = group.Type,
                        FontNames = new List<string>(group.FontNames ?? new List<string>())
                    };
                    _workingFontGroups.Add(groupCopy);
                    FontGroups.Add(groupCopy);
                }

                _workingHiddenFonts.Clear();
                HiddenFonts.Clear();
                foreach (var font in _originalHiddenFonts)
                {
                    var fontCopy = new FontItem { Name = font.Name, IsHidden = true };
                    _workingHiddenFonts.Add(fontCopy);
                    HiddenFonts.Add(fontCopy);
                }

                LogManager.WriteLog("設定データを元の状態に復元しました");
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "元データ復元");
            }
        }
        #endregion
    }
}