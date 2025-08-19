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

        private Dictionary<string, FontCustomSettings> originalSettings = new();
        private List<FontGroup> originalGroups = new();
        private List<FontItem> originalHiddenFonts = new();

        private FontSettingsItem? currentSelectedFont;
        private FontGroup? currentSelectedGroup;
        private bool isUpdatingUI = false;

        public FontSettingsWindow()
        {
            try
            {
                InitializeComponent();
                DataContext = this;
                FontListBox.ItemsSource = FontItems;
                GroupListBox.ItemsSource = FontGroups;
                HiddenFontListBox.ItemsSource = HiddenFonts;

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
                originalSettings = customSettings.ToDictionary(kvp => kvp.Key, kvp => kvp.Value.Clone());

                FontItems.Clear();
                foreach (var font in fonts.OrderBy(f => f.Name))
                {
                    var settingsItem = new FontSettingsItem { Name = font.Name };

                    if (customSettings.ContainsKey(font.Name))
                    {
                        settingsItem.CustomSettings = customSettings[font.Name].Clone();
                    }
                    else
                    {
                        settingsItem.CustomSettings = new FontCustomSettings();
                    }

                    FontItems.Add(settingsItem);
                }
                LogManager.WriteLog($"フォント設定に{FontItems.Count}個のフォントを設定しました");
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
                originalSettings = customSettings.ToDictionary(kvp => kvp.Key, kvp => kvp.Value.Clone());

                FontItems.Clear();
                var settingsItem = new FontSettingsItem { Name = fontName };

                if (customSettings.ContainsKey(fontName))
                {
                    settingsItem.CustomSettings = customSettings[fontName].Clone();
                }
                else
                {
                    settingsItem.CustomSettings = new FontCustomSettings();
                }

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
                originalGroups = fontGroups.Select(g => new FontGroup
                {
                    Name = g.Name,
                    Type = g.Type,
                    FontNames = new List<string>(g.FontNames)
                }).ToList();

                FontGroups.Clear();
                foreach (var group in fontGroups)
                {
                    FontGroups.Add(new FontGroup
                    {
                        Name = group.Name,
                        Type = group.Type,
                        FontNames = new List<string>(group.FontNames)
                    });
                }
                LogManager.WriteLog($"フォントグループ設定に{FontGroups.Count}個のグループを設定しました");
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
                originalHiddenFonts = hiddenFonts.Select(f => new FontItem { Name = f.Name, IsHidden = true }).ToList();

                HiddenFonts.Clear();
                foreach (var font in hiddenFonts.OrderBy(f => f.Name))
                {
                    HiddenFonts.Add(new FontItem { Name = font.Name, IsHidden = true });
                }
                UpdateHiddenFontsVisibility();
                LogManager.WriteLog($"非表示フォント設定に{HiddenFonts.Count}個のフォントを設定しました");
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
                foreach (var item in FontItems)
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
                var result = FontGroups.Select(g => new FontGroup
                {
                    Name = g.Name,
                    Type = g.Type,
                    FontNames = new List<string>(g.FontNames)
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
                var result = HiddenFonts.Select(f => new FontItem { Name = f.Name, IsHidden = true }).ToList();
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
                    currentSelectedFont = selectedFont;
                    DisplayFontSettings(selectedFont);
                    SettingsScrollViewer.Visibility = Visibility.Visible;
                    NoSelectionText.Visibility = Visibility.Collapsed;
                }
                else
                {
                    currentSelectedFont = null;
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
                isUpdatingUI = true;

                CurrentFontNameText.Text = fontItem.Name;
                CurrentFontNameText.FontFamily = new FontFamily(fontItem.Name);

                UseCustomSettingsCheckBox.IsChecked = fontItem.CustomSettings.UseCustomSettings;
                UpdateCustomSettingsVisibility();

                FontSizeSlider.Value = fontItem.CustomSettings.FontSize;
                FontSizeTextBox.Text = fontItem.CustomSettings.FontSize.ToString("F0");
                UseDynamicSizeCheckBox.IsChecked = fontItem.CustomSettings.UseDynamicSize;

                ColorPreviewBrush.Color = fontItem.CustomSettings.TextColor;
                ColorCodeTextBox.Text = ColorToHex(fontItem.CustomSettings.TextColor);

                BoldCheckBox.IsChecked = fontItem.CustomSettings.Bold;
                ItalicCheckBox.IsChecked = fontItem.CustomSettings.Italic;

                UpdateFontSizeControls();
                UpdatePreview();

                isUpdatingUI = false;
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "フォント設定表示");
                isUpdatingUI = false;
            }
        }

        private void UpdateFontSizeControls()
        {
            if (currentSelectedFont == null) return;

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
                if (currentSelectedFont != null && !isUpdatingUI)
                {
                    currentSelectedFont.CustomSettings.UseDynamicSize = UseDynamicSizeCheckBox.IsChecked == true;
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
                if (currentSelectedFont != null && !isUpdatingUI)
                {
                    currentSelectedFont.CustomSettings.UseCustomSettings = UseCustomSettingsCheckBox.IsChecked == true;
                    UpdateCustomSettingsVisibility();
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "カスタム設定チェック変更");
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
                if (currentSelectedFont != null && !isUpdatingUI)
                {
                    currentSelectedFont.CustomSettings.FontSize = FontSizeSlider.Value;
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
                if (currentSelectedFont != null && !isUpdatingUI && double.TryParse(FontSizeTextBox.Text, out double value))
                {
                    value = Math.Max(10, Math.Min(1200, value));

                    if (value <= FontSizeSlider.Maximum)
                    {
                        FontSizeSlider.Value = value;
                    }

                    currentSelectedFont.CustomSettings.FontSize = value;
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
                if (currentSelectedFont == null) return;

                var colorDialog = new ColorPickerDialog();
                colorDialog.Owner = this;
                colorDialog.SelectedColor = currentSelectedFont.CustomSettings.TextColor;

                if (colorDialog.ShowDialog() == true)
                {
                    currentSelectedFont.CustomSettings.TextColor = colorDialog.SelectedColor;
                    ColorPreviewBrush.Color = colorDialog.SelectedColor;
                    ColorCodeTextBox.Text = ColorToHex(colorDialog.SelectedColor);
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
                if (currentSelectedFont != null && !isUpdatingUI)
                {
                    if (TryParseHexColor(ColorCodeTextBox.Text, out Color color))
                    {
                        currentSelectedFont.CustomSettings.TextColor = color;
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
                if (currentSelectedFont != null && !isUpdatingUI)
                {
                    currentSelectedFont.CustomSettings.Bold = BoldCheckBox.IsChecked == true;
                    currentSelectedFont.CustomSettings.Italic = ItalicCheckBox.IsChecked == true;
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
            if (currentSelectedFont == null) return;

            try
            {
                PreviewTextBlock.FontFamily = new FontFamily(currentSelectedFont.Name);
                PreviewTextBlock.Foreground = new SolidColorBrush(currentSelectedFont.CustomSettings.TextColor);
                PreviewTextBlock.FontWeight = currentSelectedFont.CustomSettings.Bold ? FontWeights.Bold : FontWeights.Normal;
                PreviewTextBlock.FontStyle = currentSelectedFont.CustomSettings.Italic ? FontStyles.Italic : FontStyles.Normal;

                var displaySize = Math.Min(currentSelectedFont.CustomSettings.FontSize * 0.6, 36);
                PreviewTextBlock.FontSize = displaySize;
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
                if (currentSelectedFont == null) return;

                if (originalSettings.ContainsKey(currentSelectedFont.Name))
                {
                    currentSelectedFont.CustomSettings = originalSettings[currentSelectedFont.Name].Clone();
                }
                else
                {
                    currentSelectedFont.CustomSettings = new FontCustomSettings();
                }
                DisplayFontSettings(currentSelectedFont);
                LogManager.WriteLog($"フォント「{currentSelectedFont.Name}」の設定をリセットしました");
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
                if (currentSelectedFont == null) return;

                currentSelectedFont.CustomSettings = new FontCustomSettings();
                DisplayFontSettings(currentSelectedFont);
                LogManager.WriteLog($"フォント「{currentSelectedFont.Name}」の設定をデフォルトに戻しました");
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
                    currentSelectedGroup = selectedGroup;
                    DisplayGroupSettings(selectedGroup);
                    GroupEditPanel.Visibility = Visibility.Visible;
                    NoGroupSelectionText.Visibility = Visibility.Collapsed;
                }
                else
                {
                    currentSelectedGroup = null;
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
                isUpdatingUI = true;
                GroupNameTextBox.Text = group.Name;
                GroupFontListBox.ItemsSource = group.FontNames;
                isUpdatingUI = false;
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "グループ設定表示");
                isUpdatingUI = false;
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
                var groupToDelete = currentSelectedGroup;
                if (groupToDelete != null)
                {
                    var result = MessageBox.Show($"グループ「{groupToDelete.Name}」を削除しますか？",
                        "グループ削除", MessageBoxButton.YesNo, MessageBoxImage.Question);

                    if (result == MessageBoxResult.Yes)
                    {
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
                if (currentSelectedGroup != null && !isUpdatingUI)
                {
                    currentSelectedGroup.Name = GroupNameTextBox.Text;
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
                if (currentSelectedGroup != null)
                {
                    MessageBox.Show("グループ設定を保存しました。", "保存完了", MessageBoxButton.OK, MessageBoxImage.Information);
                    LogManager.WriteLog($"グループ「{currentSelectedGroup.Name}」の設定を保存しました");
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
                if (sender is Button button && button.DataContext is string fontName && currentSelectedGroup != null)
                {
                    currentSelectedGroup.FontNames.Remove(fontName);
                    CollectionViewSource.GetDefaultView(GroupFontListBox.ItemsSource).Refresh();
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
                if (HiddenFonts.Count > 0)
                {
                    var result = MessageBox.Show($"すべての非表示フォント（{HiddenFonts.Count}個）を表示に戻しますか？",
                        "すべて表示に戻す", MessageBoxButton.YesNo, MessageBoxImage.Question);

                    if (result == MessageBoxResult.Yes)
                    {
                        var count = HiddenFonts.Count;
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
                    NoHiddenFontsText.Visibility = HiddenFonts.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "非表示フォント表示更新");
            }
        }

        #endregion

        #region 共通メソッド

        private string ColorToHex(Color color)
        {
            return $"#{color.R:X2}{color.G:X2}{color.B:X2}";
        }

        private bool TryParseHexColor(string hex, out Color color)
        {
            color = Colors.White;

            try
            {
                if (string.IsNullOrEmpty(hex))
                    return false;

                hex = hex.Trim();
                if (hex.StartsWith("#"))
                    hex = hex.Substring(1);

                if (hex.Length != 6)
                    return false;

                if (!Regex.IsMatch(hex, @"^[0-9A-Fa-f]{6}$"))
                    return false;

                byte r = Convert.ToByte(hex.Substring(0, 2), 16);
                byte g = Convert.ToByte(hex.Substring(2, 2), 16);
                byte b = Convert.ToByte(hex.Substring(4, 2), 16);
                color = Color.FromRgb(r, g, b);
                return true;
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "16進カラー解析");
                return false;
            }
        }

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
                DialogResult = false;
                Close();
                LogManager.WriteLog("フォント設定ウィンドウをキャンセルで閉じました");
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "キャンセル ボタン");
            }
        }

        #endregion
    }
}