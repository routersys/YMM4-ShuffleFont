using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using YukkuriMovieMaker.Commons;
using System.Linq;
using System.Threading.Tasks;
using System.Collections.Generic;

namespace FontShuffle
{
    public partial class FontShuffleControl : UserControl, IPropertyEditorControl, INotifyPropertyChanged
    {
        public static readonly DependencyProperty EffectProperty =
            DependencyProperty.Register(nameof(Effect), typeof(FontShuffleEffect), typeof(FontShuffleControl),
                new PropertyMetadata(null, OnEffectChanged));

        public FontShuffleEffect? Effect
        {
            get => (FontShuffleEffect?)GetValue(EffectProperty);
            set => SetValue(EffectProperty, value);
        }

        public string SearchText { get; set; } = "";
        public bool IsSearchEmpty => string.IsNullOrEmpty(SearchText);

        public ObservableCollection<string> OrderedFontNames { get; set; } = new();
        public ObservableCollection<FontGroup> FontGroups { get; set; } = new();

        bool isUpdatingUI = false;
        ObservableCollection<FontItem> allFonts = new();
        ICollectionView filteredFontsView;

        public event EventHandler? BeginEdit;
        public event EventHandler? EndEdit;
        public event PropertyChangedEventHandler? PropertyChanged;

        #region Update Check Properties
        private string _latestVersion = "";
        public string LatestVersion { get => _latestVersion; set { _latestVersion = value; OnPropertyChanged(nameof(LatestVersion)); } }

        private string _currentVersion = "";
        public string CurrentVersion { get => _currentVersion; set { _currentVersion = value; OnPropertyChanged(nameof(CurrentVersion)); } }

        private string _releasePageUrl = "";
        public string ReleasePageUrl { get => _releasePageUrl; set { _releasePageUrl = value; OnPropertyChanged(nameof(ReleasePageUrl)); } }

        private bool _isUpdateAvailable;
        public bool IsUpdateAvailable { get => _isUpdateAvailable; set { _isUpdateAvailable = value; OnPropertyChanged(nameof(IsUpdateAvailable)); } }

        private bool _isUpToDate = true;
        public bool IsUpToDate { get => _isUpToDate; set { _isUpToDate = value; OnPropertyChanged(nameof(IsUpToDate)); } }

        private bool _newUpdateAfterIgnore;
        public bool NewUpdateAfterIgnore { get => _newUpdateAfterIgnore; set { _newUpdateAfterIgnore = value; OnPropertyChanged(nameof(NewUpdateAfterIgnore)); } }
        #endregion

        public FontShuffleControl()
        {
            try
            {
                InitializeComponent();
                filteredFontsView = CollectionViewSource.GetDefaultView(allFonts);
                FontListView.ItemsSource = filteredFontsView;
                FontGroupComboBox.ItemsSource = FontGroups;
                SetupContextMenu();
                InitializeFontGroups();
                LogManager.WriteLog("FontShuffleControlが初期化されました");
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "FontShuffleControl初期化");
            }
        }

        private void InitializeFontGroups()
        {
            try
            {
                FontGroups.Add(new FontGroup { Name = "すべて", Type = FontGroupType.All });
                FontGroups.Add(new FontGroup { Name = "日本語フォント", Type = FontGroupType.Japanese });
                FontGroups.Add(new FontGroup { Name = "英数字フォント", Type = FontGroupType.English });
                FontGroupComboBox.SelectedIndex = 0;
                LogManager.WriteLog("フォントグループが初期化されました");
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "フォントグループ初期化");
            }
        }

        private void SetupContextMenu()
        {
            try
            {
                var contextMenu = new ContextMenu();

                var settingsMenuItem = new MenuItem { Header = "フォント設定" };
                settingsMenuItem.Click += FontContextMenu_Settings_Click;
                contextMenu.Items.Add(settingsMenuItem);

                var hideMenuItem = new MenuItem { Header = "非表示" };
                hideMenuItem.Click += FontContextMenu_Hide_Click;
                contextMenu.Items.Add(hideMenuItem);

                FontListView.ContextMenu = contextMenu;
                LogManager.WriteLog("コンテキストメニューが設定されました");
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "コンテキストメニュー設定");
            }
        }

        private void FontContextMenu_Settings_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (FontListView.SelectedItem is FontItem selectedFont)
                {
                    var settingsWindow = new FontSettingsWindow();
                    settingsWindow.Owner = Window.GetWindow(this);
                    settingsWindow.SetSingleFont(selectedFont.Name, Effect?.FontCustomSettings ?? new Dictionary<string, FontCustomSettings>());

                    if (settingsWindow.ShowDialog() == true)
                    {
                        if (Effect != null)
                        {
                            BeginEdit?.Invoke(this, EventArgs.Empty);
                            var updatedSettings = settingsWindow.GetUpdatedSettings();

                            if (updatedSettings.ContainsKey(selectedFont.Name))
                            {
                                Effect.FontCustomSettings[selectedFont.Name] = updatedSettings[selectedFont.Name];
                            }
                            else if (Effect.FontCustomSettings.ContainsKey(selectedFont.Name))
                            {
                                Effect.FontCustomSettings.Remove(selectedFont.Name);
                            }

                            foreach (var font in allFonts)
                            {
                                font.UpdateCustomSettingsCache(Effect.FontCustomSettings);
                            }

                            EndEdit?.Invoke(this, EventArgs.Empty);

                            selectedFont.OnPropertyChanged(nameof(FontItem.Name));
                            ApplyFilter();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "フォント設定ダイアログ");
            }
        }

        private void FontContextMenu_Hide_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (FontListView.SelectedItem is FontItem selectedFont && Effect != null)
                {
                    BeginEdit?.Invoke(this, EventArgs.Empty);

                    selectedFont.IsHidden = true;
                    if (!Effect.HiddenFonts.Contains(selectedFont.Name))
                    {
                        Effect.HiddenFonts.Add(selectedFont.Name);
                    }

                    selectedFont.IsSelected = false;
                    Effect.SelectedFonts.Remove(selectedFont.Name);
                    Effect.FavoriteFonts.Remove(selectedFont.Name);
                    OrderedFontNames.Remove(selectedFont.Name);

                    EndEdit?.Invoke(this, EventArgs.Empty);
                    ApplyFilter();
                    UpdateCounts();
                    LogManager.WriteLog($"フォント「{selectedFont.Name}」を非表示にしました");
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "フォント非表示処理");
            }
        }

        void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                if (allFonts.Count == 0)
                {
                    LoadSystemFonts();
                }
                CheckForUpdatesAsync();
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "UserControl_Loaded");
            }
        }

        async void CheckForUpdatesAsync()
        {
            IsUpToDate = true;
            IsUpdateAvailable = false;
            NewUpdateAfterIgnore = false;

            try
            {
                CurrentVersion = Assembly.GetExecutingAssembly().GetName().Version?.ToString(3) ?? "0.0.3";
                string url = "https://api.github.com/repos/routersys/YMM4-ShuffleFont/releases/latest";

                using var client = new HttpClient();
                client.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("YMM4-ShuffleFont", CurrentVersion));

                var response = await client.GetStringAsync(url);
                var release = JsonSerializer.Deserialize<GitHubRelease>(response);

                if (release != null && !string.IsNullOrEmpty(release.tag_name))
                {
                    LatestVersion = release.tag_name.TrimStart('v');
                    ReleasePageUrl = release.html_url;

                    var currentVer = new Version(CurrentVersion);
                    var latestVer = new Version(LatestVersion);

                    if (latestVer > currentVer)
                    {
                        if (Effect != null && Effect.IgnoredVersion == LatestVersion)
                        {
                            IsUpToDate = true;
                        }
                        else
                        {
                            if (Effect != null && !string.IsNullOrEmpty(Effect.IgnoredVersion))
                            {
                                NewUpdateAfterIgnore = true;
                            }
                            else
                            {
                                IsUpdateAvailable = true;
                            }
                            IsUpToDate = false;
                        }
                    }
                    else
                    {
                        IsUpToDate = true;
                    }
                }
                LogManager.WriteLog($"更新チェック完了（現在: {CurrentVersion}, 最新: {LatestVersion}）");
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "更新チェック");
                IsUpToDate = true;
                IsUpdateAvailable = false;
                NewUpdateAfterIgnore = false;
            }
        }

        private void DownloadButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!string.IsNullOrEmpty(ReleasePageUrl))
                {
                    Process.Start(new ProcessStartInfo(ReleasePageUrl) { UseShellExecute = true });
                    LogManager.WriteLog("ダウンロードページを開きました");
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "ダウンロードページ表示");
            }
        }

        private void DownloadIcon_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DownloadButton_Click(sender, e);
        }

        private void IgnoreButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (Effect != null)
                {
                    BeginEdit?.Invoke(this, EventArgs.Empty);
                    Effect.IgnoredVersion = LatestVersion;
                    EndEdit?.Invoke(this, EventArgs.Empty);

                    IsUpdateAvailable = false;
                    NewUpdateAfterIgnore = false;
                    IsUpToDate = true;
                    LogManager.WriteLog($"バージョン {LatestVersion} を無視しました");
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "バージョン無視処理");
            }
        }

        private void FontSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var settingsWindow = new FontSettingsWindow();
                settingsWindow.Owner = Window.GetWindow(this);
                settingsWindow.SetFonts(allFonts.Where(f => f.IsSelected), Effect?.FontCustomSettings ?? new Dictionary<string, FontCustomSettings>());
                settingsWindow.SetFontGroups(FontGroups.Where(g => g.Type == FontGroupType.Custom).ToList());
                settingsWindow.SetHiddenFonts(allFonts.Where(f => f.IsHidden).ToList());

                if (settingsWindow.ShowDialog() == true)
                {
                    if (Effect != null)
                    {
                        BeginEdit?.Invoke(this, EventArgs.Empty);
                        Effect.FontCustomSettings = settingsWindow.GetUpdatedSettings();

                        var updatedGroups = settingsWindow.GetUpdatedGroups();
                        var customGroups = FontGroups.Where(g => g.Type == FontGroupType.Custom).ToList();
                        foreach (var group in customGroups)
                        {
                            FontGroups.Remove(group);
                        }
                        foreach (var group in updatedGroups)
                        {
                            FontGroups.Add(group);
                        }

                        var updatedHiddenFonts = settingsWindow.GetUpdatedHiddenFonts();
                        Effect.HiddenFonts = updatedHiddenFonts.Select(f => f.Name).ToList();
                        foreach (var font in allFonts)
                        {
                            font.IsHidden = Effect.HiddenFonts.Contains(font.Name);
                            font.UpdateCustomSettingsCache(Effect.FontCustomSettings);
                        }

                        EndEdit?.Invoke(this, EventArgs.Empty);

                        foreach (var font in allFonts)
                        {
                            font.OnPropertyChanged(nameof(FontItem.Name));
                        }
                        ApplyFilter();
                        UpdateCounts();
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "高度な設定ダイアログ");
            }
        }

        private void IndexButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                CreateFontIndexAsync();
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "インデックス作成ボタン");
            }
        }

        private async void CreateFontIndexAsync()
        {
            var progressDialog = new FontIndexProgressDialog
            {
                Owner = Window.GetWindow(this)
            };

            var progress = new Progress<FontIndexProgress>(progressInfo =>
            {
                progressDialog.UpdateProgress(progressInfo);
            });

            try
            {
                LogManager.WriteLog("フォントインデックス作成を開始します");

                var createIndexTask = FontIndexer.CreateIndexAsync(progress);
                progressDialog.Show();

                await createIndexTask;

                if (progressDialog.WasCanceled)
                {
                    LogManager.WriteLog("フォントインデックス作成がキャンセルされました");
                }
                else
                {
                    LogManager.WriteLog("フォントインデックス作成が完了しました");
                    LoadSystemFonts();
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "フォントインデックス作成");
                MessageBox.Show("フォントインデックスの作成中にエラーが発生しました。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                if (progressDialog.IsVisible)
                {
                    progressDialog.Close();
                }
            }
        }

        private void FontGroupComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (FontGroupComboBox.SelectedItem is FontGroup selectedGroup)
                {
                    ApplyGroupFilter(selectedGroup);
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "フォントグループ選択変更");
            }
        }

        private void ApplyGroupFilter(FontGroup group)
        {
            try
            {
                isUpdatingUI = true;
                BeginEdit?.Invoke(this, EventArgs.Empty);

                foreach (var font in allFonts)
                {
                    if (font.IsHidden) continue;

                    bool shouldSelect = group.Type switch
                    {
                        FontGroupType.All => false,
                        FontGroupType.Japanese => font.IsJapanese,
                        FontGroupType.English => !font.IsJapanese,
                        FontGroupType.Custom => group.FontNames.Contains(font.Name),
                        _ => false
                    };

                    if (shouldSelect && !font.IsSelected)
                    {
                        font.IsSelected = true;
                        if (!OrderedFontNames.Contains(font.Name))
                            OrderedFontNames.Add(font.Name);
                    }
                }

                isUpdatingUI = false;
                EndEdit?.Invoke(this, EventArgs.Empty);
                UpdateEffectLists();
                LogManager.WriteLog($"フォントグループ「{group.Name}」を適用しました");
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "フォントグループフィルタ適用");
                isUpdatingUI = false;
            }
        }

        private void CreateGroupButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var selectedFonts = allFonts.Where(f => f.IsSelected && !f.IsHidden).Select(f => f.Name).ToList();
                if (selectedFonts.Count == 0)
                {
                    MessageBox.Show("グループを作成するにはフォントを選択してください。", "グループ作成", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                var inputDialog = new InputDialog();
                inputDialog.Owner = Window.GetWindow(this);
                inputDialog.Title = "グループ作成";
                inputDialog.Message = "グループ名を入力してください:";
                inputDialog.Value = "新しいグループ";

                if (inputDialog.ShowDialog() == true && !string.IsNullOrWhiteSpace(inputDialog.Value))
                {
                    var newGroup = new FontGroup
                    {
                        Name = inputDialog.Value,
                        Type = FontGroupType.Custom,
                        FontNames = new List<string>(selectedFonts)
                    };
                    FontGroups.Add(newGroup);
                    FontGroupComboBox.SelectedItem = newGroup;
                    LogManager.WriteLog($"フォントグループ「{inputDialog.Value}」を作成しました（{selectedFonts.Count}フォント）");
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "グループ作成");
            }
        }

        private void ManageGroupsButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var groupManager = new FontGroupManagerWindow();
                groupManager.Owner = Window.GetWindow(this);
                groupManager.SetGroups(FontGroups.Where(g => g.Type == FontGroupType.Custom).ToList());

                if (groupManager.ShowDialog() == true)
                {
                    var customGroups = FontGroups.Where(g => g.Type == FontGroupType.Custom).ToList();
                    foreach (var group in customGroups)
                    {
                        FontGroups.Remove(group);
                    }

                    foreach (var group in groupManager.GetUpdatedGroups())
                    {
                        FontGroups.Add(group);
                    }
                    LogManager.WriteLog("フォントグループ管理を更新しました");
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "グループ管理");
            }
        }

        private void ShowOrderNumbersCheckBox_Changed(object sender, RoutedEventArgs e)
        {
        }

        void LoadSystemFonts()
        {
            try
            {
                LoadingText.Visibility = Visibility.Visible;
                FontListView.Visibility = Visibility.Hidden;

                var effectReference = Effect;
                var hiddenFonts = effectReference?.HiddenFonts?.ToList() ?? new List<string>();

                Task.Run(() =>
                {
                    try
                    {
                        var fontIndex = FontIndexer.LoadIndex();
                        if (fontIndex == null)
                        {
                            LogManager.WriteLog("インデックスが見つからないため、基本的なフォント読み込みを実行");
                            return LoadBasicFonts(hiddenFonts);
                        }

                        LogManager.WriteLog($"インデックスからフォントを読み込み中（{fontIndex.AllFonts.Count}フォント）");
                        var fonts = new List<FontItem>();
                        var hiddenSet = new HashSet<string>(hiddenFonts);

                        foreach (var fontName in fontIndex.AllFonts)
                        {
                            var font = new FontItem
                            {
                                Name = fontName,
                                IsJapanese = fontIndex.FontSupportMap.ContainsKey(fontName) ? fontIndex.FontSupportMap[fontName] : false,
                                IsHidden = hiddenSet.Contains(fontName)
                            };
                            fonts.Add(font);
                        }

                        return fonts;
                    }
                    catch (Exception ex)
                    {
                        LogManager.WriteException(ex, "フォント読み込みタスク");
                        return LoadBasicFonts(hiddenFonts);
                    }
                }).ContinueWith(task =>
                {
                    try
                    {
                        Dispatcher.Invoke(() =>
                        {
                            foreach (var font in allFonts)
                            {
                                font.PropertyChanged -= FontItem_PropertyChanged;
                            }
                            allFonts.Clear();
                            OrderedFontNames.Clear();

                            foreach (var font in task.Result)
                            {
                                font.PropertyChanged += FontItem_PropertyChanged;
                                font.UpdateCustomSettingsCache(effectReference?.FontCustomSettings);
                                allFonts.Add(font);
                            }
                            UpdateFromEffect();
                            LoadingText.Visibility = Visibility.Collapsed;
                            FontListView.Visibility = Visibility.Visible;
                            UpdateCounts();
                            LogManager.WriteLog($"フォント読み込み完了（{allFonts.Count}フォント）");
                        });
                    }
                    catch (Exception ex)
                    {
                        LogManager.WriteException(ex, "フォント読み込み完了処理");
                    }
                }, TaskScheduler.FromCurrentSynchronizationContext());
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "LoadSystemFonts");
            }
        }

        private List<FontItem> LoadBasicFonts(List<string> hiddenFonts)
        {
            try
            {
                var fontSet = new HashSet<string>();
                var fonts = new List<FontItem>();
                var hiddenSet = new HashSet<string>(hiddenFonts);

                foreach (var fontFamily in Fonts.SystemFontFamilies)
                {
                    try
                    {
                        string fontName = GetFontFamilyName(fontFamily);
                        if (!string.IsNullOrEmpty(fontName) && fontSet.Add(fontName))
                        {
                            fonts.Add(new FontItem
                            {
                                Name = fontName,
                                IsJapanese = IsJapaneseFontBasic(fontName),
                                IsHidden = hiddenSet.Contains(fontName)
                            });
                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.WriteException(ex, "個別フォント処理（基本）");
                    }
                }

                fonts.Sort((a, b) => string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase));
                return fonts;
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "基本フォント読み込み");
                return new List<FontItem>();
            }
        }

        void FontItem_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (isUpdatingUI) return;

            try
            {
                if (sender is FontItem font)
                {
                    if (e.PropertyName == nameof(FontItem.IsSelected))
                    {
                        if (font.IsSelected)
                        {
                            if (!OrderedFontNames.Contains(font.Name))
                                OrderedFontNames.Add(font.Name);
                        }
                        else
                        {
                            OrderedFontNames.Remove(font.Name);
                        }
                        UpdateEffectLists();
                    }
                    else if (e.PropertyName == nameof(FontItem.IsFavorite))
                    {
                        UpdateEffectLists();
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "フォントアイテムプロパティ変更");
            }
        }

        void UpdateEffectLists()
        {
            if (Effect == null || isUpdatingUI) return;

            try
            {
                BeginEdit?.Invoke(this, EventArgs.Empty);
                Effect.SelectedFonts = allFonts.Where(f => f.IsSelected && !f.IsHidden).Select(f => f.Name).ToList();
                Effect.FavoriteFonts = allFonts.Where(f => f.IsFavorite && !f.IsHidden).Select(f => f.Name).ToList();
                Effect.OrderedFonts = new List<string>(OrderedFontNames.Where(name => !Effect.HiddenFonts.Contains(name)));
                EndEdit?.Invoke(this, EventArgs.Empty);
                UpdateCounts();
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "エフェクトリスト更新");
            }
        }

        void ApplyFilter()
        {
            try
            {
                filteredFontsView.Filter = obj =>
                {
                    if (obj is not FontItem font) return false;

                    bool searchMatch = string.IsNullOrEmpty(SearchText) || font.Name.Contains(SearchText, StringComparison.OrdinalIgnoreCase);
                    bool selectedMatch = ShowSelectedCheckBox.IsChecked != true || font.IsSelected;
                    bool favoriteMatch = ShowFavoritesCheckBox.IsChecked != true || font.IsFavorite;
                    bool hiddenMatch = !font.IsHidden;

                    return searchMatch && selectedMatch && favoriteMatch && hiddenMatch;
                };
                filteredFontsView.Refresh();
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "フィルタ適用");
            }
        }

        void UpdateCounts()
        {
            try
            {
                var visibleFonts = allFonts.Where(f => !f.IsHidden);
                TotalCountText.Text = visibleFonts.Count().ToString();
                SelectedCountText.Text = visibleFonts.Count(f => f.IsSelected).ToString();
                FavoriteCountText.Text = visibleFonts.Count(f => f.IsFavorite).ToString();
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "カウント更新");
            }
        }

        static void OnEffectChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            try
            {
                if (d is FontShuffleControl control)
                {
                    control.UpdateFromEffect();
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "エフェクト変更処理");
            }
        }

        void UpdateFromEffect()
        {
            if (Effect == null || allFonts.Count == 0) return;

            try
            {
                isUpdatingUI = true;

                var selectedSet = new HashSet<string>(Effect.SelectedFonts);
                var favoriteSet = new HashSet<string>(Effect.FavoriteFonts);
                var hiddenSet = new HashSet<string>(Effect.HiddenFonts);

                foreach (var font in allFonts)
                {
                    font.UpdateCustomSettingsCache(Effect.FontCustomSettings);
                    font.IsSelected = selectedSet.Contains(font.Name);
                    font.IsFavorite = favoriteSet.Contains(font.Name);
                    font.IsHidden = hiddenSet.Contains(font.Name);
                }

                OrderedFontNames.Clear();
                if (Effect.OrderedFonts != null && Effect.OrderedFonts.Count > 0)
                {
                    foreach (var name in Effect.OrderedFonts)
                    {
                        if (selectedSet.Contains(name) && !hiddenSet.Contains(name))
                        {
                            OrderedFontNames.Add(name);
                        }
                    }
                }
                else
                {
                    foreach (var name in Effect.SelectedFonts)
                    {
                        if (!OrderedFontNames.Contains(name) && !hiddenSet.Contains(name))
                        {
                            OrderedFontNames.Add(name);
                        }
                    }
                }

                isUpdatingUI = false;
                UpdateCounts();
                ApplyFilter();
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "エフェクトからの更新");
                isUpdatingUI = false;
            }
        }

        void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                SearchText = SearchTextBox.Text;
                OnPropertyChanged(nameof(IsSearchEmpty));
                ApplyFilter();
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "検索テキスト変更");
            }
        }

        void FilterCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            try
            {
                ApplyFilter();
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "フィルタチェックボックス変更");
            }
        }

        void SelectAllButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                isUpdatingUI = true;
                BeginEdit?.Invoke(this, EventArgs.Empty);
                foreach (FontItem font in filteredFontsView)
                {
                    if (!font.IsHidden)
                    {
                        font.IsSelected = true;
                        if (!OrderedFontNames.Contains(font.Name))
                            OrderedFontNames.Add(font.Name);
                    }
                }
                isUpdatingUI = false;
                EndEdit?.Invoke(this, EventArgs.Empty);
                UpdateEffectLists();
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "全選択");
                isUpdatingUI = false;
            }
        }

        void DeselectAllButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                isUpdatingUI = true;
                BeginEdit?.Invoke(this, EventArgs.Empty);
                foreach (FontItem font in filteredFontsView)
                {
                    font.IsSelected = false;
                    OrderedFontNames.Remove(font.Name);
                }

                FontGroupComboBox.SelectedIndex = 0;

                isUpdatingUI = false;
                EndEdit?.Invoke(this, EventArgs.Empty);
                UpdateEffectLists();
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "全解除");
                isUpdatingUI = false;
            }
        }

        void FavoriteButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (sender is Button button && button.DataContext is FontItem font)
                {
                    font.IsFavorite = !font.IsFavorite;
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "お気に入りボタン");
            }
        }

        void UserControl_Unloaded(object sender, RoutedEventArgs e)
        {
        }

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        static string GetFontFamilyName(FontFamily fontFamily)
        {
            try
            {
                if (fontFamily.FamilyNames.Count > 0)
                {
                    return fontFamily.FamilyNames.Values.First();
                }
                return fontFamily.Source ?? "Unknown Font";
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "フォント名取得");
                return "Unknown Font";
            }
        }

        static bool IsJapaneseFontBasic(string fontName)
        {
            try
            {
                var japaneseKeywords = new[] { "Gothic", "Mincho", "游", "Yu", "MS", "メイリオ", "Meiryo", "UD", "游ゴシック", "游明朝", "ヒラギノ", "小塚", "源ノ角", "Noto" };
                return japaneseKeywords.Any(keyword => fontName.Contains(keyword, StringComparison.OrdinalIgnoreCase));
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, $"日本語フォント判定（{fontName}）");
                return false;
            }
        }

        Point? dragStartPoint;
        private void OrderListBox_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            dragStartPoint = e.GetPosition(null);
        }

        private void OrderListBox_MouseMove(object sender, MouseEventArgs e)
        {
            try
            {
                if (e.LeftButton == MouseButtonState.Pressed && dragStartPoint.HasValue)
                {
                    Point position = e.GetPosition(null);
                    if (Math.Abs(position.X - dragStartPoint.Value.X) > SystemParameters.MinimumHorizontalDragDistance ||
                        Math.Abs(position.Y - dragStartPoint.Value.Y) > SystemParameters.MinimumVerticalDragDistance)
                    {
                        var item = FindAncestor<ListBoxItem>((DependencyObject)e.OriginalSource);
                        if (item != null)
                        {
                            DragDrop.DoDragDrop(item, item.DataContext, DragDropEffects.Move);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "ドラッグ開始");
            }
        }

        private void OrderListBox_Drop(object sender, DragEventArgs e)
        {
            try
            {
                if (e.Data.GetData(typeof(string)) is string droppedData)
                {
                    var target = FindAncestor<ListBoxItem>((DependencyObject)e.OriginalSource);
                    if (target != null)
                    {
                        string targetData = (string)target.DataContext;
                        int removeIndex = OrderedFontNames.IndexOf(droppedData);
                        int targetIndex = OrderedFontNames.IndexOf(targetData);

                        if (removeIndex >= 0 && targetIndex >= 0 && removeIndex != targetIndex)
                        {
                            BeginEdit?.Invoke(this, EventArgs.Empty);
                            OrderedFontNames.Move(removeIndex, targetIndex);

                            Dispatcher.BeginInvoke(new Action(() =>
                            {
                                OrderListBox.Items.Refresh();
                            }), System.Windows.Threading.DispatcherPriority.Render);

                            UpdateEffectLists();
                            EndEdit?.Invoke(this, EventArgs.Empty);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "ドロップ処理");
            }
        }

        private void OrderListBox_DragOver(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(typeof(string)))
            {
                e.Effects = DragDropEffects.None;
                e.Handled = true;
            }
        }

        private static T? FindAncestor<T>(DependencyObject? current) where T : DependencyObject
        {
            do
            {
                if (current is T ancestor)
                {
                    return ancestor;
                }
                current = VisualTreeHelper.GetParent(current);
            }
            while (current != null);
            return null;
        }
    }
}