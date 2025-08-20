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
using System.Threading;

namespace FontShuffle
{
    public partial class FontShuffleControl : UserControl, IPropertyEditorControl, INotifyPropertyChanged
    {
        public static new readonly DependencyProperty EffectProperty =
            DependencyProperty.Register(nameof(Effect), typeof(FontShuffleEffect), typeof(FontShuffleControl),
                new PropertyMetadata(null, OnEffectChanged));

        public new FontShuffleEffect? Effect
        {
            get => (FontShuffleEffect?)GetValue(EffectProperty);
            set => SetValue(EffectProperty, value);
        }

        public string SearchText { get; set; } = "";
        public bool IsSearchEmpty => string.IsNullOrEmpty(SearchText);

        public ObservableCollection<string> OrderedFontNames { get; set; } = new();
        public ObservableCollection<FontGroup> DisplayFontGroups { get; set; } = new();

        private bool _isUpdatingUI = false;
        private ObservableCollection<FontItem> _allFonts = new();
        private ICollectionView _filteredFontsView = null!;
        private CancellationTokenSource? _indexCreationCancellationTokenSource;
        private readonly SemaphoreSlim _fontLoadingSemaphore = new(1, 1);
        private volatile bool _isLoaded = false;
        private volatile bool _isDisposed = false;

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
                _filteredFontsView = CollectionViewSource.GetDefaultView(_allFonts);
                FontListView.ItemsSource = _filteredFontsView;
                FontGroupComboBox.ItemsSource = DisplayFontGroups;
                SetupContextMenu();
                UpdateFontGroups();
                LogManager.WriteLog("FontShuffleControlが初期化されました");
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "FontShuffleControl初期化");
            }
        }

        private void UpdateFontGroups()
        {
            try
            {
                if (_isDisposed) return;

                DisplayFontGroups.Clear();
                DisplayFontGroups.Add(new FontGroup { Name = "すべて", Type = FontGroupType.All });
                DisplayFontGroups.Add(new FontGroup { Name = "日本語フォント", Type = FontGroupType.Japanese });
                DisplayFontGroups.Add(new FontGroup { Name = "英数字フォント", Type = FontGroupType.English });

                foreach (var group in FontShuffleEffect.GlobalFontGroups)
                {
                    DisplayFontGroups.Add(group);
                }

                if (FontGroupComboBox.SelectedIndex < 0)
                {
                    FontGroupComboBox.SelectedIndex = 0;
                }

                LogManager.WriteLog("フォントグループが更新されました");
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "フォントグループ更新");
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

                            foreach (var font in _allFonts)
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
                    if (!FontShuffleEffect.GlobalHiddenFonts.Contains(selectedFont.Name))
                    {
                        FontShuffleEffect.GlobalHiddenFonts.Add(selectedFont.Name);
                        FontShuffleEffect.SaveGlobalSettings();
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

        private async void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_isLoaded || _isDisposed) return;
                _isLoaded = true;

                if (_allFonts.Count == 0)
                {
                    _ = Task.Run(async () =>
                    {
                        try
                        {
                            await LoadSystemFontsAsync();
                        }
                        catch (Exception ex)
                        {
                            LogManager.WriteException(ex, "フォント読み込み非同期処理");
                        }
                    });
                }

                _ = Task.Run(async () =>
                {
                    try
                    {
                        await CheckForUpdatesAsync();
                    }
                    catch (Exception ex)
                    {
                        LogManager.WriteException(ex, "更新チェック非同期処理");
                    }
                });
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "UserControl_Loaded");
            }
        }

        private async Task CheckForUpdatesAsync()
        {
            if (_isDisposed) return;

            try
            {
                await Dispatcher.BeginInvoke(new Action(() =>
                {
                    IsUpToDate = true;
                    IsUpdateAvailable = false;
                    NewUpdateAfterIgnore = false;
                }));

                CurrentVersion = Assembly.GetExecutingAssembly().GetName().Version?.ToString(3) ?? "1.0.0";
                string url = "https://api.github.com/repos/routersys/YMM4-ShuffleFont/releases/latest";

                using var client = new HttpClient();
                client.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("YMM4-ShuffleFont", CurrentVersion));
                client.Timeout = TimeSpan.FromSeconds(10);

                var response = await client.GetStringAsync(url);
                var release = JsonSerializer.Deserialize<GitHubRelease>(response);

                if (release != null && !string.IsNullOrEmpty(release.tag_name))
                {
                    FontShuffleEffect? currentEffect = null;
                    await Dispatcher.BeginInvoke(new Action(() =>
                    {
                        if (_isDisposed) return;
                        currentEffect = Effect;
                    }));

                    await Dispatcher.BeginInvoke(new Action(() =>
                    {
                        if (_isDisposed) return;

                        LatestVersion = release.tag_name.TrimStart('v');
                        ReleasePageUrl = release.html_url;

                        if (Version.TryParse(CurrentVersion, out var currentVer) && Version.TryParse(LatestVersion, out var latestVer))
                        {
                            if (latestVer > currentVer)
                            {
                                if (currentEffect != null && currentEffect.IgnoredVersion == LatestVersion)
                                {
                                    IsUpToDate = true;
                                }
                                else
                                {
                                    if (currentEffect != null && !string.IsNullOrEmpty(currentEffect.IgnoredVersion))
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
                    }));
                }
                LogManager.WriteLog($"更新チェック完了（現在: {CurrentVersion}, 最新: {LatestVersion}）");
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "更新チェック");
                if (!_isDisposed)
                {
                    await Dispatcher.BeginInvoke(new Action(() =>
                    {
                        IsUpToDate = true;
                        IsUpdateAvailable = false;
                        NewUpdateAfterIgnore = false;
                    }));
                }
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
                settingsWindow.SetFonts(_allFonts.Where(f => f.IsSelected), Effect?.FontCustomSettings ?? new Dictionary<string, FontCustomSettings>());
                settingsWindow.SetFontGroups(FontShuffleEffect.GlobalFontGroups);
                settingsWindow.SetHiddenFonts(_allFonts.Where(f => f.IsHidden).ToList());

                if (settingsWindow.ShowDialog() == true)
                {
                    if (Effect != null)
                    {
                        BeginEdit?.Invoke(this, EventArgs.Empty);

                        Effect.FontCustomSettings = settingsWindow.GetUpdatedSettings();

                        var updatedGroups = settingsWindow.GetUpdatedGroups();
                        FontShuffleEffect.GlobalFontGroups.Clear();
                        updatedGroups.ForEach(g => FontShuffleEffect.GlobalFontGroups.Add(g));

                        var updatedHiddenFonts = settingsWindow.GetUpdatedHiddenFonts();
                        FontShuffleEffect.GlobalHiddenFonts.Clear();
                        updatedHiddenFonts.Select(f => f.Name).ToList().ForEach(name => FontShuffleEffect.GlobalHiddenFonts.Add(name));

                        FontShuffleEffect.SaveGlobalSettings();

                        UpdateFontGroups();

                        foreach (var font in _allFonts)
                        {
                            font.IsHidden = FontShuffleEffect.GlobalHiddenFonts.Contains(font.Name);
                            font.UpdateCustomSettingsCache(Effect.FontCustomSettings);
                        }

                        EndEdit?.Invoke(this, EventArgs.Empty);

                        foreach (var font in _allFonts)
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
                FontIndexer.ClearIndex();
                _ = Task.Run(CreateFontIndexAsync);
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "インデックス作成ボタン");
            }
        }

        private async Task CreateFontIndexAsync()
        {
            _indexCreationCancellationTokenSource?.Cancel();
            _indexCreationCancellationTokenSource?.Dispose();
            _indexCreationCancellationTokenSource = new CancellationTokenSource();

            FontIndexProgressDialog? progressDialog = null;

            try
            {
                LogManager.WriteLog("フォントインデックス作成を開始します");

                await Dispatcher.BeginInvoke(new Action(() =>
                {
                    if (_isDisposed) return;

                    progressDialog = new FontIndexProgressDialog
                    {
                        Owner = Window.GetWindow(this)
                    };
                    progressDialog.Show();
                }));

                if (progressDialog == null) return;

                var progress = new Progress<FontIndexProgress>(progressInfo =>
                {
                    try
                    {
                        if (!_isDisposed && progressDialog != null)
                        {
                            progressDialog.UpdateProgress(progressInfo);
                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.WriteException(ex, "進行状況更新");
                    }
                });

                var createIndexTask = FontIndexer.CreateIndexAsync(progress, progressDialog.CancellationToken);
                await createIndexTask;

                if (progressDialog.WasCanceled || _indexCreationCancellationTokenSource.Token.IsCancellationRequested)
                {
                    LogManager.WriteLog("フォントインデックス作成がキャンセルされました");
                }
                else
                {
                    LogManager.WriteLog("フォントインデックス作成が完了しました");

                    if (!_isDisposed)
                    {
                        await Dispatcher.BeginInvoke(new Action(async () => {
                            try
                            {
                                if (!_isDisposed)
                                {
                                    await LoadSystemFontsAsync();
                                }
                            }
                            catch (Exception ex)
                            {
                                LogManager.WriteException(ex, "フォント再読み込み");
                            }
                        }));
                    }
                }
            }
            catch (OperationCanceledException)
            {
                LogManager.WriteLog("フォントインデックス作成がキャンセルされました");
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "フォントインデックス作成");

                if (!_isDisposed)
                {
                    await Dispatcher.BeginInvoke(new Action(() => {
                        MessageBox.Show("フォントインデックスの作成中にエラーが発生しました。\n詳細はログファイルを確認してください。",
                            "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                    }));
                }
            }
            finally
            {
                if (progressDialog?.IsVisible == true)
                {
                    await Dispatcher.BeginInvoke(new Action(() =>
                    {
                        progressDialog?.Close();
                    }));
                }

                _indexCreationCancellationTokenSource?.Dispose();
                _indexCreationCancellationTokenSource = null;
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
                _isUpdatingUI = true;
                BeginEdit?.Invoke(this, EventArgs.Empty);

                foreach (var font in _allFonts)
                {
                    if (font.IsHidden) continue;

                    bool shouldSelect = group.Type switch
                    {
                        FontGroupType.All => false,
                        FontGroupType.Japanese => font.IsJapanese,
                        FontGroupType.English => !font.IsJapanese,
                        FontGroupType.Custom => group.FontNames?.Contains(font.Name) == true,
                        _ => false
                    };

                    if (shouldSelect && !font.IsSelected)
                    {
                        font.IsSelected = true;
                        if (!string.IsNullOrWhiteSpace(font.Name) && !OrderedFontNames.Contains(font.Name))
                            OrderedFontNames.Add(font.Name);
                    }
                }

                _isUpdatingUI = false;
                EndEdit?.Invoke(this, EventArgs.Empty);
                UpdateEffectLists();
                LogManager.WriteLog($"フォントグループ「{group.Name}」を適用しました");
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "フォントグループフィルタ適用");
                _isUpdatingUI = false;
            }
        }

        private void CreateGroupButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (Effect == null) return;

                var selectedFonts = _allFonts.Where(f => f.IsSelected && !f.IsHidden).Select(f => f.Name).ToList();
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

                    BeginEdit?.Invoke(this, EventArgs.Empty);
                    FontShuffleEffect.GlobalFontGroups.Add(newGroup);
                    FontShuffleEffect.SaveGlobalSettings();
                    EndEdit?.Invoke(this, EventArgs.Empty);

                    UpdateFontGroups();
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
                if (Effect == null) return;

                var groupManager = new FontGroupManagerWindow();
                groupManager.Owner = Window.GetWindow(this);
                groupManager.SetGroups(FontShuffleEffect.GlobalFontGroups);

                if (groupManager.ShowDialog() == true)
                {
                    BeginEdit?.Invoke(this, EventArgs.Empty);
                    var updatedGroups = groupManager.GetUpdatedGroups();
                    FontShuffleEffect.GlobalFontGroups.Clear();
                    updatedGroups.ForEach(g => FontShuffleEffect.GlobalFontGroups.Add(g));
                    FontShuffleEffect.SaveGlobalSettings();
                    EndEdit?.Invoke(this, EventArgs.Empty);

                    UpdateFontGroups();
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

        private async Task LoadSystemFontsAsync()
        {
            if (_isDisposed) return;

            try
            {
                await _fontLoadingSemaphore.WaitAsync();

                if (_isDisposed) return;

                FontShuffleEffect? effectReference = null;
                List<string> hiddenFonts = new List<string>();
                Dictionary<string, FontCustomSettings>? customSettings = null;

                await Dispatcher.BeginInvoke(new Action(() =>
                {
                    if (_isDisposed) return;

                    LoadingText.Visibility = Visibility.Visible;
                    FontListView.Visibility = Visibility.Hidden;

                    effectReference = Effect;
                    hiddenFonts = FontShuffleEffect.GlobalHiddenFonts.ToList();
                    customSettings = effectReference?.FontCustomSettings;
                }));

                var fonts = await Task.Run(() =>
                {
                    try
                    {
                        if (_isDisposed) return new List<FontItem>();

                        if (FontIndexer.ShouldRecreateIndex())
                        {
                            LogManager.WriteLog("インデックスの再作成が必要です");
                            return LoadBasicFonts(hiddenFonts);
                        }

                        var fontIndex = FontIndexer.LoadIndex();
                        if (fontIndex == null)
                        {
                            LogManager.WriteLog("インデックスが見つからないため、基本的なフォント読み込みを実行");
                            return LoadBasicFonts(hiddenFonts);
                        }

                        LogManager.WriteLog($"インデックスからフォントを読み込み中（{fontIndex.AllFonts.Count}フォント）");
                        var fontList = new List<FontItem>();
                        var hiddenSet = new HashSet<string>(hiddenFonts, StringComparer.OrdinalIgnoreCase);

                        foreach (var fontName in fontIndex.AllFonts)
                        {
                            if (_isDisposed) break;
                            if (string.IsNullOrEmpty(fontName)) continue;

                            var font = new FontItem(this)
                            {
                                Name = fontName,
                                IsJapanese = fontIndex.FontSupportMap.ContainsKey(fontName) ? fontIndex.FontSupportMap[fontName] : false,
                                IsHidden = hiddenSet.Contains(fontName)
                            };
                            fontList.Add(font);
                        }

                        return fontList;
                    }
                    catch (Exception ex)
                    {
                        LogManager.WriteException(ex, "フォント読み込みタスク");
                        return LoadBasicFonts(hiddenFonts);
                    }
                });

                if (_isDisposed) return;

                await Dispatcher.BeginInvoke(new Action(() =>
                {
                    if (_isDisposed) return;

                    foreach (var font in _allFonts)
                    {
                        font.PropertyChanged -= FontItem_PropertyChanged;
                    }
                    _allFonts.Clear();
                    OrderedFontNames.Clear();

                    foreach (var font in fonts)
                    {
                        font.PropertyChanged += FontItem_PropertyChanged;
                        font.UpdateCustomSettingsCache(customSettings);
                        _allFonts.Add(font);
                    }

                    UpdateFromEffect();
                    LoadingText.Visibility = Visibility.Collapsed;
                    FontListView.Visibility = Visibility.Visible;
                    UpdateCounts();
                }));

                LogManager.WriteLog($"フォント読み込み完了（{_allFonts.Count}フォント）");
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "LoadSystemFontsAsync");
                if (!_isDisposed)
                {
                    await Dispatcher.BeginInvoke(new Action(() =>
                    {
                        LoadingText.Visibility = Visibility.Collapsed;
                        FontListView.Visibility = Visibility.Visible;
                    }));
                }
            }
            finally
            {
                _fontLoadingSemaphore.Release();
            }
        }

        private List<FontItem> LoadBasicFonts(List<string> hiddenFonts)
        {
            try
            {
                var fontSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var fonts = new List<FontItem>();
                var hiddenSet = new HashSet<string>(hiddenFonts, StringComparer.OrdinalIgnoreCase);

                var systemFonts = System.Windows.Media.Fonts.SystemFontFamilies.ToArray();

                foreach (var fontFamily in systemFonts)
                {
                    if (_isDisposed) break;

                    try
                    {
                        string fontName = FontHelper.GetFontFamilyName(fontFamily);
                        if (!string.IsNullOrEmpty(fontName) && fontSet.Add(fontName))
                        {
                            fonts.Add(new FontItem(this)
                            {
                                Name = fontName,
                                IsJapanese = FontHelper.IsJapaneseFontUnified(fontFamily, fontName),
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

        private void FontItem_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (_isUpdatingUI || _isDisposed) return;

            try
            {
                if (sender is FontItem font)
                {
                    if (e.PropertyName == nameof(FontItem.IsSelected))
                    {
                        if (font.IsSelected)
                        {
                            if (!string.IsNullOrWhiteSpace(font.Name) && !OrderedFontNames.Contains(font.Name))
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

        private void UpdateEffectLists()
        {
            if (Effect == null || _isUpdatingUI || _isDisposed) return;

            try
            {
                BeginEdit?.Invoke(this, EventArgs.Empty);
                Effect.SelectedFonts = _allFonts.Where(f => f.IsSelected && !f.IsHidden && !string.IsNullOrWhiteSpace(f.Name)).Select(f => f.Name).ToList();
                Effect.FavoriteFonts = _allFonts.Where(f => f.IsFavorite && !f.IsHidden && !string.IsNullOrWhiteSpace(f.Name)).Select(f => f.Name).ToList();
                Effect.OrderedFonts = OrderedFontNames.Where(name => !string.IsNullOrWhiteSpace(name) && !FontShuffleEffect.GlobalHiddenFonts.Contains(name)).ToList();
                EndEdit?.Invoke(this, EventArgs.Empty);
                UpdateCounts();
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "エフェクトリスト更新");
            }
        }

        private void ApplyFilter()
        {
            if (_isDisposed) return;

            try
            {
                _filteredFontsView.Filter = obj =>
                {
                    if (obj is not FontItem font) return false;

                    bool searchMatch = string.IsNullOrEmpty(SearchText) || font.Name.Contains(SearchText, StringComparison.OrdinalIgnoreCase);
                    bool selectedMatch = ShowSelectedCheckBox.IsChecked != true || font.IsSelected;
                    bool favoriteMatch = ShowFavoritesCheckBox.IsChecked != true || font.IsFavorite;
                    bool hiddenMatch = !font.IsHidden;

                    return searchMatch && selectedMatch && favoriteMatch && hiddenMatch;
                };
                _filteredFontsView.Refresh();
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "フィルタ適用");
            }
        }

        private void UpdateCounts()
        {
            if (_isDisposed) return;

            try
            {
                var visibleFonts = _allFonts.Where(f => !f.IsHidden);
                TotalCountText.Text = visibleFonts.Count().ToString();
                SelectedCountText.Text = visibleFonts.Count(f => f.IsSelected).ToString();
                FavoriteCountText.Text = visibleFonts.Count(f => f.IsFavorite).ToString();
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "カウント更新");
            }
        }

        private static void OnEffectChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            try
            {
                if (d is FontShuffleControl control && !control._isDisposed)
                {
                    control.UpdateFromEffect();
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "エフェクト変更処理");
            }
        }

        private void UpdateFromEffect()
        {
            if (Effect == null || _allFonts.Count == 0 || _isDisposed) return;

            try
            {
                _isUpdatingUI = true;

                UpdateFontGroups();

                var selectedSet = new HashSet<string>(Effect.SelectedFonts ?? new List<string>(), StringComparer.OrdinalIgnoreCase);
                var favoriteSet = new HashSet<string>(Effect.FavoriteFonts ?? new List<string>(), StringComparer.OrdinalIgnoreCase);
                var hiddenSet = new HashSet<string>(FontShuffleEffect.GlobalHiddenFonts ?? new List<string>(), StringComparer.OrdinalIgnoreCase);

                foreach (var font in _allFonts)
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
                        if (!string.IsNullOrWhiteSpace(name) && selectedSet.Contains(name) && !hiddenSet.Contains(name))
                        {
                            OrderedFontNames.Add(name);
                        }
                    }
                }
                else
                {
                    foreach (var name in Effect.SelectedFonts ?? new List<string>())
                    {
                        if (!string.IsNullOrWhiteSpace(name) && !OrderedFontNames.Contains(name) && !hiddenSet.Contains(name))
                        {
                            OrderedFontNames.Add(name);
                        }
                    }
                }

                _isUpdatingUI = false;
                UpdateCounts();
                ApplyFilter();
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "エフェクトからの更新");
                _isUpdatingUI = false;
            }
        }

        private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                SearchText = SearchTextBox.Text ?? "";
                OnPropertyChanged(nameof(IsSearchEmpty));
                ApplyFilter();
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "検索テキスト変更");
            }
        }

        private void FilterCheckBox_Changed(object sender, RoutedEventArgs e)
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

        private void SelectAllButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _isUpdatingUI = true;
                BeginEdit?.Invoke(this, EventArgs.Empty);
                foreach (FontItem font in _filteredFontsView)
                {
                    if (!font.IsHidden)
                    {
                        font.IsSelected = true;
                        if (!string.IsNullOrWhiteSpace(font.Name) && !OrderedFontNames.Contains(font.Name))
                            OrderedFontNames.Add(font.Name);
                    }
                }
                _isUpdatingUI = false;
                EndEdit?.Invoke(this, EventArgs.Empty);
                UpdateEffectLists();
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "全選択");
                _isUpdatingUI = false;
            }
        }

        private void DeselectAllButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _isUpdatingUI = true;
                BeginEdit?.Invoke(this, EventArgs.Empty);
                foreach (FontItem font in _filteredFontsView)
                {
                    font.IsSelected = false;
                    OrderedFontNames.Remove(font.Name);
                }

                FontGroupComboBox.SelectedIndex = 0;

                _isUpdatingUI = false;
                EndEdit?.Invoke(this, EventArgs.Empty);
                UpdateEffectLists();
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "全解除");
                _isUpdatingUI = false;
            }
        }

        private void FavoriteButton_Click(object sender, RoutedEventArgs e)
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

        private void UserControl_Unloaded(object sender, RoutedEventArgs e)
        {
            try
            {
                _isDisposed = true;

                _indexCreationCancellationTokenSource?.Cancel();
                _indexCreationCancellationTokenSource?.Dispose();
                _indexCreationCancellationTokenSource = null;
                _fontLoadingSemaphore?.Dispose();

                foreach (var font in _allFonts)
                {
                    font.PropertyChanged -= FontItem_PropertyChanged;
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "UserControl_Unloaded");
            }
        }

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private Point? dragStartPoint;
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
                                try
                                {
                                    OrderListBox.Items.Refresh();
                                }
                                catch (Exception ex)
                                {
                                    LogManager.WriteException(ex, "リストボックス更新");
                                }
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