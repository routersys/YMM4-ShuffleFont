using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using YukkuriMovieMaker.Commons;

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
            InitializeComponent();
            filteredFontsView = CollectionViewSource.GetDefaultView(allFonts);
            FontListView.ItemsSource = filteredFontsView;
            FontGroupComboBox.ItemsSource = FontGroups;
            SetupContextMenu();
            InitializeFontGroups();
        }

        private void InitializeFontGroups()
        {
            FontGroups.Add(new FontGroup { Name = "すべて", Type = FontGroupType.All });
            FontGroups.Add(new FontGroup { Name = "日本語フォント", Type = FontGroupType.Japanese });
            FontGroups.Add(new FontGroup { Name = "英数字フォント", Type = FontGroupType.English });
            FontGroupComboBox.SelectedIndex = 0;
        }

        private void SetupContextMenu()
        {
            var contextMenu = new ContextMenu();
            var settingsMenuItem = new MenuItem { Header = "フォント設定" };
            settingsMenuItem.Click += FontContextMenu_Settings_Click;
            contextMenu.Items.Add(settingsMenuItem);

            FontListView.ContextMenu = contextMenu;
        }

        private void FontContextMenu_Settings_Click(object sender, RoutedEventArgs e)
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

                        EndEdit?.Invoke(this, EventArgs.Empty);

                        selectedFont.OnPropertyChanged(nameof(FontItem.Name));
                        ApplyFilter();
                    }
                }
            }
        }

        void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            if (allFonts.Count == 0)
            {
                LoadSystemFonts();
            }
            CheckForUpdatesAsync();
        }

        async void CheckForUpdatesAsync()
        {
            IsUpToDate = true;
            IsUpdateAvailable = false;
            NewUpdateAfterIgnore = false;

            try
            {
                CurrentVersion = Assembly.GetExecutingAssembly().GetName().Version?.ToString(3) ?? "0.0.2";
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
            }
            catch (Exception)
            {
                IsUpToDate = true;
                IsUpdateAvailable = false;
                NewUpdateAfterIgnore = false;
            }
        }

        private void DownloadButton_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(ReleasePageUrl))
            {
                Process.Start(new ProcessStartInfo(ReleasePageUrl) { UseShellExecute = true });
            }
        }

        private void DownloadIcon_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DownloadButton_Click(sender, e);
        }

        private void IgnoreButton_Click(object sender, RoutedEventArgs e)
        {
            if (Effect != null)
            {
                BeginEdit?.Invoke(this, EventArgs.Empty);
                Effect.IgnoredVersion = LatestVersion;
                EndEdit?.Invoke(this, EventArgs.Empty);

                IsUpdateAvailable = false;
                NewUpdateAfterIgnore = false;
                IsUpToDate = true;
            }
        }

        private void FontSettingsButton_Click(object sender, RoutedEventArgs e)
        {
            var settingsWindow = new FontSettingsWindow();
            settingsWindow.Owner = Window.GetWindow(this);
            settingsWindow.SetFonts(allFonts.Where(f => f.IsSelected), Effect?.FontCustomSettings ?? new Dictionary<string, FontCustomSettings>());

            if (settingsWindow.ShowDialog() == true)
            {
                if (Effect != null)
                {
                    BeginEdit?.Invoke(this, EventArgs.Empty);
                    Effect.FontCustomSettings = settingsWindow.GetUpdatedSettings();
                    EndEdit?.Invoke(this, EventArgs.Empty);

                    foreach (var font in allFonts)
                    {
                        font.OnPropertyChanged(nameof(FontItem.Name));
                    }
                    ApplyFilter();
                }
            }
        }

        private void FontGroupComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (FontGroupComboBox.SelectedItem is FontGroup selectedGroup)
            {
                ApplyGroupFilter(selectedGroup);
            }
        }

        private void ApplyGroupFilter(FontGroup group)
        {
            isUpdatingUI = true;
            BeginEdit?.Invoke(this, EventArgs.Empty);

            foreach (var font in allFonts)
            {
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
        }

        private void CreateGroupButton_Click(object sender, RoutedEventArgs e)
        {
            var selectedFonts = allFonts.Where(f => f.IsSelected).Select(f => f.Name).ToList();
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
            }
        }

        private void ManageGroupsButton_Click(object sender, RoutedEventArgs e)
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
            }
        }

        private void ShowOrderNumbersCheckBox_Changed(object sender, RoutedEventArgs e)
        {
        }

        private class GitHubRelease
        {
            [JsonPropertyName("tag_name")]
            public string tag_name { get; set; } = "";
            [JsonPropertyName("html_url")]
            public string html_url { get; set; } = "";
        }

        void LoadSystemFonts()
        {
            LoadingText.Visibility = Visibility.Visible;
            FontListView.Visibility = Visibility.Hidden;

            Task.Run(() =>
            {
                var fontSet = new HashSet<string>();
                var fonts = new List<FontItem>();

                foreach (var fontFamily in Fonts.SystemFontFamilies)
                {
                    string fontName = GetFontFamilyName(fontFamily);
                    if (!string.IsNullOrEmpty(fontName) && fontSet.Add(fontName))
                    {
                        fonts.Add(new FontItem(this) { Name = fontName, IsJapanese = IsJapaneseFont(fontName) });
                    }
                }

                fonts.Sort((a, b) => string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase));

                return fonts;
            }).ContinueWith(task =>
            {
                foreach (var font in task.Result)
                {
                    font.PropertyChanged += FontItem_PropertyChanged;
                    allFonts.Add(font);
                }
                UpdateFromEffect();
                LoadingText.Visibility = Visibility.Collapsed;
                FontListView.Visibility = Visibility.Visible;
                UpdateCounts();
            }, TaskScheduler.FromCurrentSynchronizationContext());
        }

        void FontItem_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (isUpdatingUI) return;

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

        void UpdateEffectLists()
        {
            if (Effect == null || isUpdatingUI) return;

            BeginEdit?.Invoke(this, EventArgs.Empty);
            Effect.SelectedFonts = allFonts.Where(f => f.IsSelected).Select(f => f.Name).ToList();
            Effect.FavoriteFonts = allFonts.Where(f => f.IsFavorite).Select(f => f.Name).ToList();
            Effect.OrderedFonts = new List<string>(OrderedFontNames);
            EndEdit?.Invoke(this, EventArgs.Empty);
            UpdateCounts();
        }

        void ApplyFilter()
        {
            filteredFontsView.Filter = obj =>
            {
                if (obj is not FontItem font) return false;

                bool searchMatch = string.IsNullOrEmpty(SearchText) || font.Name.Contains(SearchText, StringComparison.OrdinalIgnoreCase);
                bool selectedMatch = ShowSelectedCheckBox.IsChecked != true || font.IsSelected;
                bool favoriteMatch = ShowFavoritesCheckBox.IsChecked != true || font.IsFavorite;

                return searchMatch && selectedMatch && favoriteMatch;
            };
            filteredFontsView.Refresh();
        }

        void UpdateCounts()
        {
            TotalCountText.Text = allFonts.Count.ToString();
            SelectedCountText.Text = allFonts.Count(f => f.IsSelected).ToString();
            FavoriteCountText.Text = allFonts.Count(f => f.IsFavorite).ToString();
        }

        static void OnEffectChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is FontShuffleControl control)
            {
                control.UpdateFromEffect();
            }
        }

        void UpdateFromEffect()
        {
            if (Effect == null || allFonts.Count == 0) return;
            isUpdatingUI = true;

            var selectedSet = new HashSet<string>(Effect.SelectedFonts);
            var favoriteSet = new HashSet<string>(Effect.FavoriteFonts);

            foreach (var font in allFonts)
            {
                font.IsSelected = selectedSet.Contains(font.Name);
                font.IsFavorite = favoriteSet.Contains(font.Name);
            }

            OrderedFontNames.Clear();
            if (Effect.OrderedFonts != null && Effect.OrderedFonts.Count > 0)
            {
                foreach (var name in Effect.OrderedFonts)
                {
                    if (selectedSet.Contains(name))
                    {
                        OrderedFontNames.Add(name);
                    }
                }
            }
            else
            {
                foreach (var name in Effect.SelectedFonts)
                {
                    if (!OrderedFontNames.Contains(name))
                    {
                        OrderedFontNames.Add(name);
                    }
                }
            }

            isUpdatingUI = false;
            UpdateCounts();
            ApplyFilter();
        }

        void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            SearchText = SearchTextBox.Text;
            OnPropertyChanged(nameof(IsSearchEmpty));
            ApplyFilter();
        }

        void FilterCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            ApplyFilter();
        }

        void SelectAllButton_Click(object sender, RoutedEventArgs e)
        {
            isUpdatingUI = true;
            BeginEdit?.Invoke(this, EventArgs.Empty);
            foreach (FontItem font in filteredFontsView)
            {
                font.IsSelected = true;
                if (!OrderedFontNames.Contains(font.Name)) OrderedFontNames.Add(font.Name);
            }
            isUpdatingUI = false;
            EndEdit?.Invoke(this, EventArgs.Empty);
            UpdateEffectLists();
        }

        void DeselectAllButton_Click(object sender, RoutedEventArgs e)
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

        void FavoriteButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.DataContext is FontItem font)
            {
                font.IsFavorite = !font.IsFavorite;
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
            catch
            {
                return "Unknown Font";
            }
        }

        static bool IsJapaneseFont(string fontName)
        {
            var japaneseKeywords = new[] { "Gothic", "Mincho", "游", "Yu", "MS", "メイリオ", "Meiryo", "UD", "游ゴシック", "游明朝", "ヒラギノ", "小塚", "源ノ角", "Noto" };
            return japaneseKeywords.Any(keyword => fontName.Contains(keyword, StringComparison.OrdinalIgnoreCase));
        }

        Point? dragStartPoint;
        private void OrderListBox_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            dragStartPoint = e.GetPosition(null);
        }

        private void OrderListBox_MouseMove(object sender, MouseEventArgs e)
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

        private void OrderListBox_Drop(object sender, DragEventArgs e)
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

    public class InputDialog : Window
    {
        public string Value { get; set; } = "";
        public string Message { get; set; } = "";

        private TextBox textBox;

        public InputDialog()
        {
            Width = 400;
            Height = 200;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ResizeMode = ResizeMode.NoResize;

            var grid = new Grid();
            grid.Margin = new Thickness(12);
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var messageLabel = new TextBlock();
            messageLabel.SetBinding(TextBlock.TextProperty, new Binding("Message") { Source = this });
            messageLabel.Margin = new Thickness(0, 0, 0, 12);
            Grid.SetRow(messageLabel, 0);
            grid.Children.Add(messageLabel);

            textBox = new TextBox();
            textBox.SetBinding(TextBox.TextProperty, new Binding("Value") { Source = this });
            textBox.Margin = new Thickness(0, 0, 0, 12);
            Grid.SetRow(textBox, 1);
            grid.Children.Add(textBox);

            var buttonPanel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            var okButton = new Button { Content = "OK", Width = 70, Height = 24, Margin = new Thickness(0, 0, 8, 0), IsDefault = true };
            var cancelButton = new Button { Content = "キャンセル", Width = 70, Height = 24, IsCancel = true };

            okButton.Click += (s, e) => { Value = textBox.Text; DialogResult = true; Close(); };
            cancelButton.Click += (s, e) => { DialogResult = false; Close(); };

            buttonPanel.Children.Add(okButton);
            buttonPanel.Children.Add(cancelButton);

            Grid.SetRow(buttonPanel, 3);
            grid.Children.Add(buttonPanel);

            Content = grid;

            Loaded += (s, e) => textBox.Focus();
        }
    }

    public class FontGroup : INotifyPropertyChanged
    {
        public string Name { get; set; } = "";
        public FontGroupType Type { get; set; }
        public List<string> FontNames { get; set; } = new();

        public event PropertyChangedEventHandler? PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public enum FontGroupType
    {
        All,
        Japanese,
        English,
        Custom
    }

    public class FontGroupManagerWindow : Window
    {
        private ObservableCollection<FontGroup> groups = new();
        private ListBox groupListBox;

        public FontGroupManagerWindow()
        {
            Title = "グループ管理";
            Width = 400;
            Height = 300;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;

            var grid = new Grid();
            grid.Margin = new Thickness(12);
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            groupListBox = new ListBox { ItemsSource = groups };
            groupListBox.DisplayMemberPath = "Name";
            Grid.SetRow(groupListBox, 0);
            grid.Children.Add(groupListBox);

            var buttonPanel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 12, 0, 0) };
            var deleteButton = new Button { Content = "削除", Width = 70, Height = 24, Margin = new Thickness(0, 0, 8, 0) };
            var okButton = new Button { Content = "OK", Width = 70, Height = 24, Margin = new Thickness(0, 0, 8, 0) };
            var cancelButton = new Button { Content = "キャンセル", Width = 70, Height = 24 };

            deleteButton.Click += (s, e) => {
                if (groupListBox.SelectedItem is FontGroup selected)
                {
                    groups.Remove(selected);
                }
            };
            okButton.Click += (s, e) => { DialogResult = true; Close(); };
            cancelButton.Click += (s, e) => { DialogResult = false; Close(); };

            buttonPanel.Children.Add(deleteButton);
            buttonPanel.Children.Add(okButton);
            buttonPanel.Children.Add(cancelButton);

            Grid.SetRow(buttonPanel, 1);
            grid.Children.Add(buttonPanel);

            Content = grid;
        }

        public void SetGroups(List<FontGroup> fontGroups)
        {
            groups.Clear();
            foreach (var group in fontGroups)
            {
                groups.Add(group);
            }
        }

        public List<FontGroup> GetUpdatedGroups()
        {
            return groups.ToList();
        }
    }

    internal class FontShuffleControlAttribute : PropertyEditorAttribute2
    {
        public FontShuffleControlAttribute()
        {
            PropertyEditorSize = PropertyEditorSize.FullWidth;
        }

        public override FrameworkElement Create() => new FontShuffleControl();

        public override void SetBindings(FrameworkElement control, ItemProperty[] itemProperties)
        {
            if (control is FontShuffleControl selector && itemProperties.Length > 0)
            {
                selector.Effect = itemProperties[0].PropertyOwner as FontShuffleEffect;
            }
        }

        public override void ClearBindings(FrameworkElement control)
        {
            if (control is FontShuffleControl selector)
            {
                BindingOperations.ClearBinding(selector, FontShuffleControl.EffectProperty);
            }
        }
    }

    public class FontItem : INotifyPropertyChanged
    {
        private readonly FontShuffleControl? _control;

        public FontItem(FontShuffleControl? control = null)
        {
            _control = control;
        }

        public string Name { get; set; } = "";
        public bool IsJapanese { get; set; }

        bool isSelected;
        public bool IsSelected
        {
            get => isSelected;
            set { if (isSelected != value) { isSelected = value; OnPropertyChanged(nameof(IsSelected)); } }
        }

        bool isFavorite;
        public bool IsFavorite
        {
            get => isFavorite;
            set { if (isFavorite != value) { isFavorite = value; OnPropertyChanged(nameof(IsFavorite)); } }
        }

        public bool HasCustomSettings
        {
            get
            {
                if (_control?.Effect?.FontCustomSettings == null) return false;
                return _control.Effect.FontCustomSettings.ContainsKey(Name) &&
                       _control.Effect.FontCustomSettings[Name].UseCustomSettings;
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        public virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            if (propertyName == nameof(Name))
            {
                OnPropertyChanged(nameof(HasCustomSettings));
            }
        }
    }

    public class BoolToColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return (value is bool v && v) ? (SolidColorBrush)new BrushConverter().ConvertFrom("#FFD700")! : (SolidColorBrush)new BrushConverter().ConvertFrom("#FF808080")!;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class HasCustomSettingsConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is FontItem font)
            {
                return font.HasCustomSettings ? Visibility.Visible : Visibility.Collapsed;
            }
            if (value is string fontName)
            {
                return Visibility.Collapsed;
            }
            return Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class ColorToBrushConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is System.Windows.Media.Color color)
            {
                return new SolidColorBrush(color);
            }
            return new SolidColorBrush(Colors.White);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is SolidColorBrush brush)
            {
                return brush.Color;
            }
            return Colors.White;
        }
    }

    public class OrderNumberConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is ListBoxItem item)
            {
                var listBox = ItemsControl.ItemsControlFromItemContainer(item) as ListBox;
                if (listBox != null)
                {
                    int index = -1;
                    for (int i = 0; i < listBox.Items.Count; i++)
                    {
                        if (listBox.Items[i] == item.DataContext)
                        {
                            index = i;
                            break;
                        }
                    }
                    return index >= 0 ? $"{index + 1}." : "";
                }
            }
            return "";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}