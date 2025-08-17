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
                CurrentVersion = Assembly.GetExecutingAssembly().GetName().Version?.ToString(3) ?? "0.0.1";
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
                        fonts.Add(new FontItem { Name = fontName, IsJapanese = IsJapaneseFont(fontName) });
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
                    OrderedFontNames.Add(name);
                }
            }
            else
            {
                foreach (var name in Effect.SelectedFonts.OrderBy(x => OrderedFontNames.IndexOf(x)))
                {
                    OrderedFontNames.Add(name);
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
            isUpdatingUI = false;
            EndEdit?.Invoke(this, EventArgs.Empty);
            UpdateEffectLists();
        }

        void SelectJapaneseButton_Click(object sender, RoutedEventArgs e)
        {
            isUpdatingUI = true;
            BeginEdit?.Invoke(this, EventArgs.Empty);
            foreach (FontItem font in filteredFontsView)
            {
                font.IsSelected = font.IsJapanese;
                if (font.IsJapanese && !OrderedFontNames.Contains(font.Name)) OrderedFontNames.Add(font.Name);
                else if (!font.IsJapanese) OrderedFontNames.Remove(font.Name);
            }
            isUpdatingUI = false;
            EndEdit?.Invoke(this, EventArgs.Empty);
            UpdateEffectLists();
        }

        void SelectEnglishButton_Click(object sender, RoutedEventArgs e)
        {
            isUpdatingUI = true;
            BeginEdit?.Invoke(this, EventArgs.Empty);
            foreach (FontItem font in filteredFontsView)
            {
                font.IsSelected = !font.IsJapanese;
                if (!font.IsJapanese && !OrderedFontNames.Contains(font.Name)) OrderedFontNames.Add(font.Name);
                else if (font.IsJapanese) OrderedFontNames.Remove(font.Name);
            }
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
                        OrderedFontNames.Move(removeIndex, targetIndex);
                        UpdateEffectLists();
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

        public event PropertyChangedEventHandler? PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
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
}
