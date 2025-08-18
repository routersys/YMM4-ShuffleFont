using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;

namespace FontShuffle
{
    public partial class FontSettingsWindow : Window
    {
        public ObservableCollection<FontSettingsItem> FontItems { get; set; } = new();
        private Dictionary<string, FontCustomSettings> originalSettings = new();
        private FontSettingsItem? currentSelectedFont;
        private bool isUpdatingUI = false;

        public FontSettingsWindow()
        {
            InitializeComponent();
            DataContext = this;
            FontListBox.ItemsSource = FontItems;
        }

        public void SetFonts(IEnumerable<FontItem> fonts, Dictionary<string, FontCustomSettings> customSettings)
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
        }

        public void SetSingleFont(string fontName, Dictionary<string, FontCustomSettings> customSettings)
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
        }

        public Dictionary<string, FontCustomSettings> GetUpdatedSettings()
        {
            var result = new Dictionary<string, FontCustomSettings>();
            foreach (var item in FontItems)
            {
                if (item.CustomSettings.UseCustomSettings)
                {
                    result[item.Name] = item.CustomSettings.Clone();
                }
            }
            return result;
        }

        private void FontListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
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

        private void DisplayFontSettings(FontSettingsItem fontItem)
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

        private void UpdateFontSizeControls()
        {
            if (currentSelectedFont == null) return;

            bool isDynamic = UseDynamicSizeCheckBox.IsChecked == true;
            FontSizeSlider.IsEnabled = !isDynamic;
            FontSizeTextBox.IsEnabled = !isDynamic;
        }

        private void UseDynamicSizeCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            if (currentSelectedFont != null && !isUpdatingUI)
            {
                currentSelectedFont.CustomSettings.UseDynamicSize = UseDynamicSizeCheckBox.IsChecked == true;
                UpdateFontSizeControls();
                UpdatePreview();
            }
        }

        private void UseCustomSettingsCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            if (currentSelectedFont != null && !isUpdatingUI)
            {
                currentSelectedFont.CustomSettings.UseCustomSettings = UseCustomSettingsCheckBox.IsChecked == true;
                UpdateCustomSettingsVisibility();
            }
        }

        private void UpdateCustomSettingsVisibility()
        {
            CustomSettingsPanel.Visibility = UseCustomSettingsCheckBox.IsChecked == true
                ? Visibility.Visible : Visibility.Collapsed;
        }

        private void FontSizeSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (currentSelectedFont != null && !isUpdatingUI)
            {
                currentSelectedFont.CustomSettings.FontSize = FontSizeSlider.Value;
                FontSizeTextBox.Text = FontSizeSlider.Value.ToString("F0");
                UpdatePreview();
            }
        }

        private void FontSizeTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (currentSelectedFont != null && !isUpdatingUI && double.TryParse(FontSizeTextBox.Text, out double value))
            {
                value = Math.Max(FontSizeSlider.Minimum, Math.Min(FontSizeSlider.Maximum, value));
                FontSizeSlider.Value = value;
                currentSelectedFont.CustomSettings.FontSize = value;
                UpdatePreview();
            }
        }

        private void SelectColorButton_Click(object sender, RoutedEventArgs e)
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

        private void ColorCodeTextBox_TextChanged(object sender, TextChangedEventArgs e)
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

        private void StyleCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            if (currentSelectedFont != null && !isUpdatingUI)
            {
                currentSelectedFont.CustomSettings.Bold = BoldCheckBox.IsChecked == true;
                currentSelectedFont.CustomSettings.Italic = ItalicCheckBox.IsChecked == true;
                UpdatePreview();
            }
        }

        private void UpdatePreview()
        {
            if (currentSelectedFont == null) return;

            PreviewTextBlock.FontFamily = new FontFamily(currentSelectedFont.Name);
            PreviewTextBlock.Foreground = new SolidColorBrush(currentSelectedFont.CustomSettings.TextColor);
            PreviewTextBlock.FontWeight = currentSelectedFont.CustomSettings.Bold ? FontWeights.Bold : FontWeights.Normal;
            PreviewTextBlock.FontStyle = currentSelectedFont.CustomSettings.Italic ? FontStyles.Italic : FontStyles.Normal;
        }

        private string ColorToHex(Color color)
        {
            return $"#{color.R:X2}{color.G:X2}{color.B:X2}";
        }

        private bool TryParseHexColor(string hex, out Color color)
        {
            color = Colors.White;

            if (string.IsNullOrEmpty(hex))
                return false;

            hex = hex.Trim();
            if (hex.StartsWith("#"))
                hex = hex.Substring(1);

            if (hex.Length != 6)
                return false;

            if (!Regex.IsMatch(hex, @"^[0-9A-Fa-f]{6}$"))
                return false;

            try
            {
                byte r = Convert.ToByte(hex.Substring(0, 2), 16);
                byte g = Convert.ToByte(hex.Substring(2, 2), 16);
                byte b = Convert.ToByte(hex.Substring(4, 2), 16);
                color = Color.FromRgb(r, g, b);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void ResetButton_Click(object sender, RoutedEventArgs e)
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
        }

        private void DefaultButton_Click(object sender, RoutedEventArgs e)
        {
            if (currentSelectedFont == null) return;

            currentSelectedFont.CustomSettings = new FontCustomSettings();
            DisplayFontSettings(currentSelectedFont);
        }

        private void OKButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }

    public class FontSettingsItem : INotifyPropertyChanged
    {
        public string Name { get; set; } = "";

        private FontCustomSettings _customSettings = new();
        public FontCustomSettings CustomSettings
        {
            get => _customSettings;
            set
            {
                if (_customSettings != value)
                {
                    if (_customSettings != null)
                        _customSettings.PropertyChanged -= OnCustomSettingsPropertyChanged;

                    _customSettings = value;

                    if (_customSettings != null)
                        _customSettings.PropertyChanged += OnCustomSettingsPropertyChanged;

                    OnPropertyChanged(nameof(CustomSettings));
                }
            }
        }

        private void OnCustomSettingsPropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            OnPropertyChanged($"CustomSettings.{e.PropertyName}");
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class ColorPickerDialog : Window
    {
        public Color SelectedColor { get; set; } = Colors.White;

        private readonly Slider redSlider;
        private readonly Slider greenSlider;
        private readonly Slider blueSlider;
        private readonly Slider alphaSlider;
        private readonly Rectangle colorPreview;
        private readonly TextBox colorCodeTextBox;
        private bool isUpdatingFromCode = false;

        public ColorPickerDialog()
        {
            Title = "色の選択";
            Width = 400;
            Height = 320;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ResizeMode = ResizeMode.NoResize;

            var grid = new Grid();
            grid.Margin = new Thickness(12);

            for (int i = 0; i < 7; i++)
            {
                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            }
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(80) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(50) });

            var redLabel = new TextBlock { Text = "赤:", VerticalAlignment = VerticalAlignment.Center };
            Grid.SetRow(redLabel, 0);
            Grid.SetColumn(redLabel, 0);
            grid.Children.Add(redLabel);

            redSlider = new Slider { Minimum = 0, Maximum = 255, Value = SelectedColor.R };
            Grid.SetRow(redSlider, 0);
            Grid.SetColumn(redSlider, 1);
            grid.Children.Add(redSlider);

            var redValue = new TextBox { Text = ((int)SelectedColor.R).ToString(), Width = 40 };
            Grid.SetRow(redValue, 0);
            Grid.SetColumn(redValue, 2);
            grid.Children.Add(redValue);

            var greenLabel = new TextBlock { Text = "緑:", VerticalAlignment = VerticalAlignment.Center };
            Grid.SetRow(greenLabel, 1);
            Grid.SetColumn(greenLabel, 0);
            grid.Children.Add(greenLabel);

            greenSlider = new Slider { Minimum = 0, Maximum = 255, Value = SelectedColor.G };
            Grid.SetRow(greenSlider, 1);
            Grid.SetColumn(greenSlider, 1);
            grid.Children.Add(greenSlider);

            var greenValue = new TextBox { Text = ((int)SelectedColor.G).ToString(), Width = 40 };
            Grid.SetRow(greenValue, 1);
            Grid.SetColumn(greenValue, 2);
            grid.Children.Add(greenValue);

            var blueLabel = new TextBlock { Text = "青:", VerticalAlignment = VerticalAlignment.Center };
            Grid.SetRow(blueLabel, 2);
            Grid.SetColumn(blueLabel, 0);
            grid.Children.Add(blueLabel);

            blueSlider = new Slider { Minimum = 0, Maximum = 255, Value = SelectedColor.B };
            Grid.SetRow(blueSlider, 2);
            Grid.SetColumn(blueSlider, 1);
            grid.Children.Add(blueSlider);

            var blueValue = new TextBox { Text = ((int)SelectedColor.B).ToString(), Width = 40 };
            Grid.SetRow(blueValue, 2);
            Grid.SetColumn(blueValue, 2);
            grid.Children.Add(blueValue);

            var alphaLabel = new TextBlock { Text = "透明度:", VerticalAlignment = VerticalAlignment.Center };
            Grid.SetRow(alphaLabel, 3);
            Grid.SetColumn(alphaLabel, 0);
            grid.Children.Add(alphaLabel);

            alphaSlider = new Slider { Minimum = 0, Maximum = 255, Value = SelectedColor.A };
            Grid.SetRow(alphaSlider, 3);
            Grid.SetColumn(alphaSlider, 1);
            grid.Children.Add(alphaSlider);

            var alphaValue = new TextBox { Text = ((int)SelectedColor.A).ToString(), Width = 40 };
            Grid.SetRow(alphaValue, 3);
            Grid.SetColumn(alphaValue, 2);
            grid.Children.Add(alphaValue);

            var colorCodeLabel = new TextBlock { Text = "カラーコード:", VerticalAlignment = VerticalAlignment.Center };
            Grid.SetRow(colorCodeLabel, 4);
            Grid.SetColumn(colorCodeLabel, 0);
            grid.Children.Add(colorCodeLabel);

            colorCodeTextBox = new TextBox { Text = ColorToHex(SelectedColor), Width = 100 };
            Grid.SetRow(colorCodeTextBox, 4);
            Grid.SetColumn(colorCodeTextBox, 1);
            Grid.SetColumnSpan(colorCodeTextBox, 2);
            grid.Children.Add(colorCodeTextBox);

            colorPreview = new Rectangle
            {
                Height = 40,
                Stroke = Brushes.Black,
                StrokeThickness = 1,
                Margin = new Thickness(0, 8, 0, 8)
            };
            Grid.SetRow(colorPreview, 5);
            Grid.SetColumn(colorPreview, 0);
            Grid.SetColumnSpan(colorPreview, 3);
            grid.Children.Add(colorPreview);

            var buttonPanel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            var okButton = new Button { Content = "OK", Width = 70, Height = 24, Margin = new Thickness(0, 0, 8, 0) };
            var cancelButton = new Button { Content = "キャンセル", Width = 70, Height = 24 };

            okButton.Click += (s, e) => { DialogResult = true; Close(); };
            cancelButton.Click += (s, e) => { DialogResult = false; Close(); };

            buttonPanel.Children.Add(okButton);
            buttonPanel.Children.Add(cancelButton);

            Grid.SetRow(buttonPanel, 6);
            Grid.SetColumn(buttonPanel, 0);
            Grid.SetColumnSpan(buttonPanel, 3);
            grid.Children.Add(buttonPanel);

            redSlider.ValueChanged += UpdateColor;
            greenSlider.ValueChanged += UpdateColor;
            blueSlider.ValueChanged += UpdateColor;
            alphaSlider.ValueChanged += UpdateColor;

            redValue.TextChanged += (s, e) => { if (int.TryParse(redValue.Text, out int val)) redSlider.Value = Math.Max(0, Math.Min(255, val)); };
            greenValue.TextChanged += (s, e) => { if (int.TryParse(greenValue.Text, out int val)) greenSlider.Value = Math.Max(0, Math.Min(255, val)); };
            blueValue.TextChanged += (s, e) => { if (int.TryParse(blueValue.Text, out int val)) blueSlider.Value = Math.Max(0, Math.Min(255, val)); };
            alphaValue.TextChanged += (s, e) => { if (int.TryParse(alphaValue.Text, out int val)) alphaSlider.Value = Math.Max(0, Math.Min(255, val)); };

            redSlider.ValueChanged += (s, e) => redValue.Text = ((int)redSlider.Value).ToString();
            greenSlider.ValueChanged += (s, e) => greenValue.Text = ((int)greenSlider.Value).ToString();
            blueSlider.ValueChanged += (s, e) => blueValue.Text = ((int)blueSlider.Value).ToString();
            alphaSlider.ValueChanged += (s, e) => alphaValue.Text = ((int)alphaSlider.Value).ToString();

            colorCodeTextBox.TextChanged += ColorCodeTextBox_TextChanged;

            Content = grid;
            UpdateColor(null, null);
        }

        private void ColorCodeTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (isUpdatingFromCode) return;

            if (TryParseHexColor(colorCodeTextBox.Text, out Color color))
            {
                isUpdatingFromCode = true;
                redSlider.Value = color.R;
                greenSlider.Value = color.G;
                blueSlider.Value = color.B;
                alphaSlider.Value = color.A;
                isUpdatingFromCode = false;
            }
        }

        private void UpdateColor(object? sender, RoutedPropertyChangedEventArgs<double>? e)
        {
            SelectedColor = Color.FromArgb(
                (byte)alphaSlider.Value,
                (byte)redSlider.Value,
                (byte)greenSlider.Value,
                (byte)blueSlider.Value
            );

            colorPreview.Fill = new SolidColorBrush(SelectedColor);

            if (!isUpdatingFromCode)
            {
                colorCodeTextBox.Text = ColorToHex(SelectedColor);
            }
        }

        private string ColorToHex(Color color)
        {
            return $"#{color.R:X2}{color.G:X2}{color.B:X2}";
        }

        private bool TryParseHexColor(string hex, out Color color)
        {
            color = Colors.White;

            if (string.IsNullOrEmpty(hex))
                return false;

            hex = hex.Trim();
            if (hex.StartsWith("#"))
                hex = hex.Substring(1);

            if (hex.Length != 6)
                return false;

            try
            {
                byte r = Convert.ToByte(hex.Substring(0, 2), 16);
                byte g = Convert.ToByte(hex.Substring(2, 2), 16);
                byte b = Convert.ToByte(hex.Substring(4, 2), 16);
                color = Color.FromArgb(255, r, g, b);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }

    public class FontCustomSettings : INotifyPropertyChanged
    {
        private bool _useCustomSettings = false;
        public bool UseCustomSettings
        {
            get => _useCustomSettings;
            set { if (_useCustomSettings != value) { _useCustomSettings = value; OnPropertyChanged(nameof(UseCustomSettings)); } }
        }

        private double _fontSize = 48;
        public double FontSize
        {
            get => _fontSize;
            set { if (Math.Abs(_fontSize - value) > 0.1) { _fontSize = value; OnPropertyChanged(nameof(FontSize)); } }
        }

        private bool _useDynamicSize = false;
        public bool UseDynamicSize
        {
            get => _useDynamicSize;
            set { if (_useDynamicSize != value) { _useDynamicSize = value; OnPropertyChanged(nameof(UseDynamicSize)); } }
        }

        private Color _textColor = Colors.White;
        public Color TextColor
        {
            get => _textColor;
            set { if (_textColor != value) { _textColor = value; OnPropertyChanged(nameof(TextColor)); } }
        }

        private bool _bold = false;
        public bool Bold
        {
            get => _bold;
            set { if (_bold != value) { _bold = value; OnPropertyChanged(nameof(Bold)); } }
        }

        private bool _italic = false;
        public bool Italic
        {
            get => _italic;
            set { if (_italic != value) { _italic = value; OnPropertyChanged(nameof(Italic)); } }
        }

        public FontCustomSettings Clone()
        {
            return new FontCustomSettings
            {
                UseCustomSettings = this.UseCustomSettings,
                FontSize = this.FontSize,
                UseDynamicSize = this.UseDynamicSize,
                TextColor = this.TextColor,
                Bold = this.Bold,
                Italic = this.Italic
            };
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}