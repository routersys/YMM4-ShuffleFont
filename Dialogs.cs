using System.Collections.ObjectModel;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Threading;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace FontShuffle
{
    public class InputDialog : Window
    {
        public string Value { get; set; } = "";
        public string Message { get; set; } = "";

        private TextBox textBox = null!;

        public InputDialog()
        {
            try
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
                messageLabel.TextWrapping = TextWrapping.Wrap;
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

                okButton.Click += (s, e) => { Value = textBox.Text ?? ""; DialogResult = true; Close(); };
                cancelButton.Click += (s, e) => { DialogResult = false; Close(); };

                buttonPanel.Children.Add(okButton);
                buttonPanel.Children.Add(cancelButton);

                Grid.SetRow(buttonPanel, 3);
                grid.Children.Add(buttonPanel);

                Content = grid;

                Loaded += (s, e) => {
                    try
                    {
                        textBox.Focus();
                        textBox.SelectAll();
                    }
                    catch (Exception ex)
                    {
                        _ = Task.Run(() => LogManager.WriteException(ex, "InputDialog フォーカス設定"));
                    }
                };
                _ = Task.Run(() => LogManager.WriteLog("入力ダイアログが初期化されました"));
            }
            catch (Exception ex)
            {
                _ = Task.Run(() => LogManager.WriteException(ex, "InputDialog初期化"));
            }
        }
    }

    public class FontGroupManagerWindow : Window
    {
        private ObservableCollection<FontGroup> groups = new();
        private ListBox groupListBox = null!;

        public FontGroupManagerWindow()
        {
            try
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
                    try
                    {
                        if (groupListBox.SelectedItem is FontGroup selected)
                        {
                            var result = MessageBox.Show($"グループ「{selected.Name}」を削除しますか？",
                                "グループ削除", MessageBoxButton.YesNo, MessageBoxImage.Question);

                            if (result == MessageBoxResult.Yes)
                            {
                                groups.Remove(selected);
                                _ = Task.Run(() => LogManager.WriteLog($"グループ「{selected.Name}」を削除しました"));
                            }
                        }
                        else
                        {
                            MessageBox.Show("削除するグループを選択してください。", "グループ削除", MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                    }
                    catch (Exception ex)
                    {
                        _ = Task.Run(() => LogManager.WriteException(ex, "グループ削除"));
                        MessageBox.Show("グループの削除中にエラーが発生しました。", "エラー", MessageBoxButton.OK, MessageBoxImage.Error);
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
                _ = Task.Run(() => LogManager.WriteLog("グループ管理ダイアログが初期化されました"));
            }
            catch (Exception ex)
            {
                _ = Task.Run(() => LogManager.WriteException(ex, "FontGroupManagerWindow初期化"));
            }
        }

        public void SetGroups(List<FontGroup> fontGroups)
        {
            try
            {
                groups.Clear();
                foreach (var group in fontGroups ?? new List<FontGroup>())
                {
                    groups.Add(group);
                }
                _ = Task.Run(() => LogManager.WriteLog($"グループ管理に{fontGroups?.Count ?? 0}個のグループを設定しました"));
            }
            catch (Exception ex)
            {
                _ = Task.Run(() => LogManager.WriteException(ex, "グループ設定"));
            }
        }

        public List<FontGroup> GetUpdatedGroups()
        {
            try
            {
                var result = groups.ToList();
                _ = Task.Run(() => LogManager.WriteLog($"更新されたグループを取得（{result.Count}個）"));
                return result;
            }
            catch (Exception ex)
            {
                _ = Task.Run(() => LogManager.WriteException(ex, "更新グループ取得"));
                return new List<FontGroup>();
            }
        }
    }

    public class ColorPickerDialog : Window
    {
        public Color SelectedColor { get; set; } = Colors.White;

        private readonly Slider redSlider = null!;
        private readonly Slider greenSlider = null!;
        private readonly Slider blueSlider = null!;
        private readonly Slider alphaSlider = null!;
        private readonly Rectangle colorPreview = null!;
        private readonly TextBox colorCodeTextBox = null!;
        private bool isUpdatingFromCode = false;

        public ColorPickerDialog()
        {
            try
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

                colorCodeTextBox = new TextBox { Text = ColorHelper.ColorToHex(SelectedColor), Width = 100 };
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

                redValue.TextChanged += (s, e) => {
                    try
                    {
                        if (int.TryParse(redValue.Text, out int val))
                            redSlider.Value = Math.Max(0, Math.Min(255, val));
                    }
                    catch (Exception ex)
                    {
                        _ = Task.Run(() => LogManager.WriteException(ex, "赤値変更"));
                    }
                };
                greenValue.TextChanged += (s, e) => {
                    try
                    {
                        if (int.TryParse(greenValue.Text, out int val))
                            greenSlider.Value = Math.Max(0, Math.Min(255, val));
                    }
                    catch (Exception ex)
                    {
                        _ = Task.Run(() => LogManager.WriteException(ex, "緑値変更"));
                    }
                };
                blueValue.TextChanged += (s, e) => {
                    try
                    {
                        if (int.TryParse(blueValue.Text, out int val))
                            blueSlider.Value = Math.Max(0, Math.Min(255, val));
                    }
                    catch (Exception ex)
                    {
                        _ = Task.Run(() => LogManager.WriteException(ex, "青値変更"));
                    }
                };
                alphaValue.TextChanged += (s, e) => {
                    try
                    {
                        if (int.TryParse(alphaValue.Text, out int val))
                            alphaSlider.Value = Math.Max(0, Math.Min(255, val));
                    }
                    catch (Exception ex)
                    {
                        _ = Task.Run(() => LogManager.WriteException(ex, "透明度値変更"));
                    }
                };

                redSlider.ValueChanged += (s, e) => redValue.Text = ((int)redSlider.Value).ToString();
                greenSlider.ValueChanged += (s, e) => greenValue.Text = ((int)greenSlider.Value).ToString();
                blueSlider.ValueChanged += (s, e) => blueValue.Text = ((int)blueSlider.Value).ToString();
                alphaSlider.ValueChanged += (s, e) => alphaValue.Text = ((int)alphaSlider.Value).ToString();

                colorCodeTextBox.TextChanged += ColorCodeTextBox_TextChanged;

                Content = grid;
                UpdateColor(null, null);
                _ = Task.Run(() => LogManager.WriteLog("カラーピッカーダイアログが初期化されました"));
            }
            catch (Exception ex)
            {
                _ = Task.Run(() => LogManager.WriteException(ex, "ColorPickerDialog初期化"));
            }
        }

        private void ColorCodeTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            if (isUpdatingFromCode) return;

            try
            {
                if (ColorHelper.TryParseHexColor(colorCodeTextBox.Text, out Color color))
                {
                    isUpdatingFromCode = true;
                    redSlider.Value = color.R;
                    greenSlider.Value = color.G;
                    blueSlider.Value = color.B;
                    alphaSlider.Value = color.A;
                    isUpdatingFromCode = false;
                }
            }
            catch (Exception ex)
            {
                _ = Task.Run(() => LogManager.WriteException(ex, "カラーコードテキスト変更"));
            }
            finally
            {
                isUpdatingFromCode = false;
            }
        }

        private void UpdateColor(object? sender, RoutedPropertyChangedEventArgs<double>? e)
        {
            try
            {
                SelectedColor = Color.FromArgb(
                    (byte)Math.Round(alphaSlider.Value),
                    (byte)Math.Round(redSlider.Value),
                    (byte)Math.Round(greenSlider.Value),
                    (byte)Math.Round(blueSlider.Value)
                );

                colorPreview.Fill = new SolidColorBrush(SelectedColor);

                if (!isUpdatingFromCode)
                {
                    colorCodeTextBox.Text = ColorHelper.ColorToHex(SelectedColor);
                }
            }
            catch (Exception ex)
            {
                _ = Task.Run(() => LogManager.WriteException(ex, "色更新"));
            }
        }
    }

    public class FontIndexProgressDialog : Window
    {
        private readonly ProgressBar progressBar = null!;
        private readonly TextBlock statusText = null!;
        private readonly TextBlock currentFontText = null!;
        private readonly Button cancelButton = null!;
        private CancellationTokenSource? cancellationTokenSource;

        public bool WasCanceled { get; private set; }
        public CancellationToken CancellationToken => cancellationTokenSource?.Token ?? CancellationToken.None;

        public FontIndexProgressDialog()
        {
            try
            {
                Title = "フォントインデックス作成中";
                Width = 400;
                Height = 200;
                WindowStartupLocation = WindowStartupLocation.CenterOwner;
                ResizeMode = ResizeMode.NoResize;

                cancellationTokenSource = new CancellationTokenSource();

                var grid = new Grid();
                grid.Margin = new Thickness(20);
                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

                var titleText = new TextBlock
                {
                    Text = "フォントの詳細情報を解析しています...",
                    FontSize = 14,
                    FontWeight = FontWeights.Bold,
                    Margin = new Thickness(0, 0, 0, 12),
                    TextWrapping = TextWrapping.Wrap
                };
                Grid.SetRow(titleText, 0);
                grid.Children.Add(titleText);

                statusText = new TextBlock
                {
                    Text = "初期化中...",
                    Margin = new Thickness(0, 0, 0, 8),
                    TextWrapping = TextWrapping.Wrap
                };
                Grid.SetRow(statusText, 1);
                grid.Children.Add(statusText);

                progressBar = new ProgressBar
                {
                    Height = 20,
                    Minimum = 0,
                    Maximum = 100,
                    Margin = new Thickness(0, 0, 0, 8)
                };
                Grid.SetRow(progressBar, 2);
                grid.Children.Add(progressBar);

                currentFontText = new TextBlock
                {
                    Text = "",
                    FontSize = 10,
                    Foreground = Brushes.Gray,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    TextWrapping = TextWrapping.NoWrap
                };
                Grid.SetRow(currentFontText, 3);
                grid.Children.Add(currentFontText);

                cancelButton = new Button
                {
                    Content = "キャンセル",
                    Width = 80,
                    Height = 24,
                    HorizontalAlignment = HorizontalAlignment.Right
                };
                cancelButton.Click += CancelButton_Click;
                Grid.SetRow(cancelButton, 4);
                grid.Children.Add(cancelButton);

                Content = grid;

                Closing += (s, e) => {
                    try
                    {
                        if (!WasCanceled && cancellationTokenSource != null && !cancellationTokenSource.Token.IsCancellationRequested)
                        {
                            RequestCancel();
                        }
                    }
                    catch (Exception ex)
                    {
                        _ = Task.Run(() => LogManager.WriteException(ex, "FontIndexProgressDialog 終了処理"));
                    }
                };

                _ = Task.Run(() => LogManager.WriteLog("フォントインデックス進行ダイアログが初期化されました"));
            }
            catch (Exception ex)
            {
                _ = Task.Run(() => LogManager.WriteException(ex, "FontIndexProgressDialog初期化"));
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                RequestCancel();
            }
            catch (Exception ex)
            {
                _ = Task.Run(() => LogManager.WriteException(ex, "FontIndexProgressDialog キャンセルボタン"));
            }
        }

        private void RequestCancel()
        {
            try
            {
                WasCanceled = true;
                cancellationTokenSource?.Cancel();

                Dispatcher.BeginInvoke(new Action(() => {
                    try
                    {
                        statusText.Text = "キャンセル中...";
                        cancelButton.IsEnabled = false;
                        cancelButton.Content = "キャンセル中";
                    }
                    catch (Exception ex)
                    {
                        _ = Task.Run(() => LogManager.WriteException(ex, "キャンセル UI 更新"));
                    }
                }));

                _ = Task.Run(() => LogManager.WriteLog("フォントインデックス作成のキャンセルを要求しました"));
            }
            catch (Exception ex)
            {
                _ = Task.Run(() => LogManager.WriteException(ex, "キャンセル要求処理"));
            }
        }

        public void UpdateProgress(FontIndexProgress progress)
        {
            try
            {
                if (Dispatcher.CheckAccess())
                {
                    UpdateProgressInternal(progress);
                }
                else
                {
                    Dispatcher.BeginInvoke(new Action(() => {
                        try
                        {
                            UpdateProgressInternal(progress);
                        }
                        catch (Exception ex)
                        {
                            _ = Task.Run(() => LogManager.WriteException(ex, "フォントインデックス進行更新（Dispatcher）"));
                        }
                    }));
                }
            }
            catch (Exception ex)
            {
                _ = Task.Run(() => LogManager.WriteException(ex, "フォントインデックス進行更新"));
            }
        }

        private void UpdateProgressInternal(FontIndexProgress progress)
        {
            try
            {
                if (progress.IsCompleted)
                {
                    statusText.Text = "完了しました";
                    progressBar.Value = 100;
                    currentFontText.Text = "";
                    cancelButton.Content = "閉じる";
                    cancelButton.IsEnabled = true;
                }
                else if (WasCanceled)
                {
                    statusText.Text = "キャンセルされました";
                    currentFontText.Text = "";
                    cancelButton.Content = "閉じる";
                    cancelButton.IsEnabled = true;
                }
                else
                {
                    statusText.Text = $"処理中... ({progress.Current}/{progress.Total})";
                    double progressPercent = progress.Total > 0 ? (double)progress.Current / progress.Total * 100 : 0;
                    progressBar.Value = Math.Max(0, Math.Min(100, progressPercent));
                    currentFontText.Text = $"現在の処理: {progress.CurrentFont ?? ""}";
                }
            }
            catch (Exception ex)
            {
                _ = Task.Run(() => LogManager.WriteException(ex, "フォントインデックス進行更新内部処理"));
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            try
            {
                cancellationTokenSource?.Dispose();
                cancellationTokenSource = null;
            }
            catch (Exception ex)
            {
                _ = Task.Run(() => LogManager.WriteException(ex, "FontIndexProgressDialog リソース解放"));
            }
            finally
            {
                base.OnClosed(e);
            }
        }
    }
}