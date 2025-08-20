using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json.Serialization;
using System.Windows.Media;
using System.Linq;
using System.Collections.Concurrent;
using System.Threading.Tasks;
using System.Windows;

namespace FontShuffle
{
    public class FontItem : INotifyPropertyChanged
    {
        private readonly ConcurrentDictionary<string, FontCustomSettings>? _customSettingsCache;
        private readonly WeakReference<FontShuffleControl>? _controlRef;

        private string _name = "";
        private bool _isJapanese;
        private bool _isSelected;
        private bool _isFavorite;
        private bool _isHidden;
        private volatile bool _hasCustomSettings;

        public FontItem(FontShuffleControl? control = null)
        {
            _controlRef = control != null ? new WeakReference<FontShuffleControl>(control) : null;
            _customSettingsCache = new ConcurrentDictionary<string, FontCustomSettings>();
        }

        public string Name
        {
            get => _name;
            set
            {
                if (_name != value)
                {
                    _name = value ?? "";
                    OnPropertyChanged(nameof(Name));
                    _ = UpdateHasCustomSettingsAsync();
                }
            }
        }

        public bool IsJapanese
        {
            get => _isJapanese;
            set
            {
                if (_isJapanese != value)
                {
                    _isJapanese = value;
                    OnPropertyChanged(nameof(IsJapanese));
                }
            }
        }

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    OnPropertyChanged(nameof(IsSelected));
                }
            }
        }

        public bool IsFavorite
        {
            get => _isFavorite;
            set
            {
                if (_isFavorite != value)
                {
                    _isFavorite = value;
                    OnPropertyChanged(nameof(IsFavorite));
                }
            }
        }

        public bool IsHidden
        {
            get => _isHidden;
            set
            {
                if (_isHidden != value)
                {
                    _isHidden = value;
                    OnPropertyChanged(nameof(IsHidden));
                }
            }
        }

        public bool HasCustomSettings
        {
            get => _hasCustomSettings;
            private set
            {
                if (_hasCustomSettings != value)
                {
                    _hasCustomSettings = value;
                    OnPropertyChanged(nameof(HasCustomSettings));
                }
            }
        }

        private async Task UpdateHasCustomSettingsAsync()
        {
            try
            {
                if (string.IsNullOrEmpty(Name))
                {
                    HasCustomSettings = false;
                    return;
                }

                if (_customSettingsCache?.TryGetValue(Name, out var cachedSettings) == true)
                {
                    HasCustomSettings = cachedSettings?.UseCustomSettings == true;
                    return;
                }

                if (_controlRef?.TryGetTarget(out var control) == true && control.Dispatcher != null)
                {
                    await control.Dispatcher.InvokeAsync(() =>
                    {
                        try
                        {
                            var settings = control.Effect?.FontCustomSettings;
                            if (settings?.ContainsKey(Name) == true)
                            {
                                var setting = settings[Name];
                                _customSettingsCache?.TryAdd(Name, setting);
                                HasCustomSettings = setting?.UseCustomSettings == true;
                            }
                            else
                            {
                                HasCustomSettings = false;
                            }
                        }
                        catch (Exception ex)
                        {
                            LogManager.WriteException(ex, $"HasCustomSettings UIスレッド確認エラー（{Name}）");
                            HasCustomSettings = false;
                        }
                    });
                }
                else
                {
                    HasCustomSettings = false;
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, $"HasCustomSettings更新エラー（{Name}）");
                HasCustomSettings = false;
            }
        }


        public void UpdateCustomSettingsCache(Dictionary<string, FontCustomSettings>? customSettings)
        {
            try
            {
                _customSettingsCache?.Clear();

                if (customSettings != null && !string.IsNullOrEmpty(Name) && customSettings.ContainsKey(Name))
                {
                    _customSettingsCache?.TryAdd(Name, customSettings[Name]);
                    HasCustomSettings = customSettings[Name]?.UseCustomSettings == true;
                }
                else
                {
                    HasCustomSettings = false;
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, $"カスタム設定キャッシュ更新エラー（{Name}）");
                HasCustomSettings = false;
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        public virtual void OnPropertyChanged(string propertyName)
        {
            try
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, $"PropertyChanged発火エラー（{Name}, {propertyName}）");
            }
        }
    }

    public class FontGroup : INotifyPropertyChanged
    {
        private string _name = "";
        private FontGroupType _type;
        private List<string> _fontNames = new();

        public string Name
        {
            get => _name;
            set
            {
                if (_name != value)
                {
                    _name = value ?? "";
                    OnPropertyChanged(nameof(Name));
                }
            }
        }

        public FontGroupType Type
        {
            get => _type;
            set
            {
                if (_type != value)
                {
                    _type = value;
                    OnPropertyChanged(nameof(Type));
                }
            }
        }

        public List<string> FontNames
        {
            get => _fontNames;
            set
            {
                if (_fontNames != value)
                {
                    _fontNames = value ?? new List<string>();
                    OnPropertyChanged(nameof(FontNames));
                }
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            try
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, $"FontGroup PropertyChanged発火エラー（{Name}, {propertyName}）");
            }
        }
    }

    public enum FontGroupType
    {
        All,
        Japanese,
        English,
        Custom
    }

    public class FontSettingsItem : INotifyPropertyChanged
    {
        private string _name = "";
        private FontCustomSettings _customSettings = new();

        public string Name
        {
            get => _name;
            set
            {
                if (_name != value)
                {
                    _name = value ?? "";
                    OnPropertyChanged(nameof(Name));
                }
            }
        }

        public FontCustomSettings CustomSettings
        {
            get => _customSettings;
            set
            {
                if (_customSettings != value)
                {
                    try
                    {
                        if (_customSettings != null)
                            _customSettings.PropertyChanged -= OnCustomSettingsPropertyChanged;

                        _customSettings = value ?? new FontCustomSettings();

                        if (_customSettings != null)
                            _customSettings.PropertyChanged += OnCustomSettingsPropertyChanged;

                        OnPropertyChanged(nameof(CustomSettings));
                    }
                    catch (Exception ex)
                    {
                        LogManager.WriteException(ex, $"FontSettingsItem CustomSettings設定エラー（{Name}）");
                        _customSettings = value ?? new FontCustomSettings();
                    }
                }
            }
        }

        private void OnCustomSettingsPropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            try
            {
                OnPropertyChanged($"CustomSettings.{e.PropertyName}");
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, $"FontSettingsItem カスタム設定PropertyChanged転送エラー（{Name}）");
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            try
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, $"FontSettingsItem PropertyChanged発火エラー（{Name}, {propertyName}）");
            }
        }
    }

    public class FontCustomSettings : INotifyPropertyChanged
    {
        private bool _useCustomSettings = false;
        private double _fontSize = 48;
        private bool _useDynamicSize = false;
        private Color _textColor = Colors.White;
        private bool _bold = false;
        private bool _italic = false;
        private bool _isProjectSpecific = true;

        public bool UseCustomSettings
        {
            get => _useCustomSettings;
            set
            {
                if (_useCustomSettings != value)
                {
                    _useCustomSettings = value;
                    OnPropertyChanged(nameof(UseCustomSettings));
                }
            }
        }

        public double FontSize
        {
            get => _fontSize;
            set
            {
                var newValue = Math.Max(1, Math.Min(1200, value));
                if (Math.Abs(_fontSize - newValue) > 0.1)
                {
                    _fontSize = newValue;
                    OnPropertyChanged(nameof(FontSize));
                }
            }
        }

        public bool UseDynamicSize
        {
            get => _useDynamicSize;
            set
            {
                if (_useDynamicSize != value)
                {
                    _useDynamicSize = value;
                    OnPropertyChanged(nameof(UseDynamicSize));
                }
            }
        }

        public Color TextColor
        {
            get => _textColor;
            set
            {
                if (_textColor != value)
                {
                    _textColor = value;
                    OnPropertyChanged(nameof(TextColor));
                }
            }
        }

        public bool Bold
        {
            get => _bold;
            set
            {
                if (_bold != value)
                {
                    _bold = value;
                    OnPropertyChanged(nameof(Bold));
                }
            }
        }

        public bool Italic
        {
            get => _italic;
            set
            {
                if (_italic != value)
                {
                    _italic = value;
                    OnPropertyChanged(nameof(Italic));
                }
            }
        }

        public bool IsProjectSpecific
        {
            get => _isProjectSpecific;
            set
            {
                if (_isProjectSpecific != value)
                {
                    _isProjectSpecific = value;
                    OnPropertyChanged(nameof(IsProjectSpecific));
                }
            }
        }

        public FontCustomSettings Clone()
        {
            try
            {
                return new FontCustomSettings
                {
                    UseCustomSettings = this.UseCustomSettings,
                    FontSize = this.FontSize,
                    UseDynamicSize = this.UseDynamicSize,
                    TextColor = this.TextColor,
                    Bold = this.Bold,
                    Italic = this.Italic,
                    IsProjectSpecific = this.IsProjectSpecific
                };
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "FontCustomSettings クローン作成エラー");
                return new FontCustomSettings();
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            try
            {
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, $"FontCustomSettings PropertyChanged発火エラー（{propertyName}）");
            }
        }
    }

    public class GitHubRelease
    {
        [JsonPropertyName("tag_name")]
        public string tag_name { get; set; } = "";

        [JsonPropertyName("html_url")]
        public string html_url { get; set; } = "";

        [JsonPropertyName("name")]
        public string name { get; set; } = "";

        [JsonPropertyName("body")]
        public string body { get; set; } = "";

        [JsonPropertyName("published_at")]
        public DateTime published_at { get; set; }

        [JsonPropertyName("prerelease")]
        public bool prerelease { get; set; }
    }

    public enum ShuffleModeType
    {
        [Display(Name = "自動（順次）")]
        Auto,
        [Display(Name = "ランダム")]
        Random,
        [Display(Name = "選択フォントのみ")]
        Selected,
        [Display(Name = "お気に入りのみ")]
        Favorites,
        [Display(Name = "日本語フォントのみ")]
        Japanese,
        [Display(Name = "英数字フォントのみ")]
        English,
        [Display(Name = "順番通り")]
        Ordered
    }

    public enum TextAlignmentType
    {
        [Display(Name = "左揃え")]
        Left,
        [Display(Name = "中央揃え")]
        Center,
        [Display(Name = "右揃え")]
        Right
    }

    public class FontIndexProgress
    {
        private int _current;
        private int _total = 1;
        private string _currentFont = "";

        public int Current
        {
            get => _current;
            set => _current = Math.Max(0, value);
        }

        public int Total
        {
            get => _total;
            set => _total = Math.Max(1, value);
        }

        public string CurrentFont
        {
            get => _currentFont;
            set => _currentFont = value ?? "";
        }

        public bool IsCompleted { get; set; }

        public double ProgressPercent => Total > 0 ? Math.Max(0, Math.Min(100, (double)Current / Total * 100)) : 0;
    }

    public class FontIndexData
    {
        public List<string> AllFonts { get; set; } = new();
        public List<string> JapaneseFonts { get; set; } = new();
        public List<string> EnglishFonts { get; set; } = new();
        public Dictionary<string, bool> FontSupportMap { get; set; } = new();
        public DateTime LastUpdated { get; set; } = DateTime.Now;
        public DateTime FontsDirectoryLastModified { get; set; } = DateTime.Now;
        public int Version { get; set; } = 2;

        public bool IsValid()
        {
            try
            {
                if (AllFonts == null || JapaneseFonts == null || EnglishFonts == null || FontSupportMap == null)
                    return false;

                if (Version < 1)
                    return false;

                if (AllFonts.Count == 0)
                    return false;

                var totalMappedFonts = JapaneseFonts.Count + EnglishFonts.Count;
                if (Math.Abs(totalMappedFonts - AllFonts.Count) > AllFonts.Count * 0.3)
                {
                    LogManager.WriteLog($"フォント分類の整合性が疑わしい（全体: {AllFonts.Count}, 日本語: {JapaneseFonts.Count}, 英語: {EnglishFonts.Count}）", LogLevel.Warning);
                }

                return true;
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "FontIndexData 有効性チェック");
                return false;
            }
        }

        public void Normalize()
        {
            try
            {
                AllFonts = AllFonts ?? new List<string>();
                JapaneseFonts = JapaneseFonts ?? new List<string>();
                EnglishFonts = EnglishFonts ?? new List<string>();
                FontSupportMap = FontSupportMap ?? new Dictionary<string, bool>();

                var fontSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                AllFonts = AllFonts.Where(f => !string.IsNullOrWhiteSpace(f) && fontSet.Add(f)).ToList();

                fontSet.Clear();
                JapaneseFonts = JapaneseFonts.Where(f => !string.IsNullOrWhiteSpace(f) && fontSet.Add(f) && AllFonts.Contains(f, StringComparer.OrdinalIgnoreCase)).ToList();

                fontSet.Clear();
                EnglishFonts = EnglishFonts.Where(f => !string.IsNullOrWhiteSpace(f) && fontSet.Add(f) && AllFonts.Contains(f, StringComparer.OrdinalIgnoreCase)).ToList();

                AllFonts.Sort(StringComparer.OrdinalIgnoreCase);
                JapaneseFonts.Sort(StringComparer.OrdinalIgnoreCase);
                EnglishFonts.Sort(StringComparer.OrdinalIgnoreCase);

                var validKeys = new HashSet<string>(AllFonts, StringComparer.OrdinalIgnoreCase);
                var keysToRemove = FontSupportMap.Keys.Where(k => !validKeys.Contains(k)).ToList();
                foreach (var key in keysToRemove)
                {
                    FontSupportMap.Remove(key);
                }

                LogManager.WriteLog($"FontIndexData正規化完了（全体: {AllFonts.Count}, 日本語: {JapaneseFonts.Count}, 英語: {EnglishFonts.Count}）");
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "FontIndexData 正規化");
            }
        }
    }

    public static class ColorHelper
    {
        public static string ColorToHex(Color color)
        {
            try
            {
                return $"#{color.A:X2}{color.R:X2}{color.G:X2}{color.B:X2}";
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "Color to Hex変換");
                return "#FFFFFFFF";
            }
        }

        public static bool TryParseHexColor(string hex, out Color color)
        {
            color = Colors.White;

            try
            {
                if (string.IsNullOrWhiteSpace(hex))
                    return false;

                hex = hex.Trim();
                if (hex.StartsWith("#"))
                    hex = hex.Substring(1);

                if (hex.Length != 8 && hex.Length != 6)
                    return false;

                if (!System.Text.RegularExpressions.Regex.IsMatch(hex, @"^[0-9A-Fa-f]{6,8}$"))
                    return false;

                if (hex.Length == 8)
                {
                    byte a = Convert.ToByte(hex.Substring(0, 2), 16);
                    byte r = Convert.ToByte(hex.Substring(2, 2), 16);
                    byte g = Convert.ToByte(hex.Substring(4, 2), 16);
                    byte b = Convert.ToByte(hex.Substring(6, 2), 16);
                    color = Color.FromArgb(a, r, g, b);
                }
                else
                {
                    byte r = Convert.ToByte(hex.Substring(0, 2), 16);
                    byte g = Convert.ToByte(hex.Substring(2, 2), 16);
                    byte b = Convert.ToByte(hex.Substring(4, 2), 16);
                    color = Color.FromRgb(r, g, b);
                }
                return true;
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "Hex Color解析");
                return false;
            }
        }
    }
}