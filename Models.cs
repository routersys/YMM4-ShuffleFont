using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Text.Json.Serialization;
using System.Windows.Media;

namespace FontShuffle
{
    public class FontItem : INotifyPropertyChanged
    {
        private readonly FontShuffleControl? _control;
        private Dictionary<string, FontCustomSettings>? _customSettingsCache;

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

        bool isHidden;
        public bool IsHidden
        {
            get => isHidden;
            set { if (isHidden != value) { isHidden = value; OnPropertyChanged(nameof(IsHidden)); } }
        }

        public bool HasCustomSettings
        {
            get
            {
                try
                {
                    // キャッシュされた設定を使用するか、コントロール参照が安全な場合のみアクセス
                    var settings = _customSettingsCache ?? _control?.Effect?.FontCustomSettings;
                    if (settings == null) return false;

                    return settings.ContainsKey(Name) && settings[Name].UseCustomSettings;
                }
                catch
                {
                    // スレッドアクセス例外が発生した場合はfalseを返す
                    return false;
                }
            }
        }

        public void UpdateCustomSettingsCache(Dictionary<string, FontCustomSettings>? customSettings)
        {
            _customSettingsCache = customSettings;
            OnPropertyChanged(nameof(HasCustomSettings));
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

    public class GitHubRelease
    {
        [JsonPropertyName("tag_name")]
        public string tag_name { get; set; } = "";
        [JsonPropertyName("html_url")]
        public string html_url { get; set; } = "";
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
        public int Current { get; set; }
        public int Total { get; set; }
        public string CurrentFont { get; set; } = "";
        public bool IsCompleted { get; set; }
    }

    public class FontIndexData
    {
        public List<string> AllFonts { get; set; } = new();
        public List<string> JapaneseFonts { get; set; } = new();
        public List<string> EnglishFonts { get; set; } = new();
        public Dictionary<string, bool> FontSupportMap { get; set; } = new();
        public DateTime LastUpdated { get; set; } = DateTime.Now;
        public int Version { get; set; } = 1;
    }
}