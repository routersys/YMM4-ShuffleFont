using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Numerics;
using Vortice.DCommon;
using Vortice.Direct2D1;
using Vortice.DirectWrite;
using Vortice.Mathematics;
using YukkuriMovieMaker.Commons;
using YukkuriMovieMaker.Controls;
using YukkuriMovieMaker.Exo;
using YukkuriMovieMaker.Player.Video;
using YukkuriMovieMaker.Plugin.Effects;
using System.Threading;
using System.Threading.Tasks;

namespace FontShuffle
{
    [VideoEffect("シャッフルフォント", new[] { "アニメーション" }, new[] { "shuffle font", "sf" }, IsAviUtlSupported = false)]
    public class FontShuffleEffect : VideoEffectBase
    {
        public override string Label => "フォントシャッフル";

        [Display(Name = "表示テキスト", Description = "表示するテキスト", GroupName = "テキスト設定")]
        [TextEditor]
        public string DisplayText { get => _displayText; set => Set(ref _displayText, value); }
        private string _displayText = "サンプルテキスト";

        [Display(Name = "幅", Description = "描画領域の幅", GroupName = "テキスト設定")]
        [AnimationSlider("F0", "px", 1, 7680)]
        public Animation Width { get; } = new(1920, 1, 7680);

        [Display(Name = "高さ", Description = "描画領域の高さ", GroupName = "テキスト設定")]
        [AnimationSlider("F0", "px", 1, 4320)]
        public Animation Height { get; } = new(1080, 1, 4320);

        [Display(Name = "文字揃え", Description = "テキストの配置", GroupName = "テキスト設定")]
        [EnumComboBox]
        public TextAlignmentType TextAlignment { get => _textAlignment; set => Set(ref _textAlignment, value); }
        private TextAlignmentType _textAlignment = TextAlignmentType.Center;

        [Display(Name = "文字間隔", Description = "文字間の間隔", GroupName = "テキスト設定")]
        [AnimationSlider("F1", "px", -10, 50)]
        public Animation LetterSpacing { get; } = new(0, -10, 50);

        [Display(Name = "フォント変更間隔", Description = "フォントを変更するフレーム間隔")]
        [AnimationSlider("F0", "フレーム", 1, 120)]
        public Animation Interval { get; } = new(30, 1, 600);

        [Display(Name = "シャッフルモード", Description = "フォントの選択方法")]
        [EnumComboBox]
        public ShuffleModeType ShuffleMode { get => _shuffleMode; set => Set(ref _shuffleMode, value); }
        private ShuffleModeType _shuffleMode = ShuffleModeType.Auto;

        [Display(Name = "ランダムシード", Description = "ランダムモード用のシード値")]
        [AnimationSlider("F0", "", 1, 99999)]
        public Animation RandomSeed { get; } = new(12345, 1, 99999);

        [Display(Name = "フォントサイズ", Description = "テキストのサイズ（デフォルト）")]
        [AnimationSlider("F0", "px", 10, 200)]
        public Animation FontSize { get; } = new(48, 10, 500);

        [Display(Name = "文字色", Description = "テキストの色（デフォルト）")]
        [ColorPicker]
        public System.Windows.Media.Color TextColor { get => _textColor; set => Set(ref _textColor, value); }
        private System.Windows.Media.Color _textColor = System.Windows.Media.Colors.White;

        [Display(Name = "太字", Description = "太字にする（デフォルト）")]
        [ToggleSlider]
        public bool Bold { get => _bold; set => Set(ref _bold, value); }
        private bool _bold = false;

        [Display(Name = "イタリック", Description = "斜体にする（デフォルト）")]
        [ToggleSlider]
        public bool Italic { get => _italic; set => Set(ref _italic, value); }
        private bool _italic = false;

        [Display(Name = "フォント管理", Description = "使用するフォントを選択", GroupName = "フォント設定")]
        [FontShuffleControl]
        public bool FontManagement { get; set; } = true;

        public string CurrentFont { get => _currentFont; set => Set(ref _currentFont, value); }
        private string _currentFont = "Yu Gothic UI";

        public List<string> SelectedFonts { get => _selectedFonts; set => Set(ref _selectedFonts, value); }
        private List<string> _selectedFonts = new();

        public List<string> FavoriteFonts { get => _favoriteFonts; set => Set(ref _favoriteFonts, value); }
        private List<string> _favoriteFonts = new();

        public List<string> OrderedFonts { get => _orderedFonts; set => Set(ref _orderedFonts, value); }
        private List<string> _orderedFonts = new();

        public Dictionary<string, FontCustomSettings> FontCustomSettings { get => _fontCustomSettings; set => Set(ref _fontCustomSettings, value); }
        private Dictionary<string, FontCustomSettings> _fontCustomSettings = new();

        public string IgnoredVersion { get => _ignoredVersion; set => Set(ref _ignoredVersion, value); }
        private string _ignoredVersion = "";

        private static readonly object _globalLock = new object();
        private static List<string>? _globalHiddenFonts;
        private static List<FontGroup>? _globalFontGroups;
        private static Dictionary<string, FontCustomSettings>? _globalFontCustomSettings;
        private static FontIndexData? _cachedSystemFonts;
        private static readonly SemaphoreSlim _initializationSemaphore = new(1, 1);
        private static volatile bool _isInitialized = false;

        public static List<string> GlobalHiddenFonts
        {
            get
            {
                lock (_globalLock)
                {
                    if (_globalHiddenFonts == null)
                    {
                        _globalHiddenFonts = new List<string>();
                        Task.Run(async () =>
                        {
                            try
                            {
                                var hiddenFonts = await FontIndexer.LoadHiddenFontsAsync();
                                lock (_globalLock)
                                {
                                    _globalHiddenFonts = hiddenFonts;
                                }
                            }
                            catch (Exception ex)
                            {
                                LogManager.WriteException(ex, "非表示フォント非同期読み込み");
                            }
                        });
                    }
                    return _globalHiddenFonts;
                }
            }
            private set
            {
                lock (_globalLock)
                {
                    _globalHiddenFonts = value;
                }
            }
        }

        public static List<FontGroup> GlobalFontGroups
        {
            get
            {
                lock (_globalLock)
                {
                    if (_globalFontGroups == null)
                    {
                        _globalFontGroups = new List<FontGroup>();
                        Task.Run(async () =>
                        {
                            try
                            {
                                var fontGroups = await FontIndexer.LoadFontGroupsAsync();
                                lock (_globalLock)
                                {
                                    _globalFontGroups = fontGroups;
                                }
                            }
                            catch (Exception ex)
                            {
                                LogManager.WriteException(ex, "フォントグループ非同期読み込み");
                            }
                        });
                    }
                    return _globalFontGroups;
                }
            }
            private set
            {
                lock (_globalLock)
                {
                    _globalFontGroups = value;
                }
            }
        }

        public static Dictionary<string, FontCustomSettings> GlobalFontCustomSettings
        {
            get
            {
                lock (_globalLock)
                {
                    if (_globalFontCustomSettings == null)
                    {
                        _globalFontCustomSettings = new Dictionary<string, FontCustomSettings>();
                        Task.Run(async () =>
                        {
                            try
                            {
                                var settings = await FontIndexer.LoadCustomFontSettingsAsync();
                                lock (_globalLock)
                                {
                                    _globalFontCustomSettings = settings;
                                }
                            }
                            catch (Exception ex)
                            {
                                LogManager.WriteException(ex, "個別フォント設定の非同期読み込み");
                            }
                        });
                    }
                    return _globalFontCustomSettings;
                }
            }
            private set
            {
                lock (_globalLock)
                {
                    _globalFontCustomSettings = value;
                }
            }
        }


        static FontShuffleEffect()
        {
            try
            {
                LogManager.WriteLog("FontShuffleEffect静的初期化開始");
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "FontShuffleEffect静的初期化");
            }
        }

        public FontShuffleEffect()
        {
            try
            {
                Task.Run(InitializeGlobalSettingsAsync);
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "FontShuffleEffectコンストラクタ");
            }
        }

        private static async Task InitializeGlobalSettingsAsync()
        {
            try
            {
                await _initializationSemaphore.WaitAsync();
                if (_isInitialized)
                    return;

                var hiddenFontsTask = FontIndexer.LoadHiddenFontsAsync();
                var fontGroupsTask = FontIndexer.LoadFontGroupsAsync();
                var customSettingsTask = FontIndexer.LoadCustomFontSettingsAsync();


                await Task.WhenAll(hiddenFontsTask, fontGroupsTask, customSettingsTask);

                lock (_globalLock)
                {
                    _globalHiddenFonts = hiddenFontsTask.Result;
                    _globalFontGroups = fontGroupsTask.Result;
                    _globalFontCustomSettings = customSettingsTask.Result;
                }

                _isInitialized = true;
                LogManager.WriteLog("グローバル設定初期化完了");
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "グローバル設定初期化");
            }
            finally
            {
                _initializationSemaphore.Release();
            }
        }

        public static void SaveGlobalSettings()
        {
            try
            {
                Task.Run(async () =>
                {
                    try
                    {
                        List<string> hiddenFonts;
                        List<FontGroup> fontGroups;
                        Dictionary<string, FontCustomSettings> fontSettings;

                        lock (_globalLock)
                        {
                            hiddenFonts = _globalHiddenFonts?.ToList() ?? new List<string>();
                            fontGroups = _globalFontGroups?.ToList() ?? new List<FontGroup>();
                            fontSettings = _globalFontCustomSettings?
                                .Where(kvp => !kvp.Value.IsProjectSpecific)
                                .ToDictionary(kvp => kvp.Key, kvp => kvp.Value)
                                ?? new Dictionary<string, FontCustomSettings>();
                        }

                        var saveHiddenTask = FontIndexer.SaveHiddenFontsAsync(hiddenFonts);
                        var saveGroupsTask = FontIndexer.SaveFontGroupsAsync(fontGroups);
                        var saveSettingsTask = FontIndexer.SaveCustomFontSettingsAsync(fontSettings);


                        await Task.WhenAll(saveHiddenTask, saveGroupsTask, saveSettingsTask);
                    }
                    catch (Exception ex)
                    {
                        LogManager.WriteException(ex, "グローバル設定保存非同期処理");
                    }
                });
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "グローバル設定保存");
            }
        }

        public override IVideoEffectProcessor CreateVideoEffect(IGraphicsDevicesAndContext devices)
        {
            return new FontShuffleProcessor(devices, this);
        }

        public override IEnumerable<string> CreateExoVideoFilters(int keyFrameIndex, ExoOutputDescription exoOutputDescription)
        {
            return Array.Empty<string>();
        }

        protected override IEnumerable<IAnimatable> GetAnimatables() => new IAnimatable[] { Width, Height, Interval, RandomSeed, FontSize, LetterSpacing };

        public string GetFontForFrame(int frame, int length, int fps)
        {
            try
            {
                var interval = Math.Max(1, (int)Interval.GetValue(frame, length, fps));
                var currentInterval = frame / interval;
                var seed = (int)RandomSeed.GetValue(frame, length, fps);

                var fonts = GetActiveFontList();
                if (fonts.Count > 0)
                {
                    var selectedFont = ShuffleMode switch
                    {
                        ShuffleModeType.Auto => fonts[currentInterval % fonts.Count],
                        ShuffleModeType.Random => fonts[new Random(seed + currentInterval).Next(fonts.Count)],
                        ShuffleModeType.Selected => fonts[currentInterval % fonts.Count],
                        ShuffleModeType.Favorites => fonts[currentInterval % fonts.Count],
                        ShuffleModeType.Japanese => fonts[currentInterval % fonts.Count],
                        ShuffleModeType.English => fonts[currentInterval % fonts.Count],
                        ShuffleModeType.Ordered => fonts[currentInterval % fonts.Count],
                        _ => fonts[currentInterval % fonts.Count]
                    };
                    CurrentFont = selectedFont;
                    return selectedFont;
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "フォント選択処理");
            }
            return GetFallbackFont();
        }

        public string GetFontAtIndex(int index, int frame, int length, int fps)
        {
            try
            {
                var seed = (int)RandomSeed.GetValue(frame, length, fps);

                var fonts = GetActiveFontList();
                if (fonts.Count > 0)
                {
                    var selectedFont = ShuffleMode switch
                    {
                        ShuffleModeType.Auto => fonts[index % fonts.Count],
                        ShuffleModeType.Random => fonts[new Random(seed + index).Next(fonts.Count)],
                        ShuffleModeType.Selected => fonts[index % fonts.Count],
                        ShuffleModeType.Favorites => fonts[index % fonts.Count],
                        ShuffleModeType.Japanese => fonts[index % fonts.Count],
                        ShuffleModeType.English => fonts[index % fonts.Count],
                        ShuffleModeType.Ordered => fonts[index % fonts.Count],
                        _ => fonts[index % fonts.Count]
                    };
                    CurrentFont = selectedFont;
                    return selectedFont;
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "フォント選択処理");
            }
            return GetFallbackFont();
        }

        private string GetFallbackFont()
        {
            try
            {
                var fallbackFonts = new[] { "Yu Gothic UI", "Meiryo", "MS Gothic", "Arial", "Times New Roman", "Segoe UI" };
                foreach (var font in fallbackFonts)
                {
                    if (IsFontAvailable(font))
                        return font;
                }
                return "Arial";
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "フォールバックフォント取得");
                return "Arial";
            }
        }

        private bool IsFontAvailable(string fontName)
        {
            try
            {
                var fontFamily = new System.Windows.Media.FontFamily(fontName);
                return fontFamily.FamilyNames.Count > 0;
            }
            catch
            {
                return false;
            }
        }

        public List<string> GetActiveFontList()
        {
            try
            {
                var systemFonts = GetSystemFonts();

                lock (_globalLock)
                {
                    var hiddenSet = new HashSet<string>(_globalHiddenFonts ?? new List<string>(), StringComparer.OrdinalIgnoreCase);

                    var filterFonts = (List<string> fonts) => fonts?.Where(f => !string.IsNullOrEmpty(f) && !hiddenSet.Contains(f)).ToList() ?? new List<string>();

                    var result = ShuffleMode switch
                    {
                        ShuffleModeType.Ordered when _orderedFonts?.Count > 0 => filterFonts(_orderedFonts),
                        ShuffleModeType.Selected when _selectedFonts?.Count > 0 => filterFonts(_selectedFonts),
                        ShuffleModeType.Favorites when _favoriteFonts?.Count > 0 => filterFonts(_favoriteFonts),
                        ShuffleModeType.Japanese => filterFonts(systemFonts.JapaneseFonts.Count > 0 ? systemFonts.JapaneseFonts : systemFonts.AllFonts),
                        ShuffleModeType.English => filterFonts(systemFonts.EnglishFonts.Count > 0 ? systemFonts.EnglishFonts : systemFonts.AllFonts),
                        ShuffleModeType.Selected when _selectedFonts?.Count == 0 => new List<string>(),
                        ShuffleModeType.Favorites when _favoriteFonts?.Count == 0 => new List<string>(),
                        ShuffleModeType.Ordered when _orderedFonts?.Count == 0 => new List<string>(),
                        _ => filterFonts(systemFonts.AllFonts)
                    };

                    if (result.Count == 0)
                    {
                        var fallback = GetFallbackFont();
                        LogManager.WriteLog($"アクティブフォントリストが空のため、フォールバック「{fallback}」を使用");
                        return new List<string> { fallback };
                    }

                    return result;
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "アクティブフォントリスト取得");
                return new List<string> { GetFallbackFont() };
            }
        }

        public FontCustomSettings? GetFontSettings(string fontName)
        {
            try
            {
                if (string.IsNullOrEmpty(fontName)) return null;

                if (_fontCustomSettings?.TryGetValue(fontName, out var projectSetting) == true && projectSetting?.UseCustomSettings == true)
                {
                    return projectSetting;
                }

                if (GlobalFontCustomSettings?.TryGetValue(fontName, out var globalSetting) == true && globalSetting?.UseCustomSettings == true)
                {
                    return globalSetting;
                }

                return null;
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, $"フォント設定取得（{fontName}）");
                return null;
            }
        }


        private FontIndexData GetSystemFonts()
        {
            if (_cachedSystemFonts != null)
                return _cachedSystemFonts;

            try
            {
                if (FontIndexer.HasValidIndex())
                {
                    var existingIndex = FontIndexer.LoadIndex();
                    if (existingIndex != null)
                    {
                        _cachedSystemFonts = existingIndex;
                        return existingIndex;
                    }
                }

                var basicIndex = CreateBasicFontIndex();
                _cachedSystemFonts = basicIndex;
                return basicIndex;
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "システムフォント取得");
                var fallbackIndex = CreateFallbackFontIndex();
                _cachedSystemFonts = fallbackIndex;
                return fallbackIndex;
            }
        }

        private static FontIndexData CreateBasicFontIndex()
        {
            var index = new FontIndexData
            {
                Version = 2,
                LastUpdated = DateTime.Now,
                FontsDirectoryLastModified = DateTime.Now
            };

            try
            {
                var fontResults = new List<(string name, bool isJapanese)>();
                var systemFonts = System.Windows.Media.Fonts.SystemFontFamilies.ToArray();

                foreach (var fontFamily in systemFonts)
                {
                    try
                    {
                        string fontName = FontHelper.GetFontFamilyName(fontFamily);
                        if (!string.IsNullOrEmpty(fontName))
                        {
                            var isJapanese = FontHelper.IsJapaneseFontUnified(fontFamily, fontName);
                            fontResults.Add((fontName, isJapanese));
                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.WriteException(ex, "個別フォント処理");
                    }
                }

                var uniqueFonts = fontResults
                    .GroupBy(f => f.name, StringComparer.OrdinalIgnoreCase)
                    .Select(g => g.First())
                    .OrderBy(f => f.name, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                foreach (var (name, isJapanese) in uniqueFonts)
                {
                    index.AllFonts.Add(name);
                    index.FontSupportMap[name] = isJapanese;

                    if (isJapanese)
                        index.JapaneseFonts.Add(name);
                    else
                        index.EnglishFonts.Add(name);
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "基本フォントインデックス作成");
                return CreateFallbackFontIndex();
            }

            if (index.AllFonts.Count == 0)
            {
                return CreateFallbackFontIndex();
            }

            return index;
        }

        private static FontIndexData CreateFallbackFontIndex()
        {
            LogManager.WriteLog("フォールバックフォントインデックスを作成", LogLevel.Warning);

            var fallbackFonts = new Dictionary<string, bool>
            {
                { "Yu Gothic UI", true },
                { "Meiryo", true },
                { "MS Gothic", true },
                { "Arial", false },
                { "Times New Roman", false },
                { "Calibri", false },
                { "Verdana", false },
                { "Segoe UI", false },
                { "Tahoma", false }
            };

            var index = new FontIndexData
            {
                Version = 2,
                LastUpdated = DateTime.Now,
                FontsDirectoryLastModified = DateTime.Now
            };

            foreach (var font in fallbackFonts)
            {
                try
                {
                    if (IsFontAvailableStatic(font.Key))
                    {
                        index.AllFonts.Add(font.Key);
                        index.FontSupportMap[font.Key] = font.Value;

                        if (font.Value)
                            index.JapaneseFonts.Add(font.Key);
                        else
                            index.EnglishFonts.Add(font.Key);
                    }
                }
                catch (Exception ex)
                {
                    LogManager.WriteException(ex, $"フォールバックフォント確認（{font.Key}）");
                }
            }

            if (index.AllFonts.Count == 0)
            {
                index.AllFonts.Add("Arial");
                index.EnglishFonts.Add("Arial");
                index.FontSupportMap["Arial"] = false;
                LogManager.WriteLog("最終フォールバックとしてArialを追加");
            }

            return index;
        }

        private static bool IsFontAvailableStatic(string fontName)
        {
            try
            {
                var fontFamily = new System.Windows.Media.FontFamily(fontName);
                return fontFamily.FamilyNames.Count > 0;
            }
            catch
            {
                return false;
            }
        }
    }

    public class FontShuffleProcessor : IVideoEffectProcessor
    {
        private readonly IGraphicsDevicesAndContext _devices;
        private readonly FontShuffleEffect _effect;

        private int _lastProcessedFrame = -1;
        private int _fontIndex = -1;
        private int _nextChangeFrame = 0;
        private string? _currentFont;

        private readonly ConcurrentDictionary<string, IDWriteTextFormat> _textFormatCache = new();
        private ID2D1SolidColorBrush? _textBrush;
        private IDWriteFactory? _dwriteFactory;
        private ID2D1CommandList? _commandList;
        private ID2D1Image? _inputImage;

        private string _lastUsedFont = "";
        private float _lastFontSize = 0;
        private Color4 _lastColor = new(0, 0, 0, 0);
        private string _lastText = "";
        private bool _lastBold = false;
        private bool _lastItalic = false;
        private float _lastLetterSpacing = 0;
        private TextAlignmentType _lastTextAlignment = TextAlignmentType.Center;
        private bool _isInitialized = false;
        private SizeI _currentSize = new(1920, 1080);

        public ID2D1Image? Output => _commandList;

        public FontShuffleProcessor(IGraphicsDevicesAndContext devices, FontShuffleEffect effect)
        {
            _devices = devices;
            _effect = effect;
            Initialize();
        }

        private void Initialize()
        {
            try
            {
                if (_devices?.DeviceContext != null)
                {
                    _textBrush = _devices.DeviceContext.CreateSolidColorBrush(new Color4(1f, 1f, 1f, 1f));
                    var result = DWrite.DWriteCreateFactory(Vortice.DirectWrite.FactoryType.Shared, out _dwriteFactory);
                    if (result.Success && _dwriteFactory != null)
                    {
                        _isInitialized = true;
                        LogManager.WriteLog("FontShuffleProcessor初期化完了");
                    }
                    else
                    {
                        LogManager.WriteLog($"DWriteファクトリ作成に失敗: {result}", LogLevel.Error);
                    }
                }
                else
                {
                    LogManager.WriteLog("DeviceContextがnull", LogLevel.Error);
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "FontShuffleProcessor初期化");
                _isInitialized = false;
            }
        }

        public DrawDescription Update(EffectDescription effectDescription)
        {
            if (!_isInitialized || _devices?.DeviceContext == null)
                return effectDescription.DrawDescription;

            try
            {
                var frame = effectDescription.ItemPosition.Frame;
                var length = effectDescription.ItemDuration.Frame;
                var fps = effectDescription.FPS;

                if (frame < _lastProcessedFrame)
                {
                    _fontIndex = -1;
                    _nextChangeFrame = 0;
                }
                _lastProcessedFrame = frame;

                if (frame >= _nextChangeFrame)
                {
                    while (frame >= _nextChangeFrame)
                    {
                        _fontIndex++;
                        var interval = Math.Max(1, (int)_effect.Interval.GetValue(_nextChangeFrame, length, fps));
                        _nextChangeFrame += interval;
                    }
                }
                _currentFont = _effect.GetFontAtIndex(_fontIndex, frame, length, (int)fps);

                var width = Math.Max(1, (int)_effect.Width.GetValue(frame, length, fps));
                var height = Math.Max(1, (int)_effect.Height.GetValue(frame, length, fps));
                _currentSize = new SizeI(width, height);

                var fonts = _effect.GetActiveFontList();
                if (fonts.Count == 0)
                {
                    CreateEmptyCommandList();
                    return effectDescription.DrawDescription;
                }

                var customSettings = _effect.GetFontSettings(_currentFont ?? "");

                float fontSize;
                if (customSettings?.UseDynamicSize == true)
                {
                    fontSize = Math.Max(1, (float)_effect.FontSize.GetValue(frame, length, fps));
                }
                else
                {
                    fontSize = Math.Max(1, (float)(customSettings?.FontSize ?? _effect.FontSize.GetValue(frame, length, fps)));
                }

                var mediaColor = customSettings?.TextColor ?? _effect.TextColor;
                var color = new Color4(
                    Math.Max(0, Math.Min(1, mediaColor.R / 255f)),
                    Math.Max(0, Math.Min(1, mediaColor.G / 255f)),
                    Math.Max(0, Math.Min(1, mediaColor.B / 255f)),
                    Math.Max(0, Math.Min(1, mediaColor.A / 255f))
                );

                var bold = customSettings?.Bold ?? _effect.Bold;
                var italic = customSettings?.Italic ?? _effect.Italic;
                var letterSpacing = (float)_effect.LetterSpacing.GetValue(frame, length, fps);
                var textAlignment = _effect.TextAlignment;
                var displayText = _effect.DisplayText ?? "";

                bool needsUpdate = _currentFont != _lastUsedFont || Math.Abs(fontSize - _lastFontSize) > 0.1f ||
                                   !ColorsEqual(color, _lastColor) || displayText != _lastText ||
                                   bold != _lastBold || italic != _lastItalic ||
                                   Math.Abs(letterSpacing - _lastLetterSpacing) > 0.1f ||
                                   textAlignment != _lastTextAlignment;

                if (needsUpdate)
                {
                    if (!ColorsEqual(color, _lastColor))
                    {
                        UpdateTextBrush(color);
                    }
                    _lastUsedFont = _currentFont ?? "Arial";
                    _lastFontSize = fontSize;
                    _lastColor = color;
                    _lastText = displayText;
                    _lastBold = bold;
                    _lastItalic = italic;
                    _lastLetterSpacing = letterSpacing;
                    _lastTextAlignment = textAlignment;
                }

                RenderText(frame, length, fps, fontSize, bold, italic, textAlignment);
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "エフェクト更新処理");
                CreateEmptyCommandList();
            }

            return effectDescription.DrawDescription;
        }

        private void CreateEmptyCommandList()
        {
            try
            {
                _commandList?.Dispose();
                if (_devices?.DeviceContext != null)
                {
                    _commandList = _devices.DeviceContext.CreateCommandList();
                    var dc = _devices.DeviceContext;
                    var prevTarget = dc.Target;
                    dc.Target = _commandList;
                    dc.BeginDraw();
                    dc.Clear(new Color4(0, 0, 0, 0));
                    dc.EndDraw();
                    _commandList.Close();
                    if (prevTarget != null)
                        dc.Target = prevTarget;
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "空のコマンドリスト作成");
            }
        }

        private IDWriteTextFormat? GetOrCreateTextFormat(string fontName, float fontSize, bool bold, bool italic, TextAlignmentType textAlignment)
        {
            var key = $"{fontName}_{fontSize}_{bold}_{italic}_{textAlignment}";

            if (_textFormatCache.TryGetValue(key, out var existingFormat))
            {
                return existingFormat;
            }

            try
            {
                if (_dwriteFactory != null)
                {
                    IDWriteTextFormat? newFormat = null;
                    try
                    {
                        newFormat = _dwriteFactory.CreateTextFormat(
                            fontName ?? "Arial",
                            null,
                            bold ? FontWeight.Bold : FontWeight.Normal,
                            italic ? FontStyle.Italic : FontStyle.Normal,
                            FontStretch.Normal,
                            Math.Max(1, fontSize)
                        );
                    }
                    catch (Exception ex)
                    {
                        LogManager.WriteException(ex, $"フォント作成失敗（{fontName}）、フォールバック");
                        try
                        {
                            newFormat = _dwriteFactory.CreateTextFormat(
                                "Arial",
                                null,
                                bold ? FontWeight.Bold : FontWeight.Normal,
                                italic ? FontStyle.Italic : FontStyle.Normal,
                                FontStretch.Normal,
                                Math.Max(1, fontSize)
                            );
                        }
                        catch (Exception ex2)
                        {
                            LogManager.WriteException(ex2, "Arialフォールバック作成失敗");
                            return null;
                        }
                    }

                    if (newFormat != null)
                    {
                        newFormat.TextAlignment = textAlignment switch
                        {
                            TextAlignmentType.Left => TextAlignment.Leading,
                            TextAlignmentType.Right => TextAlignment.Trailing,
                            TextAlignmentType.Center => TextAlignment.Center,
                            _ => TextAlignment.Center
                        };
                        newFormat.ParagraphAlignment = ParagraphAlignment.Center;

                        if (_textFormatCache.Count > 50)
                        {
                            ClearOldTextFormats();
                        }

                        _textFormatCache.TryAdd(key, newFormat);
                        return newFormat;
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "テキストフォーマット作成");
            }

            return null;
        }

        private void ClearOldTextFormats()
        {
            try
            {
                var keysToRemove = _textFormatCache.Keys.Take(_textFormatCache.Count / 2).ToList();
                foreach (var key in keysToRemove)
                {
                    if (_textFormatCache.TryRemove(key, out var format))
                    {
                        format?.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "古いテキストフォーマットクリア");
            }
        }

        private void UpdateTextBrush(Color4 color)
        {
            try
            {
                if (_textBrush != null && _devices?.DeviceContext != null)
                {
                    _textBrush.Color = color;
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "テキストブラシ更新");
                try
                {
                    _textBrush?.Dispose();
                    if (_devices?.DeviceContext != null)
                        _textBrush = _devices.DeviceContext.CreateSolidColorBrush(color);
                }
                catch (Exception ex2)
                {
                    LogManager.WriteException(ex2, "テキストブラシ再作成");
                    _textBrush = null;
                }
            }
        }

        private void RenderText(int frame, int length, float fps, float fontSize, bool bold, bool italic, TextAlignmentType textAlignment)
        {
            try
            {
                _commandList?.Dispose();
                if (_devices?.DeviceContext != null)
                    _commandList = _devices.DeviceContext.CreateCommandList();

                var dc = _devices?.DeviceContext;
                var prevTarget = dc?.Target;

                if (dc != null && _commandList != null)
                {
                    try
                    {
                        dc.Target = _commandList;
                        dc.BeginDraw();
                        dc.Clear(new Color4(0, 0, 0, 0));

                        var displayText = _effect.DisplayText ?? "";
                        if (!string.IsNullOrEmpty(displayText) && _dwriteFactory != null && _textBrush != null)
                        {
                            var textFormat = GetOrCreateTextFormat(_currentFont ?? "Arial", fontSize, bold, italic, textAlignment);
                            if (textFormat != null)
                            {
                                using var textLayout = _dwriteFactory.CreateTextLayout(
                                    displayText,
                                    textFormat,
                                    _currentSize.Width,
                                    _currentSize.Height
                                );

                                var letterSpacing = (float)_effect.LetterSpacing.GetValue(frame, length, (int)fps);
                                if (Math.Abs(letterSpacing) > 0.001f)
                                {
                                    try
                                    {
                                        using var textLayout1 = textLayout.QueryInterfaceOrNull<IDWriteTextLayout1>();
                                        if (textLayout1 != null)
                                        {
                                            textLayout1.SetCharacterSpacing(letterSpacing, 0, 0, new TextRange(0, displayText.Length));
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        LogManager.WriteException(ex, "文字間隔設定");
                                    }
                                }

                                var origin = new Vector2(-_currentSize.Width / 2f, -_currentSize.Height / 2f);
                                dc.DrawTextLayout(origin, textLayout, _textBrush);
                            }
                        }
                        dc.EndDraw();
                        _commandList.Close();
                    }
                    catch (Exception ex)
                    {
                        LogManager.WriteException(ex, "テキスト描画");
                    }
                    finally
                    {
                        if (dc != null && prevTarget != null)
                            dc.Target = prevTarget;
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "テキスト描画処理");
            }
        }

        private static bool ColorsEqual(Color4 a, Color4 b)
        {
            return Math.Abs(a.R - b.R) < 0.001f && Math.Abs(a.G - b.G) < 0.001f &&
                   Math.Abs(a.B - b.B) < 0.001f && Math.Abs(a.A - b.A) < 0.001f;
        }

        public void SetInput(ID2D1Image? input)
        {
            _inputImage = input;
        }

        public void ClearInput()
        {
            _inputImage = null;
        }

        public void Dispose()
        {
            try
            {
                foreach (var format in _textFormatCache.Values)
                {
                    format?.Dispose();
                }
                _textFormatCache.Clear();

                _textBrush?.Dispose();
                _commandList?.Dispose();
                _dwriteFactory?.Dispose();
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "FontShuffleProcessor解放");
            }
            finally
            {
                _textBrush = null;
                _commandList = null;
                _dwriteFactory = null;
                _inputImage = null;
                _isInitialized = false;
            }
        }
    }
}