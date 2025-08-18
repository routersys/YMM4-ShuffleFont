using System.ComponentModel.DataAnnotations;
using System.Numerics;
using Vortice;
using Vortice.DCommon;
using Vortice.Direct2D1;
using Vortice.DirectWrite;
using Vortice.Mathematics;
using YukkuriMovieMaker.Commons;
using YukkuriMovieMaker.Controls;
using YukkuriMovieMaker.Exo;
using YukkuriMovieMaker.Player.Video;
using YukkuriMovieMaker.Plugin.Effects;

namespace FontShuffle
{
    [VideoEffect("シャッフルフォント", ["アニメーション"], ["shuffle font", "sf"], IsAviUtlSupported = false)]
    public class FontShuffleEffect : VideoEffectBase
    {
        public override string Label => "フォントシャッフル";

        [Display(Name = "表示テキスト", Description = "表示するテキスト")]
        [TextEditor]
        public string DisplayText { get => displayText; set => Set(ref displayText, value); }
        string displayText = "サンプルテキスト";

        [Display(Name = "幅", Description = "描画領域の幅")]
        [AnimationSlider("F0", "px", 1, 7680)]
        public Animation Width { get; } = new Animation(1920, 1, 7680);

        [Display(Name = "高さ", Description = "描画領域の高さ")]
        [AnimationSlider("F0", "px", 1, 4320)]
        public Animation Height { get; } = new Animation(1080, 1, 4320);

        [Display(Name = "フォント変更間隔", Description = "フォントを変更するフレーム間隔")]
        [AnimationSlider("F0", "フレーム", 1, 120)]
        public Animation Interval { get; } = new Animation(30, 1, 600);

        [Display(Name = "シャッフルモード", Description = "フォントの選択方法")]
        [EnumComboBox]
        public ShuffleModeType ShuffleMode { get => shuffleMode; set => Set(ref shuffleMode, value); }
        ShuffleModeType shuffleMode = ShuffleModeType.Auto;

        [Display(Name = "ランダムシード", Description = "ランダムモード用のシード値")]
        [AnimationSlider("F0", "", 1, 99999)]
        public Animation RandomSeed { get; } = new Animation(12345, 1, 99999);

        [Display(Name = "フォントサイズ", Description = "テキストのサイズ（デフォルト）")]
        [AnimationSlider("F0", "px", 10, 200)]
        public Animation FontSize { get; } = new Animation(48, 10, 500);

        [Display(Name = "文字色", Description = "テキストの色（デフォルト）")]
        [ColorPicker]
        public System.Windows.Media.Color TextColor { get => textColor; set => Set(ref textColor, value); }
        System.Windows.Media.Color textColor = System.Windows.Media.Colors.White;

        [Display(Name = "太字", Description = "太字にする（デフォルト）")]
        [ToggleSlider]
        public bool Bold { get => bold; set => Set(ref bold, value); }
        bool bold = false;

        [Display(Name = "イタリック", Description = "斜体にする（デフォルト）")]
        [ToggleSlider]
        public bool Italic { get => italic; set => Set(ref italic, value); }
        bool italic = false;

        [Display(Name = "フォント管理", Description = "使用するフォントを選択", GroupName = "フォント設定")]
        [FontShuffleControl]
        public bool FontManagement { get; set; } = true;

        public string CurrentFont { get => currentFont; set => Set(ref currentFont, value); }
        string currentFont = "Yu Gothic UI";

        public List<string> SelectedFonts { get => selectedFonts; set => Set(ref selectedFonts, value); }
        List<string> selectedFonts = new();

        public List<string> FavoriteFonts { get => favoriteFonts; set => Set(ref favoriteFonts, value); }
        List<string> favoriteFonts = new();

        public List<string> OrderedFonts { get => orderedFonts; set => Set(ref orderedFonts, value); }
        List<string> orderedFonts = new();

        public Dictionary<string, FontCustomSettings> FontCustomSettings { get => fontCustomSettings; set => Set(ref fontCustomSettings, value); }
        Dictionary<string, FontCustomSettings> fontCustomSettings = new();

        public string IgnoredVersion { get => ignoredVersion; set => Set(ref ignoredVersion, value); }
        string ignoredVersion = "";

        static readonly Lazy<(List<string> AllFonts, List<string> JapaneseFonts, List<string> EnglishFonts)> _systemFonts = new(LoadSystemFonts);

        public override IVideoEffectProcessor CreateVideoEffect(IGraphicsDevicesAndContext devices)
        {
            return new FontShuffleProcessor(devices, this);
        }

        public override IEnumerable<string> CreateExoVideoFilters(int keyFrameIndex, ExoOutputDescription exoOutputDescription)
        {
            return [];
        }

        protected override IEnumerable<IAnimatable> GetAnimatables() => [Width, Height, Interval, RandomSeed, FontSize];

        public string GetFontForFrame(int frame)
        {
            try
            {
                var interval = Math.Max(1, (int)Interval.GetValue(frame, 1, 30));
                var currentInterval = frame / interval;
                var seed = (int)RandomSeed.GetValue(frame, 1, 30);

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
            catch
            {
            }
            return "Yu Gothic UI";
        }

        public List<string> GetActiveFontList()
        {
            var systemFonts = _systemFonts.Value;

            return ShuffleMode switch
            {
                ShuffleModeType.Ordered when OrderedFonts.Count > 0 => OrderedFonts,
                ShuffleModeType.Selected when SelectedFonts.Count > 0 => SelectedFonts,
                ShuffleModeType.Favorites when FavoriteFonts.Count > 0 => FavoriteFonts,
                ShuffleModeType.Japanese => systemFonts.JapaneseFonts.Count > 0 ? systemFonts.JapaneseFonts : systemFonts.AllFonts,
                ShuffleModeType.English => systemFonts.EnglishFonts.Count > 0 ? systemFonts.EnglishFonts : systemFonts.AllFonts,
                ShuffleModeType.Selected when SelectedFonts.Count == 0 => new List<string>(),
                ShuffleModeType.Favorites when FavoriteFonts.Count == 0 => new List<string>(),
                ShuffleModeType.Ordered when OrderedFonts.Count == 0 => new List<string>(),
                _ => systemFonts.AllFonts
            };
        }

        public FontCustomSettings? GetFontSettings(string fontName)
        {
            if (FontCustomSettings.ContainsKey(fontName) && FontCustomSettings[fontName].UseCustomSettings)
            {
                return FontCustomSettings[fontName];
            }
            return null;
        }

        static (List<string> AllFonts, List<string> JapaneseFonts, List<string> EnglishFonts) LoadSystemFonts()
        {
            var allFonts = new List<string>();
            var japaneseFonts = new List<string>();
            var englishFonts = new List<string>();

            try
            {
                var fontSet = new HashSet<string>();
                foreach (var fontFamily in System.Windows.Media.Fonts.SystemFontFamilies)
                {
                    string fontName = GetFontFamilyName(fontFamily);
                    if (!string.IsNullOrEmpty(fontName) && fontSet.Add(fontName))
                    {
                        allFonts.Add(fontName);

                        if (IsJapaneseFont(fontName))
                            japaneseFonts.Add(fontName);
                        else
                            englishFonts.Add(fontName);
                    }
                }

                allFonts.Sort();
                japaneseFonts.Sort();
                englishFonts.Sort();
            }
            catch
            {
                allFonts.AddRange(["Yu Gothic UI", "Meiryo", "MS Gothic", "Arial", "Times New Roman", "Calibri", "Verdana"]);
                japaneseFonts.AddRange(["Yu Gothic UI", "Meiryo", "MS Gothic"]);
                englishFonts.AddRange(["Arial", "Times New Roman", "Calibri", "Verdana"]);
            }

            return (allFonts, japaneseFonts, englishFonts);
        }

        static string GetFontFamilyName(System.Windows.Media.FontFamily fontFamily)
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

    public class FontShuffleProcessor : IVideoEffectProcessor
    {
        readonly IGraphicsDevicesAndContext devices;
        readonly FontShuffleEffect effect;

        int lastFrame = -1;
        string? currentFont;

        IDWriteTextFormat? currentTextFormat;
        ID2D1SolidColorBrush? textBrush;
        IDWriteFactory? dwriteFactory;
        ID2D1CommandList? commandList;
        ID2D1Image? inputImage;

        string lastUsedFont = "";
        float lastFontSize = 0;
        Color4 lastColor = new(0, 0, 0, 0);
        string lastText = "";
        bool lastBold = false;
        bool lastItalic = false;
        bool isInitialized = false;
        SizeI currentSize = new(1920, 1080);

        public ID2D1Image? Output => commandList;

        public FontShuffleProcessor(IGraphicsDevicesAndContext devices, FontShuffleEffect effect)
        {
            this.devices = devices;
            this.effect = effect;
            Initialize();
        }

        void Initialize()
        {
            try
            {
                if (devices?.DeviceContext != null)
                {
                    textBrush = devices.DeviceContext.CreateSolidColorBrush(new Color4(1f, 1f, 1f, 1f));
                    var result = DWrite.DWriteCreateFactory(Vortice.DirectWrite.FactoryType.Shared, out dwriteFactory);
                    if (result.Success)
                    {
                        isInitialized = true;
                    }
                }
            }
            catch
            {
                isInitialized = false;
            }
        }

        public DrawDescription Update(EffectDescription effectDescription)
        {
            if (!isInitialized || devices?.DeviceContext == null)
                return effectDescription.DrawDescription;

            try
            {
                var frame = effectDescription.ItemPosition.Frame;
                var length = effectDescription.ItemDuration.Frame;
                var fps = effectDescription.FPS;

                var width = (int)effect.Width.GetValue(frame, length, fps);
                var height = (int)effect.Height.GetValue(frame, length, fps);
                currentSize = new SizeI(width, height);

                var interval = Math.Max(1, (int)effect.Interval.GetValue(frame, length, fps));
                var currentInterval = frame / interval;

                if (lastFrame != currentInterval)
                {
                    currentFont = effect.GetFontForFrame(frame);
                    lastFrame = currentInterval;
                }

                var fonts = effect.GetActiveFontList();
                if (fonts.Count == 0)
                {
                    commandList?.Dispose();
                    commandList = devices.DeviceContext.CreateCommandList();
                    var dc = devices.DeviceContext;
                    var prevTarget = dc.Target;
                    dc.Target = commandList;
                    dc.BeginDraw();
                    dc.Clear(new Color4(0, 0, 0, 0));
                    dc.EndDraw();
                    commandList.Close();
                    if (prevTarget != null)
                        dc.Target = prevTarget;
                    return effectDescription.DrawDescription;
                }

                var customSettings = effect.GetFontSettings(currentFont ?? "");

                float fontSize;
                if (customSettings?.UseDynamicSize == true)
                {
                    fontSize = (float)effect.FontSize.GetValue(frame, length, fps);
                }
                else
                {
                    fontSize = (float)(customSettings?.FontSize ?? effect.FontSize.GetValue(frame, length, fps));
                }

                var mediaColor = customSettings?.TextColor ?? effect.TextColor;
                var color = new Color4(mediaColor.R / 255f, mediaColor.G / 255f, mediaColor.B / 255f, mediaColor.A / 255f);
                var bold = customSettings?.Bold ?? effect.Bold;
                var italic = customSettings?.Italic ?? effect.Italic;

                bool needsUpdate = currentFont != lastUsedFont || Math.Abs(fontSize - lastFontSize) > 0.1f ||
                                   !ColorsEqual(color, lastColor) || effect.DisplayText != lastText ||
                                   bold != lastBold || italic != lastItalic;

                if (needsUpdate)
                {
                    if (currentFont != lastUsedFont || Math.Abs(fontSize - lastFontSize) > 0.1f ||
                        bold != lastBold || italic != lastItalic)
                    {
                        UpdateTextFormat(currentFont ?? "Yu Gothic UI", fontSize, bold, italic);
                    }
                    if (!ColorsEqual(color, lastColor))
                    {
                        UpdateTextBrush(color);
                    }
                    lastUsedFont = currentFont ?? "Yu Gothic UI";
                    lastFontSize = fontSize;
                    lastColor = color;
                    lastText = effect.DisplayText;
                    lastBold = bold;
                    lastItalic = italic;
                }

                RenderText();
            }
            catch
            {
            }

            return effectDescription.DrawDescription;
        }

        void UpdateTextFormat(string fontName, float fontSize, bool bold, bool italic)
        {
            try
            {
                currentTextFormat?.Dispose();
                currentTextFormat = null;
                if (dwriteFactory != null && devices?.DeviceContext != null)
                {
                    try
                    {
                        currentTextFormat = dwriteFactory.CreateTextFormat(fontName, null,
                            bold ? FontWeight.Bold : FontWeight.Normal,
                            italic ? FontStyle.Italic : FontStyle.Normal,
                            FontStretch.Normal, fontSize);
                    }
                    catch
                    {
                        currentTextFormat = dwriteFactory.CreateTextFormat("Arial", null,
                            bold ? FontWeight.Bold : FontWeight.Normal,
                            italic ? FontStyle.Italic : FontStyle.Normal,
                            FontStretch.Normal, fontSize);
                    }
                    if (currentTextFormat != null)
                    {
                        currentTextFormat.TextAlignment = TextAlignment.Center;
                        currentTextFormat.ParagraphAlignment = ParagraphAlignment.Center;
                    }
                }
            }
            catch
            {
                currentTextFormat = null;
            }
        }

        void UpdateTextBrush(Color4 color)
        {
            try
            {
                if (textBrush != null && devices?.DeviceContext != null)
                    textBrush.Color = color;
            }
            catch
            {
                try
                {
                    textBrush?.Dispose();
                    if (devices?.DeviceContext != null)
                        textBrush = devices.DeviceContext.CreateSolidColorBrush(color);
                }
                catch
                {
                    textBrush = null;
                }
            }
        }

        void RenderText()
        {
            try
            {
                commandList?.Dispose();
                if (devices?.DeviceContext != null)
                    commandList = devices.DeviceContext.CreateCommandList();

                var dc = devices?.DeviceContext;
                var prevTarget = dc?.Target;

                if (dc != null && commandList != null)
                {
                    try
                    {
                        dc.Target = commandList;
                        dc.BeginDraw();
                        dc.Clear(new Color4(0, 0, 0, 0));

                        if (!string.IsNullOrEmpty(effect.DisplayText) && dwriteFactory != null &&
                            currentTextFormat != null && textBrush != null)
                        {
                            using var textLayout = dwriteFactory.CreateTextLayout(
                                effect.DisplayText,
                                currentTextFormat,
                                currentSize.Width,
                                currentSize.Height
                            );

                            var origin = new Vector2(-currentSize.Width / 2f, -currentSize.Height / 2f);
                            dc.DrawTextLayout(origin, textLayout, textBrush);
                        }
                        dc.EndDraw();
                        commandList.Close();
                    }
                    catch
                    {
                    }
                    finally
                    {
                        if (dc != null && prevTarget != null)
                            dc.Target = prevTarget;
                    }
                }
            }
            catch
            {
            }
        }

        static bool ColorsEqual(Color4 a, Color4 b)
        {
            return Math.Abs(a.R - b.R) < 0.001f && Math.Abs(a.G - b.G) < 0.001f &&
                   Math.Abs(a.B - b.B) < 0.001f && Math.Abs(a.A - b.A) < 0.001f;
        }

        public void SetInput(ID2D1Image? input)
        {
            inputImage = input;
        }

        public void ClearInput()
        {
            inputImage = null;
        }

        public void Dispose()
        {
            try
            {
                currentTextFormat?.Dispose();
                textBrush?.Dispose();
                commandList?.Dispose();
                dwriteFactory?.Dispose();
            }
            catch
            {
            }
            currentTextFormat = null;
            textBrush = null;
            commandList = null;
            dwriteFactory = null;
            inputImage = null;
            isInitialized = false;
        }
    }
}