using System;
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

namespace FontShuffle
{
    [VideoEffect("シャッフルフォント", new[] { "アニメーション" }, new[] { "shuffle font", "sf" }, IsAviUtlSupported = false)]
    public class FontShuffleEffect : VideoEffectBase
    {
        public override string Label => "フォントシャッフル";

        [Display(Name = "表示テキスト", Description = "表示するテキスト", GroupName = "テキスト設定")]
        [TextEditor]
        public string DisplayText { get => displayText; set => Set(ref displayText, value); }
        private string displayText = "サンプルテキスト";

        [Display(Name = "幅", Description = "描画領域の幅", GroupName = "テキスト設定")]
        [AnimationSlider("F0", "px", 1, 7680)]
        public Animation Width { get; } = new(1920, 1, 7680);

        [Display(Name = "高さ", Description = "描画領域の高さ", GroupName = "テキスト設定")]
        [AnimationSlider("F0", "px", 1, 4320)]
        public Animation Height { get; } = new(1080, 1, 4320);

        [Display(Name = "文字揃え", Description = "テキストの配置", GroupName = "テキスト設定")]
        [EnumComboBox]
        public TextAlignmentType TextAlignment { get => textAlignment; set => Set(ref textAlignment, value); }
        private TextAlignmentType textAlignment = TextAlignmentType.Center;

        [Display(Name = "文字間隔", Description = "文字間の間隔", GroupName = "テキスト設定")]
        [AnimationSlider("F1", "px", -10, 50)]
        public Animation LetterSpacing { get; } = new(0, -10, 50);

        [Display(Name = "フォント変更間隔", Description = "フォントを変更するフレーム間隔")]
        [AnimationSlider("F0", "フレーム", 1, 120)]
        public Animation Interval { get; } = new(30, 1, 600);

        [Display(Name = "シャッフルモード", Description = "フォントの選択方法")]
        [EnumComboBox]
        public ShuffleModeType ShuffleMode { get => shuffleMode; set => Set(ref shuffleMode, value); }
        private ShuffleModeType shuffleMode = ShuffleModeType.Auto;

        [Display(Name = "ランダムシード", Description = "ランダムモード用のシード値")]
        [AnimationSlider("F0", "", 1, 99999)]
        public Animation RandomSeed { get; } = new(12345, 1, 99999);

        [Display(Name = "フォントサイズ", Description = "テキストのサイズ（デフォルト）")]
        [AnimationSlider("F0", "px", 10, 200)]
        public Animation FontSize { get; } = new(48, 10, 500);

        [Display(Name = "文字色", Description = "テキストの色（デフォルト）")]
        [ColorPicker]
        public System.Windows.Media.Color TextColor { get => textColor; set => Set(ref textColor, value); }
        private System.Windows.Media.Color textColor = System.Windows.Media.Colors.White;

        [Display(Name = "太字", Description = "太字にする（デフォルト）")]
        [ToggleSlider]
        public bool Bold { get => bold; set => Set(ref bold, value); }
        private bool bold = false;

        [Display(Name = "イタリック", Description = "斜体にする（デフォルト）")]
        [ToggleSlider]
        public bool Italic { get => italic; set => Set(ref italic, value); }
        private bool italic = false;

        [Display(Name = "フォント管理", Description = "使用するフォントを選択", GroupName = "フォント設定")]
        [FontShuffleControl]
        public bool FontManagement { get; set; } = true;

        public string CurrentFont { get => currentFont; set => Set(ref currentFont, value); }
        private string currentFont = "Yu Gothic UI";

        public List<string> SelectedFonts { get => selectedFonts; set => Set(ref selectedFonts, value); }
        private List<string> selectedFonts = new();

        public List<string> FavoriteFonts { get => favoriteFonts; set => Set(ref favoriteFonts, value); }
        private List<string> favoriteFonts = new();

        public List<string> OrderedFonts { get => orderedFonts; set => Set(ref orderedFonts, value); }
        private List<string> orderedFonts = new();

        public List<string> HiddenFonts { get => hiddenFonts; set => Set(ref hiddenFonts, value); }
        private List<string> hiddenFonts = new();

        public Dictionary<string, FontCustomSettings> FontCustomSettings { get => fontCustomSettings; set => Set(ref fontCustomSettings, value); }
        private Dictionary<string, FontCustomSettings> fontCustomSettings = new();

        public string IgnoredVersion { get => ignoredVersion; set => Set(ref ignoredVersion, value); }
        private string ignoredVersion = "";

        private static readonly Lazy<FontIndexData> _systemFonts = new(LoadSystemFontsWithIndex);

        public override IVideoEffectProcessor CreateVideoEffect(IGraphicsDevicesAndContext devices)
        {
            return new FontShuffleProcessor(devices, this);
        }

        public override IEnumerable<string> CreateExoVideoFilters(int keyFrameIndex, ExoOutputDescription exoOutputDescription)
        {
            return Array.Empty<string>();
        }

        protected override IEnumerable<IAnimatable> GetAnimatables() => new IAnimatable[] { Width, Height, Interval, RandomSeed, FontSize, LetterSpacing };

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
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "フォント選択処理");
            }
            return "Yu Gothic UI";
        }

        public List<string> GetActiveFontList()
        {
            try
            {
                var systemFonts = _systemFonts.Value;
                var hiddenSet = new HashSet<string>(HiddenFonts);

                var filterFonts = (List<string> fonts) => fonts.Where(f => !hiddenSet.Contains(f)).ToList();

                return ShuffleMode switch
                {
                    ShuffleModeType.Ordered when OrderedFonts.Count > 0 => filterFonts(OrderedFonts),
                    ShuffleModeType.Selected when SelectedFonts.Count > 0 => filterFonts(SelectedFonts),
                    ShuffleModeType.Favorites when FavoriteFonts.Count > 0 => filterFonts(FavoriteFonts),
                    ShuffleModeType.Japanese => filterFonts(systemFonts.JapaneseFonts.Count > 0 ? systemFonts.JapaneseFonts : systemFonts.AllFonts),
                    ShuffleModeType.English => filterFonts(systemFonts.EnglishFonts.Count > 0 ? systemFonts.EnglishFonts : systemFonts.AllFonts),
                    ShuffleModeType.Selected when SelectedFonts.Count == 0 => new List<string>(),
                    ShuffleModeType.Favorites when FavoriteFonts.Count == 0 => new List<string>(),
                    ShuffleModeType.Ordered when OrderedFonts.Count == 0 => new List<string>(),
                    _ => filterFonts(systemFonts.AllFonts)
                };
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "アクティブフォントリスト取得");
                return new List<string> { "Yu Gothic UI" };
            }
        }

        public FontCustomSettings? GetFontSettings(string fontName)
        {
            try
            {
                if (FontCustomSettings.ContainsKey(fontName) && FontCustomSettings[fontName].UseCustomSettings)
                {
                    return FontCustomSettings[fontName];
                }
                return null;
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, $"フォント設定取得（{fontName}）");
                return null;
            }
        }

        private static FontIndexData LoadSystemFontsWithIndex()
        {
            try
            {
                LogManager.WriteLog("システムフォント読み込み開始");

                var existingIndex = FontIndexer.LoadIndex();
                if (existingIndex != null)
                {
                    LogManager.WriteLog($"既存インデックスを使用（フォント数: {existingIndex.AllFonts.Count}）");
                    return existingIndex;
                }

                LogManager.WriteLog("新しいインデックスを作成します");
                var newIndex = CreateBasicFontIndex();
                LogManager.WriteLog($"基本インデックス作成完了（フォント数: {newIndex.AllFonts.Count}）");
                return newIndex;
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "システムフォント読み込み");
                return CreateFallbackFontIndex();
            }
        }

        private static FontIndexData CreateBasicFontIndex()
        {
            var index = new FontIndexData();
            var fontSet = new HashSet<string>();

            try
            {
                foreach (var fontFamily in System.Windows.Media.Fonts.SystemFontFamilies)
                {
                    try
                    {
                        string fontName = GetFontFamilyName(fontFamily);
                        if (!string.IsNullOrEmpty(fontName) && fontSet.Add(fontName))
                        {
                            index.AllFonts.Add(fontName);

                            if (IsJapaneseFontBasic(fontName))
                                index.JapaneseFonts.Add(fontName);
                            else
                                index.EnglishFonts.Add(fontName);
                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.WriteException(ex, "個別フォント処理");
                    }
                }

                index.AllFonts.Sort();
                index.JapaneseFonts.Sort();
                index.EnglishFonts.Sort();
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "基本フォントインデックス作成");
                return CreateFallbackFontIndex();
            }

            return index;
        }

        private static FontIndexData CreateFallbackFontIndex()
        {
            LogManager.WriteLog("フォールバックフォントインデックスを作成", LogLevel.Warning);

            return new FontIndexData
            {
                AllFonts = new List<string> { "Yu Gothic UI", "Meiryo", "MS Gothic", "Arial", "Times New Roman", "Calibri", "Verdana" },
                JapaneseFonts = new List<string> { "Yu Gothic UI", "Meiryo", "MS Gothic" },
                EnglishFonts = new List<string> { "Arial", "Times New Roman", "Calibri", "Verdana" }
            };
        }

        private static string GetFontFamilyName(System.Windows.Media.FontFamily fontFamily)
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

        private static bool IsJapaneseFontBasic(string fontName)
        {
            try
            {
                var japaneseKeywords = new[] {
                    "Gothic", "Mincho", "游", "Yu", "MS", "メイリオ", "Meiryo", "UD",
                    "游ゴシック", "游明朝", "ヒラギノ", "小塚", "源ノ角", "Noto"
                };
                return japaneseKeywords.Any(keyword => fontName.Contains(keyword, StringComparison.OrdinalIgnoreCase));
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, $"日本語フォント判定（{fontName}）");
                return false;
            }
        }
    }

    public class FontShuffleProcessor : IVideoEffectProcessor
    {
        private readonly IGraphicsDevicesAndContext devices;
        private readonly FontShuffleEffect effect;

        private int lastFrame = -1;
        private string? currentFont;

        private IDWriteTextFormat? currentTextFormat;
        private ID2D1SolidColorBrush? textBrush;
        private IDWriteFactory? dwriteFactory;
        private ID2D1CommandList? commandList;
        private ID2D1Image? inputImage;

        private string lastUsedFont = "";
        private float lastFontSize = 0;
        private Color4 lastColor = new(0, 0, 0, 0);
        private string lastText = "";
        private bool lastBold = false;
        private bool lastItalic = false;
        private float lastLetterSpacing = 0;
        private TextAlignmentType lastTextAlignment = TextAlignmentType.Center;
        private bool isInitialized = false;
        private SizeI currentSize = new(1920, 1080);

        public ID2D1Image? Output => commandList;

        public FontShuffleProcessor(IGraphicsDevicesAndContext devices, FontShuffleEffect effect)
        {
            this.devices = devices;
            this.effect = effect;
            Initialize();
        }

        private void Initialize()
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
                        LogManager.WriteLog("FontShuffleProcessor初期化完了");
                    }
                    else
                    {
                        LogManager.WriteLog("DWriteファクトリ作成に失敗", LogLevel.Error);
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
                var letterSpacing = (float)effect.LetterSpacing.GetValue(frame, length, fps);
                var textAlignment = effect.TextAlignment;

                bool needsUpdate = currentFont != lastUsedFont || Math.Abs(fontSize - lastFontSize) > 0.1f ||
                                   !ColorsEqual(color, lastColor) || effect.DisplayText != lastText ||
                                   bold != lastBold || italic != lastItalic ||
                                   Math.Abs(letterSpacing - lastLetterSpacing) > 0.1f ||
                                   textAlignment != lastTextAlignment;

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
                    lastLetterSpacing = letterSpacing;
                    lastTextAlignment = textAlignment;
                }

                RenderText(frame, length, fps);
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "エフェクト更新処理");
            }

            return effectDescription.DrawDescription;
        }

        private void UpdateTextFormat(string fontName, float fontSize, bool bold, bool italic)
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
                    catch (Exception ex)
                    {
                        LogManager.WriteException(ex, $"フォント作成失敗（{fontName}）、フォールバック");
                        currentTextFormat = dwriteFactory.CreateTextFormat("Arial", null,
                            bold ? FontWeight.Bold : FontWeight.Normal,
                            italic ? FontStyle.Italic : FontStyle.Normal,
                            FontStretch.Normal, fontSize);
                    }
                    if (currentTextFormat != null)
                    {
                        currentTextFormat.TextAlignment = TextAlignment.Leading;
                        currentTextFormat.ParagraphAlignment = ParagraphAlignment.Near;
                    }
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "テキストフォーマット更新");
                currentTextFormat = null;
            }
        }

        private void UpdateTextBrush(Color4 color)
        {
            try
            {
                if (textBrush != null && devices?.DeviceContext != null)
                    textBrush.Color = color;
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "テキストブラシ更新");
                try
                {
                    textBrush?.Dispose();
                    if (devices?.DeviceContext != null)
                        textBrush = devices.DeviceContext.CreateSolidColorBrush(color);
                }
                catch (Exception ex2)
                {
                    LogManager.WriteException(ex2, "テキストブラシ再作成");
                    textBrush = null;
                }
            }
        }

        private void RenderText(int frame, int length, float fps)
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
                                float.MaxValue,
                                float.MaxValue
                            );

                            var letterSpacing = (float)effect.LetterSpacing.GetValue(frame, length, (int)fps);
                            if (Math.Abs(letterSpacing) > 0.001f)
                            {
                                using var textLayout1 = textLayout.QueryInterfaceOrNull<IDWriteTextLayout1>();
                                if (textLayout1 != null)
                                {
                                    textLayout1.SetCharacterSpacing(letterSpacing, 0, 0, new TextRange(0, effect.DisplayText.Length));
                                }
                            }

                            var textMetrics = textLayout.Metrics;

                            float x, y;

                            y = -textMetrics.Height / 2f - textMetrics.Top;

                            switch (effect.TextAlignment)
                            {
                                case TextAlignmentType.Left:
                                    x = -currentSize.Width / 2f - textMetrics.Left;
                                    break;
                                case TextAlignmentType.Right:
                                    x = currentSize.Width / 2f - textMetrics.Left - textMetrics.Width;
                                    break;
                                case TextAlignmentType.Center:
                                default:
                                    x = -textMetrics.Width / 2f - textMetrics.Left;
                                    break;
                            }

                            var origin = new Vector2(x, y);
                            dc.DrawTextLayout(origin, textLayout, textBrush);
                        }
                        dc.EndDraw();
                        commandList.Close();
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
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "FontShuffleProcessor解放");
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