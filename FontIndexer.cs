using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Media;
using System.Globalization;

namespace FontShuffle
{
    public static class FontIndexer
    {
        private static readonly string IndexFilePath;
        private static FontIndexData? _cachedIndex;
        private const int CurrentIndexVersion = 1;

        static FontIndexer()
        {
            try
            {
                var assemblyLocation = Assembly.GetExecutingAssembly().Location;
                var pluginDirectory = Path.GetDirectoryName(assemblyLocation) ?? Path.GetTempPath();
                IndexFilePath = Path.Combine(pluginDirectory, "fontindex.json");
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "FontIndexer初期化");
                IndexFilePath = Path.Combine(Path.GetTempPath(), "fontindex.json");
            }
        }

        public static bool HasValidIndex()
        {
            try
            {
                if (!File.Exists(IndexFilePath))
                    return false;

                var json = File.ReadAllText(IndexFilePath);
                var index = JsonSerializer.Deserialize<FontIndexData>(json);

                if (index == null || index.Version != CurrentIndexVersion)
                    return false;

                var daysSinceUpdate = (DateTime.Now - index.LastUpdated).TotalDays;
                return daysSinceUpdate < 7;
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "インデックス有効性チェック");
                return false;
            }
        }

        public static FontIndexData? LoadIndex()
        {
            try
            {
                if (_cachedIndex != null)
                    return _cachedIndex;

                if (!File.Exists(IndexFilePath))
                    return null;

                var json = File.ReadAllText(IndexFilePath);
                var index = JsonSerializer.Deserialize<FontIndexData>(json);

                if (index?.Version == CurrentIndexVersion)
                {
                    _cachedIndex = index;
                    LogManager.WriteLog($"フォントインデックスを読み込みました（フォント数: {index.AllFonts.Count}）");
                    return index;
                }

                return null;
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "インデックス読み込み");
                return null;
            }
        }

        public static async Task<FontIndexData> CreateIndexAsync(IProgress<FontIndexProgress> progress = null)
        {
            LogManager.WriteLog("フォントインデックス作成を開始します");

            return await Task.Run(() =>
            {
                try
                {
                    var index = new FontIndexData
                    {
                        Version = CurrentIndexVersion,
                        LastUpdated = DateTime.Now
                    };

                    var systemFonts = Fonts.SystemFontFamilies.ToArray();
                    var fontSet = new HashSet<string>();
                    var processed = 0;

                    foreach (var fontFamily in systemFonts)
                    {
                        try
                        {
                            var fontName = GetFontFamilyName(fontFamily);
                            if (!string.IsNullOrEmpty(fontName) && fontSet.Add(fontName))
                            {
                                index.AllFonts.Add(fontName);

                                var isJapanese = IsJapaneseFontAdvanced(fontFamily, fontName);
                                index.FontSupportMap[fontName] = isJapanese;

                                if (isJapanese)
                                    index.JapaneseFonts.Add(fontName);
                                else
                                    index.EnglishFonts.Add(fontName);

                                progress?.Report(new FontIndexProgress
                                {
                                    Current = processed + 1,
                                    Total = systemFonts.Length,
                                    CurrentFont = fontName,
                                    IsCompleted = false
                                });
                            }
                        }
                        catch (Exception ex)
                        {
                            LogManager.WriteException(ex, $"フォント処理中にエラー（フォント#{processed}）");
                        }

                        processed++;
                    }

                    index.AllFonts.Sort();
                    index.JapaneseFonts.Sort();
                    index.EnglishFonts.Sort();

                    SaveIndex(index);
                    _cachedIndex = index;

                    progress?.Report(new FontIndexProgress
                    {
                        Current = systemFonts.Length,
                        Total = systemFonts.Length,
                        CurrentFont = "",
                        IsCompleted = true
                    });

                    LogManager.WriteLog($"フォントインデックス作成完了（総数: {index.AllFonts.Count}, 日本語: {index.JapaneseFonts.Count}, 英語: {index.EnglishFonts.Count}）");

                    return index;
                }
                catch (Exception ex)
                {
                    LogManager.WriteException(ex, "フォントインデックス作成");
                    throw;
                }
            });
        }

        private static void SaveIndex(FontIndexData index)
        {
            try
            {
                var json = JsonSerializer.Serialize(index, new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                });

                File.WriteAllText(IndexFilePath, json);
                LogManager.WriteLog("フォントインデックスを保存しました");
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "インデックス保存");
            }
        }

        private static string GetFontFamilyName(FontFamily fontFamily)
        {
            try
            {
                if (fontFamily.FamilyNames.Count > 0)
                {
                    var japaneseName = fontFamily.FamilyNames
                        .Where(kvp => kvp.Key.IetfLanguageTag.StartsWith("ja"))
                        .Select(kvp => kvp.Value)
                        .FirstOrDefault();

                    if (!string.IsNullOrEmpty(japaneseName))
                        return japaneseName;

                    var englishName = fontFamily.FamilyNames
                        .Where(kvp => kvp.Key.IetfLanguageTag.StartsWith("en"))
                        .Select(kvp => kvp.Value)
                        .FirstOrDefault();

                    if (!string.IsNullOrEmpty(englishName))
                        return englishName;

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

        private static bool IsJapaneseFontAdvanced(FontFamily fontFamily, string fontName)
        {
            try
            {
                var japaneseKeywords = new[] {
                    "Gothic", "Mincho", "游", "Yu", "MS", "メイリオ", "Meiryo", "UD",
                    "游ゴシック", "游明朝", "ヒラギノ", "小塚", "源ノ角", "Noto", "BIZ",
                    "UD デジタル", "HGS", "HG", "DFP", "DF", "VL", "IPAex", "IPA",
                    "TakaoEx", "Takao", "さざなみ", "Sazanami", "M+", "Migu",
                    "Rounded", "Kochi", "東風", "VLゴシック", "ＭＳ ゴシック", "ＭＳ 明朝"
                };

                if (japaneseKeywords.Any(keyword => fontName.Contains(keyword, StringComparison.OrdinalIgnoreCase)))
                    return true;

                if (fontFamily.FamilyNames.Any(kvp => kvp.Key.IetfLanguageTag.StartsWith("ja")))
                    return true;

                try
                {
                    var typeface = fontFamily.GetTypefaces().FirstOrDefault();
                    if (typeface != null)
                    {
                        if (typeface.TryGetGlyphTypeface(out var glyphTypeface))
                        {
                            var hiraganaA = 'あ';
                            var codepoint = (int)hiraganaA;
                            if (glyphTypeface.CharacterToGlyphMap.ContainsKey(codepoint))
                                return true;
                        }
                    }
                }
                catch
                {
                }

                return false;
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, $"日本語フォント判定（{fontName}）");
                return fontName.Contains("Gothic") || fontName.Contains("Mincho") || fontName.Contains("MS");
            }
        }

        public static void ClearIndex()
        {
            try
            {
                if (File.Exists(IndexFilePath))
                {
                    File.Delete(IndexFilePath);
                    LogManager.WriteLog("フォントインデックスを削除しました");
                }
                _cachedIndex = null;
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "インデックス削除");
            }
        }
    }
}