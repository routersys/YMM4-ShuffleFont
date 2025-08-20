using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Media;
using System.Globalization;

namespace FontShuffle
{
    public static class FontIndexer
    {
        private static readonly string IndexFilePath;
        private static readonly string HiddenFontsFilePath;
        private static readonly string FontGroupsFilePath;
        private static readonly string CustomFontSettingsFilePath;
        private static readonly string FontsDirectory;
        private static FontIndexData? _cachedIndex;
        private static readonly object _cacheLock = new object();
        private static readonly SemaphoreSlim _fileLock = new SemaphoreSlim(1, 1);

        private const int CurrentIndexVersion = 2;
        private const int IndexValidityDays = 1;
        private const int ProcessingTimeoutSeconds = 60;
        private const int MaxParallelThreads = 4;

        static FontIndexer()
        {
            try
            {
                var assemblyLocation = Assembly.GetExecutingAssembly().Location;
                var pluginDirectory = !string.IsNullOrEmpty(assemblyLocation)
                    ? Path.GetDirectoryName(assemblyLocation)
                    : Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);

                var baseDirectory = pluginDirectory ?? Path.GetTempPath();
                IndexFilePath = Path.Combine(baseDirectory, "fontindex.json");
                HiddenFontsFilePath = Path.Combine(baseDirectory, "hidden.json");
                FontGroupsFilePath = Path.Combine(baseDirectory, "groups.json");
                CustomFontSettingsFilePath = Path.Combine(baseDirectory, "fontsettings.json");

                FontsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
                if (!Directory.Exists(FontsDirectory))
                {
                    FontsDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "FontIndexer初期化");
                var tempPath = Path.GetTempPath();
                IndexFilePath = Path.Combine(tempPath, "fontindex.json");
                HiddenFontsFilePath = Path.Combine(tempPath, "hidden.json");
                FontGroupsFilePath = Path.Combine(tempPath, "groups.json");
                CustomFontSettingsFilePath = Path.Combine(tempPath, "fontsettings.json");
                FontsDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
            }
        }

        public static bool HasValidIndex()
        {
            try
            {
                if (!File.Exists(IndexFilePath))
                    return false;

                using var fileStream = new FileStream(IndexFilePath, FileMode.Open, FileAccess.Read, FileShare.Read);
                var index = JsonSerializer.Deserialize<FontIndexData>(fileStream);

                if (index?.IsValid() != true || index.Version != CurrentIndexVersion)
                    return false;

                var daysSinceUpdate = (DateTime.Now - index.LastUpdated).TotalDays;
                if (daysSinceUpdate >= IndexValidityDays)
                {
                    LogManager.WriteLog($"インデックスが古すぎます（{daysSinceUpdate:F1}日前）");
                    return false;
                }

                if (HasFontsDirectoryChanged(index.FontsDirectoryLastModified))
                {
                    LogManager.WriteLog("フォントディレクトリに変更が検出されました");
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "インデックス有効性チェック");
                return false;
            }
        }

        private static bool HasFontsDirectoryChanged(DateTime lastRecorded)
        {
            try
            {
                if (!Directory.Exists(FontsDirectory))
                    return true;

                var directoryInfo = new DirectoryInfo(FontsDirectory);
                var currentModified = directoryInfo.LastWriteTime;

                return Math.Abs((currentModified - lastRecorded).TotalSeconds) > 1;
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "フォントディレクトリ変更チェック");
                return true;
            }
        }

        public static FontIndexData? LoadIndex()
        {
            lock (_cacheLock)
            {
                if (_cachedIndex != null)
                    return _cachedIndex;
            }

            try
            {
                if (!File.Exists(IndexFilePath))
                    return null;

                using var fileStream = new FileStream(IndexFilePath, FileMode.Open, FileAccess.Read, FileShare.Read);
                var index = JsonSerializer.Deserialize<FontIndexData>(fileStream);

                if (index?.IsValid() == true && index.Version == CurrentIndexVersion)
                {
                    index.Normalize();
                    lock (_cacheLock)
                    {
                        _cachedIndex = index;
                    }
                    LogManager.WriteLog($"フォントインデックスを読み込みました（フォント数: {index.AllFonts.Count}）");
                    return index;
                }

                LogManager.WriteLog($"インデックスバージョンが古いか無効（現在: {index?.Version}, 必要: {CurrentIndexVersion}）");
                return null;
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "インデックス読み込み");
                return null;
            }
        }

        public static async Task<FontIndexData> CreateIndexAsync(IProgress<FontIndexProgress>? progress = null, CancellationToken cancellationToken = default)
        {
            LogManager.WriteLog("フォントインデックス作成を開始します");

            using var timeoutCts = new CancellationTokenSource(TimeSpan.FromSeconds(ProcessingTimeoutSeconds));
            using var combinedCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, timeoutCts.Token);

            try
            {
                var index = new FontIndexData
                {
                    Version = CurrentIndexVersion,
                    LastUpdated = DateTime.Now,
                    FontsDirectoryLastModified = GetFontsDirectoryModified()
                };

                var systemFonts = Fonts.SystemFontFamilies.ToArray();
                var fontResults = new ConcurrentBag<(string name, bool isJapanese)>();
                var processed = 0;

                progress?.Report(new FontIndexProgress
                {
                    Current = 0,
                    Total = systemFonts.Length,
                    CurrentFont = "初期化中...",
                    IsCompleted = false
                });

                var parallelOptions = new ParallelOptions
                {
                    CancellationToken = combinedCts.Token,
                    MaxDegreeOfParallelism = Math.Min(MaxParallelThreads, Environment.ProcessorCount)
                };

                await Task.Run(() =>
                {
                    try
                    {
                        Parallel.ForEach(systemFonts, parallelOptions, fontFamily =>
                        {
                            combinedCts.Token.ThrowIfCancellationRequested();

                            try
                            {
                                var fontName = FontHelper.GetFontFamilyName(fontFamily);
                                if (!string.IsNullOrEmpty(fontName))
                                {
                                    var isJapanese = FontHelper.IsJapaneseFontAdvanced(fontFamily, fontName);
                                    fontResults.Add((fontName, isJapanese));

                                    var currentProcessed = Interlocked.Increment(ref processed);
                                    if (currentProcessed % 10 == 0 || currentProcessed == systemFonts.Length)
                                    {
                                        progress?.Report(new FontIndexProgress
                                        {
                                            Current = currentProcessed,
                                            Total = systemFonts.Length,
                                            CurrentFont = fontName,
                                            IsCompleted = false
                                        });
                                    }
                                }
                            }
                            catch (OperationCanceledException)
                            {
                                throw;
                            }
                            catch (Exception ex)
                            {
                                LogManager.WriteException(ex, $"フォント処理中にエラー");
                            }
                        });
                    }
                    catch (OperationCanceledException)
                    {
                        throw;
                    }
                    catch (Exception ex)
                    {
                        LogManager.WriteException(ex, "並列フォント処理");
                        throw;
                    }
                }, combinedCts.Token);

                combinedCts.Token.ThrowIfCancellationRequested();

                var uniqueFonts = fontResults
                    .GroupBy(f => f.name, StringComparer.OrdinalIgnoreCase)
                    .Select(g => g.First())
                    .OrderBy(f => f.name, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                foreach (var (name, isJapanese) in uniqueFonts)
                {
                    combinedCts.Token.ThrowIfCancellationRequested();

                    index.AllFonts.Add(name);
                    index.FontSupportMap[name] = isJapanese;

                    if (isJapanese)
                        index.JapaneseFonts.Add(name);
                    else
                        index.EnglishFonts.Add(name);
                }

                combinedCts.Token.ThrowIfCancellationRequested();

                await SaveIndexAsync(index);

                lock (_cacheLock)
                {
                    _cachedIndex = index;
                }

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
            catch (OperationCanceledException) when (timeoutCts.Token.IsCancellationRequested)
            {
                LogManager.WriteLog("フォントインデックス作成がタイムアウトしました", LogLevel.Warning);
                throw new TimeoutException("フォントインデックス作成がタイムアウトしました");
            }
            catch (OperationCanceledException)
            {
                LogManager.WriteLog("フォントインデックス作成がキャンセルされました");
                throw;
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "フォントインデックス作成");
                throw;
            }
        }

        private static DateTime GetFontsDirectoryModified()
        {
            try
            {
                if (Directory.Exists(FontsDirectory))
                {
                    return new DirectoryInfo(FontsDirectory).LastWriteTime;
                }
                return DateTime.MinValue;
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "フォントディレクトリ更新日時取得");
                return DateTime.MinValue;
            }
        }

        private static async Task SaveIndexAsync(FontIndexData index)
        {
            var timeoutCts = new CancellationTokenSource(TimeSpan.FromSeconds(10));

            try
            {
                await _fileLock.WaitAsync(timeoutCts.Token);

                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                };

                var json = JsonSerializer.Serialize(index, options);
                var tempFile = IndexFilePath + ".tmp";

                await File.WriteAllTextAsync(tempFile, json, timeoutCts.Token);

                if (File.Exists(IndexFilePath))
                {
                    File.Replace(tempFile, IndexFilePath, null);
                }
                else
                {
                    File.Move(tempFile, IndexFilePath);
                }

                LogManager.WriteLog("フォントインデックスを保存しました");
            }
            catch (OperationCanceledException)
            {
                LogManager.WriteLog("インデックス保存がタイムアウトしました", LogLevel.Warning);
                throw new TimeoutException("インデックス保存がタイムアウトしました");
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "インデックス保存");
                throw;
            }
            finally
            {
                _fileLock.Release();
                timeoutCts.Dispose();
            }
        }

        public static void ClearIndex()
        {
            try
            {
                lock (_cacheLock)
                {
                    _cachedIndex = null;
                }

                if (File.Exists(IndexFilePath))
                {
                    File.Delete(IndexFilePath);
                    LogManager.WriteLog("フォントインデックスを削除しました");
                }
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "インデックス削除");
            }
        }

        public static void InvalidateCache()
        {
            try
            {
                lock (_cacheLock)
                {
                    _cachedIndex = null;
                }
                LogManager.WriteLog("フォントインデックスキャッシュを無効化しました");
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "キャッシュ無効化");
            }
        }

        public static bool ShouldRecreateIndex()
        {
            try
            {
                if (!HasValidIndex())
                    return true;

                var index = LoadIndex();
                if (index == null)
                    return true;

                if (index.AllFonts.Count < 5)
                {
                    LogManager.WriteLog($"フォント数が少なすぎます（{index.AllFonts.Count}個）");
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "インデックス再作成判定");
                return true;
            }
        }

        public static async Task<List<string>> LoadHiddenFontsAsync()
        {
            var timeoutCts = new CancellationTokenSource(TimeSpan.FromSeconds(5));

            try
            {
                await _fileLock.WaitAsync(timeoutCts.Token);

                if (!File.Exists(HiddenFontsFilePath))
                {
                    return new List<string>();
                }

                using var fileStream = new FileStream(HiddenFontsFilePath, FileMode.Open, FileAccess.Read, FileShare.Read);
                var hiddenFonts = await JsonSerializer.DeserializeAsync<List<string>>(fileStream, cancellationToken: timeoutCts.Token);
                LogManager.WriteLog($"非表示フォント設定を読み込みました ({hiddenFonts?.Count ?? 0}件)");
                return hiddenFonts ?? new List<string>();
            }
            catch (OperationCanceledException)
            {
                LogManager.WriteLog("非表示フォント読み込みがタイムアウトしました", LogLevel.Warning);
                return new List<string>();
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "非表示フォント設定の読み込み");
                return new List<string>();
            }
            finally
            {
                _fileLock.Release();
                timeoutCts.Dispose();
            }
        }

        public static List<string> LoadHiddenFonts()
        {
            try
            {
                return LoadHiddenFontsAsync().GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "非表示フォント設定の読み込み（同期）");
                return new List<string>();
            }
        }

        public static async Task SaveHiddenFontsAsync(List<string> hiddenFonts)
        {
            var timeoutCts = new CancellationTokenSource(TimeSpan.FromSeconds(5));

            try
            {
                await _fileLock.WaitAsync(timeoutCts.Token);

                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                };

                var json = JsonSerializer.Serialize(hiddenFonts ?? new List<string>(), options);
                var tempFile = HiddenFontsFilePath + ".tmp";

                await File.WriteAllTextAsync(tempFile, json, timeoutCts.Token);

                if (File.Exists(HiddenFontsFilePath))
                {
                    File.Replace(tempFile, HiddenFontsFilePath, null);
                }
                else
                {
                    File.Move(tempFile, HiddenFontsFilePath);
                }

                LogManager.WriteLog($"非表示フォント設定を保存しました ({hiddenFonts?.Count ?? 0}件)");
            }
            catch (OperationCanceledException)
            {
                LogManager.WriteLog("非表示フォント保存がタイムアウトしました", LogLevel.Warning);
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "非表示フォント設定の保存");
            }
            finally
            {
                _fileLock.Release();
                timeoutCts.Dispose();
            }
        }

        public static void SaveHiddenFonts(List<string> hiddenFonts)
        {
            try
            {
                SaveHiddenFontsAsync(hiddenFonts).GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "非表示フォント設定の保存（同期）");
            }
        }

        public static async Task<List<FontGroup>> LoadFontGroupsAsync()
        {
            var timeoutCts = new CancellationTokenSource(TimeSpan.FromSeconds(5));

            try
            {
                await _fileLock.WaitAsync(timeoutCts.Token);

                if (!File.Exists(FontGroupsFilePath))
                {
                    return new List<FontGroup>();
                }

                using var fileStream = new FileStream(FontGroupsFilePath, FileMode.Open, FileAccess.Read, FileShare.Read);
                var fontGroups = await JsonSerializer.DeserializeAsync<List<FontGroup>>(fileStream, cancellationToken: timeoutCts.Token);
                LogManager.WriteLog($"フォントグループ設定を読み込みました ({fontGroups?.Count ?? 0}件)");
                return fontGroups ?? new List<FontGroup>();
            }
            catch (OperationCanceledException)
            {
                LogManager.WriteLog("フォントグループ読み込みがタイムアウトしました", LogLevel.Warning);
                return new List<FontGroup>();
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "フォントグループ設定の読み込み");
                return new List<FontGroup>();
            }
            finally
            {
                _fileLock.Release();
                timeoutCts.Dispose();
            }
        }

        public static List<FontGroup> LoadFontGroups()
        {
            try
            {
                return LoadFontGroupsAsync().GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "フォントグループ設定の読み込み（同期）");
                return new List<FontGroup>();
            }
        }

        public static async Task SaveFontGroupsAsync(List<FontGroup> fontGroups)
        {
            var timeoutCts = new CancellationTokenSource(TimeSpan.FromSeconds(5));

            try
            {
                await _fileLock.WaitAsync(timeoutCts.Token);

                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                };

                var json = JsonSerializer.Serialize(fontGroups ?? new List<FontGroup>(), options);
                var tempFile = FontGroupsFilePath + ".tmp";

                await File.WriteAllTextAsync(tempFile, json, timeoutCts.Token);

                if (File.Exists(FontGroupsFilePath))
                {
                    File.Replace(tempFile, FontGroupsFilePath, null);
                }
                else
                {
                    File.Move(tempFile, FontGroupsFilePath);
                }

                LogManager.WriteLog($"フォントグループ設定を保存しました ({fontGroups?.Count ?? 0}件)");
            }
            catch (OperationCanceledException)
            {
                LogManager.WriteLog("フォントグループ保存がタイムアウトしました", LogLevel.Warning);
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "フォントグループ設定の保存");
            }
            finally
            {
                _fileLock.Release();
                timeoutCts.Dispose();
            }
        }

        public static void SaveFontGroups(List<FontGroup> fontGroups)
        {
            try
            {
                SaveFontGroupsAsync(fontGroups).GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "フォントグループ設定の保存（同期）");
            }
        }

        public static async Task<Dictionary<string, FontCustomSettings>> LoadCustomFontSettingsAsync()
        {
            var timeoutCts = new CancellationTokenSource(TimeSpan.FromSeconds(5));

            try
            {
                await _fileLock.WaitAsync(timeoutCts.Token);

                if (!File.Exists(CustomFontSettingsFilePath))
                {
                    return new Dictionary<string, FontCustomSettings>();
                }

                using var fileStream = new FileStream(CustomFontSettingsFilePath, FileMode.Open, FileAccess.Read, FileShare.Read);
                var settings = await JsonSerializer.DeserializeAsync<Dictionary<string, FontCustomSettings>>(fileStream, cancellationToken: timeoutCts.Token);
                LogManager.WriteLog($"個別フォント設定を読み込みました ({settings?.Count ?? 0}件)");
                return settings ?? new Dictionary<string, FontCustomSettings>();
            }
            catch (OperationCanceledException)
            {
                LogManager.WriteLog("個別フォント設定の読み込みがタイムアウトしました", LogLevel.Warning);
                return new Dictionary<string, FontCustomSettings>();
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "個別フォント設定の読み込み");
                return new Dictionary<string, FontCustomSettings>();
            }
            finally
            {
                _fileLock.Release();
                timeoutCts.Dispose();
            }
        }

        public static async Task SaveCustomFontSettingsAsync(Dictionary<string, FontCustomSettings> settings)
        {
            var timeoutCts = new CancellationTokenSource(TimeSpan.FromSeconds(5));

            try
            {
                await _fileLock.WaitAsync(timeoutCts.Token);

                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                };

                var json = JsonSerializer.Serialize(settings ?? new Dictionary<string, FontCustomSettings>(), options);
                var tempFile = CustomFontSettingsFilePath + ".tmp";

                await File.WriteAllTextAsync(tempFile, json, timeoutCts.Token);

                if (File.Exists(CustomFontSettingsFilePath))
                {
                    File.Replace(tempFile, CustomFontSettingsFilePath, null);
                }
                else
                {
                    File.Move(tempFile, CustomFontSettingsFilePath);
                }

                LogManager.WriteLog($"個別フォント設定を保存しました ({settings?.Count ?? 0}件)");
            }
            catch (OperationCanceledException)
            {
                LogManager.WriteLog("個別フォント設定の保存がタイムアウトしました", LogLevel.Warning);
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "個別フォント設定の保存");
            }
            finally
            {
                _fileLock.Release();
                timeoutCts.Dispose();
            }
        }
    }

    public static class FontHelper
    {
        private static readonly string[] JapaneseKeywords = {
            "Gothic", "Mincho", "明朝", "ゴシック",
            "游", "Yu", "MS", "メイリオ", "Meiryo", "UD",
            "游ゴシック", "游明朝", "ヒラギノ", "Hiragino",
            "小塚", "Kozuka", "源ノ角", "Source Han", "Noto Sans CJK", "Noto Serif CJK",
            "BIZ", "UD デジタル", "HGS", "HG", "DFP", "DF", "VL",
            "IPAex", "IPA", "TakaoEx", "Takao", "さざなみ", "Sazanami",
            "M+", "Migu", "Rounded", "Kochi", "東風", "VLゴシック",
            "ＭＳ ゴシック", "ＭＳ 明朝", "ＭＳ Ｐゴシック", "ＭＳ Ｐ明朝",
            "Arial Unicode MS",
            "SimSun", "SimHei", "NSimSun", "FangSong", "KaiTi",
            "Malgun Gothic", "Batang", "Dotum", "Gulim",
            "Apple SD Gothic Neo", "Apple Color Emoji",
            "Segoe UI Historic", "Segoe UI Symbol"
        };

        private static readonly char[] JapaneseTestChars = { 'あ', 'い', 'う', 'ア', 'イ', 'ウ' };

        public static string GetFontFamilyName(FontFamily fontFamily)
        {
            try
            {
                if (fontFamily?.FamilyNames?.Count > 0)
                {
                    var japaneseName = fontFamily.FamilyNames
                        .Where(kvp => kvp.Key.IetfLanguageTag.StartsWith("ja", StringComparison.OrdinalIgnoreCase))
                        .Select(kvp => kvp.Value)
                        .FirstOrDefault();

                    if (!string.IsNullOrEmpty(japaneseName))
                        return japaneseName;

                    var englishName = fontFamily.FamilyNames
                        .Where(kvp => kvp.Key.IetfLanguageTag.StartsWith("en", StringComparison.OrdinalIgnoreCase))
                        .Select(kvp => kvp.Value)
                        .FirstOrDefault();

                    if (!string.IsNullOrEmpty(englishName))
                        return englishName;

                    return fontFamily.FamilyNames.Values.FirstOrDefault() ?? "";
                }
                return fontFamily?.Source ?? "Unknown Font";
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, "フォント名取得");
                return "Unknown Font";
            }
        }

        public static bool IsJapaneseFontAdvanced(FontFamily fontFamily, string fontName)
        {
            try
            {
                using var timeoutCts = new CancellationTokenSource(TimeSpan.FromMilliseconds(100));

                return IsJapaneseFontUnified(fontFamily, fontName);
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, $"日本語フォント判定（{fontName}）");
                return IsJapaneseFontByKeywordsOnly(fontName);
            }
        }

        public static bool IsJapaneseFontUnified(FontFamily? fontFamily, string fontName)
        {
            try
            {
                if (JapaneseKeywords.Any(keyword => fontName.Contains(keyword, StringComparison.OrdinalIgnoreCase)))
                    return true;

                if (fontFamily?.FamilyNames?.Any(kvp => kvp.Key.IetfLanguageTag.StartsWith("ja", StringComparison.OrdinalIgnoreCase)) == true)
                    return true;

                try
                {
                    if (fontFamily != null)
                    {
                        var typefaces = fontFamily.GetTypefaces();
                        if (typefaces?.Any() == true)
                        {
                            var firstTypeface = typefaces.FirstOrDefault();
                            if (firstTypeface != null)
                            {
                                try
                                {
                                    if (firstTypeface.TryGetGlyphTypeface(out var glyphTypeface))
                                    {
                                        try
                                        {
                                            var supportedCount = JapaneseTestChars.Count(c =>
                                                glyphTypeface.CharacterToGlyphMap.ContainsKey((int)c));

                                            return supportedCount >= 3;
                                        }
                                        finally
                                        {
                                            try
                                            {
                                                if (glyphTypeface is IDisposable disposable)
                                                {
                                                    disposable.Dispose();
                                                }
                                            }
                                            catch { }
                                        }
                                    }
                                }
                                catch { }
                            }
                        }
                    }
                }
                catch { }

                return false;
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, $"統一日本語フォント判定（{fontName}）");
                return IsJapaneseFontByKeywordsOnly(fontName);
            }
        }

        private static bool IsJapaneseFontByKeywordsOnly(string fontName)
        {
            try
            {
                var basicKeywords = new[] {
                    "Gothic", "Mincho", "明朝", "ゴシック", "游", "Yu", "MS",
                    "メイリオ", "Meiryo", "UD", "游ゴシック", "游明朝"
                };
                return basicKeywords.Any(keyword => fontName.Contains(keyword, StringComparison.OrdinalIgnoreCase));
            }
            catch (Exception ex)
            {
                LogManager.WriteException(ex, $"キーワードベース日本語フォント判定（{fontName}）");
                return false;
            }
        }
    }
}