using System;
using System.Collections.Concurrent;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace FontShuffle
{
    public static class LogManager
    {
        private static readonly string LogDirectory;
        private static readonly object LogLock = new object();
        private static readonly ConcurrentQueue<LogEntry> LogQueue = new();
        private static Timer? LogTimer;
        private const int MaxLogFiles = 5;
        private const int LogBatchSize = 20;
        private const int FlushIntervalSeconds = 10;

        private static volatile bool isShuttingDown = false;
        private static volatile bool isInitialized = false;
        private static readonly SemaphoreSlim initializationSemaphore = new(1, 1);

        static LogManager()
        {
            try
            {
                var assemblyLocation = Assembly.GetExecutingAssembly().Location;
                var pluginDirectory = !string.IsNullOrEmpty(assemblyLocation)
                    ? Path.GetDirectoryName(assemblyLocation)
                    : Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);

                LogDirectory = Path.Combine(pluginDirectory ?? Path.GetTempPath(), "log");

                Task.Run(InitializeAsync);
            }
            catch (Exception ex)
            {
                try
                {
                    LogDirectory = Path.Combine(Path.GetTempPath(), "log");
                    Task.Run(InitializeAsync);
                }
                catch
                {
                    LogDirectory = Path.GetTempPath();
                }
            }
        }

        private static async Task InitializeAsync()
        {
            try
            {
                await initializationSemaphore.WaitAsync();

                if (isInitialized)
                    return;

                await Task.Run(() =>
                {
                    try
                    {
                        if (!Directory.Exists(LogDirectory))
                            Directory.CreateDirectory(LogDirectory);

                        CleanupOldLogs();
                    }
                    catch (Exception ex)
                    {
                        try
                        {
                            var fallbackMessage = $"ログ初期化エラー: {ex.Message}";
                            File.AppendAllText(Path.Combine(Path.GetTempPath(), "FontShuffle_InitError.log"),
                                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {fallbackMessage}\n");
                        }
                        catch { }
                    }
                });

                LogTimer = new Timer(FlushLogQueue, null,
                    TimeSpan.FromSeconds(FlushIntervalSeconds),
                    TimeSpan.FromSeconds(FlushIntervalSeconds));

                AppDomain.CurrentDomain.ProcessExit += OnProcessExit;
                isInitialized = true;
            }
            catch (Exception ex)
            {
                try
                {
                    var fallbackMessage = $"ログマネージャー非同期初期化エラー: {ex.Message}";
                    File.AppendAllText(Path.Combine(Path.GetTempPath(), "FontShuffle_InitError.log"),
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {fallbackMessage}\n");
                }
                catch { }
            }
            finally
            {
                initializationSemaphore.Release();
            }
        }

        private static void OnProcessExit(object? sender, EventArgs e)
        {
            isShuttingDown = true;
            FlushLogQueue(null);
            LogTimer?.Dispose();
            initializationSemaphore?.Dispose();
        }

        public static void WriteLog(string message, LogLevel level = LogLevel.Info, Exception? exception = null)
        {
            if (isShuttingDown) return;

            try
            {
                var entry = new LogEntry
                {
                    Timestamp = DateTime.Now,
                    Level = level,
                    Message = message ?? string.Empty,
                    Exception = exception
                };

                LogQueue.Enqueue(entry);

                if (level == LogLevel.Critical)
                {
                    Task.Run(() => FlushLogQueue(null));
                    _ = Task.Run(() => ShowCriticalError(message, exception));
                }
                else if (LogQueue.Count >= LogBatchSize)
                {
                    Task.Run(() => FlushLogQueue(null));
                }
            }
            catch (Exception ex)
            {
                try
                {
                    var fallbackMessage = $"ログエラー: {ex.Message}\n元のメッセージ: {message}";
                    File.AppendAllText(Path.Combine(LogDirectory ?? Path.GetTempPath(), "emergency.log"),
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {fallbackMessage}\n");
                }
                catch { }
            }
        }

        private static void FlushLogQueue(object? state)
        {
            if (LogQueue.IsEmpty || !isInitialized) return;

            var entries = new List<LogEntry>();
            while (entries.Count < LogBatchSize && LogQueue.TryDequeue(out var entry))
            {
                entries.Add(entry);
            }

            if (entries.Count == 0) return;

            try
            {
                var logFileName = $"{DateTime.Now:yyyy-MM-dd_HH}.log";
                var logFilePath = Path.Combine(LogDirectory, logFileName);

                using var fileStream = new FileStream(logFilePath, FileMode.Append, FileAccess.Write, FileShare.Read);
                using var writer = new StreamWriter(fileStream);

                foreach (var entry in entries)
                {
                    var levelText = entry.Level switch
                    {
                        LogLevel.Info => "情報",
                        LogLevel.Warning => "警告",
                        LogLevel.Error => "エラー",
                        LogLevel.Critical => "重大",
                        _ => "不明"
                    };

                    var logEntry = $"[{entry.Timestamp:yyyy-MM-dd HH:mm:ss.fff}] [{levelText}] {entry.Message}";
                    if (entry.Exception != null)
                    {
                        logEntry += $"\n例外: {entry.Exception.Message}";
                        if (!string.IsNullOrEmpty(entry.Exception.StackTrace))
                        {
                            logEntry += $"\nスタックトレース: {entry.Exception.StackTrace}";
                        }
                    }
                    writer.WriteLine(logEntry);
                }
            }
            catch (Exception ex)
            {
                try
                {
                    var fallbackPath = Path.Combine(Path.GetTempPath(), "FontShuffle_Emergency.log");
                    File.AppendAllText(fallbackPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] ログ書き込みエラー: {ex.Message}\n");
                }
                catch { }
            }
        }

        public static void WriteException(Exception exception, string context = "")
        {
            var message = string.IsNullOrEmpty(context) ?
                $"例外が発生しました: {exception?.Message}" :
                $"{context}で例外が発生しました: {exception?.Message}";
            WriteLog(message, LogLevel.Error, exception);
        }

        private static async Task ShowCriticalError(string message, Exception? exception)
        {
            try
            {
                await Task.Delay(100);

                if (Application.Current?.Dispatcher != null)
                {
                    await Application.Current.Dispatcher.BeginInvoke(new Action(() =>
                    {
                        try
                        {
                            var errorMessage = $"重大なエラーが発生しました:\n\n{message}";
                            if (exception != null)
                            {
                                errorMessage += $"\n\n詳細: {exception.Message}";
                            }
                            MessageBox.Show(errorMessage, "FontShuffle 重大エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                        catch { }
                    }));
                }
                else
                {
                    MessageBox.Show($"重大なエラーが発生しました: {message}", "FontShuffle 重大エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch { }
        }

        private static void CleanupOldLogs()
        {
            try
            {
                if (!Directory.Exists(LogDirectory))
                    return;

                var logFiles = Directory.GetFiles(LogDirectory, "*.log")
                    .Select(f => new FileInfo(f))
                    .Where(f => f.Exists)
                    .OrderByDescending(f => f.CreationTime)
                    .ToArray();

                if (logFiles.Length > MaxLogFiles)
                {
                    var filesToDelete = logFiles.Skip(MaxLogFiles).ToArray();

                    Task.Run(() =>
                    {
                        try
                        {
                            Parallel.ForEach(filesToDelete, new ParallelOptions
                            {
                                MaxDegreeOfParallelism = 2
                            }, file =>
                            {
                                try
                                {
                                    if (file.Exists)
                                    {
                                        file.Delete();
                                    }
                                }
                                catch (Exception ex)
                                {
                                    WriteLog($"古いログファイルの削除に失敗: {file.Name} - {ex.Message}", LogLevel.Warning);
                                }
                            });
                        }
                        catch (Exception ex)
                        {
                            WriteLog($"ログファイルのクリーンアップに失敗: {ex.Message}", LogLevel.Warning, ex);
                        }
                    });
                }
            }
            catch (Exception ex)
            {
                try
                {
                    var fallbackPath = Path.Combine(Path.GetTempPath(), "FontShuffle_CleanupError.log");
                    File.AppendAllText(fallbackPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] クリーンアップエラー: {ex.Message}\n");
                }
                catch { }
            }
        }

        public static string GetLogDirectory()
        {
            return LogDirectory ?? Path.GetTempPath();
        }

        public static void OpenLogDirectory()
        {
            try
            {
                Task.Run(() =>
                {
                    try
                    {
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = LogDirectory,
                            UseShellExecute = true
                        });
                    }
                    catch (Exception ex)
                    {
                        WriteLog($"ログディレクトリを開けませんでした: {ex.Message}", LogLevel.Warning, ex);
                    }
                });
            }
            catch (Exception ex)
            {
                WriteLog($"ログディレクトリを開く処理でエラー: {ex.Message}", LogLevel.Warning, ex);
            }
        }

        public static void Shutdown()
        {
            isShuttingDown = true;
            FlushLogQueue(null);
            LogTimer?.Dispose();
            initializationSemaphore?.Dispose();
        }
    }

    public class LogEntry
    {
        public DateTime Timestamp { get; set; }
        public LogLevel Level { get; set; }
        public string Message { get; set; } = string.Empty;
        public Exception? Exception { get; set; }
    }

    public enum LogLevel
    {
        Info,
        Warning,
        Error,
        Critical
    }
}