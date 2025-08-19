using System;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Globalization;
using System.Linq;

namespace FontShuffle
{
    public static class LogManager
    {
        private static readonly string LogDirectory;
        private static readonly object LogLock = new object();
        private const int MaxLogFiles = 10;

        static LogManager()
        {
            try
            {
                var assemblyLocation = Assembly.GetExecutingAssembly().Location;
                var pluginDirectory = Path.GetDirectoryName(assemblyLocation) ?? Path.GetTempPath();
                LogDirectory = Path.Combine(pluginDirectory, "log");

                if (!Directory.Exists(LogDirectory))
                    Directory.CreateDirectory(LogDirectory);

                CleanupOldLogs();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ログシステムの初期化に失敗しました: {ex.Message}", "FontShuffle エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
                LogDirectory = Path.Combine(Path.GetTempPath(), "FontShuffle_Log");
                Directory.CreateDirectory(LogDirectory);
            }
        }

        public static void WriteLog(string message, LogLevel level = LogLevel.Info, Exception exception = null)
        {
            try
            {
                lock (LogLock)
                {
                    var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture);
                    var logFileName = $"{DateTime.Now:yyyy-MM-dd_HH}.log";
                    var logFilePath = Path.Combine(LogDirectory, logFileName);

                    var levelText = level switch
                    {
                        LogLevel.Info => "情報",
                        LogLevel.Warning => "警告",
                        LogLevel.Error => "エラー",
                        LogLevel.Critical => "重大",
                        _ => "不明"
                    };

                    var logEntry = $"[{timestamp}] [{levelText}] {message}";
                    if (exception != null)
                    {
                        logEntry += $"\n例外: {exception.Message}\nスタックトレース: {exception.StackTrace}";
                    }
                    logEntry += "\n";

                    File.AppendAllText(logFilePath, logEntry);

                    if (level == LogLevel.Critical)
                    {
                        ShowCriticalError(message, exception);
                    }
                }
            }
            catch (Exception ex)
            {
                try
                {
                    MessageBox.Show($"ログ書き込みエラー: {ex.Message}\n元のメッセージ: {message}",
                        "FontShuffle ログエラー", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                catch
                {
                    // 最終的なフォールバック - 何もしない
                }
            }
        }

        public static void WriteException(Exception exception, string context = "")
        {
            var message = string.IsNullOrEmpty(context) ?
                $"例外が発生しました: {exception.Message}" :
                $"{context}で例外が発生しました: {exception.Message}";
            WriteLog(message, LogLevel.Error, exception);
        }

        private static void ShowCriticalError(string message, Exception exception)
        {
            try
            {
                Application.Current?.Dispatcher.BeginInvoke(new Action(() =>
                {
                    var errorMessage = $"重大なエラーが発生しました:\n\n{message}";
                    if (exception != null)
                    {
                        errorMessage += $"\n\n詳細: {exception.Message}";
                    }
                    MessageBox.Show(errorMessage, "FontShuffle 重大エラー", MessageBoxButton.OK, MessageBoxImage.Error);
                }));
            }
            catch
            {
                // フォールバック
                MessageBox.Show($"重大なエラーが発生しました: {message}", "FontShuffle 重大エラー", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private static void CleanupOldLogs()
        {
            try
            {
                var logFiles = Directory.GetFiles(LogDirectory, "*.log")
                    .Select(f => new FileInfo(f))
                    .OrderByDescending(f => f.CreationTime)
                    .ToArray();

                if (logFiles.Length > MaxLogFiles)
                {
                    for (int i = MaxLogFiles; i < logFiles.Length; i++)
                    {
                        try
                        {
                            logFiles[i].Delete();
                        }
                        catch (Exception ex)
                        {
                            WriteLog($"古いログファイルの削除に失敗: {logFiles[i].Name} - {ex.Message}", LogLevel.Warning);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLog($"ログファイルのクリーンアップに失敗: {ex.Message}", LogLevel.Warning, ex);
            }
        }

        public static string GetLogDirectory()
        {
            return LogDirectory;
        }

        public static void OpenLogDirectory()
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
        }
    }

    public enum LogLevel
    {
        Info,
        Warning,
        Error,
        Critical
    }
}