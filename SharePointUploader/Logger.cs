using System;
using System.IO;
using System.Linq;

namespace SharePointUploader;

public class Logger
{
  private readonly string _logDirectory;
  private readonly string _logFilePath;
  private readonly object _lockObject = new object();

  public Logger()
  {
    _logDirectory = Path.Combine(Directory.GetCurrentDirectory(), "Logs");
    Directory.CreateDirectory(_logDirectory);

    // 実行ごとに新しいログファイルを作成
    var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
    _logFilePath = Path.Combine(_logDirectory, $"SharePointUploader_{timestamp}.log");

    // 古いログファイルを削除（最新10件を保持）
    RotateLogs();
  }

  private void RotateLogs()
  {
    try
    {
      var logFiles = Directory.GetFiles(_logDirectory, "SharePointUploader_*.log")
        .Select(f => new FileInfo(f))
        .OrderByDescending(f => f.CreationTime)
        .ToList();

      // 最新10件を保持し、それ以外を削除
      if (logFiles.Count > 10)
      {
        foreach (var file in logFiles.Skip(10))
        {
          try
          {
            file.Delete();
          }
          catch
          {
            // 削除に失敗しても続行
          }
        }
      }
    }
    catch
    {
      // ローテーションに失敗しても続行
    }
  }

  public void LogInfo(string message)
  {
    Log("INFO", message);
  }

  public void LogError(string message)
  {
    Log("ERROR", message);
  }

  public void LogWarning(string message)
  {
    Log("WARNING", message);
  }

  private void Log(string level, string message)
  {
    var logMessage = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{level}] {message}";

    // コンソールに出力
    Console.WriteLine(logMessage);

    // ログファイルに出力
    lock (_lockObject)
    {
      try
      {
        File.AppendAllText(_logFilePath, logMessage + Environment.NewLine);
      }
      catch
      {
        // ログファイルへの書き込みに失敗しても続行
      }
    }
  }
}
