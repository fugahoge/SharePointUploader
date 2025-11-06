using System;
using System.IO;
using Microsoft.Extensions.Logging;
using Serilog;

namespace SharePointUploader;

class Program
{
  static async Task Main(string[] args)
  {
    var config = Config.Load();
    var logger = CreateLogger(config.Log);

    try
    {
      // 設定の検証
      ValidateConfig(config.SharePoint, logger);

      // 引数の確認
      if (args.Length == 0)
      {
        logger.LogError("ファイル名が指定されていません");
        logger.LogInformation("使用方法: SharePointUploader.exe <ファイル名>");
        Environment.Exit(1);
        return;
      }

      var fileName = args[0];

      // ファイルの存在確認
      if (!File.Exists(fileName))
      {
        logger.LogError($"ファイルが見つかりません: {fileName}");
        Environment.Exit(1);
        return;
      }

      // Graph API クライアントの作成
      var client = new GraphApiClient(
        config.SharePoint.TenantId,
        config.SharePoint.ClientId,
        config.SharePoint.CertificatePath,
        config.SharePoint.CertificatePassword,
        logger
      );

      logger.LogInformation($"アップロード開始: {fileName}");

      try
      {
        // ファイルをアップロード
        var webUrl = await client.UploadFileAsync(
          config.SharePoint.SiteUrl,
          config.SharePoint.LibraryName,
          config.SharePoint.FolderPath,
          fileName
        );

        logger.LogInformation("アップロードが正常に完了しました");
        logger.LogInformation($"ファイルURL: {webUrl}");
      }
      finally
      {
        client.Dispose();
      }
    }
    catch (Exception ex)
    {
      logger.LogError(ex, "エラーが発生しました");
      Environment.Exit(1);
    }
    finally
    {
      // Serilogのクリーンアップ
      Log.CloseAndFlush();
    }
  }

  private static ILogger<Program> CreateLogger(LogConfig? logConfig)
  {
    // ログディレクトリの設定
    var logDirectory = Path.Combine(Directory.GetCurrentDirectory(), "Logs");
    Directory.CreateDirectory(logDirectory);

    // ログ設定の取得（デフォルト値を使用）
    var logLevel = logConfig?.Level ?? "Information";
    var retainedFileCountLimit = logConfig?.RetainedFileCountLimit ?? 10;

    // ログレベルのパース
    var minimumLevel = logLevel switch
    {
      "Verbose" => Serilog.Events.LogEventLevel.Verbose,
      "Debug" => Serilog.Events.LogEventLevel.Debug,
      "Information" => Serilog.Events.LogEventLevel.Information,
      "Warning" => Serilog.Events.LogEventLevel.Warning,
      "Error" => Serilog.Events.LogEventLevel.Error,
      "Fatal" => Serilog.Events.LogEventLevel.Fatal,
      _ => Serilog.Events.LogEventLevel.Information
    };

    // Loggerの作成
    Log.Logger = new LoggerConfiguration()
      .MinimumLevel.Is(minimumLevel)
      .WriteTo.Console()
      .WriteTo.File(
        Path.Combine(logDirectory, "SharePointUploader_.log"),
        rollingInterval: RollingInterval.Day,
        retainedFileCountLimit: retainedFileCountLimit,
        outputTemplate: "[{Timestamp:yyyy-MM-dd HH:mm:ss}] [{Level:u3}] {Message:lj}{NewLine}{Exception}")
      .CreateLogger();

    // ILoggerFactoryの作成
    using var loggerFactory = LoggerFactory.Create(builder =>
    {
      builder.AddSerilog();
    });

    return loggerFactory.CreateLogger<Program>();
  }

  private static void ValidateConfig(SharePointConfig config, Microsoft.Extensions.Logging.ILogger logger)
  {
    if (string.IsNullOrWhiteSpace(config.SiteUrl))
    {
      throw new Exception("設定エラー: SiteUrlが指定されていません");
    }

    if (string.IsNullOrWhiteSpace(config.LibraryName))
    {
      throw new Exception("設定エラー: LibraryNameが指定されていません");
    }

    if (string.IsNullOrWhiteSpace(config.TenantId))
    {
      throw new Exception("設定エラー: TenantIdが指定されていません");
    }

    if (string.IsNullOrWhiteSpace(config.ClientId))
    {
      throw new Exception("設定エラー: ClientIdが指定されていません");
    }

    if (string.IsNullOrWhiteSpace(config.CertificatePath))
    {
      throw new Exception("設定エラー: CertificatePathが指定されていません");
    }

    if (!File.Exists(config.CertificatePath))
    {
      throw new Exception($"設定エラー: 証明書ファイルが見つかりません: {config.CertificatePath}");
    }

    logger.LogInformation("設定の検証が完了しました");
  }
}
