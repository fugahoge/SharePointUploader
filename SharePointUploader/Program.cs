using System;
using System.IO;
using Microsoft.Extensions.Logging;
using Serilog;

namespace SharePointUploader;

class Program
{
  static async Task Main(string[] args)
  {
    // ログディレクトリの設定
    var logDirectory = Path.Combine(Directory.GetCurrentDirectory(), "Logs");
    Directory.CreateDirectory(logDirectory);

    // Serilogの設定
    Log.Logger = new LoggerConfiguration()
      .MinimumLevel.Information()
      .WriteTo.Console()
      .WriteTo.File(
        Path.Combine(logDirectory, "SharePointUploader_.log"),
        rollingInterval: RollingInterval.Day,
        retainedFileCountLimit: 10,
        outputTemplate: "[{Timestamp:yyyy-MM-dd HH:mm:ss}] [{Level:u3}] {Message:lj}{NewLine}{Exception}")
      .CreateLogger();

    // ILoggerFactoryの作成
    using var loggerFactory = LoggerFactory.Create(builder =>
    {
      builder.AddSerilog();
    });

    var logger = loggerFactory.CreateLogger<Program>();

    try
    {
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

      // 設定ファイルの読み込み
      logger.LogInformation("設定ファイル読み込み...");
      var config = Config.Load();

      // 設定の検証
      ValidateConfig(config.SharePoint, logger);

      // Graph API クライアントの作成
      var client = new GraphApiClient(
        config.SharePoint.TenantId,
        config.SharePoint.ClientId,
        config.SharePoint.CertificatePath,
        config.SharePoint.CertificatePassword
      );

      logger.LogInformation($"アップロード開始: {fileName}");

      try
      {
        // ファイルをアップロード
        var webUrl = await client.UploadFileAsync(
          config.SharePoint.SiteUrl,
          config.SharePoint.LibraryName,
          config.SharePoint.FolderPath,
          fileName,
          logger
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

  private static void ValidateConfig(SharePointConfig config, ILogger logger)
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
