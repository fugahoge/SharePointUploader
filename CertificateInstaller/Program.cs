using System.Reflection;
using System.Security.Cryptography.X509Certificates;
using Microsoft.Extensions.Logging;
using Serilog;

namespace CertificateInstaller;

class Program
{
  static int Main(string[] args)
  {
    Config? config = null;
    ILogger<Program>? logger = null;

    try
    {
      // 設定ファイルを読み込む
      try
      {
        config = Config.Load();
      }
      catch (Exception ex)
      {
        Console.Error.WriteLine($"エラー: 設定ファイルの読み込みに失敗しました: {ex.Message}");
        return 1;
      }

      // ロガーの作成
      logger = CreateLogger(config.Log);

      logger.LogInformation("証明書インストーラーを開始します...");

      // 埋め込まれたPFXファイルを検索
      var assembly = Assembly.GetExecutingAssembly();
      var pfxResourceName = assembly.GetManifestResourceNames()
        .FirstOrDefault(name => name.EndsWith("certificate.pfx", StringComparison.OrdinalIgnoreCase));

      if (string.IsNullOrEmpty(pfxResourceName))
      {
        logger?.LogError("インストール情報が見つかりません。");
        return 1;
      }

      logger?.LogInformation("インストール情報を検出: {ResourceName}", pfxResourceName);

      // PFXファイルを読み込む
      byte[] pfxData;
      using (var stream = assembly.GetManifestResourceStream(pfxResourceName)!)
      {
        if (stream == null)
        {
          logger?.LogError("インストール情報を開けませんでした。");
          return 1;
        }

        using (var memoryStream = new MemoryStream())
        {
          stream.CopyTo(memoryStream);
          pfxData = memoryStream.ToArray();
        }
      }

      logger?.LogInformation("インストール情報サイズ: {Size} バイト", pfxData.Length);

      // パスワードの入力
      Console.Write("インストール情報のパスワードを入力してください（空の場合はEnterキー）: ");
      string? password = Console.ReadLine();
      if (string.IsNullOrWhiteSpace(password))
      {
        password = null;
      }

      // 証明書を読み込む
      X509Certificate2? certificate = null;
      try
      {
        certificate = new X509Certificate2(pfxData, password, X509KeyStorageFlags.Exportable);
        logger?.LogInformation("証明書を読み込みました。");
        logger?.LogInformation("  サブジェクト: {Subject}", certificate.Subject);
        logger?.LogInformation("  発行者: {Issuer}", certificate.Issuer);
        logger?.LogInformation("  有効期限: {NotAfter}", certificate.NotAfter.ToString("yyyy/MM/dd HH:mm:ss"));
      }
      catch (Exception ex)
      {
        logger?.LogError(ex, "インストール情報の読み込みに失敗しました");
        if (ex.Message.Contains("password") || ex.Message.Contains("パスワード"))
        {
          logger?.LogError("パスワードが正しくない可能性があります。");
        }
        return 1;
      }

      // キーストアに登録するストアを設定ファイルから取得
      StoreName storeName;
      if (!Enum.TryParse<StoreName>(config.StoreName, true, out storeName))
      {
        logger?.LogError("無効なStoreNameです: {StoreName}", config.StoreName);
        logger?.LogError("有効な値: My, Root, Trust, CertificateAuthority, etc.");
        return 1;
      }

      StoreLocation storeLocation;
      if (!Enum.TryParse<StoreLocation>(config.StoreLocation, true, out storeLocation))
      {
        logger?.LogError("無効なStoreLocationです: {StoreLocation}", config.StoreLocation);
        logger?.LogError("有効な値: CurrentUser, LocalMachine");
        return 1;
      }

      logger?.LogInformation("ストア名: {StoreName}", storeName);
      logger?.LogInformation("ストア位置: {StoreLocation}", storeLocation);
      if (storeLocation == StoreLocation.LocalMachine)
      {
        logger?.LogInformation("  (管理者権限が必要です)");
      }

      // 証明書をキーストアに追加
      try
      {
        using (var store = new X509Store(storeName, storeLocation))
        {
          store.Open(OpenFlags.ReadWrite);
          try
          {
            // 既存の証明書をチェック
            var existingCerts = store.Certificates.Find(
              X509FindType.FindByThumbprint,
              certificate.Thumbprint,
              false);

            if (existingCerts.Count > 0)
            {
              logger?.LogWarning("同じ証明書（サムプリント: {Thumbprint}）が既に存在します。", certificate.Thumbprint);
              Console.Write("上書きしますか？ (y/N): ");
              var response = Console.ReadLine();
              if (response?.ToLowerInvariant() != "y")
              {
                logger?.LogInformation("インストールをキャンセルしました。");
                return 0;
              }

              // 既存の証明書を削除
              foreach (var existingCert in existingCerts)
              {
                store.Remove(existingCert);
                existingCert.Dispose();
              }
              logger?.LogInformation("既存の証明書を削除しました。");
            }

            // 新しい証明書を追加
            store.Add(certificate);
            logger?.LogInformation("証明書をキーストアに正常に登録しました。");
          }
          finally
          {
            store.Close();
          }
        }
      }
      catch (UnauthorizedAccessException)
      {
        logger?.LogError("キーストアへのアクセス権限がありません。");
        logger?.LogError("LocalMachineストアを使用する場合は、管理者権限で実行してください。");
        return 1;
      }
      catch (Exception ex)
      {
        logger?.LogError(ex, "キーストアへの登録に失敗しました");
        return 1;
      }

      logger?.LogInformation("証明書のインストールが完了しました。");
      return 0;
    }
    catch (Exception ex)
    {
      logger?.LogError(ex, "予期しないエラーが発生しました");
      return 1;
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
    var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
    var logFileName = $"CertificateInstaller_{timestamp}.log";

    Log.Logger = new LoggerConfiguration()
      .MinimumLevel.Is(minimumLevel)
      .WriteTo.Console()
      .WriteTo.File(
        Path.Combine(logDirectory, logFileName),
        rollingInterval: RollingInterval.Infinite,
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
}
