using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
using Azure.Core;
using Azure.Identity;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;

namespace SharePointUploader;

public class GraphApiClient : IDisposable
{
  private readonly GraphServiceClient _graphClient;
  private readonly ILogger _logger;

  public GraphApiClient(SharePointConfig config, ILogger logger)
  {
    _logger = logger;
    
    // スコープを設定
    var scopes = new[] { 
      "https://graph.microsoft.com/Files.ReadWrite.All",
      "https://graph.microsoft.com/Sites.ReadWrite.All",
      "https://graph.microsoft.com/User.Read",
      "https://graph.microsoft.com/.default"
    };

    // 証明書を読み込む
    X509Certificate2 certificate;
    try
    {
      if (!string.IsNullOrWhiteSpace(config.CertificatePath))
      {
        // 証明書を読み込む（ファイル）
        _logger.LogInformation($"証明書を読み込みます: {config.CertificatePath}");
        if (string.IsNullOrEmpty(config.CertificatePassword))
        {
          certificate = new X509Certificate2(config.CertificatePath);
        }
        else
        {
          certificate = new X509Certificate2(config.CertificatePath, config.CertificatePassword);
        }
      }
      else
      {
        // 証明書を読み込む（Windowsキーストア）
        _logger.LogInformation("証明書を読み込みます");
        certificate = LoadCertificateFromStore(config.Thumbprint, config.StoreName, config.StoreLocation);
      }

      // 秘密鍵が含まれていることを確認
      if (!certificate.HasPrivateKey)
      {
        certificate.Dispose();
        throw new InvalidOperationException(
          $"証明書に秘密鍵が含まれていません。");
      }
    }
    catch (Exception ex)
    {
      throw new Exception($"証明書の読み込みに失敗しました: {ex.Message}", ex);
    }

    // 証明書認証（アプリケーション認証）
    var certificateCredential = new ClientCertificateCredential(
      config.TenantId,
      config.ClientId,
      certificate
    );

    // ユーザー認証（Interactive Browser認証）
    // 初回のみブラウザでログインし、次回以降はキャッシュされたトークンを使用
    var interactiveCredential = new InteractiveBrowserCredential(
      new InteractiveBrowserCredentialOptions
      {
        // 初回ログイン: ブラウザでログインし、アクセストークンとリフレッシュトークンを取得
        // 2回目以降: キャッシュからリフレッシュトークンを読み込み、アクセストークンを自動更新
        // リフレッシュトークンの有効期限経過後は再度ブラウザでログインが必要（通常は90日経過後）
        
        TenantId = config.TenantId,
        ClientId = config.ClientId,
        RedirectUri = new Uri("http://localhost"),

        // トークンキャッシュを有効化（Windows Credential Managerへ保存）
        TokenCachePersistenceOptions = new TokenCachePersistenceOptions
        {
          Name = "SharePointUploaderTokenCache"
        }
      }
    );

    // 証明書でアプリケーション認証を行い、ユーザーがログインしてそのユーザーの権限でアクセス
    var credential = new ChainedTokenCredential(
      certificateCredential,
      interactiveCredential
    );

    _logger.LogInformation("証明書認証とユーザー認証を組み合わせて認証を行います。");
    _logger.LogInformation("初回のみブラウザでログインが必要です。次回以降はキャッシュされたトークンを使用します。");

    // GraphServiceClientの作成
    _graphClient = new GraphServiceClient(credential, scopes);
  }

  /// <summary>
  /// Windowsキーストアから証明書を読み込む
  /// </summary>
  private X509Certificate2 LoadCertificateFromStore(string thumbprint, string storeName, string storeLocation)
  {
    // StoreNameをパース
    if (!Enum.TryParse<StoreName>(storeName, true, out var parsedStoreName))
    {
      throw new ArgumentException($"無効なStoreNameです: {storeName}");
    }

    // StoreLocationをパース
    if (!Enum.TryParse<StoreLocation>(storeLocation, true, out var parsedStoreLocation))
    {
      throw new ArgumentException($"無効なStoreLocationです: {storeLocation}");
    }

    _logger.LogInformation($"キーストア: {parsedStoreName}, 場所: {parsedStoreLocation}, サムプリント: {thumbprint}");

    // キーストアを開く
    using var store = new X509Store(parsedStoreName, parsedStoreLocation);
    store.Open(OpenFlags.ReadOnly);

    // サムプリントで証明書を検索（大文字小文字を区別しない）
    var certificates = store.Certificates.Find(
      X509FindType.FindByThumbprint,
      thumbprint,
      false // validOnly: false にすることで、有効期限切れの証明書も検索可能
    );

    if (certificates.Count == 0)
    {
      throw new Exception(
        $"指定されたサムプリント '{thumbprint}' の証明書がキーストア '{parsedStoreName}' ({parsedStoreLocation}) に見つかりませんでした");
    }

    // 最初の証明書を取得（通常は1つのはず）
    var certificate = certificates[0];
    
    // 取得した証明書を返す
    return new X509Certificate2(certificate);
  }

  /// <summary>
  /// ファイルをアップロード
  /// </summary>
  public async Task<string> UploadFileAsync(string siteUrl, string libraryName, string folderPath, string localFilePath)
  {
    // SharePointサイトのIDを取得
    var siteId = await GetSiteIdAsync(siteUrl);

    // ドライブIDを取得
    var driveId = await GetDriveIdAsync(siteId, libraryName);

    // ファイル名を取得
    var fileName = Path.GetFileName(localFilePath);

    _logger.LogInformation($"ファイルをアップロード中: {fileName}");

    // フォルダパスを正規化
    var targetFolderPath = string.IsNullOrEmpty(folderPath) ? string.Empty : folderPath.Trim('/');

    // フォルダが存在することを確認（存在しない場合は作成）
    var folderId = await EnsureFolderAsync(driveId, targetFolderPath);
    var fileSize = new FileInfo(localFilePath).Length;

    DriveItem? driveItem;

    // ファイルをアップロード
    if (fileSize < 4 * 1024 * 1024)
    {
      using var fileStream = new FileStream(localFilePath, FileMode.Open, FileAccess.Read);
      driveItem = await _graphClient.Drives[driveId]
        .Items[folderId]
        .ItemWithPath(fileName)
        .Content
        .PutAsync(fileStream);
    }
    else
    {
      var uploadSession = await _graphClient.Drives[driveId]
        .Items[folderId]
        .ItemWithPath(fileName)
        .CreateUploadSession
        .PostAsync(new Microsoft.Graph.Drives.Item.Items.Item.CreateUploadSession.CreateUploadSessionPostRequestBody
        {
          Item = new DriveItemUploadableProperties
          {
            AdditionalData = new Dictionary<string, object>
            {
              { "@microsoft.graph.conflictBehavior", "replace" }
            }
          }
        });

      if (uploadSession?.UploadUrl == null)
      {
        throw new Exception("アップロードセッションの作成に失敗しました");
      }

      using var fileStream = new FileStream(localFilePath, FileMode.Open, FileAccess.Read);
      var uploadTask = new LargeFileUploadTask<DriveItem>(uploadSession, fileStream, 320 * 1024);
      var uploadResult = await uploadTask.UploadAsync();

      if (!uploadResult.UploadSucceeded)
      {
        throw new Exception("大容量ファイルのアップロードに失敗しました");
      }

      driveItem = uploadResult.ItemResponse;
    }

    if (driveItem?.WebUrl == null)
    {
      throw new Exception("ファイルのアップロードに失敗しました: WebUrlが取得できませんでした");
    }

    _logger.LogInformation($"ファイルのアップロードが完了しました: {driveItem.WebUrl}");
    return driveItem.WebUrl;
  }

  /// <summary>
  /// SharePointサイトIDを取得
  /// </summary>
  private async Task<string> GetSiteIdAsync(string siteUrl)
  {
    _logger.LogInformation("SharePointサイトIDを取得中...");

    // URLからホスト名とパスを抽出
    var uri = new Uri(siteUrl);
    var hostname = uri.Host;
    var sitePath = uri.AbsolutePath.TrimStart('/');
    var graphApiPath = $"{hostname}:/{sitePath}";

    _logger.LogInformation($"リクエストURL: sites/{graphApiPath}");

    // Graph APIを使用してサイト情報を取得
    Site? site;
    try
    {
      site = await _graphClient
        .Sites[graphApiPath]
        .GetAsync();
    }
    catch (ODataError oDataError)
    {
      // ODataErrorの場合は詳細情報をログに記録
      var errorMessage = GetODataErrorMessage(oDataError);
      _logger.LogError(oDataError, $"サイトIDの取得中にODataErrorが発生しました:\n{errorMessage}");
      
      // スタックトレースも含めて詳細を記録
      _logger.LogError($"スタックトレース:\n{oDataError.StackTrace}");
      
      throw new Exception($"サイトIDの取得に失敗しました: {errorMessage}", oDataError);
    }
    catch (Exception ex)
    {
      // その他の例外も詳細を記録
      _logger.LogError(ex, $"サイトIDの取得中に例外が発生しました: {ex.GetType().Name}");
      _logger.LogError($"例外メッセージ: {ex.Message}");
      _logger.LogError($"スタックトレース:\n{ex.StackTrace}");
      
      if (ex.InnerException != null)
      {
        _logger.LogError($"内部例外: {ex.InnerException.GetType().Name} - {ex.InnerException.Message}");
        _logger.LogError($"内部例外スタックトレース:\n{ex.InnerException.StackTrace}");
      }
      
      throw HandleGraphApiException(ex, "サイトIDの取得");
    }

    if (site?.Id == null)
    {
      throw new Exception("サイトIDが取得できませんでした");
    }

    _logger.LogInformation($"サイトIDを取得しました: {site.Id}");
    return site.Id;
  }

  /// <summary>
  /// ドライブ（ライブラリ）IDを取得
  /// </summary>
  private async Task<string> GetDriveIdAsync(string siteId, string libraryName)
  {
    _logger.LogInformation($"ドライブ（ライブラリ）IDを取得中: {libraryName}");

    // ドライブ一覧を取得
    DriveCollectionResponse? drives;
    try
    {
      drives = await _graphClient
        .Sites[siteId]
        .Drives
        .GetAsync();
    }
    catch (Exception ex)
    {
      throw HandleGraphApiException(ex, "ドライブIDの取得");
    }

    if (drives?.Value == null)
    {
      throw new Exception("ドライブ一覧が取得できませんでした");
    }

    // 指定されたライブラリ名に一致するドライブを検索
    var drive = drives.Value.FirstOrDefault(d => d.Name == libraryName);

    if (drive?.Id == null)
    {
      throw new Exception($"指定されたライブラリ '{libraryName}' が見つかりませんでした");
    }

    _logger.LogInformation($"ドライブIDを取得しました: {drive.Id}");
    return drive.Id;
  }

  /// <summary>
  /// フォルダが存在することを確認し、存在しない場合は作成
  /// </summary>
  private async Task<string> EnsureFolderAsync(string driveId, string folderPath)
  {
    // フォルダパスが空の場合はルートを返す
    if (string.IsNullOrEmpty(folderPath))
    {
      return "root";
    }

    var folders = folderPath.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
    var currentParentId = "root";

    foreach (var folderName in folders)
    {
      try
      {
        // フォルダが存在するか確認
        var existingFolder = await _graphClient.Drives[driveId]
          .Items[currentParentId]
          .ItemWithPath(folderName)
          .GetAsync();

        if (existingFolder?.Id != null)
        {
          currentParentId = existingFolder.Id;
          continue;
        }
      }
      catch
      {
        // フォルダが存在しない場合は作成
      }

      // フォルダを作成
      var newFolder = new DriveItem
      {
        Name = folderName,
        Folder = new Folder(),
        AdditionalData = new Dictionary<string, object>
        {
          { "@microsoft.graph.conflictBehavior", "fail" }
        }
      };

      try
      {
        var createdFolder = await _graphClient.Drives[driveId]
          .Items[currentParentId]
          .Children
          .PostAsync(newFolder);

        if (createdFolder?.Id != null)
        {
          _logger.LogInformation($"フォルダを作成しました: {folderName}");
          currentParentId = createdFolder.Id;
        }
        else
        {
          throw new Exception($"フォルダの作成に失敗しました: {folderName}");
        }
      }
      catch (Exception ex)
      {
        // 既に存在する場合は無視して取得
        if (ex.Message.Contains("nameAlreadyExists") || ex.Message.Contains("resourceAlreadyExists"))
        {
          var existingFolder = await _graphClient.Drives[driveId]
            .Items[currentParentId]
            .ItemWithPath(folderName)
            .GetAsync();

          if (existingFolder?.Id != null)
          {
            currentParentId = existingFolder.Id;
          }
          else
          {
            throw;
          }
        }
        else
        {
          throw;
        }
      }
    }

    return currentParentId;
  }

  public void Dispose()
  {
    _graphClient?.Dispose();
  }
  
  /// <summary>
  /// ODataError例外から詳細なエラーメッセージを取得する
  /// </summary>
  private string GetODataErrorMessage(ODataError oDataError)
  {
    if (oDataError?.Error == null)
    {
      return "ODataError: エラー情報が取得できませんでした";
    }

    var error = oDataError.Error;
    var message = $"ODataError: Code={error.Code}, Message={error.Message}";

    if (error.Details != null && error.Details.Count > 0)
    {
      var details = string.Join("; ", error.Details.Select(d => $"{d.Target}: {d.Message}"));
      message += $", Details=[{details}]";
    }

    // 内部エラーを再帰的に取得
    if (error.InnerError != null)
    {
      message += GetInnerErrorDetails(error.InnerError, 1);
    }

    return message;
  }

  /// <summary>
  /// 内部エラーの詳細を再帰的に取得する
  /// </summary>
  private string GetInnerErrorDetails(MainError? innerError, int depth)
  {
    if (innerError == null)
    {
      return string.Empty;
    }

    var indent = new string(' ', depth * 2);
    var message = $"\n{indent}InnerError[{depth}]: Code={innerError.Code}, Message={innerError.Message}";

    // 内部エラーの内部エラーも再帰的に取得（最大5階層まで）
    if (innerError.InnerError != null && depth < 5)
    {
      message += GetInnerErrorDetails(innerError.InnerError, depth + 1);
    }

    return message;
  }

  /// <summary>
  /// Graph API呼び出し時の例外を処理し、適切なエラーメッセージを返す
  /// </summary>
  private Exception HandleGraphApiException(Exception ex, string operation)
  {
    if (ex is ODataError oDataError)
    {
      var errorMessage = GetODataErrorMessage(oDataError);
      _logger.LogError(ex, $"{operation}中にODataErrorが発生しました: {errorMessage}");
      return new Exception($"{operation}に失敗しました: {errorMessage}", ex);
    }

    _logger.LogError(ex, $"{operation}中に例外が発生しました");
    return new Exception($"{operation}に失敗しました: {ex.Message}", ex);
  }
}
