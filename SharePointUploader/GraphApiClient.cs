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

namespace SharePointUploader;

public class GraphApiClient : IDisposable
{
  private readonly GraphServiceClient _graphClient;
  private readonly ILogger _logger;

  public GraphApiClient(string tenantId, string clientId, string certificatePath, string certificatePassword, ILogger logger)
  {
    _logger = logger;

    // 証明書を読み込む
    X509Certificate2 certificate;
    try
    {
      if (string.IsNullOrEmpty(certificatePassword))
      {
        certificate = new X509Certificate2(certificatePath);
      }
      else
      {
        certificate = new X509Certificate2(certificatePath, certificatePassword);
      }
    }
    catch (Exception ex)
    {
      throw new Exception($"証明書の読み込みに失敗しました: {ex.Message}", ex);
    }

    // 証明書認証とユーザー認証を組み合わせた認証
    var scopes = new[] { 
      "https://graph.microsoft.com/Files.ReadWrite.All",
      "https://graph.microsoft.com/Sites.ReadWrite.All",
      "https://graph.microsoft.com/User.Read"
    };

    // 証明書認証（アプリケーション認証）
    var certificateCredential = new ClientCertificateCredential(
      tenantId,
      clientId,
      certificate
    );

    // ユーザー認証（Interactive Browser認証）
    // 初回のみブラウザでログインし、次回以降はキャッシュされたトークンを使用
    var interactiveCredential = new InteractiveBrowserCredential(
      new InteractiveBrowserCredentialOptions
      {
        // 初回ログイン: ブラウザでログインし、アクセストークンとリフレッシュトークンを取得
        // 2回目以降: キャッシュからリフレッシュトークンを読み込み、アクセストークンを自動更新
        // リフレッシュトークンの有効期限90日経過後は再度ブラウザでログインが必要
        
        TenantId = tenantId,
        ClientId = clientId,
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

  private async Task<string> GetSiteIdAsync(string siteUrl)
  {
    _logger.LogInformation("SharePointサイトIDを取得中...");

    // URLからホスト名とパスを抽出
    var uri = new Uri(siteUrl);
    var hostname = uri.Host;
    var sitePath = uri.AbsolutePath.TrimStart('/');

    // Graph APIを使用してサイト情報を取得
    Site? site;
    try
    {
      site = await _graphClient
        .Sites[$"{hostname}:/{sitePath}"]
        .GetAsync();
    }
    catch (Exception ex)
    {
      throw new Exception($"サイトIDの取得に失敗しました: {ex.Message}", ex);
    }

    if (site?.Id == null)
    {
      throw new Exception("サイトIDが取得できませんでした");
    }

    _logger.LogInformation($"サイトIDを取得しました: {site.Id}");
    return site.Id;
  }

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
      throw new Exception($"ドライブIDの取得に失敗しました: {ex.Message}", ex);
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
}
