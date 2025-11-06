using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
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

    // ClientCertificateCredentialを使用して認証
    var credential = new ClientCertificateCredential(
      tenantId,
      clientId,
      certificate
    );

    // GraphServiceClientの作成
    _graphClient = new GraphServiceClient(credential);
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

    // ファイルをアップロード
    using var fileStream = new FileStream(localFilePath, FileMode.Open, FileAccess.Read);
    
    var driveItem = await _graphClient.Drives[driveId]
      .Items[folderId]
      .ItemWithPath(fileName)
      .Content
      .PutAsync(fileStream);

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
