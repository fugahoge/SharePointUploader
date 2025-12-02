using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Threading;
using System.Threading.Tasks;
using Azure.Core;
using Azure.Identity;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using System.IdentityModel.Tokens.Jwt;

namespace SharePointUploader;

public class GraphApiClient : IDisposable
{
  private readonly ILogger _logger;

  private readonly GraphServiceClient _graphClient;

  public GraphApiClient(string tenantId, string clientId, string authRecordFile, ILogger logger)
  {
    _logger = logger;
    
    // 権限スコープ
    var scopes = new[] { 
      "Files.ReadWrite.All",
      "Sites.ReadWrite.All",
      "User.Read"
    };

    AuthenticationRecord? authRecord = null;

    // キャッシュされた認証情報を読み込む
    if (File.Exists(authRecordFile))
    {
      using var stream = File.OpenRead(authRecordFile);
      authRecord = AuthenticationRecord.Deserialize(stream);
    }

    // ユーザー認証（Interactive Browser認証）
    // 初回のみブラウザでログインし、次回以降はキャッシュされたトークンを使用
    var interactiveCredential = new InteractiveBrowserCredential(
      new InteractiveBrowserCredentialOptions
      {
        TenantId = tenantId,
        ClientId = clientId,
        RedirectUri = new Uri("http://localhost"),

        // トークンキャッシュを有効化（Windows Credential Managerへ保存）
        TokenCachePersistenceOptions = new TokenCachePersistenceOptions
        {
          Name = "SharePointUploaderTokenCache"
        },

        // キャッシュされた認証情報を使用
        AuthenticationRecord = authRecord
      }
    );

    // 認証情報がない場合はユーザー認証を行い、認証情報を保存
    if(authRecord == null)
    {
      _logger.LogWarning("ユーザー認証を行います。");

      var context = new TokenRequestContext(scopes);

      authRecord = interactiveCredential
        .AuthenticateAsync(context)
        .GetAwaiter()
        .GetResult();

      using var stream = File.Create(authRecordFile);
      authRecord.Serialize(stream);
    }

    // GraphServiceClientの作成
    _graphClient = new GraphServiceClient(interactiveCredential, scopes);

    // トークンの権限（スコープ）を表示
    try
    {
      LogTokenScopesAsync(interactiveCredential, scopes).GetAwaiter().GetResult();
    }
    catch (Exception ex)
    {
      _logger.LogWarning(ex, "トークンの権限情報の表示中にエラーが発生しました（処理は続行します）");
    }
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

    _logger.LogInformation($"ファイルをアップロードします: {fileName}");

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
      throw new Exception("ファイルのアップロードに失敗しました: アップロード結果が取得できませんでした");
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

    // Graph APIを使用してサイト情報を取得
    Site? site;
    try
    {
      site = await _graphClient
        .Sites[graphApiPath]
        .GetAsync();
    }
    catch (Exception ex)
    {
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
  /// 内部エラーの詳細を取得する
  /// </summary>
  private string GetInnerErrorDetails(InnerError? innerError, int depth)
  {
    if (innerError == null)
    {
      return string.Empty;
    }

    var indent = new string(' ', depth * 2);
    return $"\n{indent}InnerError[{depth}]: {innerError}";
  }
  
  /// <summary>
  /// トークンの権限（スコープ）をログに表示する
  /// </summary>
  private async Task LogTokenScopesAsync(TokenCredential credential, string[] scopes)
  {
    try
    {
      var tokenRequestContext = new TokenRequestContext(scopes);
      var token = await credential.GetTokenAsync(tokenRequestContext, CancellationToken.None);

      if (string.IsNullOrEmpty(token.Token))
      {
        _logger.LogWarning("トークンの取得に失敗しました");
        return;
      }

      // JWTをデコード
      var handler = new JwtSecurityTokenHandler();
      if (!handler.CanReadToken(token.Token))
      {
        _logger.LogWarning("トークンのデコードに失敗しました");
        return;
      }

      var jsonToken = handler.ReadJwtToken(token.Token);

      _logger.LogInformation("=== トークンの権限情報 ===");

      // スコープ（委任アクセスの場合）
      var scpClaim = jsonToken.Claims.FirstOrDefault(c => c.Type == "scp");
      if (scpClaim != null)
      {
        var scopesList = scpClaim.Value.Split(' ');
        _logger.LogInformation($"委任アクセスのスコープ ({scopesList.Length}個):");
        foreach (var scope in scopesList)
        {
          _logger.LogInformation($"  - {scope}");
        }
      }

      // ロール（アプリケーションアクセスの場合）
      var rolesClaim = jsonToken.Claims.FirstOrDefault(c => c.Type == "roles");
      if (rolesClaim != null)
      {
        var rolesList = rolesClaim.Value.Split(' ');
        _logger.LogInformation($"アプリケーションアクセスのロール ({rolesList.Length}個):");
        foreach (var role in rolesList)
        {
          _logger.LogInformation($"  - {role}");
        }
      }

      // その他の重要なクレーム
      var appIdClaim = jsonToken.Claims.FirstOrDefault(c => c.Type == "appid");
      if (appIdClaim != null)
      {
//        _logger.LogInformation($"アプリケーションID: {appIdClaim.Value}");
      }

      var oidClaim = jsonToken.Claims.FirstOrDefault(c => c.Type == "oid");
      if (oidClaim != null)
      {
//        _logger.LogInformation($"オブジェクトID: {oidClaim.Value}");
      }

      var upnClaim = jsonToken.Claims.FirstOrDefault(c => c.Type == "upn");
      if (upnClaim != null)
      {
//        _logger.LogInformation($"ユーザープリンシパル名: {upnClaim.Value}");
      }

      var expClaim = jsonToken.Claims.FirstOrDefault(c => c.Type == "exp");
      if (expClaim != null && long.TryParse(expClaim.Value, out var expUnixTime))
      {
        var expirationTime = DateTimeOffset.FromUnixTimeSeconds(expUnixTime).LocalDateTime;
        _logger.LogInformation($"トークンの有効期限: {expirationTime:yyyy-MM-dd HH:mm:ss}");
      }

      _logger.LogInformation("========================");
    }
    catch (Exception ex)
    {
      _logger.LogWarning(ex, "トークンの権限情報の取得に失敗しました");
    }
  }
}
