using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;

namespace SharePointUploader;

public class GraphApiClient
{
  private readonly string _tenantId;
  private readonly string _clientId;
  private readonly string _certificatePath;
  private readonly string _certificatePassword;
  private readonly HttpClient _httpClient;
  private string? _accessToken;
  private DateTime _tokenExpiry;

  public GraphApiClient(string tenantId, string clientId, string certificatePath, string certificatePassword)
  {
    _tenantId = tenantId;
    _clientId = clientId;
    _certificatePath = certificatePath;
    _certificatePassword = certificatePassword;
    _httpClient = new HttpClient();
  }

  public async Task<string> GetAccessTokenAsync(ILogger logger)
  {
    // トークンが有効な場合は再利用
    if (!string.IsNullOrEmpty(_accessToken) && DateTime.UtcNow < _tokenExpiry)
    {
      return _accessToken;
    }

    logger.LogInformation("アクセストークンを取得中...");

    // 証明書を読み込む
    X509Certificate2? certificate;
    try
    {
      if (string.IsNullOrEmpty(_certificatePassword))
      {
        certificate = new X509Certificate2(_certificatePath);
      }
      else
      {
        certificate = new X509Certificate2(_certificatePath, _certificatePassword);
      }
    }
    catch (Exception ex)
    {
      throw new Exception($"証明書の読み込みに失敗しました: {ex.Message}", ex);
    }

    // OAuth 2.0 トークンエンドポイント
    var tokenEndpoint = $"https://login.microsoftonline.com/{_tenantId}/oauth2/v2.0/token";

    // JWT アサーションを作成
    var assertion = CreateJwtAssertion(certificate);

    // トークンリクエスト
    var requestContent = new FormUrlEncodedContent(new[]
    {
      new KeyValuePair<string, string>("client_id", _clientId),
      new KeyValuePair<string, string>("scope", "https://graph.microsoft.com/.default"),
      new KeyValuePair<string, string>("client_assertion_type", "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"),
      new KeyValuePair<string, string>("client_assertion", assertion),
      new KeyValuePair<string, string>("grant_type", "client_credentials")
    });

    var response = await _httpClient.PostAsync(tokenEndpoint, requestContent);
    var responseContent = await response.Content.ReadAsStringAsync();

    if (!response.IsSuccessStatusCode)
    {
      throw new Exception($"トークン取得に失敗しました: {response.StatusCode} - {responseContent}");
    }

    var tokenResponse = JsonSerializer.Deserialize<JsonElement>(responseContent);
    _accessToken = tokenResponse.GetProperty("access_token").GetString();
    var expiresIn = tokenResponse.GetProperty("expires_in").GetInt32();
    _tokenExpiry = DateTime.UtcNow.AddSeconds(expiresIn - 300); // 5分前に期限切れとして扱う

    logger.LogInformation("アクセストークンを取得しました");
    return _accessToken ?? throw new Exception("アクセストークンが取得できませんでした");
  }

  private string CreateJwtAssertion(X509Certificate2 certificate)
  {
    // JWT ヘッダー
    var header = new
    {
      alg = "RS256",
      typ = "JWT",
      x5t = Base64UrlEncode(certificate.GetCertHash())
    };

    // JWT ペイロード
    var now = DateTimeOffset.UtcNow.ToUnixTimeSeconds();
    var payload = new
    {
      aud = $"https://login.microsoftonline.com/{_tenantId}/oauth2/v2.0/token",
      exp = now + 3600,
      iss = _clientId,
      jti = Guid.NewGuid().ToString(),
      nbf = now,
      sub = _clientId
    };

    var headerJson = JsonSerializer.Serialize(header);
    var payloadJson = JsonSerializer.Serialize(payload);

    var headerBase64 = Base64UrlEncode(Encoding.UTF8.GetBytes(headerJson));
    var payloadBase64 = Base64UrlEncode(Encoding.UTF8.GetBytes(payloadJson));

    var unsignedToken = $"{headerBase64}.{payloadBase64}";

    // 署名
    using var rsa = certificate.GetRSAPrivateKey();
    if (rsa == null)
    {
      throw new Exception("証明書からRSA秘密鍵を取得できませんでした");
    }

    var signature = rsa.SignData(Encoding.UTF8.GetBytes(unsignedToken), System.Security.Cryptography.HashAlgorithmName.SHA256, System.Security.Cryptography.RSASignaturePadding.Pkcs1);
    var signatureBase64 = Base64UrlEncode(signature);

    return $"{unsignedToken}.{signatureBase64}";
  }

  private static string Base64UrlEncode(byte[] input)
  {
    return Convert.ToBase64String(input)
      .TrimEnd('=')
      .Replace('+', '-')
      .Replace('/', '_');
  }

  public async Task<string> UploadFileAsync(string siteUrl, string libraryName, string folderPath, string localFilePath, ILogger logger)
  {
    var accessToken = await GetAccessTokenAsync(logger);

    // SharePointサイトのIDを取得
    var siteId = await GetSiteIdAsync(siteUrl, accessToken, logger);

    // ドライブIDを取得
    var driveId = await GetDriveIdAsync(siteId, libraryName, accessToken, logger);

    // ファイル名を取得
    var fileName = Path.GetFileName(localFilePath);
    var fileContent = await File.ReadAllBytesAsync(localFilePath);

    // アップロード先のパスを構築
    var uploadPath = string.IsNullOrEmpty(folderPath) ? fileName : $"{folderPath}/{fileName}";
    
    // パスをURLエンコード
    var encodedPath = Uri.EscapeDataString(uploadPath).Replace("%2F", "/");

    logger.LogInformation($"ファイルをアップロード中: {uploadPath}");

    // ファイルをアップロード
    var uploadUrl = $"https://graph.microsoft.com/v1.0/sites/{siteId}/drives/{driveId}/root:/{encodedPath}:/content";
    
    _httpClient.DefaultRequestHeaders.Clear();
    _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

    using var content = new ByteArrayContent(fileContent);
    content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");

    var response = await _httpClient.PutAsync(uploadUrl, content);

    if (!response.IsSuccessStatusCode)
    {
      var errorContent = await response.Content.ReadAsStringAsync();
      throw new Exception($"ファイルアップロードに失敗しました: {response.StatusCode} - {errorContent}");
    }

    var responseContent = await response.Content.ReadAsStringAsync();
    var uploadResult = JsonSerializer.Deserialize<JsonElement>(responseContent);
    var webUrl = uploadResult.GetProperty("webUrl").GetString();

    logger.LogInformation($"ファイルのアップロードが完了しました: {webUrl}");
    return webUrl ?? string.Empty;
  }

  private async Task<string> GetSiteIdAsync(string siteUrl, string accessToken, ILogger logger)
  {
    logger.LogInformation("SharePointサイトIDを取得中...");

    // URLからホスト名とパスを抽出
    var uri = new Uri(siteUrl);
    var hostname = uri.Host;
    var sitePath = uri.AbsolutePath.TrimStart('/');

    var apiUrl = $"https://graph.microsoft.com/v1.0/sites/{hostname}:/{sitePath}";

    _httpClient.DefaultRequestHeaders.Clear();
    _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

    var response = await _httpClient.GetAsync(apiUrl);
    var content = await response.Content.ReadAsStringAsync();

    if (!response.IsSuccessStatusCode)
    {
      throw new Exception($"サイトIDの取得に失敗しました: {response.StatusCode} - {content}");
    }

    var site = JsonSerializer.Deserialize<JsonElement>(content);
    var siteId = site.GetProperty("id").GetString();

    logger.LogInformation($"サイトIDを取得しました: {siteId}");
    return siteId ?? throw new Exception("サイトIDが取得できませんでした");
  }

  private async Task<string> GetDriveIdAsync(string siteId, string libraryName, string accessToken, ILogger logger)
  {
    logger.LogInformation($"ドライブ（ライブラリ）IDを取得中: {libraryName}");

    var apiUrl = $"https://graph.microsoft.com/v1.0/sites/{siteId}/drives";

    _httpClient.DefaultRequestHeaders.Clear();
    _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

    var response = await _httpClient.GetAsync(apiUrl);
    var content = await response.Content.ReadAsStringAsync();

    if (!response.IsSuccessStatusCode)
    {
      throw new Exception($"ドライブIDの取得に失敗しました: {response.StatusCode} - {content}");
    }

    var drives = JsonSerializer.Deserialize<JsonElement>(content);
    var drivesArray = drives.GetProperty("value").EnumerateArray();

    foreach (var drive in drivesArray)
    {
      var name = drive.GetProperty("name").GetString();
      if (name == libraryName)
      {
        var driveId = drive.GetProperty("id").GetString();
        logger.LogInformation($"ドライブIDを取得しました: {driveId}");
        return driveId ?? throw new Exception("ドライブIDが取得できませんでした");
      }
    }

    throw new Exception($"指定されたライブラリ '{libraryName}' が見つかりませんでした");
  }

  public void Dispose()
  {
    _httpClient?.Dispose();
  }
}
