using System;
using System.IO;
using System.Text.Json;

namespace SharePointUploader;

public class Config
{
  public SharePointConfig SharePoint { get; set; } = new();
  public LogConfig Log { get; set; } = new();

  public static Config Load()
  {
    var configPath = Path.Combine(Directory.GetCurrentDirectory(), "Config.json");
    
    if (!File.Exists(configPath))
    {
      throw new FileNotFoundException($"設定ファイルが見つかりません: {configPath}");
    }

    var json = File.ReadAllText(configPath);
    var config = JsonSerializer.Deserialize<Config>(json, new JsonSerializerOptions
    {
      PropertyNameCaseInsensitive = true
    });

    if (config == null)
    {
      throw new InvalidOperationException("設定ファイルの読み込みに失敗しました");
    }

    return config;
  }
}

public class SharePointConfig
{
  public string SiteUrl { get; set; } = string.Empty;
  public string LibraryName { get; set; } = string.Empty;
  public string FolderPath { get; set; } = string.Empty;
  public string TenantId { get; set; } = string.Empty;
  public string ClientId { get; set; } = string.Empty;
  public string AuthRecordFile { get; set; } = "AuthAccount.json";
}

public class LogConfig
{
  public string Level { get; set; } = "Information";
  public int RetainedFileCountLimit { get; set; } = 10;
}
