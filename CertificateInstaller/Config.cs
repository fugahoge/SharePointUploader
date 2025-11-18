using System;
using System.IO;
using System.Text.Json;

namespace CertificateInstaller;

public class Config
{
  public string StoreName { get; set; } = "My";
  public string StoreLocation { get; set; } = "CurrentUser";
  public LogConfig? Log { get; set; }

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

public class LogConfig
{
  public string Level { get; set; } = "Information";
  public int RetainedFileCountLimit { get; set; } = 10;
}

