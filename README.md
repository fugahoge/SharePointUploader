# SharePointUploader

SharePoint OnlineにファイルをアップロードするためのC# CLIツールです。

## 機能

- Graph APIを使用した証明書認証
- SharePoint Onlineへのファイルアップロード
- ログ出力（コンソールとログファイル）
- ログローテーション（最新10件を保持）

## 要件

- .NET 8.0 SDK

## セットアップ

1. プロジェクトをビルドします：
```bash
dotnet build
```

2. `Config.json`ファイルを編集し、以下の情報を設定します：
   - `SiteUrl`: SharePointサイトのURL
   - `LibraryName`: アップロード先のライブラリ名
   - `FolderPath`: アップロード先のフォルダパス（オプション）
   - `TenantId`: Azure ADテナントID
   - `ClientId`: Azure ADアプリケーション（クライアント）ID
   - `CertificatePath`: 証明書ファイル（.pfx）のパス
   - `CertificatePassword`: 証明書のパスワード（空の場合はパスワードなし）

## 使用方法

```bash
dotnet run -- <ファイル名>
```

または、ビルド後の実行ファイルを使用：

```bash
SharePointUploader.exe <ファイル名>
```

## ログ

ログは以下の場所に保存されます：
- `Logs/SharePointUploader_YYYYMMDD_HHMMSS.log`

最新10件のログファイルが保持され、それより古いものは自動的に削除されます。

## 設定ファイル例

```json
{
  "SharePoint": {
    "SiteUrl": "https://yourtenant.sharepoint.com/sites/yoursite",
    "LibraryName": "Documents",
    "FolderPath": "Shared Documents/UploadFolder",
    "TenantId": "your-tenant-id",
    "ClientId": "your-client-id",
    "CertificatePath": "path/to/certificate.pfx",
    "CertificatePassword": "certificate-password"
  }
}
```

