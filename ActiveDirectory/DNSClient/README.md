# Check-DnsUpdateRegistry.ps1

DNS 動的更新に関するレジストリ設定を収集してレポートを出力するスクリプトです。

## 概要

以下のレジストリ値を収集し、DNS 動的更新の設定状態をレポートします:

### グローバル設定

`HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters`

| 設定名 | 既定値 | 説明 |
|--------|--------|------|
| DisableDynamicUpdate | 0 | DNS更新の登録を無効化 |
| DisableReverseAddressRegistrations | 0 | PTR登録を無効化 |
| DefaultRegistrationRefreshInterval | 86400 | 登録更新間隔（秒） |
| DefaultRegistrationTTL | 1200 | 登録TTL（秒） |
| UpdateSecurityLevel | 0 | セキュリティレベル |
| DisableReplaceAddressesInConflicts | 0 | 競合時の上書き無効化 |

### インターフェイス別設定

`HKLM\...\Parameters\Interfaces\<GUID>\DisableDynamicUpdate`

### アダプター別設定

`HKLM\...\Parameters\Adapters\<InterfaceName>\MaxNumberOfAddressesToRegister`

## 使い方

```powershell
.\Check-DnsUpdateRegistry.ps1
```

## 出力

スクリプト実行ディレクトリに `DNSUpdateRegistryReport_yyyyMMdd_HHmmss` フォルダが作成され、以下のファイルが出力されます:

- `global.csv` - グローバル設定
- `interfaces.csv` - インターフェイス別設定
- `adapters.csv` - アダプター別設定
- `README.txt` - レポートの説明

## 前提条件

- Windows PowerShell 3.0 以上
- 管理者権限で実行

## 注意事項

- これはオンプレミス Windows 用のスクリプトです
- クラウド認証は不要です
