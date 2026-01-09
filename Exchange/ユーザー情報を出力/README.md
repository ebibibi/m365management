# Get-UserInformation.ps1

指定したユーザーの情報を取得して表示するスクリプトです。

## 前提条件

- PowerShell 7 以上
- ExchangeOnlineManagement モジュール

## 使い方

```powershell
pwsh ./Get-UserInformation.ps1 -User '<ユーザーのメールアドレス>'
```

### 例

```powershell
pwsh ./Get-UserInformation.ps1 -User 'user@example.com'
```

## 出力内容

- 姓
- 名
- 表示名
- メールアドレス
- 所属グループ
- 権限を持っている共有メールボックス

## 認証について

スクリプト実行時にExchange Onlineに未接続の場合、デバイスコード認証が行われます。
表示されるコードを https://microsoft.com/devicelogin で入力して認証してください。
