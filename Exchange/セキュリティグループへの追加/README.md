# Add-UserToSecurityGroup.ps1

このスクリプトは指定されたセキュリティグループに、ユーザーを追加するスクリプトです。

グループメンバーシップの変更はグローバル管理者権限があったとしてもグループのオーナーとして登録されていないと失敗してしまうため、引数として管理者のIDを渡し、グループのオーナーでなければ追加してからメンバーを変更するようにしています。

## 使い方

```powershell
pwsh ./Add-UserToSecurityGroup.ps1 -UserId '<ユーザーのメールID>' -SecurityGroup '<グループ名>' -adminUserId '<管理者ユーザー名>'
```

### 例

```powershell
pwsh ./Add-UserToSecurityGroup.ps1 -UserId 'user@example.com' -SecurityGroup 'Sales Team' -adminUserId 'admin@example.com'
```

## 前提条件

- PowerShell 7 以上
- ExchangeOnlineManagement モジュール

## 認証について

スクリプト実行時にExchange Onlineに未接続の場合、デバイスコード認証が行われます。
表示されるコードを https://microsoft.com/devicelogin で入力して認証してください。