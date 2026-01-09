# Add-RightsToSharedMailbox.ps1

このスクリプトは指定された共有メールボックスに、指定されたユーザーのフルアクセス権限と所有者権限を追加するスクリプトです。

# Add-RightsToSharedMailboxes.ps1

このスクリプトは指定された複数の共有メールボックスに、指定されたユーザーのフルアクセス権限と所有者権限を追加するスクリプトです。

## 使用方法

### 1. 文字列で直接指定する方法
```powershell
.\Add-RightsToSharedMailboxes.ps1 -User "user@example.com" -SharedMailboxesRaw "mailbox1@example.com
mailbox2@example.com
mailbox3@example.com"
```

### 2. ファイルから読み込む方法（推奨）
```powershell
# まず、メールボックスのリストをファイルに保存します
@"
mailbox1@example.com
mailbox2@example.com
mailbox3@example.com
"@ > mailboxes.txt

# ファイルを指定して実行します
.\Add-RightsToSharedMailboxes.ps1 -User "user@example.com" -SharedMailboxesFile "mailboxes.txt"
```

この方法では、PowerShellの入力制限に悩まされることなく、大量のメールボックスを処理できます。

## 前提条件

- PowerShell 7 以上
- ExchangeOnlineManagement モジュール

## 認証について

スクリプト実行時にExchange Onlineに未接続の場合、デバイスコード認証が行われます。
表示されるコードを https://microsoft.com/devicelogin で入力して認証してください。
