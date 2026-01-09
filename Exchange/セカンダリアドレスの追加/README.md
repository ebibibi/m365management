# Add-SecondaryAddressToMailbox.ps1

このスクリプトは指定されたメールボックスに、セカンダリアドレスを追加するスクリプトです。

## 前提条件

- PowerShell 7 以上
- ExchangeOnlineManagement モジュール

## 使い方

```powershell
pwsh ./Add-SecondaryAddressToMailbox.ps1 -Mailbox '<メールボックスのアドレス>' -email '<追加するセカンダリアドレス>'
```

### 例

```powershell
pwsh ./Add-SecondaryAddressToMailbox.ps1 -Mailbox 'shared@example.com' -email 'alias@example.com'
```

## 動作

1. Exchange Onlineへの接続状態を確認
2. 未接続の場合、デバイスコード認証で接続（ブラウザでの認証が必要）
3. 指定メールボックスの現在のメールアドレス一覧を表示
4. セカンダリアドレスが既に存在するか確認
5. 存在しなければ追加

## 認証について

スクリプト実行時にExchange Onlineに未接続の場合、デバイスコード認証が行われます。
表示されるコードを https://microsoft.com/devicelogin で入力して認証してください。
