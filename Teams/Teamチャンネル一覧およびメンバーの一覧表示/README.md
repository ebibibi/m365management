# Export-TeamsMembers_IncludeChannels.ps1

このスクリプトは下記を出力するスクリプトです。

- チームの一覧
- チームのメンバーの一覧
- チャンネルの一覧
- チャンネルのメンバーの一覧

スクリプトを実行し権限を持つユーザーでログインすると、一覧をCSV形式で取得できます。

出力されるCSVはBOM付きのUTF-8なので、日本語を含む場合でもそのままExcelで開けます。

## 使い方

```powershell
pwsh ./Export-TeamsMembers_IncludeChannels.ps1
```

## 前提条件

- PowerShell 7 以上
- MicrosoftTeams モジュール

## 認証について

スクリプト実行時にデバイスコード認証が行われます。
表示されるコードを https://microsoft.com/devicelogin で入力して認証してください。
