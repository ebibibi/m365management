# Export-TeamsMembers.ps1

このスクリプトは、Teams のチーム メンバーの一覧を取得するスクリプトです。

スクリプトを実行し権限を持つユーザーでログインすると、Teams のチーム メンバーの一覧を取得できます。

出力されるCSVはBOM付きのUTF-8なので、日本語を含む場合でもそのままExcelで開けます。

## 使い方

```powershell
pwsh ./Export-TeamsMembers.ps1
```

## 前提条件

- PowerShell 7 以上
- MicrosoftTeams モジュール

## 認証について

スクリプト実行時にデバイスコード認証が行われます。
表示されるコードを https://microsoft.com/devicelogin で入力して認証してください。

# 解説
下記のYoutube動画で利用方法や中身の解説をしています。

- https://youtu.be/eonRkm77IOo
