# setattribute.ps1

Entra ID のサービスプリンシパルにカスタムセキュリティ属性を設定するスクリプトです。

## 概要

条件付きアクセスポリシーなどで特定のアプリケーションを除外対象として識別するために、カスタムセキュリティ属性 `isExcludeApp` を設定します。

## 前提条件

- PowerShell 7 以上
- Microsoft.Graph モジュール
- 以下のカスタムセキュリティ属性が事前に作成済みであること:
  - 属性セット: `isExcludeApps`
  - 属性定義: `isExcludeApp`

## 必要な権限

- Application.ReadWrite.All
- Directory.ReadWrite.All
- CustomSecAttributeAssignment.ReadWrite.All

## 使い方

```powershell
pwsh ./setattribute.ps1
```

## 設定方法

スクリプト内の `$excludedAppNames` 配列に、除外対象としたいアプリケーションの DisplayName を記載します。

```powershell
$excludedAppNames = @(
    "ConditionalAccessTest",
    "Sample App A",
    "Sample App B"
)
```

## 動作

1. Microsoft Graph に接続
2. すべてのサービスプリンシパルを取得
3. 各アプリケーションに対して:
   - 除外リストに含まれる場合: `isExcludeApp = "true"`
   - 含まれない場合: `isExcludeApp = "false"`

## 認証について

スクリプト実行時にデバイスコード認証が行われます。
表示されるコードを https://microsoft.com/devicelogin で入力して認証してください。
