# M365 インベントリ・レポート スクリプト

## 目的  
Microsoft 365 テナント内の構成情報を一括取得し、棚卸し・クリーンアップ判断に役立つレポートを生成します。  
取得対象は以下のとおりです。

| 区分 | 取得内容 |
|------|----------|
| Users | 社内ユーザー (Member) 一覧、サインイン最終日時、ProxyAddresses 等 |
| Guests | ゲストユーザー (Guest) 一覧、招待状態、サインイン最終日時 |
| Groups | メール対応グループ/セキュリティグループ (非 Teams) |
| Teams | チーム情報、作成日時、可視性 |
| TeamMembers | チームごとのメンバー (Owner/Member/Guest) |
| TeamOwners | チームごとの所有者 |
| SharedMailboxes | 共有メールボックス、最終ログオン、容量 等 |
| SharedMailboxPerms | 共有メールボックスの権限一覧（FullAccess, SendAs, SendOnBehalf） |
| DistributionGroups | 配布/セキュリティ グループ (EXO) |
| DLMembers | 配布グループごとのメンバー |

出力は **ImportExcel モジュールが存在する場合** は 1 つの Excel ファイル（複数シート）、  
無い場合は `output` フォルダー内に複数 CSV とログを生成します。

---

## 前提条件

| 項目 | バージョン / 権限 |
|------|------------------|
| PowerShell | 7.2 以上 |
| モジュール | Microsoft.Graph / ImportExcel / ExchangeOnlineManagement |
| Azure AD ロール | Global Reader または Reports Reader 以上 |
| Exchange ロール | View-Only Recipients 以上 |

> **初回実行前に** `Setup-InventoryReportPrereqs.ps1` を実行してモジュールのインストールと Graph/EXO への管理者同意 (Consent) を取得してください。

---

## セットアップ手順

```powershell
# 1. 必要モジュールとコンセントを準備 (初回のみ)
cd .\All
.\Setup-InventoryReportPrereqs.ps1
```

成功メッセージを確認したら閉じて構いません。

---

## レポート生成手順

```powershell
# 2. インベントリ収集 (毎回実行)
cd .\All
.\Generate-M365InventoryReport.ps1
```

- `.\output` フォルダーが自動生成され、`M365-Inventory-yyyymmdd_hhmm.xlsx` もしくは複数 CSV が保存されます。  
- `Run_yyyymmdd_hhmm.log` に実行ログが記録されます。  

---

## カスタマイズ例

| 追加したい対象 | 参考コマンド / 変更箇所 |
|---------------|-------------------------|
| SharePoint サイト一覧 | Graph: `Get-MgSite -All` |
| OneDrive 容量 | Graph: `reports/getOneDriveUsageAccountDetail` |
| ライセンス情報 | `Get-MgSubscribedSku`, `Get-MgUserLicenseDetail` |

`Generate-M365InventoryReport.ps1` の関数ブロックに追加し、エクスポート部にシート／CSV を追加してください。

---

## トラブルシューティング

| 症状 | 対処 |
|------|------|
| Graph 接続に失敗 | Azure ポータルでアプリ同意が拒否されていないか確認 |
| ExchangeOnlineManagement が読み込めない | `Update-Module ExchangeOnlineManagement -Force` |
| ImportExcel が無い/使えない | 自動で CSV 出力にフォールバック |

---

## ライセンス
MIT License
