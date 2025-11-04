# Invoke-ADStaleInventory.ps1

Active Directory の古くなったオブジェクトと構成の問題を検出するスクリプト

## 概要

このスクリプトは、Active Directory 環境内の以下の問題を自動的に検出します：

- **DC メタデータ**: 削除されていないドメインコントローラーの残骸
- **古いコンピューターアカウント**: 長期間ログオンしていない、または無効化されたコンピューター
- **古いユーザーアカウント**: 長期間ログオンしていない、または無効化されたユーザー
- **DNS レコードの残骸**: 削除済み DC の DNS レコードが残存している問題
- **AD サイト/サブネット/サイトリンクの構成問題**:
  - サーバー/サブネット/リンクから参照されていない孤立サイト
  - サブネットが割り当てられていないサイト
  - サイトに割り当てられていないサブネット
  - 存在しないサイトを参照しているサブネット
  - サブネットの重複・オーバーラップ
  - サイトリンクの構成エラー

## 前提条件

### 必要なモジュール
- Active Directory PowerShell モジュール

### 必要な権限
- **Domain Admins** または **Enterprise Admins** グループのメンバーシップ
- DNS レコードチェックを有効にする場合は、ドメインコントローラー上の DNS サーバーへの読み取り権限

### 実行環境
- Windows Server 2012 R2 以降
- PowerShell 5.1 以降
- ドメインコントローラーまたは RSAT (リモートサーバー管理ツール) がインストールされた端末

## パラメータ

### -InactiveDays
非アクティブと判定する日数を指定します。

- **型**: `int`
- **デフォルト**: `90`
- **説明**: 最終ログオンから指定した日数が経過したオブジェクトを「古い」と判定します

### -OutputPath
レポートを出力するディレクトリパスを指定します。

- **型**: `string`
- **デフォルト**: なし（指定しない場合は CSV 出力なし）
- **説明**: 指定したディレクトリに以下のファイルが出力されます
  - `AD-StaleInventory-YYYYMMDD-HHMMSS.csv` (全結果)
  - `AD-StaleInventory-{Category}-YYYYMMDD-HHMMSS.csv` (カテゴリ別)

### -CheckDns
削除済み DC の DNS レコード残骸をチェックします。

- **型**: `switch`
- **デフォルト**: `$true`
- **説明**: 降格済み DC の DNS レコードが残存していないかをチェックします。DNSサーバーは稼働中のドメインコントローラーから自動的に選択されます

### -CheckSites
サイト/サブネット/サイトリンクの健全性チェックを実行するかを指定します。

- **型**: `switch`
- **デフォルト**: `$true`
- **説明**: AD サイトトポロジーの構成問題をチェックします

### -Help
スクリプトのヘルプメッセージを表示します。

- **型**: `switch`
- **デフォルト**: なし
- **説明**: パラメータの説明と使用例を表示して終了します

## 使用例

### 例 0: ヘルプを表示
```powershell
.\Invoke-ADStaleInventory.ps1 -Help
```

スクリプトのヘルプメッセージを表示します。

### 例 1: 基本的な実行（CSV 出力あり）
```powershell
.\Invoke-ADStaleInventory.ps1 -OutputPath "C:\Temp\AD-StaleReport"
```

デフォルト設定（90日間非アクティブ）で実行し、結果を `C:\Temp\AD-StaleReport` に出力します。

### 例 2: 非アクティブ期間を変更
```powershell
.\Invoke-ADStaleInventory.ps1 -InactiveDays 180 -OutputPath "C:\Reports"
```

180日間非アクティブなオブジェクトを検出します。

### 例 3: DNS チェックをスキップ
```powershell
.\Invoke-ADStaleInventory.ps1 -CheckDns:$false -OutputPath "C:\Temp\AD-StaleReport"
```

DNS レコードのチェックをスキップします（既定では有効）。

### 例 4: サイトチェックのみをスキップ
```powershell
.\Invoke-ADStaleInventory.ps1 -CheckSites:$false -OutputPath "C:\Reports"
```

サイト/サブネット/サイトリンクのチェックをスキップします。

### 例 5: コンソール出力のみ（CSV なし）
```powershell
.\Invoke-ADStaleInventory.ps1
```

結果を画面に表示のみで、CSV ファイルは出力しません。

## 出力形式

### コンソール出力
スクリプトは以下の情報をコンソールに表示します：

```
[INFO] Active Directory健全性チェックを開始します...
[INFO] InactiveDays: 90, CheckDns: True, CheckSites: True
VERBOSE: DNS残骸チェック用のDNSサーバーとして [dc01.contoso.local] を使用します

===== 検出結果サマリー =====
  DC-Metadata: 2 件
  Stale-Computer: 15 件
  Stale-User: 20 件
  AD-Site: 3 件
  AD-Subnet: 2 件

===== 重要度別 =====
  High: 10 件
  Medium: 20 件
  Low: 12 件

[完了] 全 42 件の問題候補を検出しました。
[INFO] 詳細はCSVファイルをご確認ください: C:\Temp\AD-StaleReport
```

詳細は以下のコマンドで確認できます:
```powershell
$result | Out-GridView
$result | Format-Table Category, Name, Confidence, Detail -AutoSize
```

### CSV 出力
以下のフィールドを含む CSV ファイルが生成されます：

| フィールド | 説明 |
|----------|------|
| Category | 問題のカテゴリ (`DC-Metadata`, `Stale-Computer`, `Stale-User`, など) |
| Name | オブジェクト名 |
| DN | 識別名 (Distinguished Name) |
| Detail | 問題の詳細説明 |
| LastSeen | 最終確認日時 |
| Confidence | 信頼度 (`High`, `Medium`, `Low`) |
| RecommendedAction | 推奨される対処方法 |

## 検出カテゴリの詳細

### DC-Metadata
削除されていない DC のメタデータが Configuration パーティションに残存している問題を検出します。

**推奨アクション**: `ntdsutil` を使用してメタデータをクリーンアップします。

### Stale-Computer
以下のいずれかに該当するコンピューターアカウントを検出します：
- アカウントが無効化されている
- 最終ログオンが指定日数を超えている
- 作成後一度もログオンしていない（古いアカウント）

**推奨アクション**: 管理者に確認の上、削除または専用 OU に移動します。

### Stale-User
以下のいずれかに該当するユーザーアカウントを検出します：
- アカウントが無効化されている
- 最終ログオンが指定日数を超えている
- 作成後一度もログオンしていない（古いアカウント）
- パスワードが無期限に設定されている（サービスアカウントの可能性）

**推奨アクション**: 人事・管理部門と調整の上、削除またはアーカイブ OU に移動します。

### Stale-DC-DNS
削除済み DC の DNS レコードが残存している問題を検出します。

**推奨アクション**: DNS の A/SRV/CNAME レコードを手動で確認・削除します。

### AD-Site
以下のサイト構成の問題を検出します：
- サーバー/サブネット/リンクから一切参照されていない孤立サイト
- サブネットは割り当てられているがサーバーがないサイト
- サーバーはあるがサブネットが割り当てられていないサイト
- 複数サイト構成なのにサイトリンクに未参加のサイト

**推奨アクション**: 意図的な構成でなければ、サイトを削除または適切に構成します。

### AD-Subnet
以下のサブネット構成の問題を検出します：
- どのサイトにも割り当てられていないサブネット
- 存在しないサイトを参照しているサブネット
- 重複する CIDR のサブネット
- オーバーラップする IP 範囲のサブネット

**推奨アクション**: CIDR を見直し、重複・包含関係を解消します。

### AD-SiteLink
以下のサイトリンク構成の問題を検出します：
- リンクに含まれるサイトが2未満（レプリケーション経路として無意味）
- 存在しないサイトを参照しているリンク

**推奨アクション**: `siteList` から不正な参照を削除、または適切なサイトを追加します。

## 結果の活用方法

### スクリプト実行後の変数利用
スクリプトは結果を `$result` 変数に格納します。以下のコマンドで詳細確認が可能です：

```powershell
# グリッドビューで確認
$result | Out-GridView

# カテゴリ別に表示
$result | Group-Object Category | Format-Table Count, Name -AutoSize

# 信頼度 High のみ表示
$result | Where-Object { $_.Confidence -eq 'High' } | Format-Table

# 特定カテゴリのみ表示
$result | Where-Object { $_.Category -eq 'DC-Metadata' } | Format-List *
```

## トラブルシューティング

### エラー: "Get-ADDomainController : 用語 'Get-ADDomainController' は、コマンドレットの名前として認識されません"

**原因**: Active Directory PowerShell モジュールがインストールされていません。

**解決方法**:
```powershell
# Windows Server の場合
Install-WindowsFeature -Name RSAT-AD-PowerShell

# Windows 10/11 の場合
Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
```

### エラー: "アクセスが拒否されました"

**原因**: 実行ユーザーに適切な権限がありません。

**解決方法**: Domain Admins または Enterprise Admins グループのメンバーアカウントで実行してください。

### 警告: "稼働DCが見つからないため、DNSチェックをスキップします"

**原因**: 稼働中のドメインコントローラーが見つからないため、DNSサーバーを自動選択できません。

**解決方法**:
- ドメインコントローラーが正常に稼働しているか確認
- ネットワーク接続を確認
- Active Directory PowerShell モジュールが正しくロードされているか確認

### エラー: "DNSレコード取得エラー"

**原因**: DNS サーバーへのアクセスまたはDNSレコードの取得に失敗しました。

**解決方法**:
- ファイアウォールで DNS ポート（53）が開いているか確認
- DNS サーバーへの WinRM アクセスが可能か確認
- `-Verbose` パラメータを付けて実行し、詳細なエラー情報を確認

## ライセンス

このスクリプトは MIT ライセンスの下で提供されます。

## 作者

m365management プロジェクト

## 更新履歴

- 2025-01-XX: 初版リリース
  - DC メタデータチェック
  - 古いコンピューター/ユーザーチェック
  - サイト/サブネット/サイトリンクチェック
  - DNS レコード残骸チェック
