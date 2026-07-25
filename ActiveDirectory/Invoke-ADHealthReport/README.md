# Invoke-ADHealthReport.ps1

Active Directory を読み取り専用で検査し、**単一ファイルの HTML レポート**を生成します。

検出結果は「**利用者への影響があるかどうか**」を軸に 3 つのフェーズへ自動的に振り分けられ、
件数の羅列ではなく「まず何から手を付けるか」が伝わる形で出力されます。

## 特徴

- **読み取り専用**。使用するのは `Get-AD*` 系のみで、構成を変更する処理は含みません
- **単一ファイル**。依存モジュールは Active Directory PowerShell モジュールのみ。持ち込むものが少ないので審査を通しやすい
- **HTML 1 枚で完結**。外部 CSS / JS / フォントを読みに行かないため、閉じた環境でもそのまま開けます。印刷・PDF 化も可
- **匿名化モード**あり。アカウント名を伏せたまま社外に共有できます
- **サンプルモード**あり。AD に接続せずレポートの体裁を確認できます

## 使い方

```powershell
# 基本（カレントディレクトリに出力）
.\Invoke-ADHealthReport.ps1

# 出力先を指定
.\Invoke-ADHealthReport.ps1 -OutputPath C:\Temp\ADReport

# 社外に渡す前提で、アカウント名を伏せる
.\Invoke-ADHealthReport.ps1 -Anonymize -OutputPath C:\Temp\ADReport

# 閾値を変える（非アクティブ 180 日 / パスワード未変更 365 日）
.\Invoke-ADHealthReport.ps1 -InactiveDays 180 -StalePasswordDays 365

# AD に接続せずサンプルを生成（提案時のサンプル提示用）
.\Invoke-ADHealthReport.ps1 -DemoData -CompanyName 'サンプル株式会社'
```

出力は 2 ファイルです。

| ファイル | 内容 |
|---|---|
| `AD-HealthReport-<日時>.html` | 報告用のレポート本体（そのまま共有できます） |
| `AD-HealthReport-<日時>.csv` | 検出明細の全件（作業リストとして使えます） |

## パラメーター

| 名前 | 既定値 | 説明 |
|---|---|---|
| `-InactiveDays` | 90 | 非アクティブと判定する日数 |
| `-StalePasswordDays` | 180 | 重要アカウントのパスワードが「長期間未変更」と判定される日数 |
| `-OutputPath` | `.` | 出力先ディレクトリ |
| `-CompanyName` | ドメイン名 | レポートに表示する組織名 |
| `-Anonymize` | — | アカウント名・コンピューター名を伏せ字にする |
| `-DemoData` | — | AD に接続せずサンプルデータで生成する |

## 実行要件

- Windows PowerShell 5.1 以降
- Active Directory PowerShell モジュール（RSAT）
  ```powershell
  # Windows Server
  Install-WindowsFeature -Name RSAT-AD-PowerShell
  # Windows 10 / 11
  Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
  ```
- ドメインユーザー権限で大半の項目は取得できます。特権グループのメンバー列挙など一部の項目は、
  権限が足りない場合に自動でスキップされます（エラーで止まりません）

## 検査する項目

| カテゴリ | 検査内容 | 既定のフェーズ |
|---|---|---|
| 特権・認証 | krbtgt のパスワード経過日数 | フェーズ1 |
| 特権・認証 | ビルトイン Administrator のパスワード経過日数 | フェーズ2 |
| 特権・認証 | 特権グループ（Domain/Enterprise/Schema Admins）のメンバー数 | フェーズ2 |
| 特権・認証 | 制約なし委任（DC を除く） | フェーズ2 |
| 特権・認証 | SPN 付きユーザー（Kerberoasting の標的） | フェーズ3 |
| 特権・認証 | Kerberos 事前認証が無効なアカウント（AS-REP Roasting） | フェーズ2 |
| アカウント衛生 | パスワード不要 / 可逆暗号化 / パスワード無期限 | フェーズ2〜3 |
| 棚卸し | 長期間未使用のユーザー・コンピューター | フェーズ3 |
| 基盤 | サポート終了 OS の残存 | フェーズ2 |
| 基盤 | AD ごみ箱の有効化状況 | フェーズ1 |
| 基盤 | パスワードポリシー（長さ・ロックアウト・複雑さ） | フェーズ2 |
| 基盤 | Windows LAPS の導入状況（スキーマ属性の有無で推定） | フェーズ2 |
| サイト構成 | サイト未割当のサブネット / DC のいないサイト | フェーズ3 |

### フェーズの考え方

| | 内容 |
|---|---|
| **フェーズ1** | ユーザー影響なし・すぐ着手できる。管理者側の作業だけで完了する |
| **フェーズ2** | 影響範囲の事前調査が必要。止まるものがないか確認してから進める |
| **フェーズ3** | 整理・整頓。急がないが、放置すると調査や移行のたびにコストを払い続ける |

**重要度（High/Medium/Low）ではなく、着手のしやすさで並べているのが要点です。**
実務では「危険な順」に並べても手が止まります。「今日できること」から動けるほうが、結果的に早く改善します。

## スコアについて

100 点からの減点方式です（High 12 点 / Medium 5 点 / Low 2 点）。
ただし重要度ごとに減点の上限を設けています（High 48 / Medium 25 / Low 12）。
上限がないと指摘の多い環境がすべて 0 点になり、改善しても点が動かず比較に使えないためです。

**環境間の優劣を示す指標ではありません。同じ環境の改善前後を比べる目安として使ってください。**

## チェック項目を追加する

各チェックは「スナップショットを受け取り、`New-Finding` で作った検出結果を 0 件以上返す関数」です。
関数を書いて `$script:CheckRegistry` に名前を足せば、レポートに反映されます。

```powershell
function Test-MyCheck {
    param($Snap)
    $items = @( <# 判定して該当分を集める #> )
    if ($items.Count -eq 0) { return }
    New-Finding -Category 'カテゴリ' -Title '見出し' `
        -Severity 'Medium' -Phase 'Phase2' `
        -Detail '何が問題なのか' `
        -Impact '利用者への影響' `
        -Action '推奨する対処' `
        -Reference 'https://learn.microsoft.com/...' `
        -Items $items
}
```

## 既存ツールとの関係

[PingCastle](https://github.com/netwrix/pingcastle) や [ADxRay](https://github.com/ClaudioMerola/ADxRay) など、
優れた無償ツールが既にあります。それらは検出の網羅性で優れています。

このスクリプトの狙いは網羅性ではなく、**日本語で、着手順に、利用者影響とセットで示すこと**です。
既存ツールの出力を持っている場合は、それを読み解く形でも構いません。

## ライセンス

MIT
