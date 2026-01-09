# Find-OldObjects

Active Directory から長期間ログオンしていないユーザーやコンピューターを検出するスクリプト群です。

## スクリプト一覧

### Find-OldUsers.ps1

1年以上ログオンしていない有効なユーザーアカウントを検出します。

```powershell
.\Find-OldUsers.ps1
```

#### 出力内容

- Name: ユーザー名
- SamAccountName: ログオン名
- LastLogonDate: 最終ログオン日時

#### 除外されるアカウント

以下のシステムアカウントは自動的に除外されます:
- IUSR_*
- IWAM_*
- ASPNET
- Guest
- krbtgt
- DefaultAccount
- WDAGUtilityAccount

### Find-OldComputers.ps1

1年以上ログオンしていないコンピューターアカウントを検出します。

```powershell
.\Find-OldComputers.ps1
```

#### 出力内容

- Name: コンピューター名
- LastLogonDate: 最終ログオン日時

## 前提条件

- Windows PowerShell 5.1 以上、または PowerShell 7 以上
- Active Directory PowerShell モジュール
- ドメインコントローラーまたは RSAT がインストールされた端末

## 注意事項

- これらはオンプレミス Active Directory 用のスクリプトです
- クラウド認証は不要です
- Domain Users 以上の権限で実行可能です
