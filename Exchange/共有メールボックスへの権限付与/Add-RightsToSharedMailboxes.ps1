param (
    [Parameter(Mandatory=$true)]
    [string]$User,

    [Parameter(Mandatory=$true, ParameterSetName="RawInput")]
    [string]$SharedMailboxesRaw,

    [Parameter(Mandatory=$true, ParameterSetName="FileInput")]
    [string]$SharedMailboxesFile
)

$mailboxArray = @()

# ファイルパラメータが指定されている場合はファイルから読み込む
if ($PSCmdlet.ParameterSetName -eq "FileInput" -and -not [string]::IsNullOrWhiteSpace($SharedMailboxesFile)) {
    if (Test-Path $SharedMailboxesFile) {
        $mailboxArray = Get-Content -Path $SharedMailboxesFile | 
                        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | 
                        ForEach-Object { $_.Trim() }
        Write-Host "ファイル '$SharedMailboxesFile' からメールボックスリストを読み込みました。"
    } else {
        Write-Error "指定されたファイル '$SharedMailboxesFile' が見つかりません。"
        exit 1
    }
}
# 生の文字列パラメータが指定されている場合は文字列を分割
else {
    # $SharedMailboxesRaw の最初と最後のダブルクォーテーションを削除
    $cleanedMailboxesString = $SharedMailboxesRaw.Trim('"')

    # 改行で分割して配列にする
    # 複数の種類の改行文字（CR、LF、CRLF）や '>>' で対応
    $mailboxArray = $cleanedMailboxesString -split '\r\n|\n|\r|>>|`n' | 
                    ForEach-Object { $_.Trim() } | 
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
}

Write-Host "処理する共有メールボックス: $($mailboxArray.Count)件"
$mailboxArray | ForEach-Object { Write-Host " - $_" }
Write-Host ""

# 各共有メールボックスに対して Add-RightsToSharedMailbox.ps1 を実行
foreach ($mailbox in $mailboxArray) {
    Write-Host "共有メールボックス '$mailbox' の処理を開始します..."
    try {
        # スクリプトのフルパスを取得
        $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath "Add-RightsToSharedMailbox.ps1"
        
        # Add-RightsToSharedMailbox.ps1 を呼び出し
        & $scriptPath -SharedMailbox $mailbox -User $User
        
        Write-Host "共有メールボックス '$mailbox' の処理が正常に完了しました。"
    } catch {
        Write-Error "共有メールボックス '$mailbox' の処理中にエラーが発生しました: $($_.Exception.Message)"
    }
    Write-Host "" # 区切りとして空行を出力
}

Write-Host "全ての共有メールボックスの処理が完了しました。"