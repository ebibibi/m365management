# Microsoft Graph SDK に接続（デバイスコード認証）
Connect-MgGraph -Scopes "Application.ReadWrite.All", "Directory.ReadWrite.All", "CustomSecAttributeAssignment.ReadWrite.All", "Application.Read.All", "CustomSecAttributeAssignment.Read.All" -UseDeviceCode

# 除外対象アプリの ObjectId または DisplayName のリストをここに記述
$excludedAppNames = @(
    "ConditionalAccessTest",
    "Sample App A",
    "Sample App B"
)

# 属性セットと定義名（既に作成済みのものを使用）
$attributeSetName = "isExcludeApps"
$attributeDefinitionName = "isExcludeApp"

# すべての ServicePrincipal を取得（ページング対応）
$allApps = @()
$nextLink = "https://graph.microsoft.com/v1.0/servicePrincipals"
do {
    $result = Invoke-MgGraphRequest -Method GET -Uri $nextLink
    $allApps += $result.value
    $nextLink = $result.'@odata.nextLink'
} while ($nextLink)

Write-Host "合計アプリ数: $($allApps.Count)"

# アプリごとに属性をセット
foreach ($app in $allApps) {
    $isExcluded = $false

    if ($excludedAppNames -contains $app.DisplayName) {
        $isExcluded = $true
    }

    $attributeValue = if ($isExcluded) { "true" } else { "false" }

    # 属性セットの設定
    $params = @{
        customSecurityAttributes = @{
            $attributeSetName = @{
                "@odata.type" = "#Microsoft.DirectoryServices.CustomSecurityAttributeValue"
                $attributeDefinitionName = $attributeValue
            }
        }
    }

    try {
        Update-MgServicePrincipal -ServicePrincipalId $app.Id -BodyParameter $params
        Write-Host "[OK] $($app.DisplayName): $attributeValue"
    } catch {
        Write-Warning "[NG] $($app.DisplayName): $($_.Exception.Message)"
    }
}
