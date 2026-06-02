<#
.SYNOPSIS
    M365 License Feature Usage Report
.DESCRIPTION
    Analyse l'utilisation reelle des fonctionnalites incluses dans les licences M365.
    Workloads : Exchange, Teams, SharePoint, OneDrive, M365 Apps, Entra ID P1.
    Produit un rapport CSV + HTML avec scoring par utilisateur.
.PARAMETER Period
    Periode d'analyse : 7, 30, 90 ou 180 jours (defaut : 30)
.PARAMETER OutputPath
    Dossier de sortie pour les rapports
.PARAMETER SkipConnect
    Passer la connexion Graph (si deja connecte)
.EXAMPLE
    .\M365-FeatureUsage-Report.ps1 -Period 30 -OutputPath "C:\Reports"
.NOTES
    Module requis : Install-Module Microsoft.Graph -Scope CurrentUser
    Permissions   : Reports.Read.All, User.Read.All, Directory.Read.All,
                    Policy.Read.All, UserAuthenticationMethod.Read.All, AuditLog.Read.All
#>

[CmdletBinding()]
param(
    [ValidateSet('7','30','90','180')]
    [string]$Period = '30',
    [string]$OutputPath = ".\M365_FeatureUsage_$(Get-Date -Format 'yyyyMMdd_HHmmss')",
    [switch]$SkipConnect
)

Set-StrictMode -Off
$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'

#region -- CONFIGURATION ------------------------------------------------------

$KnownSKUs = @{
    '06ebc4ee-1bb5-47dd-8120-11324bc54e06' = 'Microsoft 365 E5'
    '05e9a617-0261-4cee-bb44-138d3ef5d965' = 'Microsoft 365 E3'
    'c7df2760-2c81-4ef7-b578-5b5392b571df' = 'Office 365 E5'
    '6fd2c87f-b296-42f0-b197-1e91e994b900' = 'Office 365 E3'
    '4b9405b0-7788-4568-add1-99614e613b69' = 'Exchange Online Plan 1'
    '19ec0d23-8335-4cbd-94ac-6050e30712fa' = 'Exchange Online Plan 2'
    'efccb6f7-5641-4e0e-bd10-b4976e1bf68e' = 'Enterprise Mobility + Security E3'
    'b05e124f-c7cc-45a0-a6aa-8cf78c946968' = 'Enterprise Mobility + Security E5'
    '84a661c4-e949-4bd2-a560-ed7766fcaf2b' = 'Microsoft 365 Business Premium'
    'cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46' = 'Microsoft 365 Business Basic'
}

$ColorGreen  = '#1a7a4a'
$ColorOrange = '#e67e22'
$ColorRed    = '#c0392b'
$ColorGray   = '#7f8c8d'
$ColorBlue   = '#2980b9'

#endregion

#region -- FONCTIONS ----------------------------------------------------------

function Write-Step   { param([string]$Msg,[string]$Col='Cyan') Write-Host "  -> $Msg" -ForegroundColor $Col }
function Write-Section { param([string]$T) Write-Host "`n========================================" -ForegroundColor DarkCyan; Write-Host "  $T" -ForegroundColor White; Write-Host "========================================" -ForegroundColor DarkCyan }

function Get-UsageScore {
    param([bool]$IsActive,[bool]$IsLicensed)
    if (-not $IsLicensed) { return 'N/A' }
    if ($IsActive)        { return 'Utilise' }
    return 'Inactif'
}

function Get-ScoreColor {
    param([string]$Score)
    switch ($Score) {
        'Utilise' { return $script:ColorGreen  }
        'Partiel' { return $script:ColorOrange }
        'Inactif' { return $script:ColorRed    }
        default   { return $script:ColorGray   }
    }
}

function Build-TdScore {
    param([string]$Score)
    $c = Get-ScoreColor -Score $Score
    $label = switch ($Score) {
        'Utilise' { 'Utilise' }
        'Inactif' { 'Inactif' }
        'Partiel' { 'Partiel' }
        default   { 'N/A'     }
    }
    return "<td style='text-align:center'><span style='background:$c;color:white;padding:2px 7px;border-radius:10px;font-size:11px'>$label</span></td>"
}

function Build-PctBar {
    param([int]$Pct)
    $c = if ($Pct -ge 70) { $script:ColorGreen } elseif ($Pct -ge 40) { $script:ColorOrange } else { $script:ColorRed }
    return "<div style='background:#ecf0f1;border-radius:4px;height:14px;width:100%'><div style='background:$c;width:${Pct}%;height:14px;border-radius:4px'></div></div><small>${Pct}%</small>"
}

function Build-StatCard {
    param([string]$Title,[int]$Licensed,[int]$Active,[int]$Inactive)
    $pct   = if ($Licensed -gt 0) { [math]::Round(($Active / $Licensed) * 100) } else { 0 }
    $color = if ($pct -ge 70) { $script:ColorGreen } elseif ($pct -ge 40) { $script:ColorOrange } else { $script:ColorRed }
    $bar   = Build-PctBar -Pct $pct
    $html  = "<div class='stat-card'>"
    $html += "<div class='stat-title'>$Title</div>"
    $html += "<div class='stat-value' style='color:$color'>$Active <span style='font-size:14px;color:#7f8c8d'>/ $Licensed</span></div>"
    $html += "<div class='stat-label'>actifs / licencies</div>"
    $html += $bar
    $html += "<div style='margin-top:6px;font-size:11px;color:$script:ColorRed'>Inactifs : $Inactive</div>"
    $html += "</div>"
    return $html
}

function Get-ReportCSV {
    param([string]$Uri)
    try {
        $raw = Invoke-MgGraphRequest -Uri $Uri -Method GET -OutputType HttpResponseMessage
        return $raw.Content.ReadAsStringAsync().Result
    } catch {
        Write-Warning "  [WARN] Rapport indisponible : $Uri --- $($_.Exception.Message)"
        return $null
    }
}

function Test-ActiveByDate {
    param([string]$DateStr, [int]$DaysBack)
    if ([string]::IsNullOrWhiteSpace($DateStr)) { return $false }
    try {
        $d = [datetime]::Parse($DateStr)
        return ($d -ge (Get-Date).AddDays(-$DaysBack))
    } catch { return $false }
}

#endregion

#region -- CONNEXION ----------------------------------------------------------

Write-Section "CONNEXION MICROSOFT GRAPH"

if (-not $SkipConnect) {
    $scopes = @(
        'Reports.Read.All','User.Read.All','Directory.Read.All',
        'Policy.Read.All','UserAuthenticationMethod.Read.All','AuditLog.Read.All'
    )
    Write-Step "Connexion en cours..."
    Connect-MgGraph -Scopes $scopes -NoWelcome
    Write-Step "Connecte." -Col Green
}

$ctx = Get-MgContext
Write-Step "Tenant : $($ctx.TenantId)"
Write-Step "Compte : $($ctx.Account)"

#endregion

#region -- DOSSIER SORTIE ------------------------------------------------------

if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath | Out-Null }
Write-Step "Rapports dans : $OutputPath"

#endregion

#region -- COLLECTE DONNEES ----------------------------------------------------

Write-Section "COLLECTE DES DONNEES (Periode : $Period jours)"

# 1. Utilisateurs licencies
Write-Step "Recuperation des utilisateurs licencies..."
$AllUsers = @()
$uri = "https://graph.microsoft.com/v1.0/users?`$select=id,displayName,userPrincipalName,accountEnabled,assignedLicenses,jobTitle,department,usageLocation&`$top=999&`$filter=assignedLicenses/`$count ne 0&`$count=true"
$resp = Invoke-MgGraphRequest -Uri $uri -Method GET -Headers @{'ConsistencyLevel'='eventual'} -OutputType PSObject
$AllUsers += $resp.value
while ($resp.PSObject.Properties.Name -contains '@odata.nextLink' -and $resp.'@odata.nextLink') {
    $resp = Invoke-MgGraphRequest -Uri $resp.'@odata.nextLink' -Method GET -OutputType PSObject
    $AllUsers += $resp.value
}
Write-Step "$($AllUsers.Count) utilisateurs licencies." -Col Green

# 2. SKUs
Write-Step "Recuperation des SKUs..."
$SKUData = @{}
$skus = Invoke-MgGraphRequest -Uri 'https://graph.microsoft.com/v1.0/subscribedSkus' -Method GET -OutputType PSObject
foreach ($sku in $skus.value) {
    $SKUData[$sku.skuId] = @{
        Name  = if ($KnownSKUs[$sku.skuId]) { $KnownSKUs[$sku.skuId] } else { $sku.skuPartNumber }
        Plans = $sku.servicePlans | ForEach-Object { $_.servicePlanName }
    }
}

# 3. Exchange
Write-Step "Rapport Exchange Online..."
$ExchangeUsage = @{}
$raw = Get-ReportCSV -Uri "https://graph.microsoft.com/v1.0/reports/getMailboxUsageDetail(period='D$Period')"
if ($raw) {
    foreach ($line in ($raw -split "`n" | Select-Object -Skip 1)) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $cols = $line -split ','
        if ($cols.Count -ge 8) {
            $upn = $cols[1].Trim('"')
            $ExchangeUsage[$upn] = @{
                LastActivity = $cols[7].Trim('"')
                ItemCount    = $cols[4].Trim('"')
                IsActive     = ($cols[7].Trim('"') -ne '')
            }
        }
    }
}

# 4. Active Users (tous workloads)
Write-Step "Rapport Active Users..."
$ActiveUD = @{}
$raw = Get-ReportCSV -Uri "https://graph.microsoft.com/v1.0/reports/getOffice365ActiveUserDetail(period='D$Period')"
if ($raw) {
    foreach ($line in ($raw -split "`n" | Select-Object -Skip 1)) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $cols = $line -split ','
        if ($cols.Count -ge 15) {
            $upn = $cols[1].Trim('"')
            $ActiveUD[$upn] = @{
                HasExchange   = ($cols[4].Trim('"')  -eq 'True')
                HasOneDrive   = ($cols[5].Trim('"')  -eq 'True')
                HasSharePoint = ($cols[6].Trim('"')  -eq 'True')
                HasTeams      = ($cols[9].Trim('"')  -eq 'True')
                HasM365Apps   = ($cols[10].Trim('"') -eq 'True')
                LastExchange  = $cols[11].Trim('"')
                LastOneDrive  = $cols[12].Trim('"')
                LastSharePoint= $cols[13].Trim('"')
                LastTeams     = $cols[14].Trim('"')
            }
        }
    }
}

# 5. Teams
Write-Step "Rapport Teams..."
$TeamsUsage = @{}
$raw = Get-ReportCSV -Uri "https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserDetail(period='D$Period')"
if ($raw) {
    foreach ($line in ($raw -split "`n" | Select-Object -Skip 1)) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $cols = $line -split ','
        if ($cols.Count -ge 15) {
            $upn = $cols[1].Trim('"')
            $meetings = [long]($cols[6].Trim('"') -replace '[^0-9]','')
            $chats    = [long]($cols[3].Trim('"') -replace '[^0-9]','') + [long]($cols[4].Trim('"') -replace '[^0-9]','')
            $TeamsUsage[$upn] = @{
                Meetings     = $meetings
                Messages     = $chats
                LastActivity = $cols[14].Trim('"')
                IsActive     = ($cols[14].Trim('"') -ne '')
            }
        }
    }
}

# 6. M365 Apps Activations
Write-Step "Rapport M365 Apps Activations..."
$AppsAct = @{}
$raw = Get-ReportCSV -Uri 'https://graph.microsoft.com/v1.0/reports/getOffice365ActivationsUserDetail'
if ($raw) {
    foreach ($line in ($raw -split "`n" | Select-Object -Skip 1)) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $cols  = $line -split ','
        if ($cols.Count -ge 9) {
            $upn   = $cols[1].Trim('"')
            $total = 0
            for ($i = 4; $i -le 8; $i++) { $total += [long]($cols[$i].Trim('"') -replace '[^0-9]','') }
            $AppsAct[$upn] = @{ TotalActivations = $total; IsActive = ($total -gt 0) }
        }
    }
}

# 7. MFA / SSPR
Write-Step "Entra ID P1 : MFA & SSPR..."
$MFAStatus = @{}
try {
    $mfaResp = Invoke-MgGraphRequest -Uri 'https://graph.microsoft.com/v1.0/reports/authenticationMethods/userRegistrationDetails?$top=999' -Method GET -OutputType PSObject
    foreach ($e in $mfaResp.value) {
        $MFAStatus[$e.userPrincipalName] = @{
            IsMFARegistered  = $e.isMfaRegistered
            IsSSPRRegistered = $e.isSsprRegistered
            Methods          = ($e.methodsRegistered -join ', ')
        }
    }
} catch { Write-Warning "  [WARN] MFA details indisponibles (UserAuthenticationMethod.Read.All requis)" }

# 8. Conditional Access
Write-Step "Conditional Access policies..."
$CACount = 0; $CAEnabled = 0
try {
    $caResp    = Invoke-MgGraphRequest -Uri 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies' -Method GET -OutputType PSObject
    $CACount   = $caResp.value.Count
    $CAEnabled = ($caResp.value | Where-Object { $_.state -eq 'enabled' }).Count
} catch { Write-Warning "  [WARN] CA policies indisponibles (Policy.Read.All requis)" }

#endregion

#region -- ANALYSE PAR UTILISATEUR --------------------------------------------

Write-Section "ANALYSE PAR UTILISATEUR"

$Report  = @()
$counter = 0

foreach ($user in $AllUsers) {
    $counter++
    if ($counter % 50 -eq 0) { Write-Step "Traitement $counter / $($AllUsers.Count)..." }
    $upn = $user.userPrincipalName

    # Licences
    $licNames = @()
    $hasExPlan = $false; $hasTePlan = $false; $hasSPPlan = $false
    $hasODPlan = $false; $hasAPPlan = $false; $hasP1Plan = $false

    foreach ($lic in $user.assignedLicenses) {
        if ($SKUData[$lic.skuId]) {
            $licNames += $SKUData[$lic.skuId].Name
            $plans     = $SKUData[$lic.skuId].Plans -join ','
            if ($plans -match 'EXCHANGE')                              { $hasExPlan = $true }
            if ($plans -match 'TEAMS')                                 { $hasTePlan = $true }
            if ($plans -match 'SHAREPOINT')                            { $hasSPPlan = $true }
            if ($plans -match 'ONEDRIVE')                              { $hasODPlan = $true }
            if ($plans -match 'O365_BUSINESS|OFFICESUBSCRIPTION')      { $hasAPPlan = $true }
            if ($plans -match 'AAD_PREMIUM|MFA_PREMIUM')               { $hasP1Plan = $true }
        } else {
            $licNames += $lic.skuId
        }
    }

    # Scores workloads — bases sur les dates reelles (pas les booleens Microsoft)
    $exLastRaw = if ($ActiveUD[$upn]) { $ActiveUD[$upn].LastExchange   } else { '' }
    $spLastRaw = if ($ActiveUD[$upn]) { $ActiveUD[$upn].LastSharePoint } else { '' }
    $odLastRaw = if ($ActiveUD[$upn]) { $ActiveUD[$upn].LastOneDrive   } else { '' }
    $teLastRaw = if ($TeamsUsage[$upn]) { $TeamsUsage[$upn].LastActivity } elseif ($ActiveUD[$upn]) { $ActiveUD[$upn].LastTeams } else { '' }

    $exActive  = Test-ActiveByDate -DateStr $exLastRaw -DaysBack ([int]$Period)
    $spActive  = Test-ActiveByDate -DateStr $spLastRaw -DaysBack ([int]$Period)
    $odActive  = Test-ActiveByDate -DateStr $odLastRaw -DaysBack ([int]$Period)
    $teActive  = Test-ActiveByDate -DateStr $teLastRaw -DaysBack ([int]$Period)
    $apActive  = if ($AppsAct[$upn])  { $AppsAct[$upn].IsActive } else { $false }

    $exScore = Get-UsageScore -IsActive $exActive -IsLicensed $hasExPlan
    $teScore = Get-UsageScore -IsActive $teActive -IsLicensed $hasTePlan
    $spScore = Get-UsageScore -IsActive $spActive -IsLicensed $hasSPPlan
    $odScore = Get-UsageScore -IsActive $odActive -IsLicensed $hasODPlan
    $apScore = Get-UsageScore -IsActive $apActive -IsLicensed $hasAPPlan

    $mfa     = $MFAStatus[$upn]
    $mfaReg  = if ($mfa) { $mfa.IsMFARegistered  } else { $null }
    $ssprReg = if ($mfa) { $mfa.IsSSPRRegistered } else { $null }
    $p1Score = if (-not $hasP1Plan) { 'N/A' }
               elseif ($mfaReg -and $ssprReg) { 'Utilise' }
               elseif ($mfaReg -or  $ssprReg) { 'Partiel' }
               else { 'Inactif' }

    # Score global
    $scoredWorkloads  = @($exScore,$teScore,$spScore,$odScore,$apScore,$p1Score) | Where-Object { $_ -ne 'N/A' }
    $activeWorkloads  = ($scoredWorkloads | Where-Object { $_ -eq 'Utilise' }).Count
    $totalWorkloads   = $scoredWorkloads.Count
    $globalPct        = if ($totalWorkloads -gt 0) { [math]::Round(($activeWorkloads / $totalWorkloads) * 100) } else { 0 }

    $Report += [PSCustomObject]@{
        DisplayName       = $user.displayName
        UPN               = $upn
        Department        = $user.department
        JobTitle          = $user.jobTitle
        AccountEnabled    = $user.accountEnabled
        Licences          = $licNames -join ' | '
        Exchange_Plan     = $hasExPlan; Exchange_Score = $exScore
        Exchange_LastActivity = if ($ActiveUD[$upn]) { $ActiveUD[$upn].LastExchange } else { '' }
        Teams_Plan        = $hasTePlan; Teams_Score = $teScore
        Teams_Meetings    = if ($TeamsUsage[$upn]) { $TeamsUsage[$upn].Meetings } else { 0 }
        Teams_Messages    = if ($TeamsUsage[$upn]) { $TeamsUsage[$upn].Messages } else { 0 }
        Teams_LastActivity = if ($TeamsUsage[$upn]) { $TeamsUsage[$upn].LastActivity } else { '' }
        SharePoint_Plan   = $hasSPPlan; SharePoint_Score = $spScore
        SharePoint_LastActivity = if ($ActiveUD[$upn]) { $ActiveUD[$upn].LastSharePoint } else { '' }
        OneDrive_Plan     = $hasODPlan; OneDrive_Score = $odScore
        OneDrive_LastActivity = if ($ActiveUD[$upn]) { $ActiveUD[$upn].LastOneDrive } else { '' }
        M365Apps_Plan     = $hasAPPlan; M365Apps_Score = $apScore
        M365Apps_Activations = if ($AppsAct[$upn]) { $AppsAct[$upn].TotalActivations } else { 0 }
        EntraP1_Plan      = $hasP1Plan; EntraP1_Score = $p1Score
        MFA_Registered    = if ($null -ne $mfaReg)  { $mfaReg  } else { 'Inconnu' }
        SSPR_Registered   = if ($null -ne $ssprReg) { $ssprReg } else { 'Inconnu' }
        MFA_Methods       = if ($mfa) { $mfa.Methods } else { '' }
        GlobalUsagePct    = $globalPct
        ActiveWorkloads   = $activeWorkloads
        TotalWorkloads    = $totalWorkloads
    }
}

Write-Step "$($Report.Count) utilisateurs analyses." -Col Green

#endregion

#region -- STATISTIQUES --------------------------------------------------------

$Stats = [PSCustomObject]@{
    TotalUsers              = $Report.Count
    Exchange_Licensed       = ($Report | Where-Object { $_.Exchange_Plan   }).Count
    Exchange_Active         = ($Report | Where-Object { $_.Exchange_Score  -eq 'Utilise' }).Count
    Exchange_Inactive       = ($Report | Where-Object { $_.Exchange_Plan   -and $_.Exchange_Score  -eq 'Inactif' }).Count
    Teams_Licensed          = ($Report | Where-Object { $_.Teams_Plan      }).Count
    Teams_Active            = ($Report | Where-Object { $_.Teams_Score     -eq 'Utilise' }).Count
    Teams_Inactive          = ($Report | Where-Object { $_.Teams_Plan      -and $_.Teams_Score     -eq 'Inactif' }).Count
    SharePoint_Licensed     = ($Report | Where-Object { $_.SharePoint_Plan }).Count
    SharePoint_Active       = ($Report | Where-Object { $_.SharePoint_Score -eq 'Utilise' }).Count
    SharePoint_Inactive     = ($Report | Where-Object { $_.SharePoint_Plan -and $_.SharePoint_Score -eq 'Inactif' }).Count
    OneDrive_Licensed       = ($Report | Where-Object { $_.OneDrive_Plan   }).Count
    OneDrive_Active         = ($Report | Where-Object { $_.OneDrive_Score  -eq 'Utilise' }).Count
    OneDrive_Inactive       = ($Report | Where-Object { $_.OneDrive_Plan   -and $_.OneDrive_Score  -eq 'Inactif' }).Count
    M365Apps_Licensed       = ($Report | Where-Object { $_.M365Apps_Plan   }).Count
    M365Apps_Active         = ($Report | Where-Object { $_.M365Apps_Score  -eq 'Utilise' }).Count
    M365Apps_Inactive       = ($Report | Where-Object { $_.M365Apps_Plan   -and $_.M365Apps_Score  -eq 'Inactif' }).Count
    EntraP1_Licensed        = ($Report | Where-Object { $_.EntraP1_Plan    }).Count
    EntraP1_MFA_Registered  = ($Report | Where-Object { $_.MFA_Registered  -eq $true }).Count
    EntraP1_SSPR_Registered = ($Report | Where-Object { $_.SSPR_Registered -eq $true }).Count
    EntraP1_Inactive        = ($Report | Where-Object { $_.EntraP1_Plan    -and $_.EntraP1_Score   -eq 'Inactif' }).Count
    CA_Total                = $CACount
    CA_Enabled              = $CAEnabled
    AverageScore            = if ($Report.Count -gt 0) { [math]::Round(($Report | Measure-Object -Property GlobalUsagePct -Average).Average) } else { 0 }
}

#endregion

#region -- EXPORT CSV ----------------------------------------------------------

Write-Section "EXPORT DES RAPPORTS"

$csvDetail  = Join-Path $OutputPath 'M365_FeatureUsage_Detail.csv'
$csvSummary = Join-Path $OutputPath 'M365_FeatureUsage_Summary.csv'
$Report | Export-Csv -Path $csvDetail  -NoTypeInformation -Encoding UTF8 -Delimiter ';'
$Stats  | Export-Csv -Path $csvSummary -NoTypeInformation -Encoding UTF8 -Delimiter ';'
Write-Step "CSV detail  : $csvDetail"  -Col Green
Write-Step "CSV summary : $csvSummary" -Col Green

#endregion

#region -- RAPPORT HTML --------------------------------------------------------

Write-Step "Generation du rapport HTML..."

# Construire les lignes du tableau AVANT le here-string
$tableRowsHtml = ''
foreach ($row in ($Report | Sort-Object GlobalUsagePct)) {
    $sc     = if ($row.GlobalUsagePct -ge 70) { $ColorGreen } elseif ($row.GlobalUsagePct -ge 40) { $ColorOrange } else { $ColorRed }
    $accTxt = if ($row.AccountEnabled) { '[OK]' } else { '[OFF]' }
    $mfaTxt = switch ($row.MFA_Registered)  { $true {'Oui'} $false {'Non'} default {'?'} }
    $ssrTxt = switch ($row.SSPR_Registered) { $true {'Oui'} $false {'Non'} default {'?'} }
    $mfaC   = if ($row.MFA_Registered  -eq $true)  { $ColorGreen } else { $ColorRed }
    $ssrC   = if ($row.SSPR_Registered -eq $true)  { $ColorGreen } else { $ColorRed }
    $dept   = if ($row.Department) { $row.Department } else { '-' }
    $lic    = if ($row.Licences)   { $row.Licences   } else { '-' }
    $pct    = $row.GlobalUsagePct

    $tableRowsHtml += "<tr>"
    $tableRowsHtml += "<td>$accTxt $($row.DisplayName)</td>"
    $tableRowsHtml += "<td style='font-size:11px;color:#7f8c8d'>$dept</td>"
    $tableRowsHtml += "<td style='font-size:11px'>$lic</td>"
    $tableRowsHtml += $(Build-TdScore $row.Exchange_Score)
    $tableRowsHtml += $(Build-TdScore $row.Teams_Score)
    $tableRowsHtml += $(Build-TdScore $row.SharePoint_Score)
    $tableRowsHtml += $(Build-TdScore $row.OneDrive_Score)
    $tableRowsHtml += $(Build-TdScore $row.M365Apps_Score)
    $tableRowsHtml += $(Build-TdScore $row.EntraP1_Score)
    $tableRowsHtml += "<td style='text-align:center;color:$mfaC;font-weight:600'>$mfaTxt</td>"
    $tableRowsHtml += "<td style='text-align:center;color:$ssrC;font-weight:600'>$ssrTxt</td>"
    $tableRowsHtml += "<td style='text-align:center;font-weight:bold;color:$sc'>$pct%</td>"
    $tableRowsHtml += "</tr>`n"
}

# Cartes workloads
$cardsHtml  = Build-StatCard -Title 'Exchange Online' -Licensed $Stats.Exchange_Licensed    -Active $Stats.Exchange_Active    -Inactive $Stats.Exchange_Inactive
$cardsHtml += Build-StatCard -Title 'Teams'           -Licensed $Stats.Teams_Licensed       -Active $Stats.Teams_Active       -Inactive $Stats.Teams_Inactive
$cardsHtml += Build-StatCard -Title 'SharePoint'      -Licensed $Stats.SharePoint_Licensed  -Active $Stats.SharePoint_Active  -Inactive $Stats.SharePoint_Inactive
$cardsHtml += Build-StatCard -Title 'OneDrive'        -Licensed $Stats.OneDrive_Licensed    -Active $Stats.OneDrive_Active    -Inactive $Stats.OneDrive_Inactive
$cardsHtml += Build-StatCard -Title 'M365 Apps'       -Licensed $Stats.M365Apps_Licensed    -Active $Stats.M365Apps_Active    -Inactive $Stats.M365Apps_Inactive
$cardsHtml += Build-StatCard -Title 'Entra ID P1'     -Licensed $Stats.EntraP1_Licensed     -Active $Stats.EntraP1_MFA_Registered -Inactive $Stats.EntraP1_Inactive

$avgPct   = $Stats.AverageScore
$avgColor = if ($avgPct -ge 70) { $ColorGreen } elseif ($avgPct -ge 40) { $ColorOrange } else { $ColorRed }
$genDate  = Get-Date -Format 'dd/MM/yyyy a HH:mm'
$tenantId = $ctx.TenantId

$htmlPath = Join-Path $OutputPath 'M365_FeatureUsage_Report.html'

$htmlHeader = @'
<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<title>M365 Feature Usage Report</title>
<style>
body{font-family:'Segoe UI',sans-serif;background:#f4f6f8;margin:0;padding:20px;color:#2c3e50}
h1{color:#2c3e50;font-size:24px;margin-bottom:4px}
.sub{color:#7f8c8d;font-size:13px;margin-bottom:24px}
.kpi-row{display:flex;flex-wrap:wrap;gap:14px;margin-bottom:24px}
.kpi-box{background:white;border-radius:10px;padding:16px 22px;flex:1;min-width:150px;box-shadow:0 2px 8px rgba(0,0,0,.07)}
.kpi-label{font-size:11px;color:#95a5a6;text-transform:uppercase;letter-spacing:1px}
.kpi-value{font-size:30px;font-weight:700;margin:4px 0 2px}
.stat-row{display:flex;flex-wrap:wrap;gap:12px;margin-bottom:24px}
.stat-card{background:white;border-radius:10px;padding:14px;flex:1;min-width:140px;box-shadow:0 2px 8px rgba(0,0,0,.07)}
.stat-title{font-size:13px;font-weight:600;margin-bottom:6px}
.stat-value{font-size:26px;font-weight:700}
.stat-label{font-size:11px;color:#95a5a6;margin-bottom:6px}
table{width:100%;border-collapse:collapse;background:white;border-radius:10px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.07);font-size:12px}
th{background:#2c3e50;color:white;padding:10px 8px;text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.5px}
td{padding:7px 8px;border-bottom:1px solid #ecf0f1}
tr:hover{background:#f8f9fa}
.sec{font-size:16px;font-weight:700;margin:22px 0 10px;border-left:4px solid #2980b9;padding-left:10px}
.info{background:#eaf4fb;border-left:4px solid #2980b9;padding:10px 14px;border-radius:6px;font-size:12px;margin-bottom:18px}
.foot{text-align:center;font-size:11px;color:#bdc3c7;margin-top:28px}
</style>
</head>
<body>
'@

$htmlFooter = '<div class="foot">M365 Feature Usage Report &mdash; Microsoft Graph API &mdash; Donnees sur ' + $Period + ' jours</div></body></html>'

$htmlBody  = "<h1>M365 License Feature Usage Report</h1>`n"
$htmlBody += "<div class='sub'>Periode : $Period jours &nbsp;&bull;&nbsp; Genere le $genDate &nbsp;&bull;&nbsp; Tenant : $tenantId</div>`n"
$htmlBody += "<div class='info'>Ce rapport analyse si les fonctionnalites incluses dans les licences sont reellement utilisees &mdash; independamment du nombre de licences attribuees. Score faible = cout sans valeur.</div>`n"

$htmlBody += "<div class='sec'>Vue d'ensemble</div>`n"
$htmlBody += "<div class='kpi-row'>`n"
$htmlBody += "<div class='kpi-box'><div class='kpi-label'>Utilisateurs licencies</div><div class='kpi-value' style='color:$ColorBlue'>$($Stats.TotalUsers)</div></div>`n"
$htmlBody += "<div class='kpi-box'><div class='kpi-label'>Score global utilisation</div><div class='kpi-value' style='color:$avgColor'>$avgPct%</div></div>`n"
$htmlBody += "<div class='kpi-box'><div class='kpi-label'>Politiques CA actives</div><div class='kpi-value' style='color:$ColorBlue'>$($Stats.CA_Enabled) / $($Stats.CA_Total)</div></div>`n"
$htmlBody += "<div class='kpi-box'><div class='kpi-label'>MFA enregistres</div><div class='kpi-value' style='color:$ColorGreen'>$($Stats.EntraP1_MFA_Registered)</div></div>`n"
$htmlBody += "<div class='kpi-box'><div class='kpi-label'>SSPR enregistres</div><div class='kpi-value' style='color:$ColorGreen'>$($Stats.EntraP1_SSPR_Registered)</div></div>`n"
$htmlBody += "</div>`n"

$htmlBody += "<div class='sec'>Utilisation par Workload</div>`n"
$htmlBody += "<div class='stat-row'>$cardsHtml</div>`n"

$htmlBody += "<div class='sec'>Detail par Utilisateur</div>`n"
$htmlBody += "<table><thead><tr>"
$htmlBody += "<th>Utilisateur</th><th>Departement</th><th>Licence(s)</th>"
$htmlBody += "<th>Exchange</th><th>Teams</th><th>SharePoint</th><th>OneDrive</th><th>M365 Apps</th><th>Entra P1</th>"
$htmlBody += "<th>MFA</th><th>SSPR</th><th>Score</th>"
$htmlBody += "</tr></thead><tbody>`n"
$htmlBody += $tableRowsHtml
$htmlBody += "</tbody></table>`n"

($htmlHeader + $htmlBody + $htmlFooter) | Out-File -FilePath $htmlPath -Encoding UTF8
Write-Step "HTML : $htmlPath" -Col Green

#endregion

#region -- RESUME CONSOLE ------------------------------------------------------

Write-Section "RESULTATS"
Write-Host ""
Write-Host "  Utilisateurs analyses      : $($Report.Count)" -ForegroundColor White

$scoreDisplayColor = 'Green'
if ($avgPct -lt 70) { $scoreDisplayColor = 'Yellow' }
if ($avgPct -lt 40) { $scoreDisplayColor = 'Red'    }
Write-Host "  Score global d'utilisation : $avgPct %" -ForegroundColor $scoreDisplayColor

Write-Host ""
Write-Host "  WORKLOAD          Licencies  Actifs  Inactifs" -ForegroundColor DarkCyan
Write-Host "  --------------------------------------------"   -ForegroundColor DarkGray
Write-Host "  Exchange Online : $($Stats.Exchange_Licensed.ToString().PadRight(8)) $($Stats.Exchange_Active.ToString().PadRight(7)) $($Stats.Exchange_Inactive)"
Write-Host "  Teams           : $($Stats.Teams_Licensed.ToString().PadRight(8)) $($Stats.Teams_Active.ToString().PadRight(7)) $($Stats.Teams_Inactive)"
Write-Host "  SharePoint      : $($Stats.SharePoint_Licensed.ToString().PadRight(8)) $($Stats.SharePoint_Active.ToString().PadRight(7)) $($Stats.SharePoint_Inactive)"
Write-Host "  OneDrive        : $($Stats.OneDrive_Licensed.ToString().PadRight(8)) $($Stats.OneDrive_Active.ToString().PadRight(7)) $($Stats.OneDrive_Inactive)"
Write-Host "  M365 Apps       : $($Stats.M365Apps_Licensed.ToString().PadRight(8)) $($Stats.M365Apps_Active.ToString().PadRight(7)) $($Stats.M365Apps_Inactive)"
Write-Host "  Entra ID P1     : $($Stats.EntraP1_Licensed.ToString().PadRight(8)) $($Stats.EntraP1_MFA_Registered.ToString().PadRight(7)) $($Stats.EntraP1_Inactive)"
Write-Host ""
Write-Host "  MFA enregistres     : $($Stats.EntraP1_MFA_Registered) / $($Stats.TotalUsers)"
Write-Host "  SSPR enregistres    : $($Stats.EntraP1_SSPR_Registered) / $($Stats.TotalUsers)"
Write-Host "  CA policies actives : $($Stats.CA_Enabled) / $($Stats.CA_Total)"
Write-Host ""
Write-Host "  Rapport HTML  : $htmlPath"   -ForegroundColor Cyan
Write-Host "  Detail CSV    : $csvDetail"  -ForegroundColor Cyan
Write-Host "  Resume CSV    : $csvSummary" -ForegroundColor Cyan
Write-Host ""

#endregion

if (-not $SkipConnect) { Disconnect-MgGraph | Out-Null; Write-Step "Deconnecte." -Col Gray }
Write-Host "  OK - Analyse terminee.`n" -ForegroundColor Green
