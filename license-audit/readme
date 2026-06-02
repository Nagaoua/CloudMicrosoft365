# M365-FeatureUsage-Report

PowerShell script that queries Microsoft Graph to report **actual workload usage** 
per licensed Microsoft 365 user — and identify unused features and optimization opportunities.

## What it audits
Exchange Online · Teams · SharePoint · OneDrive · M365 Apps · MFA/SSPR (Entra P1)

> Detection is based on **last activity dates**, not Microsoft's boolean flags 
> (`HasOneDrive`, `HasTeams`, etc.) which are known to be unreliable.

## Output
- CSV — one row per user, importable in Excel
- HTML — color-coded dashboard with per-user usage score

## Requirements
- PowerShell 5.1+
- Entra ID App Registration with: `Reports.Read.All`, `User.Read.All`, 
  `UserAuthenticationMethod.Read.All`, `Policy.Read.All`

## Usage
````powershell
.\M365-FeatureUsage-Report.ps1 `
    -TenantId     "contoso.onmicrosoft.com" `
    -ClientId     "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" `
    -ClientSecret "YOUR_SECRET" `
    -Period       30 `
    -OutputPath   "C:\Audit"
````

## License
MIT
