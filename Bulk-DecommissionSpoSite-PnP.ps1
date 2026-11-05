[CmdletBinding()]
Param()

#region Variables

### DEV ###
$AdminSiteUrl              = "https://gocleverpointcom-admin.sharepoint.com"
$SharePointServicesSiteUrl = "https://gocleverpointcom.sharepoint.com/sites/experiment-SharePointServices"
$ClientId                  = "e391b4e0-0151-4aa2-8ce7-dccf4b3921fa"
$SitesFile                 = "$PSScriptRoot\spo-decom-list.csv"
############

### PRD ###
# $AdminSiteUrl              = "https://admin.admin.com"
# $SharePointServicesSiteUrl = "https://gocleverpointcom.sharepoint.com/sites/SharePointServices"
# $ClientId                  = "XXXX-XXX-XXX-XXX-XXX"
# $SitesFile                 = "$PSScriptRoot\spo-decom-list.csv"
############

$AttestationListName = "SPO Attestation Tracker"

#endregion

Function LogStep ($label, $value, $color = "Green")
{
    Write-Host ("    {0,-15}: " -f $label) -NoNewline -ForegroundColor Gray
    Write-Host $value -ForegroundColor $color
}

Function LogLine
{
    Write-Host ("-" * 78) -ForegroundColor DarkGray
}

Function GetM365Group ($SiteUrl)
{
    $tenantSite = Get-PnPTenantSite -Identity $SiteUrl -Connection $admin -ErrorAction Stop
    if ($tenantSite.GroupId -eq [guid]::Empty) { return $null }

    Get-PnPMicrosoft365Group -Identity $tenantSite.GroupId -Connection $admin -ErrorAction Stop
}

Function LockSite ($SiteUrl)
{
    Set-PnPTenantSite -Identity $SiteUrl -LockState NoAccess -Connection $admin -ErrorAction Stop -WarningAction SilentlyContinue
    (Get-PnPTenantSite -Identity $SiteUrl -Connection $admin -ErrorAction Stop).LockState
}

Function RemoveM365Group ($Group)
{
    Remove-PnPMicrosoft365Group -Identity $Group.Id -Connection $admin -ErrorAction Stop
}

Function UpdateInventory ($SiteUrl, $Notes)
{
    $matchedItem = $attestationItems | Where-Object { $_["SiteUrl"] -eq $SiteUrl } | Select-Object -First 1
    if (-not $matchedItem) { return $false }

    $values = @{
        Decomissioned   = "Yes"
        DecomissionDate = (Get-Date)
        AdminNotes      = $Notes
    }
    Set-PnPListItem -List $AttestationListName -Identity $matchedItem.Id -Values $values -Connection $svc -ErrorAction Stop | Out-Null
    return $true
}

Function Main ($SiteUrl)
{
    $group = GetM365Group -SiteUrl $SiteUrl
    if ($group) { LogStep "M365 group" "$($group.DisplayName) ($($group.Id))" }
    else        { LogStep "M365 group" "none connected" Gray }

    $lockState = LockSite -SiteUrl $SiteUrl
    LogStep "Lock" $lockState
    $script:locked++

    $notes = "LOCKED By Script"
    if ($group)
    {
        RemoveM365Group -Group $group
        LogStep "Group delete" "removed"
        $notes = "$notes; M365 group $($group.Id) deleted"
        $script:groupsRemoved++
    }

    if (UpdateInventory -SiteUrl $SiteUrl -Notes $notes) { LogStep "Inventory" "updated" }
    else { LogStep "Inventory" "no row in '$AttestationListName'" DarkYellow }
}

$admin = Connect-PnPOnline -Url $AdminSiteUrl -TenantAdminUrl $AdminSiteUrl -Interactive -ClientId $ClientId -ReturnConnection -WarningAction SilentlyContinue
$svc   = Connect-PnPOnline -Url $SharePointServicesSiteUrl -Interactive -ClientId $ClientId -ReturnConnection -WarningAction SilentlyContinue

$attestationItems = Get-PnPListItem -List $AttestationListName -PageSize 500 -Connection $svc

$sites = @(Import-Csv -Path $SitesFile)
$total = $sites.Count
$locked = 0; $groupsRemoved = 0; $failed = 0

Write-Host ""
Write-Host "Bulk decommission - $total site(s) in $(Split-Path $SitesFile -Leaf)" -ForegroundColor Cyan
LogLine

for ($i = 0; $i -lt $total; $i++)
{
    $url = $sites[$i].URL.Trim().TrimEnd('/')
    Write-Host ""
    Write-Host ("[{0} of {1}] " -f ($i + 1), $total) -NoNewline -ForegroundColor Cyan
    Write-Host $url -ForegroundColor White

    try   { Main -SiteUrl $url }
    catch { LogStep "FAILED" $_.Exception.Message Red; $failed++ }
}

Write-Host ""
LogLine
Write-Host "Summary" -ForegroundColor Cyan
LogStep "Sites" $total White
LogStep "Locked" $locked
LogStep "Groups deleted" $groupsRemoved
LogStep "Failed" $failed $(if ($failed -gt 0) { "Red" } else { "Green" })
LogLine
Write-Host "Lock changes can take a few minutes to take effect." -ForegroundColor DarkGray
if ($failed -gt 0) { exit 1 }
