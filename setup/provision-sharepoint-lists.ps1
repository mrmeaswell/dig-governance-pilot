<#
DIG Governance Pilot: Microsoft 365 setup (run by CTS / a SharePoint site owner)

WHAT THIS DOES
  Creates the four SharePoint lists the site at https://dig.bknown.pro reads and writes
  through Microsoft Graph. The data stays in RTC's Microsoft 365 tenant. The web page
  only holds code, never records.

ONE-TIME STEPS
  1. Pick or create a SharePoint site for DIG (e.g. https://<tenant>.sharepoint.com/sites/DIG).
     Grant DIG members Edit on the site. Faculty who only submit issues need
     Contribute on the "DIG Data Issues" list and Read on the other three.

  2. Register the web app in Entra ID (portal.azure.com > App registrations > New):
       - Name: DIG Governance Pilot
       - Supported account types: Accounts in this organizational directory only
       - Platform: Single-page application (SPA)
       - Redirect URI: https://dig.bknown.pro/
     Then go to API permissions and add Microsoft Graph delegated permissions:
       - User.Read
       - Sites.ReadWrite.All
     Grant admin consent if tenant policy requires it.
     (Delegated means the site can only do what the signed-in user can already do in SharePoint.)

  3. Run this script to create the lists:
       Install-Module Microsoft.Graph.Sites -Scope CurrentUser
       ./provision-sharepoint-lists.ps1 -SiteHostname "<tenant>.sharepoint.com" -SitePath "/sites/DIG"

  4. Put four values into the M365_CONFIG block at the top of the script in index.html:
       tenantId, clientId (from step 2), siteHostname, sitePath (from step 3).
     Commit and push. The site switches from "Local preview" to "Shared mode" and
     shows a "Load starter data" button the first time someone signs in.
#>
param(
  [Parameter(Mandatory = $true)] [string] $SiteHostname,   # e.g. rtcedu.sharepoint.com
  [string] $SitePath = ""                                   # e.g. /sites/DIG ; blank = root site
)

$ErrorActionPreference = "Stop"
Import-Module Microsoft.Graph.Sites
Connect-MgGraph -Scopes "Sites.Manage.All" | Out-Null

$siteRef = if ($SitePath) { "$($SiteHostname):$SitePath" } else { $SiteHostname }
$site = Get-MgSite -SiteId $siteRef
Write-Host "Using site: $($site.WebUrl)"

function Line($name)  { @{ name = $name; text = @{} } }
function Multi($name) { @{ name = $name; text = @{ allowMultipleLines = $true; linesForEditing = 6 } } }

# Column internal names must match FIELD_MAP in index.html. "Title" is built in.
$lists = @(
  @{ displayName = "DIG Data Dictionary"; columns = @(
      (Multi "Definition"), (Line "SourceSystem"), (Multi "CalcMethod"), (Line "Domain"),
      (Line "OwnerName"), (Line "LastReview"), (Line "SeedKey")) },
  @{ displayName = "DIG Workstreams"; columns = @(
      (Multi "WorkDescription"), (Line "Domain"), (Line "Phase"), (Line "Status"),
      (Line "OwnerName"), (Multi "Notes"), (Line "SeedKey")) },
  @{ displayName = "DIG Decisions"; columns = @(
      (Multi "Detail"), (Line "OwnerName"), (Line "Priority"), (Line "Status"),
      (Multi "Notes"), (Multi "Dissent"), (Line "ResolvedDate"), (Line "SeedKey")) },
  @{ displayName = "DIG Data Issues"; columns = @(
      (Line "Submitter"), (Line "Program"), (Line "Metric"), (Line "DateObserved"),
      (Multi "IssueDescription"), (Line "Status"), (Multi "IRNotes"),
      (Line "DateSubmitted"), (Line "AckDate"), (Line "ResolvedDate")) }
)

$existing = Get-MgSiteList -SiteId $site.Id -All | Select-Object -ExpandProperty DisplayName
foreach ($l in $lists) {
  if ($existing -contains $l.displayName) {
    Write-Host "Exists, skipped: $($l.displayName)"
    continue
  }
  New-MgSiteList -SiteId $site.Id -BodyParameter @{
    displayName = $l.displayName
    columns     = $l.columns
    list        = @{ template = "genericList" }
  } | Out-Null
  Write-Host "Created: $($l.displayName)"
}

Write-Host ""
Write-Host "Done. Add these to M365_CONFIG in index.html:"
Write-Host "  siteHostname: '$SiteHostname'"
Write-Host "  sitePath:     '$SitePath'"
Write-Host "  tenantId:     '$((Get-MgContext).TenantId)'"
Write-Host "  clientId:     '<Application (client) ID from the app registration>'"
