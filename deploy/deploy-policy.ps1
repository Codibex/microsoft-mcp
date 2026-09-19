#Requires -RunAsAdministrator
<#
.SYNOPSIS
  Deploys the admin-owned versioned policy.json (Outlook, Calendar + AI disclosure).

.DESCRIPTION
  Builds policy.json, atomically writes it to the admin-owned system location
  (%ProgramData%\microsoft-mcp\policy.json) and locks it down (Administrators/SYSTEM
  full, Users read-only), so a user-level LLM with file access can neither rewrite
  nor delete it. The MCP server reads this file as its ONLY policy source
  (Messaging__* env vars are ignored by design) and refuses to start with a
  user-writable restrictive policy.

.EXAMPLE
  .\deploy-policy.ps1 -AllowedRecipientDomains firma.de,tochter.firma.de `
    -AiDisclosureText "Hinweis: Dieser Entwurf wurde von einer KI erstellt und muss vor dem Versand geprüft werden."

.EXAMPLE
  .\deploy-policy.ps1 -AllowedRecipientAddresses partner@example.com `
    -AiDisclosureText "Hinweis: Dieser Entwurf wurde von einer KI erstellt und muss vor dem Versand geprüft werden."
#>
[CmdletBinding()]
param(
  [string[]]$AllowedRecipientDomains = @(),

  [string[]]$AllowedRecipientAddresses = @(),

  [string[]]$AllowedAttendeeDomains = @(),

  [string[]]$AllowedAttendeeAddresses = @(),

  [Parameter(Mandatory)]
  [string]$AiDisclosureText,

  [bool]$RequireInternalRecipients = $true,

  [Nullable[bool]]$RequireInternalAttendees = $null,

  [bool]$AiDisclosureEnabled = $true,

  [string]$PolicyPath = (Join-Path $env:ProgramData 'microsoft-mcp\policy.json')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ($RequireInternalRecipients -and $AllowedRecipientDomains.Count -eq 0 -and $AllowedRecipientAddresses.Count -eq 0) {
  throw 'RequireInternalRecipients is set but no AllowedRecipientDomains or AllowedRecipientAddresses were given.'
}
if ($AiDisclosureEnabled -and [string]::IsNullOrWhiteSpace($AiDisclosureText)) {
  throw 'AiDisclosureEnabled is set but AiDisclosureText is empty.'
}

$rawAttendeeDomains = if ($PSBoundParameters.ContainsKey('AllowedAttendeeDomains')) {
  $AllowedAttendeeDomains
} else {
  $AllowedRecipientDomains
}
$rawAttendeeAddresses = if ($PSBoundParameters.ContainsKey('AllowedAttendeeAddresses')) {
  $AllowedAttendeeAddresses
} else {
  $AllowedRecipientAddresses
}
$attendeeDomains = @($rawAttendeeDomains | ForEach-Object { $_.Trim() } | Where-Object { $_ })
$attendeeAddresses = @($rawAttendeeAddresses | ForEach-Object { $_.Trim() } | Where-Object { $_ })
$requireInternalAttendees = if ($null -eq $RequireInternalAttendees) {
  $RequireInternalRecipients
} else {
  [bool]$RequireInternalAttendees
}
if ($requireInternalAttendees -and $attendeeDomains.Count -eq 0 -and $attendeeAddresses.Count -eq 0) {
  throw 'RequireInternalAttendees is set but no AllowedAttendeeDomains or AllowedAttendeeAddresses were given.'
}

$policy = [ordered]@{
  version  = 1
  outlook  = [ordered]@{
    requireInternalRecipients = $RequireInternalRecipients
    allowedRecipientDomains   = @($AllowedRecipientDomains | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    allowedRecipientAddresses = @($AllowedRecipientAddresses | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    aiDisclosureEnabled       = $AiDisclosureEnabled
    aiDisclosureText          = $AiDisclosureText
  }
  calendar = [ordered]@{
    requireInternalAttendees = $requireInternalAttendees
    allowedAttendeeDomains   = $attendeeDomains
    allowedAttendeeAddresses = $attendeeAddresses
  }
  teams = [ordered]@{}
}

$dir = Split-Path -Parent $PolicyPath
if (-not (Test-Path $dir)) {
  New-Item -ItemType Directory -Path $dir | Out-Null
}
$tempPath = "$PolicyPath.tmp.$([guid]::NewGuid().ToString('N'))"
$policy | ConvertTo-Json -Depth 4 | Set-Content -Path $tempPath -Encoding utf8NoBOM

# Lock down file AND directory (a writable directory allows delete-and-replace
# of even a read-only file). SID form: locale-proof.
# S-1-5-32-544 Administrators, S-1-5-18 SYSTEM, S-1-5-32-545 Users.
icacls $dir /inheritance:r /grant:r '*S-1-5-32-544:(OI)(CI)F' '*S-1-5-18:(OI)(CI)F' '*S-1-5-32-545:(OI)(CI)RX' | Out-Null
icacls $tempPath /inheritance:r /grant:r '*S-1-5-32-544:F' '*S-1-5-18:F' '*S-1-5-32-545:R' | Out-Null
if ($LASTEXITCODE -ne 0) {
  throw "icacls failed for $tempPath."
}

if (Test-Path $PolicyPath) {
  [System.IO.File]::Replace($tempPath, $PolicyPath, $null, $true)
} else {
  [System.IO.File]::Move($tempPath, $PolicyPath)
}
Write-Host "Wrote $PolicyPath"

# Verify: Users must not hold any Write-ish right.
$bad = (Get-Acl -Path $PolicyPath).Access | Where-Object {
  $_.IdentityReference.Value -like '*\Users' -and
  ($_.FileSystemRights -band [System.Security.AccessControl.FileSystemRights]::Write) -and
  $_.AccessControlType -eq 'Allow'
}
if ($bad) {
  throw "Protection FAILED: Users still holds write rights on $PolicyPath. Review the ACL manually."
}

Write-Host "Protected $PolicyPath (Administrators/SYSTEM full, Users read-only)."
Get-Acl -Path $PolicyPath | Format-Table -AutoSize | Out-String | Write-Host
