[CmdletBinding()]
param(
    [Parameter()]
    [string]$ApplicationRoot = $PSScriptRoot,

    [Parameter()]
    [switch]$ReplacePat
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Resolve-ApplicationPath {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Root
    )

    $ExpandedPath = [Environment]::ExpandEnvironmentVariables($Path)
    if ([IO.Path]::IsPathRooted($ExpandedPath)) { return $ExpandedPath }
    return Join-Path -Path $Root -ChildPath $ExpandedPath
}

function ConvertTo-JiraComponentList {
    [CmdletBinding()]
    param([AllowNull()]$Response)

    if ($null -eq $Response) { return @() }
    $Items = @($Response)

    if ($Items.Count -eq 1 -and
        $Items[0].PSObject.Properties['name'] -and
        $Items[0].PSObject.Properties['id'] -and
        $Items[0].name -is [System.Collections.IEnumerable] -and
        $Items[0].name -isnot [string]) {

        $Names = @($Items[0].name)
        $Ids = @($Items[0].id)
        $Expanded = @()

        for ($Index = 0; $Index -lt $Names.Count; $Index++) {
            $Expanded += [pscustomobject]@{
                name = [string]$Names[$Index]
                id   = if ($Index -lt $Ids.Count) { [string]$Ids[$Index] } else { '' }
            }
        }
        return @($Expanded)
    }

    return @($Items)
}

function ConvertFrom-SecureStringToPlainText {
    param([Parameter(Mandatory)][Security.SecureString]$SecureString)

    $Pointer = [IntPtr]::Zero
    try {
        $Pointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
        return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($Pointer)
    }
    finally {
        if ($Pointer -ne [IntPtr]::Zero) {
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($Pointer)
        }
    }
}

$ConfigPath = Join-Path -Path $ApplicationRoot -ChildPath 'config.json'
if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) {
    throw "Configuration file not found: $ConfigPath"
}

try {
    $Config = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
}
catch {
    throw "Could not parse '$ConfigPath'. $($_.Exception.Message)"
}

foreach ($Section in @('modules', 'security', 'jira', 'notes', 'processing', 'mailboxes')) {
    if ($null -eq $Config.PSObject.Properties[$Section]) {
        throw "Required config section '$Section' is missing."
    }
}

$ModulesDirectory = Resolve-ApplicationPath -Path ([string]$Config.modules.directory) -Root $ApplicationRoot
foreach ($ModuleName in @('Configuration', 'Security', 'JiraAPI')) {
    $ModulePath = Join-Path -Path $ModulesDirectory -ChildPath "$ModuleName.psm1"
    if (-not (Test-Path -LiteralPath $ModulePath -PathType Leaf)) {
        throw "Required module not found: $ModulePath"
    }
    Import-Module $ModulePath -Force -ErrorAction Stop
}

$CredentialTarget = [string]$Config.security.credentialTarget

if (:IsNullOrWhiteSpace($CredentialTarget)) {
    throw 'security.credentialTarget is missing from config.json.'
}

Write-Host 'Jira Email Importer setup' -ForegroundColor Cyan
Write-Host "Configuration: $ConfigPath"
Write-Host "Credential target: $CredentialTarget"

if (
    -not (Test-JiraPatSecret -Target $CredentialTarget) -or
    $ReplacePat
) {
    Set-JiraPatSecret `
        -Target $CredentialTarget `
        -UserName ([string]$Config.jira.url) `
        -Force:$ReplacePat |
        Out-Null

    Write-Host `
        'Jira PAT stored in Windows Credential Manager.' `
        -ForegroundColor Green
}
else {
    Write-Host `
        'Jira PAT already exists in Windows Credential Manager.' `
        -ForegroundColor Green
}

$Client = $null
$SecurePat = $null

try {
    $SecurePat = Get-JiraPatSecret `
        -Target $CredentialTarget

    $Client = New-JiraClient `
        -BaseUrl ([string]$Config.jira.url) `
        -SecurePat $SecurePat `
        -TimeoutSeconds ([int]$Config.jira.requestTimeoutSeconds)
}
finally {
    $SecurePat = $null
}

if ($null -eq $Client) {
    throw 'The Jira client could not be initialized.'
}

try {
    $Connection = Test-JiraConnection -Client $Client
    Write-Host "Jira connection successful: $($Connection.DisplayName)" -ForegroundColor Green

    $AllFields = @(Get-JiraFields -Client $Client)
    $EpicField = Get-JiraEpicLinkField -Client $Client -Fields $AllFields

    if ($null -eq $EpicField -or [string]::IsNullOrWhiteSpace([string]$EpicField.id)) {
        throw 'The Jira Epic Link field could not be identified.'
    }
    Write-Host "Epic Link field: $($EpicField.name) ($($EpicField.id))" -ForegroundColor Green

    $ConfiguredEpicKey = [string]$Config.jira.defaultEpicKey
    if ([string]::IsNullOrWhiteSpace($ConfiguredEpicKey)) {
        throw 'jira.defaultEpicKey is missing from config.json.'
    }

    $DefaultEpic = Test-JiraEpic -Client $Client -EpicKey $ConfiguredEpicKey
    if (-not $DefaultEpic.IsValid) {
        throw "$($DefaultEpic.Key) has issue type '$($DefaultEpic.IssueType)', not 'Epic'."
    }
    Write-Host "Default Epic: [$($DefaultEpic.Key)] $($DefaultEpic.Summary)" -ForegroundColor Green

    $EnabledMailboxes = @($Config.mailboxes | Where-Object {
        $null -eq $_.PSObject.Properties['enabled'] -or [bool]$_.enabled
    })
    if ($EnabledMailboxes.Count -eq 0) {
        throw 'No enabled mailboxes are configured.'
    }

    $TestedProjects = @{}
    foreach ($Mailbox in $EnabledMailboxes) {
        $MailboxName = [string]$Mailbox.name
        $ProjectKey = [string]$Mailbox.jira.projectKey
        $ComponentName = [string]$Mailbox.jira.componentName
        $IssueTypeName = [string]$Mailbox.jira.issueType

        if ([string]::IsNullOrWhiteSpace($ProjectKey)) {
            throw "Mailbox '$MailboxName' has no jira.projectKey."
        }
        if ([string]::IsNullOrWhiteSpace($ComponentName)) {
            throw "Mailbox '$MailboxName' needs a jira.componentName in config.json."
        }
        if ([string]::IsNullOrWhiteSpace($IssueTypeName)) {
            throw "Mailbox '$MailboxName' needs a jira.issueType in config.json."
        }

        if (-not $TestedProjects.ContainsKey($ProjectKey)) {
            $EncodedKey = [Uri]::EscapeDataString($ProjectKey)
            $Project = Invoke-JiraRequest -Client $Client -Method Get -Path "rest/api/2/project/$EncodedKey"
            Write-Host "Project accessible: [$($Project.key)] $($Project.name)" -ForegroundColor Green

            $CreateMeta = Invoke-JiraRequest -Client $Client -Method Get -Path "rest/api/2/issue/createmeta/$EncodedKey/issuetypes"
            if ($null -ne $CreateMeta.PSObject.Properties['values']) {
                $IssueTypes = @($CreateMeta.values)
            }
            elseif ($CreateMeta -is [System.Collections.IEnumerable] -and $CreateMeta -isnot [string]) {
                $IssueTypes = @($CreateMeta)
            }
            else {
                $IssueTypes = @()
            }

            if ($IssueTypes.Count -eq 0) {
                throw "Jira returned no creatable issue types for project '$ProjectKey'. Check Browse Projects and Create Issues permissions."
            }

            $TestedProjects[$ProjectKey] = [pscustomobject]@{
                id         = [string]$Project.id
                key        = [string]$Project.key
                name       = [string]$Project.name
                issuetypes = $IssueTypes
            }
        }

        $ProjectMeta = $TestedProjects[$ProjectKey]
        $IssueType = @($ProjectMeta.issuetypes | Where-Object {
            ([string]$_.name).Equals($IssueTypeName, [StringComparison]::OrdinalIgnoreCase)
        }) | Select-Object -First 1

        if ($null -eq $IssueType) {
            throw "Issue type '$IssueTypeName' is unavailable in project '$ProjectKey' for mailbox '$MailboxName'."
        }

        $EncodedProjectKey = [Uri]::EscapeDataString($ProjectKey)
        $ComponentResponse = Invoke-JiraRequest -Client $Client -Method Get -Path "rest/api/2/project/$EncodedProjectKey/components"
        $ProjectComponents = @(ConvertTo-JiraComponentList -Response $ComponentResponse)
        $MatchingComponents = @($ProjectComponents | Where-Object {
            ([string]$_.name).Trim().Equals($ComponentName.Trim(), [StringComparison]::OrdinalIgnoreCase)
        })

        if ($MatchingComponents.Count -eq 0) {
            $AvailableNames = @($ProjectComponents | ForEach-Object {
                "'$(([string]$_.name).Trim())' (ID $($_.id))"
            } | Sort-Object) -join '; '
            throw "Component '$ComponentName' was not found in project '$ProjectKey' for mailbox '$MailboxName'. Available components: $AvailableNames"
        }
        if ($MatchingComponents.Count -gt 1) {
            throw "More than one component named '$ComponentName' was returned for project '$ProjectKey'. Use a unique component name."
        }

        $Component = $MatchingComponents[0]
        Write-Host "Target validated for '$MailboxName': project=$ProjectKey; issueType=$($IssueType.name); component=$($Component.name) ($($Component.id))" -ForegroundColor Green
    }

    Write-Host 'Setup and Jira target validation completed successfully.' -ForegroundColor Green
}
finally {
    if ($null -ne $Client) {
        Remove-JiraClient -Client $Client
        $Client = $null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
