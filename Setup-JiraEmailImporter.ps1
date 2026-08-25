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

$ConsoleModule = Join-Path -Path $ApplicationRoot -ChildPath 'Modules\Console.psm1'
if (-not (Test-Path -LiteralPath $ConsoleModule -PathType Leaf)) {
    throw "Required module not found: $ConsoleModule"
}
Import-Module $ConsoleModule -Force -ErrorAction Stop

Invoke-UiScript -ScriptBlock {
    $ConfigPath = Join-Path -Path $ApplicationRoot -ChildPath 'config.json'
    if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) {
        throw (
            "Configuration file not found: $ConfigPath`n`n" +
            'Place config.json in the application folder, then re-run setup.'
        )
    }

    try {
        $Config = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
    }
    catch {
        throw (
            "Could not parse '$ConfigPath'.`n`n" +
            'Check that the file is valid JSON (matching braces and quotes).' +
            "`nTechnical detail: $($_.Exception.Message)"
        )
    }

    foreach ($Section in @('modules', 'security', 'jira', 'notes', 'processing', 'mailboxes')) {
        if ($null -eq $Config.PSObject.Properties[$Section]) {
            throw (
                "Required config section '$Section' is missing.`n`n" +
                "Add a '$Section' object to config.json. See the other sections in that file for the expected shape."
            )
        }
    }

    $ModulesDirectory = Resolve-ApplicationPath -Path ([string]$Config.modules.directory) -Root $ApplicationRoot
    foreach ($ModuleName in @('Security', 'JiraAPI')) {
        $ModulePath = Join-Path -Path $ModulesDirectory -ChildPath "$ModuleName.psm1"
        if (-not (Test-Path -LiteralPath $ModulePath -PathType Leaf)) {
            throw (
                "Required module not found: $ModulePath`n`n" +
                "Confirm modules.directory in config.json points at the Modules folder."
            )
        }
        Import-Module $ModulePath -Force -ErrorAction Stop
    }

    $CredentialTarget = [string]$Config.security.credentialTarget
    $JiraUrl = [string]$Config.jira.url

    if ([string]::IsNullOrWhiteSpace($CredentialTarget)) {
        throw (
            "security.credentialTarget is missing from config.json.`n`n" +
            'Set it to a Credential Manager name, for example NotesToJira:JiraPAT.'
        )
    }

    Write-UiHeader 'Jira Email Importer - Setup'
    Write-UiPair 'Configuration' $ConfigPath
    Write-UiPair 'Jira URL' $JiraUrl
    Write-UiPair 'Credential' $CredentialTarget

    Write-UiSection 'Credential'
    if (
        -not (Test-JiraPatSecret -Target $CredentialTarget) -or
        $ReplacePat
    ) {
        Set-JiraPatSecret `
            -Target $CredentialTarget `
            -UserName $JiraUrl `
            -Force:$ReplacePat |
            Out-Null

        Write-UiSuccess 'Jira PAT stored in Windows Credential Manager.'
    }
    else {
        Write-UiSuccess 'Jira PAT already exists in Windows Credential Manager.'
        Write-UiDetail 'Use -ReplacePat if you need to store a new token.'
    }

    $Client = $null
    $SecurePat = $null

    try {
        $SecurePat = Get-JiraPatSecret -Target $CredentialTarget

        $Client = New-JiraClient `
            -BaseUrl $JiraUrl `
            -SecurePat $SecurePat `
            -TimeoutSeconds ([int]$Config.jira.requestTimeoutSeconds)
    }
    finally {
        $SecurePat = $null
    }

    if ($null -eq $Client) {
        throw (
            "The Jira client could not be initialized.`n`n" +
            'Check jira.url in config.json and that a PAT is stored, then retry.'
        )
    }

    try {
        Write-UiSection 'Jira connection'
        $Connection = Test-JiraConnection -Client $Client

        $ConnectedAs = [string]$Connection.DisplayName
        if (-not [string]::IsNullOrWhiteSpace([string]$Connection.EmailAddress)) {
            $ConnectedAs = '{0} <{1}>' -f $Connection.DisplayName, $Connection.EmailAddress
        }
        elseif (-not [string]::IsNullOrWhiteSpace([string]$Connection.UserName)) {
            $ConnectedAs = '{0} ({1})' -f $Connection.DisplayName, $Connection.UserName
        }

        Write-UiSuccess "Connected as $ConnectedAs"
        Write-UiPair 'Jira' ([string]$Connection.JiraUrl)
        Write-UiPair 'Response time' ('{0} ms' -f $Connection.DurationMs)

        $EpicField = Get-JiraEpicLinkField -Client $Client

        if ($null -eq $EpicField -or [string]::IsNullOrWhiteSpace([string]$EpicField.id)) {
            throw (
                "The Jira Epic Link field could not be identified.`n`n" +
                'Confirm Jira Software (epics) is enabled and that the PAT user can browse issues.'
            )
        }

        Write-UiSection 'Epic'
        Write-UiSuccess ("Epic Link field: {0} ({1})" -f $EpicField.name, $EpicField.id)

        $ConfiguredEpicKey = [string]$Config.jira.defaultEpicKey
        if ([string]::IsNullOrWhiteSpace($ConfiguredEpicKey)) {
            throw (
                "jira.defaultEpicKey is missing from config.json.`n`n" +
                'Set it to an Epic issue key, for example MADIP-1234.'
            )
        }

        $DefaultEpic = Test-JiraEpic -Client $Client -EpicKey $ConfiguredEpicKey
        if (-not $DefaultEpic.IsValid) {
            throw (
                "Configured issue '$($DefaultEpic.Key)' is a $($DefaultEpic.IssueTypeName), not an Epic.`n`n" +
                "Set jira.defaultEpicKey in config.json to an Epic key. Current issue: $($DefaultEpic.Summary)"
            )
        }
        Write-UiSuccess ("Default Epic: [{0}] {1}" -f $DefaultEpic.Key, $DefaultEpic.Summary)
        if (-not [string]::IsNullOrWhiteSpace([string]$DefaultEpic.Status)) {
            Write-UiPair 'Epic status' ([string]$DefaultEpic.Status)
        }
        Write-UiPair 'Epic URL' ([string]$DefaultEpic.Url)

        $EnabledMailboxes = @($Config.mailboxes | Where-Object {
            $null -eq $_.PSObject.Properties['enabled'] -or [bool]$_.enabled
        })
        if ($EnabledMailboxes.Count -eq 0) {
            throw (
                "No enabled mailboxes are configured.`n`n" +
                'Set at least one mailboxes[] entry to "enabled": true in config.json.'
            )
        }

        Write-UiSection ("Mailbox targets ({0})" -f $EnabledMailboxes.Count)
        $TestedProjects = @{}
        foreach ($Mailbox in $EnabledMailboxes) {
            $MailboxName = [string]$Mailbox.name
            $ProjectKey = [string]$Mailbox.jira.projectKey
            $ComponentName = [string]$Mailbox.jira.componentName
            $IssueTypeName = [string]$Mailbox.jira.issueType

            if ([string]::IsNullOrWhiteSpace($ProjectKey)) {
                throw (
                    "Mailbox '$MailboxName' has no jira.projectKey.`n`n" +
                    "Set mailboxes[].jira.projectKey for '$MailboxName' in config.json."
                )
            }
            if ([string]::IsNullOrWhiteSpace($ComponentName)) {
                throw (
                    "Mailbox '$MailboxName' needs a jira.componentName in config.json.`n`n" +
                    'Use the component name exactly as it appears in Jira.'
                )
            }
            if ([string]::IsNullOrWhiteSpace($IssueTypeName)) {
                throw (
                    "Mailbox '$MailboxName' needs a jira.issueType in config.json.`n`n" +
                    'Typical values are Task or Story, depending on the project.'
                )
            }

            if (-not $TestedProjects.ContainsKey($ProjectKey)) {
                $EncodedKey = [Uri]::EscapeDataString($ProjectKey)
                $Project = Invoke-JiraRequest -Client $Client -Method Get -Path "rest/api/2/project/$EncodedKey"

                $IssueTypes = @(Get-JiraProjectIssueTypes -Client $Client -ProjectKey $ProjectKey)

                if ($IssueTypes.Count -eq 0) {
                    throw (
                        "Jira returned no creatable issue types for project '$ProjectKey'.`n`n" +
                        'Check that the PAT user has Browse Projects and Create Issues permission on that project.'
                    )
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
                $AvailableTypes = @($ProjectMeta.issuetypes | ForEach-Object {
                    "'$(([string]$_.name).Trim())'"
                } | Sort-Object) -join ', '
                throw (
                    "Issue type '$IssueTypeName' is not available in project '$ProjectKey' for mailbox '$MailboxName'.`n`n" +
                    "Set mailboxes[].jira.issueType to one of: $AvailableTypes"
                )
            }

            $ProjectComponents = @(Get-JiraProjectComponents -Client $Client -ProjectKey $ProjectKey)
            $MatchingComponents = @($ProjectComponents | Where-Object {
                ([string]$_.name).Trim().Equals($ComponentName.Trim(), [StringComparison]::OrdinalIgnoreCase)
            })

            if ($MatchingComponents.Count -eq 0) {
                $AvailableNames = @($ProjectComponents | ForEach-Object {
                    "'$(([string]$_.name).Trim())'"
                } | Sort-Object)
                $AvailableText = if ($AvailableNames.Count -eq 0) {
                    '(none returned - check project permissions)'
                }
                else {
                    $AvailableNames -join ', '
                }
                throw (
                    "Component '$ComponentName' was not found in project '$ProjectKey' for mailbox '$MailboxName'.`n`n" +
                    "Set mailboxes[].jira.componentName to one of: $AvailableText"
                )
            }
            if ($MatchingComponents.Count -gt 1) {
                throw (
                    "More than one component named '$ComponentName' was returned for project '$ProjectKey'.`n`n" +
                    'Rename the duplicate in Jira, or use a unique component name in config.json.'
                )
            }

            $Component = $MatchingComponents[0]
            Write-UiInfo $MailboxName
            Write-UiPair 'Project' ('[{0}] {1}' -f $ProjectMeta.key, $ProjectMeta.name)
            Write-UiPair 'Issue type' ([string]$IssueType.name)
            Write-UiPair 'Component' ('{0} (ID {1})' -f $Component.name, $Component.id)
        }

        Write-UiSection 'Result'
        Write-UiSuccess 'Setup and Jira target validation completed successfully.'
        Write-UiDetail 'You can now run the Notes import (32-bit PowerShell / Lotus Notes host).'
    }
    finally {
        if ($null -ne $Client) {
            Remove-JiraClient -Client $Client
            $Client = $null
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}
