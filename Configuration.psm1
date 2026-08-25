Set-StrictMode -Version Latest

<#
.SYNOPSIS
    Loads and validates Jira Email Importer configuration.

.DESCRIPTION
    Reads Config\config.json, validates required settings, and resolves paths
    relative to the application root. It performs no Jira or Outlook operations.
#>

function Get-ImporterConfigurationValue {
    [CmdletBinding()]
    param(
        [Parameter()]
        [AllowNull()]
        [object]$InputObject,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Name
    )

    if ($null -eq $InputObject) {
        return $null
    }

    if ($InputObject -is [System.Collections.IDictionary]) {
        if ($InputObject.Contains($Name)) {
            return $InputObject[$Name]
        }

        return $null
    }

    $Property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $Property) {
        return $null
    }

    return $Property.Value
}

function Test-ImporterBlankValue {
    [CmdletBinding()]
    param(
        [Parameter()]
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return $true
    }

    return ([string]$Value).Trim().Length -eq 0
}

function Resolve-ImporterPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$ApplicationRoot,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path
    )

    if ($Path -match '^(?:[A-Za-z]:[\\/]|\\\\|/)') {
        return $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    }

    $CombinedPath = Join-Path -Path $ApplicationRoot -ChildPath $Path
    return $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($CombinedPath)
}

function Test-ImporterConfiguration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [psobject]$Configuration
    )

    $Errors = New-Object 'System.Collections.Generic.List[string]'

    $SchemaVersion = Get-ImporterConfigurationValue -InputObject $Configuration -Name 'schemaVersion'
    $Jira = Get-ImporterConfigurationValue -InputObject $Configuration -Name 'jira'
    $Outlook = Get-ImporterConfigurationValue -InputObject $Configuration -Name 'outlook'
    $Processing = Get-ImporterConfigurationValue -InputObject $Configuration -Name 'processing'
    $Logging = Get-ImporterConfigurationValue -InputObject $Configuration -Name 'logging'

    if ($null -eq $SchemaVersion) {
        $Errors.Add('Missing required setting: schemaVersion')
    }
    else {
        try {
            if ([int]$SchemaVersion -ne 1) {
                $Errors.Add("Unsupported schemaVersion '$SchemaVersion'. Expected version 1.")
            }
        }
        catch {
            $Errors.Add("schemaVersion '$SchemaVersion' is not a valid integer.")
        }
    }

    if ($null -eq $Jira) {
        $Errors.Add('Missing required section: jira')
    }
    else {
        $JiraUrl = Get-ImporterConfigurationValue -InputObject $Jira -Name 'url'
        $ProjectKey = Get-ImporterConfigurationValue -InputObject $Jira -Name 'projectKey'
        $IssueType = Get-ImporterConfigurationValue -InputObject $Jira -Name 'defaultIssueType'
        $EpicKey = Get-ImporterConfigurationValue -InputObject $Jira -Name 'defaultEpicKey'
        $Component = Get-ImporterConfigurationValue -InputObject $Jira -Name 'defaultComponent'
        $SecretFile = Get-ImporterConfigurationValue -InputObject $Jira -Name 'secretFile'
        $Timeout = Get-ImporterConfigurationValue -InputObject $Jira -Name 'requestTimeoutSeconds'
        $CacheMinutes = Get-ImporterConfigurationValue -InputObject $Jira -Name 'metadataCacheMinutes'

        if (Test-ImporterBlankValue -Value $JiraUrl) {
            $Errors.Add('Missing required setting: jira.url')
        }
        else {
            try {
                $Uri = New-Object System.Uri -ArgumentList ([string]$JiraUrl)
                if (-not $Uri.IsAbsoluteUri -or $Uri.Scheme -ne 'https') {
                    $Errors.Add('jira.url must be a valid HTTPS URL.')
                }
            }
            catch {
                $Errors.Add('jira.url must be a valid HTTPS URL.')
            }
        }

        if (Test-ImporterBlankValue -Value $ProjectKey) {
            $Errors.Add('Missing required setting: jira.projectKey')
        }
        elseif ([string]$ProjectKey -notmatch '^[A-Za-z][A-Za-z0-9_]*$') {
            $Errors.Add('jira.projectKey is not valid.')
        }

        if (Test-ImporterBlankValue -Value $IssueType) {
            $Errors.Add('Missing required setting: jira.defaultIssueType')
        }

        if (Test-ImporterBlankValue -Value $EpicKey) {
            $Errors.Add('Missing required setting: jira.defaultEpicKey')
        }
        elseif ([string]$EpicKey -notmatch '^[A-Za-z][A-Za-z0-9_]*-[0-9]+$') {
            $Errors.Add("jira.defaultEpicKey '$EpicKey' is not valid.")
        }

        # The default component is optional, but if present it cannot be blank.
        if ($null -ne $Component -and (Test-ImporterBlankValue -Value $Component)) {
            $Errors.Add('jira.defaultComponent must be omitted or contain a component name.')
        }

        if (Test-ImporterBlankValue -Value $SecretFile) {
            $Errors.Add('Missing required setting: jira.secretFile')
        }

        if ($null -eq $Timeout) {
            $Errors.Add('Missing required setting: jira.requestTimeoutSeconds')
        }
        else {
            try {
                if ([int]$Timeout -lt 1) {
                    $Errors.Add('jira.requestTimeoutSeconds must be at least 1.')
                }
            }
            catch {
                $Errors.Add('jira.requestTimeoutSeconds must be an integer.')
            }
        }

        if ($null -eq $CacheMinutes) {
            $Errors.Add('Missing required setting: jira.metadataCacheMinutes')
        }
        else {
            try {
                if ([int]$CacheMinutes -lt 0) {
                    $Errors.Add('jira.metadataCacheMinutes cannot be negative.')
                }
            }
            catch {
                $Errors.Add('jira.metadataCacheMinutes must be an integer.')
            }
        }
    }

    if ($null -eq $Outlook) {
        $Errors.Add('Missing required section: outlook')
    }
    else {
        $MailboxName = Get-ImporterConfigurationValue -InputObject $Outlook -Name 'mailboxName'
        $SourceFolder = Get-ImporterConfigurationValue -InputObject $Outlook -Name 'sourceFolder'
        $MaximumEmails = Get-ImporterConfigurationValue -InputObject $Outlook -Name 'maximumEmailsPerRun'

        if (Test-ImporterBlankValue -Value $MailboxName) {
            $Errors.Add('Missing required setting: outlook.mailboxName')
        }

        if (Test-ImporterBlankValue -Value $SourceFolder) {
            $Errors.Add('Missing required setting: outlook.sourceFolder')
        }

        if ($null -eq $MaximumEmails) {
            $Errors.Add('Missing required setting: outlook.maximumEmailsPerRun')
        }
        else {
            try {
                if ([int]$MaximumEmails -lt 1) {
                    $Errors.Add('outlook.maximumEmailsPerRun must be at least 1.')
                }
            }
            catch {
                $Errors.Add('outlook.maximumEmailsPerRun must be an integer.')
            }
        }
    }

    if ($null -eq $Processing) {
        $Errors.Add('Missing required section: processing')
    }
    else {
        $MaximumCharacters = Get-ImporterConfigurationValue -InputObject $Processing -Name 'maximumDescriptionCharacters'
        if ($null -eq $MaximumCharacters) {
            $Errors.Add('Missing required setting: processing.maximumDescriptionCharacters')
        }
        else {
            try {
                if ([int]$MaximumCharacters -lt 1) {
                    $Errors.Add('processing.maximumDescriptionCharacters must be at least 1.')
                }
            }
            catch {
                $Errors.Add('processing.maximumDescriptionCharacters must be an integer.')
            }
        }
    }

    if ($null -eq $Logging) {
        $Errors.Add('Missing required section: logging')
    }
    else {
        $ValidLogLevels = @('Debug', 'Information', 'Warning', 'Error')
        $LogLevel = Get-ImporterConfigurationValue -InputObject $Logging -Name 'level'
        $RetentionDays = Get-ImporterConfigurationValue -InputObject $Logging -Name 'retentionDays'

        if ([string]$LogLevel -notin $ValidLogLevels) {
            $Errors.Add("logging.level must be one of: $($ValidLogLevels -join ', ').")
        }

        if ($null -eq $RetentionDays) {
            $Errors.Add('Missing required setting: logging.retentionDays')
        }
        else {
            try {
                if ([int]$RetentionDays -lt 1) {
                    $Errors.Add('logging.retentionDays must be at least 1.')
                }
            }
            catch {
                $Errors.Add('logging.retentionDays must be an integer.')
            }
        }
    }

    return [pscustomobject]@{
        IsValid = ($Errors.Count -eq 0)
        Errors  = $Errors.ToArray()
    }
}

function Import-ImporterConfiguration {
    [CmdletBinding()]
    param(
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$ApplicationRoot = (Split-Path -Parent $PSScriptRoot),

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$ConfigPath = 'Config\config.json'
    )

    $ResolvedApplicationRoot = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ApplicationRoot)
    $ResolvedConfigPath = Resolve-ImporterPath -ApplicationRoot $ResolvedApplicationRoot -Path $ConfigPath

    if (-not (Test-Path -LiteralPath $ResolvedConfigPath -PathType Leaf)) {
        throw "Configuration file was not found: $ResolvedConfigPath"
    }

    try {
        $Configuration = Get-Content -LiteralPath $ResolvedConfigPath -Raw -ErrorAction Stop |
            ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        throw "Could not read configuration '$ResolvedConfigPath': $($_.Exception.Message)"
    }

    $Validation = Test-ImporterConfiguration -Configuration $Configuration
    if (-not $Validation.IsValid) {
        $Message = "Configuration validation failed:`r`n - " + ($Validation.Errors -join "`r`n - ")
        throw $Message
    }

    $Configuration.jira.url = ([string]$Configuration.jira.url).TrimEnd('/')

    $DefaultComponent = Get-ImporterConfigurationValue -InputObject $Configuration.jira -Name 'defaultComponent'
    if ($null -ne $DefaultComponent) {
        $Configuration.jira.defaultComponent = ([string]$DefaultComponent).Trim()
    }

    $ResolvedSecretPath = Resolve-ImporterPath `
        -ApplicationRoot $ResolvedApplicationRoot `
        -Path ([string]$Configuration.jira.secretFile)

    $Paths = [pscustomobject]@{
        ApplicationRoot  = $ResolvedApplicationRoot
        ConfigFile       = $ResolvedConfigPath
        ConfigDirectory  = Join-Path $ResolvedApplicationRoot 'Config'
        ModulesDirectory = Join-Path $ResolvedApplicationRoot 'Modules'
        StateDirectory   = Join-Path $ResolvedApplicationRoot 'State'
        LogsDirectory    = Join-Path $ResolvedApplicationRoot 'Logs'
        TestsDirectory   = Join-Path $ResolvedApplicationRoot 'Tests'
        SecretFile       = $ResolvedSecretPath
    }

    foreach ($Directory in @($Paths.StateDirectory, $Paths.LogsDirectory)) {
        if (-not (Test-Path -LiteralPath $Directory -PathType Container)) {
            New-Item -Path $Directory -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }
    }

    return [pscustomobject]@{
        Settings = $Configuration
        Paths    = $Paths
    }
}

Export-ModuleMember -Function @(
    'Import-ImporterConfiguration',
    'Resolve-ImporterPath',
    'Test-ImporterConfiguration'
)
