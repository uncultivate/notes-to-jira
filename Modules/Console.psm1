Set-StrictMode -Version Latest

<#
.SYNOPSIS
    Coloured, structured console output for the Notes-to-Jira scripts.

.DESCRIPTION
    Keeps user-facing messages consistent across Setup-JiraEmailImporter.ps1
    and Task-NotesQuery.ps1. Errors are printed as readable text instead of
    PowerShell throw dumps (script, line, CategoryInfo, FullyQualifiedErrorId).
#>

$script:UiRule = ('=' * 64)
$script:UiSubRule = ('-' * 64)

function Write-UiRule {
    [CmdletBinding()]
    param(
        [Parameter()]
        [ValidateSet('Header', 'Section')]
        [string]$Style = 'Section'
    )

    if ($Style -eq 'Header') {
        Write-Host $script:UiRule -ForegroundColor DarkCyan
    }
    else {
        Write-Host $script:UiSubRule -ForegroundColor DarkCyan
    }
}

function Write-UiHeader {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Title
    )

    Write-Host ''
    Write-UiRule -Style Header
    Write-Host " $Title" -ForegroundColor Cyan
    Write-UiRule -Style Header
}

function Write-UiSection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Title
    )

    Write-Host ''
    Write-UiRule -Style Section
    Write-Host " $Title" -ForegroundColor Cyan
}

function Write-UiSuccess {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Message
    )

    Write-Host "  $Message" -ForegroundColor Green
}

function Write-UiInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Message
    )

    Write-Host "  $Message" -ForegroundColor White
}

function Write-UiDetail {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Message
    )

    Write-Host "  $Message" -ForegroundColor DarkGray
}

function Write-UiWarning {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Message
    )

    Write-Host "  $Message" -ForegroundColor Yellow
}

function Write-UiPair {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Label,

        [Parameter()]
        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value,

        [Parameter()]
        [System.ConsoleColor]$ValueColor = [System.ConsoleColor]::White,

        [Parameter()]
        [int]$LabelWidth = 22
    )

    $DisplayValue = if ($null -eq $Value) { '' } else { [string]$Value }
    $PaddedLabel = $Label.PadRight($LabelWidth)
    Write-Host "  $PaddedLabel" -NoNewline -ForegroundColor DarkGray
    Write-Host $DisplayValue -ForegroundColor $ValueColor
}

function Write-UiMetric {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Label,

        [Parameter(Mandatory)]
        [int]$Value,

        [Parameter()]
        [ValidateSet('Neutral', 'Action', 'Skip', 'Problem')]
        [string]$Kind = 'Neutral',

        [Parameter()]
        [int]$LabelWidth = 28
    )

    $Color = [System.ConsoleColor]::White
    if ($Kind -eq 'Action' -and $Value -gt 0) {
        $Color = [System.ConsoleColor]::Green
    }
    elseif ($Kind -eq 'Skip' -and $Value -gt 0) {
        $Color = [System.ConsoleColor]::Yellow
    }
    elseif ($Kind -eq 'Problem' -and $Value -gt 0) {
        $Color = [System.ConsoleColor]::Yellow
    }
    elseif ($Value -eq 0) {
        $Color = [System.ConsoleColor]::DarkGray
    }

    Write-UiPair -Label $Label -Value $Value -ValueColor $Color -LabelWidth $LabelWidth
}

function Write-UiError {
    [CmdletBinding(DefaultParameterSetName = 'Text')]
    param(
        [Parameter(ParameterSetName = 'Record', Mandatory)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord,

        [Parameter(ParameterSetName = 'Text', Mandatory, Position = 0)]
        [AllowEmptyString()]
        [string]$Message
    )

    $Text = if ($PSCmdlet.ParameterSetName -eq 'Record') {
        [string]$ErrorRecord.Exception.Message
    }
    else {
        $Message
    }

    Write-Host ''
    Write-Host ' Error' -ForegroundColor Red
    Write-UiRule -Style Section

    foreach ($Line in ($Text -split '\r?\n')) {
        if ([string]::IsNullOrWhiteSpace($Line)) {
            Write-Host ''
            continue
        }

        if ($Line -match '^\s*Technical detail:') {
            Write-Host "  $Line" -ForegroundColor DarkGray
        }
        elseif ($Line -match '^\s*- ') {
            Write-Host "  $Line" -ForegroundColor Yellow
        }
        elseif ($Line -match '^\s*(Typical causes:|Fix:|Check |Open |Run |Set |Raise |Use |Wait |Make sure|Re-run |Type )') {
            Write-Host "  $Line" -ForegroundColor Yellow
        }
        else {
            Write-Host "  $Line" -ForegroundColor Red
        }
    }

    Write-Host ''
}

function Invoke-UiScript {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [scriptblock]$ScriptBlock
    )

    try {
        & $ScriptBlock
    }
    catch {
        Write-UiError -ErrorRecord $_
        exit 1
    }
}

Export-ModuleMember -Function @(
    'Invoke-UiScript',
    'Write-UiDetail',
    'Write-UiError',
    'Write-UiHeader',
    'Write-UiInfo',
    'Write-UiMetric',
    'Write-UiPair',
    'Write-UiSection',
    'Write-UiSuccess',
    'Write-UiWarning'
)
