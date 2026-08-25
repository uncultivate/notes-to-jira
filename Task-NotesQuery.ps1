# Title: Task-NotesQuery-Safe
# Purpose: Preview and confirm Notes-to-Jira actions, prevent repeat processing,
#          create one Jira issue per Notes conversation, and add replies as comments.

[CmdletBinding()]
param(
    [Parameter()][string]$ApplicationRoot = $PSScriptRoot,
    [Parameter()][switch]$ReplacePat,
    [Parameter()][switch]$DryRun,
    [Parameter()][ValidateRange(0, 10000)][int]$MaxIssues = 0,
    [Parameter()][switch]$Force
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0

function Resolve-ApplicationPath {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Root)
    $Expanded = [Environment]::ExpandEnvironmentVariables($Path)
    if ([IO.Path]::IsPathRooted($Expanded)) { return $Expanded }
    return Join-Path $Root $Expanded
}

function Normalize-CategoryPath {
    param([AllowNull()][string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    return (($Path.Trim() -replace '/', '\') -replace '\+', '\').TrimEnd('\')
}

function Test-CategoryPathMatch {
    param(
        [Parameter(Mandatory)][string]$ActualPath,
        [Parameter(Mandatory)][string[]]$ConfiguredPaths,
        [Parameter()][bool]$IncludeDescendants = $false
    )
    foreach ($ConfiguredPath in $ConfiguredPaths) {
        if ($ActualPath.Equals($ConfiguredPath, [StringComparison]::OrdinalIgnoreCase)) { return $true }
        if ($IncludeDescendants -and $ActualPath.StartsWith("$ConfiguredPath\", [StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

function Get-OptionalPropertyValue {
    param([AllowNull()]$Object, [Parameter(Mandatory)][string]$Name)
    if ($null -eq $Object) { return $null }
    $Property = $Object.PSObject.Properties[$Name]
    if ($null -eq $Property) { return $null }
    return $Property.Value
}

function Get-ConfiguredEpicKey {
    param([Parameter(Mandatory)]$Mailbox, [Parameter(Mandatory)]$Config)

    # Prefer a mailbox-specific Epic, then fall back to a global Jira setting.
    $MailboxJira = Get-OptionalPropertyValue -Object $Mailbox -Name 'jira'
    $Candidates = @(
        (Get-OptionalPropertyValue -Object $MailboxJira -Name 'epicKey'),
        (Get-OptionalPropertyValue -Object $MailboxJira -Name 'defaultEpicKey'),
        (Get-OptionalPropertyValue -Object $Config.jira -Name 'defaultEpicKey'),
        (Get-OptionalPropertyValue -Object $Config.jira -Name 'epicKey')
    )

    foreach ($Candidate in $Candidates) {
        $Value = [string]$Candidate
        if (-not [string]::IsNullOrWhiteSpace($Value)) { return $Value.Trim() }
    }

    throw (
        "No Epic key is configured for mailbox '$([string]$Mailbox.name)'. " +
        'Set mailboxes[].jira.epicKey (preferred) or jira.defaultEpicKey in config.json.'
    )
}

function Get-JiraComponentByName {
    param(
        [Parameter(Mandatory)][psobject]$Client,
        [Parameter(Mandatory)][string]$ProjectKey,
        [Parameter(Mandatory)][string]$ComponentName
    )

    $Components = @(Get-JiraProjectComponents -Client $Client -ProjectKey $ProjectKey)
    $Matches = @($Components | Where-Object {
        ([string]$_.name).Trim().Equals($ComponentName.Trim(), [StringComparison]::OrdinalIgnoreCase)
    })
    if ($Matches.Count -eq 0) { throw "Jira component '$ComponentName' was not found in project '$ProjectKey'." }
    if ($Matches.Count -gt 1) { throw "More than one Jira component is named '$ComponentName' in project '$ProjectKey'." }
    return $Matches[0]
}

function New-JiraIssue {
    param(
        [Parameter(Mandatory)][psobject]$Client,
        [Parameter(Mandatory)]$Mailbox,
        [Parameter(Mandatory)][string]$ComponentId,
        [Parameter(Mandatory)][string]$EpicFieldId,
        [Parameter(Mandatory)][string]$EpicKey,
        [Parameter(Mandatory)]$Email
    )

    $Title = [string]$Email.subject
    foreach ($Text in @($Config.processing.subjectTextToRemove)) {
        $Title = $Title -replace [regex]::Escape([string]$Text), ''
    }
    $Title = ($Title -replace '"', '').Trim()
    if ([string]::IsNullOrWhiteSpace($Title)) { $Title = [string]$Config.processing.emptySubjectText }

    $Created = [datetime]$Email.creationtime
    $Description = @"
<p><i>$($Config.jira.description.introduction)</i><br><br>
<b>For action: </b>$($Config.jira.description.action)<br><br>
<b>From: </b>$($Email.authors)<br>
<b>Notes category: </b>$($Email.categoryPath)<br><br>
$($Email.emailBody)<br><br>
<a href="$($Email.notesurl)">$($Config.jira.description.linkText)</a></p>
"@

    $Fields = @{
        project     = @{ key = [string]$Mailbox.jira.projectKey }
        labels      = @($Mailbox.jira.labels | ForEach-Object { [string]$_ })
        components  = @(@{ id = $ComponentId })
        summary     = $Title
        description = $Description
        duedate     = $Created.AddDays([int]$Config.jira.dueAfterDays).ToString('yyyy-MM-dd')
        issuetype   = @{ name = [string]$Mailbox.jira.issueType }
    }

    if ([string]::IsNullOrWhiteSpace($EpicFieldId)) { throw 'The Epic Link field ID is blank.' }
    if ([string]::IsNullOrWhiteSpace($EpicKey)) { throw 'The configured Epic key is blank.' }

    $Fields[$EpicFieldId] = $EpicKey
    $StartField = [string](Get-OptionalPropertyValue -Object $Config.jira -Name 'startDateFieldId')
    if (-not [string]::IsNullOrWhiteSpace($StartField)) {
        $Fields[$StartField] = $Created.ToString('yyyy-MM-dd')
    }

    return Invoke-JiraRequest -Client $Client -Method Post -Path 'rest/api/2/issue' -Body @{ fields = $Fields }
}

function Add-JiraEmailComment {
    param(
        [Parameter(Mandatory)][psobject]$Client,
        [Parameter(Mandatory)][string]$IssueKey,
        [Parameter(Mandatory)]$Email
    )
    $Comment = @"
Email reply received

From: $($Email.authors)
Received: $($Email.creationtime)
Subject: $($Email.subject)
Notes category: $($Email.categoryPath)

$($Email.emailBody)

View email in Notes: $($Email.notesurl)
"@
    return Invoke-JiraRequest -Client $Client -Method Post -Path "rest/api/2/issue/$IssueKey/comment" -Body @{ body = $Comment }
}

function Save-State {
    param([Parameter(Mandatory)][Collections.Generic.List[object]]$State, [Parameter(Mandatory)][string]$Path)
    $State | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
}

if ([Environment]::Is64BitProcess) {
    throw 'Run this script with the root-level launcher or 32-bit Windows PowerShell.'
}

$ConfigPath = Join-Path $ApplicationRoot 'config.json'
if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) { throw "Configuration not found: $ConfigPath" }
$Config = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json

$ModulesDirectory = Resolve-ApplicationPath -Path ([string]$Config.modules.directory) -Root $ApplicationRoot
foreach ($ModuleName in @('Configuration', 'Security', 'JiraAPI')) {
    $ModulePath = Join-Path $ModulesDirectory "$ModuleName.psm1"
    if (-not (Test-Path -LiteralPath $ModulePath -PathType Leaf)) { throw "Required module not found: $ModulePath" }
    Import-Module $ModulePath -Force -ErrorAction Stop
}

# Retrieve the Windows Credential Manager target name from config.json.
$CredentialTarget = [string]$Config.security.credentialTarget

if (:IsNullOrWhiteSpace($CredentialTarget)) {
    throw (
        'security.credentialTarget is missing or empty in config.json. ' +
        'For example: NotesToJira:JiraPAT'
    )
}

Write-Host "Jira credential target: $CredentialTarget"

# Store the PAT if it does not exist, or replace it when -ReplacePat is used.
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
        'Jira PAT found in Windows Credential Manager.' `
        -ForegroundColor Green
}

$SecurePat = $null
$JiraClient = $null

try {
    # Get-JiraPatSecret returns a SecureString.
    $SecurePat = Get-JiraPatSecret `
        -Target $CredentialTarget

    # Your config uses requestTimeoutSeconds, not timeoutSeconds.
    $TimeoutSeconds = 60

    if ($Config.jira.PSObject.Properties['requestTimeoutSeconds']) {
        $TimeoutSeconds =
            [int]$Config.jira.requestTimeoutSeconds
    }

    $MaximumRetryCount = 2

    if ($Config.jira.PSObject.Properties['maximumRetryCount']) {
        $MaximumRetryCount =
            [int]$Config.jira.maximumRetryCount
    }

    $JiraClient = New-JiraClient `
        -BaseUrl ([string]$Config.jira.url) `
        -SecurePat $SecurePat `
        -TimeoutSeconds $TimeoutSeconds `
        -MaximumRetryCount $MaximumRetryCount

    if ($null -eq $JiraClient) {
        throw 'The Jira client could not be initialized.'
    }
}
finally {
    # Release the script's reference to the SecureString as soon as the
    # Jira client has been initialized.
    $SecurePat = $null
}

try {
    # Resolve the Jira Epic Link custom field once. This replaces the undefined
    # $EpicField variable in the previous script.
    $EpicField = Get-JiraEpicLinkField -Client $JiraClient
    $EpicFieldId = [string]$EpicField.id
    if ([string]::IsNullOrWhiteSpace($EpicFieldId)) {
        throw 'Jira Epic Link discovery returned a field without an ID.'
    }
    Write-Host "Using Jira Epic Link field '$([string]$EpicField.name)' ($EpicFieldId)." -ForegroundColor Green

    foreach ($Mailbox in @($Config.mailboxes)) {
        if ($Mailbox.PSObject.Properties['enabled'] -and -not [bool]$Mailbox.enabled) { continue }

        $MailboxName = [string]$Mailbox.name
        $ProjectKey = [string]$Mailbox.jira.projectKey
        $EpicKey = Get-ConfiguredEpicKey -Mailbox $Mailbox -Config $Config
        $ValidatedEpic = Test-JiraEpic -Client $JiraClient -EpicKey $EpicKey
        if (-not $ValidatedEpic.IsValid) {
            throw "Configured issue '$EpicKey' is type '$($ValidatedEpic.IssueTypeName)', not Epic."
        }

        $Component = Get-JiraComponentByName `
            -Client $JiraClient `
            -ProjectKey $ProjectKey `
            -ComponentName ([string]$Mailbox.jira.componentName)

        Write-Host "Using Jira component '$($Component.name)' ($($Component.id)) in project '$ProjectKey'." -ForegroundColor Green
        Write-Host "Using Jira Epic '$EpicKey': $($ValidatedEpic.Summary)" -ForegroundColor Green

        $StatePath = Resolve-ApplicationPath -Path ([string]$Mailbox.csvPath) -Root $ApplicationRoot
        $StateDirectory = Split-Path -Parent $StatePath
        if (-not (Test-Path -LiteralPath $StateDirectory)) { $null = New-Item -ItemType Directory -Path $StateDirectory -Force }

        $State = [Collections.Generic.List[object]]::new()
        $Processed = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        $IssueByDocument = @{}
        if (Test-Path -LiteralPath $StatePath -PathType Leaf) {
            foreach ($ImportedRow in @(Import-Csv -LiteralPath $StatePath)) {
                $GetValue = {
                    param($Object, [string]$PropertyName)
                    $Property = $Object.PSObject.Properties[$PropertyName]
                    if ($null -eq $Property -or $null -eq $Property.Value) { return '' }
                    return [string]$Property.Value
                }

                $Id = & $GetValue $ImportedRow 'universalid'
                $ParentId = & $GetValue $ImportedRow 'parent'
                $RootId = & $GetValue $ImportedRow 'rootUniversalId'
                $JiraKey = & $GetValue $ImportedRow 'jiraIssueKey'
                $Row = [pscustomobject][ordered]@{
                    universalid = $Id
                    parent = $ParentId
                    rootUniversalId = $RootId
                    jiraIssueKey = $JiraKey
                    jiraAction = (& $GetValue $ImportedRow 'jiraAction')
                    categoryPath = (& $GetValue $ImportedRow 'categoryPath')
                    subject = (& $GetValue $ImportedRow 'subject')
                    creationtime = (& $GetValue $ImportedRow 'creationtime')
                    notesurl = (& $GetValue $ImportedRow 'notesurl')
                    emailBody = (& $GetValue $ImportedRow 'emailBody')
                    authors = (& $GetValue $ImportedRow 'authors')
                    processedAt = (& $GetValue $ImportedRow 'processedAt')
                }
                $State.Add($Row)
                if (-not [string]::IsNullOrWhiteSpace($Id)) { $null = $Processed.Add($Id) }
                if (-not [string]::IsNullOrWhiteSpace($JiraKey)) {
                    $IssueByDocument[$Id] = $JiraKey
                    if (-not [string]::IsNullOrWhiteSpace($RootId)) { $IssueByDocument[$RootId] = $JiraKey }
                }
            }
        }

        $IncludeDescendants = $Mailbox.PSObject.Properties['includeDescendantCategories'] -and [bool]$Mailbox.includeDescendantCategories
        $IncludedPaths = @($Mailbox.includedCategoryPaths | ForEach-Object { Normalize-CategoryPath ([string]$_) })
        $JiraPaths = @($Mailbox.jiraCategoryPaths | ForEach-Object { Normalize-CategoryPath ([string]$_) })

        $Session = $Database = $View = $Document = $null
        $Candidates = [Collections.Generic.List[object]]::new()
        $SkippedSensitive = 0
        $SkippedProcessed = 0
        try {
            Write-Host "`nOpening Notes mailbox: $MailboxName" -ForegroundColor Cyan
            $Session = New-Object -ComObject Lotus.NotesSession
            $Session.Initialize()
            $Database = $Session.GetDatabase([string]$Config.notes.server, [string]$Mailbox.databasePath)
            if ($null -eq $Database -or -not $Database.IsOpen) { throw "Could not open Notes database '$($Mailbox.databasePath)'." }
            $View = $Database.GetView([string]$Config.notes.viewName)
            if ($null -eq $View) { throw "Notes view '$($Config.notes.viewName)' was not found." }
            Write-Host "Database: $($Database.Title); view: $($View.Name); documents: $($View.AllEntries.Count)" -ForegroundColor Green

            $Document = $View.GetFirstDocument()
            while ($null -ne $Document) {
                $Next = $View.GetNextDocument($Document)
                $CategoryPath = if ($Document.ColumnValues.Count -gt 0) { Normalize-CategoryPath ([string]$Document.ColumnValues[0]) } else { '' }
                if (Test-CategoryPathMatch -ActualPath $CategoryPath -ConfiguredPaths $IncludedPaths -IncludeDescendants $IncludeDescendants) {
                    $SubjectItem = $Document.Items | Where-Object { $_.Name -eq 'Subject' } | Select-Object -First 1
                    if ($null -ne $SubjectItem) {
                        $FromItem = $Document.Items | Where-Object { $_.Name -eq 'From' } | Select-Object -First 1
                        $InetFromItem = $Document.Items | Where-Object { $_.Name -eq 'INetFrom' } | Select-Object -First 1
                        $BodyItem = $Document.Items | Where-Object { $_.Name -eq 'Body' } | Select-Object -First 1
                        $Author = if ($FromItem) { ([string]$FromItem.Text) -replace "`r", '' -replace '.*<(.*)>.*', '$1' } else { '' }
                        if ([string]::IsNullOrWhiteSpace($Author) -and $InetFromItem) { $Author = [string]$InetFromItem.Text }
                        $Body = ''
                        if ($null -ne $BodyItem) {
                            $Body = ([string]$BodyItem.Text) -replace "`r", ''
                            $Body = $Body.Replace([char]0x2019, [char]39)
                            $Body = $Body -replace '[\u2013\u2014]', '-'
                        }
                        if ($Body.Length -gt 30000) { $Body = $Body.Substring(0, 30000) }
                        $Record = [pscustomobject]@{
                            universalid = [string]$Document.UniversalID
                            parent = [string]$Document.ParentDocumentUNID
                            categoryPath = $CategoryPath
                            subject = [string]$SubjectItem.Text
                            creationtime = [datetime]$Document.Created
                            notesurl = [string]$Document.NotesURL
                            emailBody = $Body
                            authors = $Author
                        }
                        if ($Processed.Contains($Record.universalid)) { $SkippedProcessed++ }
                        else {
                            $Sensitive = $false
                            foreach ($Pattern in @($Config.processing.sensitiveSubjectPatterns)) {
                                if ($Record.subject -like [string]$Pattern) { $Sensitive = $true; break }
                            }
                            if ($Sensitive) { $SkippedSensitive++ }
                            elseif (Test-CategoryPathMatch -ActualPath $CategoryPath -ConfiguredPaths $JiraPaths -IncludeDescendants $IncludeDescendants) {
                                $Candidates.Add($Record)
                            }
                        }
                    }
                }
                $Document = $Next
            }
        }
        finally {
            foreach ($ComObject in @($Document, $View, $Database, $Session)) {
                if ($null -ne $ComObject -and [Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
                    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
                }
            }
        }

        $Plan = [Collections.Generic.List[object]]::new()
        $PlannedIssueByDocument = @{}
        foreach ($Email in @($Candidates | Sort-Object creationtime)) {
            $ParentId = [string]$Email.parent
            $ExistingKey = $null
            if (-not [string]::IsNullOrWhiteSpace($ParentId) -and $IssueByDocument.ContainsKey($ParentId)) {
                $ExistingKey = $IssueByDocument[$ParentId]
            }
            elseif (-not [string]::IsNullOrWhiteSpace($ParentId) -and $PlannedIssueByDocument.ContainsKey($ParentId)) {
                $ExistingKey = $PlannedIssueByDocument[$ParentId]
            }

            if ($ExistingKey) {
                $Action = 'AddComment'
                $IssueKey = [string]$ExistingKey
                $ConversationMarker = $IssueKey
            }
            elseif (-not [string]::IsNullOrWhiteSpace($ParentId) -and $Processed.Contains($ParentId)) {
                $Action = 'SkipUnmappedParent'
                $IssueKey = ''
                $ConversationMarker = ''
            }
            else {
                $Action = 'CreateIssue'
                $IssueKey = ''
                $ConversationMarker = "PENDING:$($Email.universalid)"
            }

            if (-not [string]::IsNullOrWhiteSpace($ConversationMarker)) {
                $PlannedIssueByDocument[$Email.universalid] = $ConversationMarker
            }
            $Plan.Add([pscustomobject]@{
                Action = $Action
                IssueKey = $IssueKey
                ConversationMarker = $ConversationMarker
                Email = $Email
            })
        }

        $CreateCount = @($Plan | Where-Object Action -eq 'CreateIssue').Count
        $CommentCount = @($Plan | Where-Object Action -eq 'AddComment').Count
        $UnmappedCount = @($Plan | Where-Object Action -eq 'SkipUnmappedParent').Count
        $PreviewPath = Join-Path $ApplicationRoot "JiraImportPreview-$($MailboxName -replace '[^A-Za-z0-9_-]', '_').csv"
        $Plan | ForEach-Object {
            [pscustomobject]@{
                Action = $_.Action
                ExistingIssueKey = $_.IssueKey
                Subject = $_.Email.subject
                Created = $_.Email.creationtime
                UniversalId = $_.Email.universalid
                ParentUniversalId = $_.Email.parent
                CategoryPath = $_.Email.categoryPath
            }
        } | Export-Csv -LiteralPath $PreviewPath -NoTypeInformation -Encoding UTF8

        Write-Host "`nNotes scan completed for '$MailboxName'." -ForegroundColor Cyan
        Write-Host "New Jira issues proposed : $CreateCount"
        Write-Host "Jira comments proposed   : $CommentCount"
        Write-Host "Already processed skipped: $SkippedProcessed"
        Write-Host "Sensitive emails skipped : $SkippedSensitive"
        Write-Host "Replies without Jira map : $UnmappedCount"
        Write-Host "Preview file              : $PreviewPath"

        if ($MaxIssues -gt 0 -and $CreateCount -gt $MaxIssues) {
            throw "Proposed issue count ($CreateCount) exceeds -MaxIssues $MaxIssues. No Jira changes were made."
        }
        if ($DryRun) {
            Write-Host 'DRY RUN: No Jira issues, comments, or state changes were made.' -ForegroundColor Yellow
            continue
        }
        if (($CreateCount + $CommentCount) -eq 0) {
            Write-Host 'No Jira changes are required.' -ForegroundColor Green
            continue
        }
        if (-not $Force) {
            $Answer = Read-Host "Proceed with $CreateCount new issue(s) and $CommentCount comment(s)? Type YES to continue"
            if ($Answer -cne 'YES') {
                Write-Host 'Cancelled. No Jira or state changes were made.' -ForegroundColor Yellow
                continue
            }
        }

        $CreatedKeyByDocument = @{}
        foreach ($Item in $Plan) {
            $Email = $Item.Email
            if ($Item.Action -eq 'SkipUnmappedParent') {
                Write-Warning "Skipped reply '$($Email.subject)': its parent was processed previously but has no Jira key in the state CSV."
                continue
            }

            if ($Item.Action -eq 'CreateIssue') {
                $Issue = New-JiraIssue `
                    -Client $JiraClient `
                    -Mailbox $Mailbox `
                    -ComponentId ([string]$Component.id) `
                    -EpicFieldId $EpicFieldId `
                    -EpicKey $EpicKey `
                    -Email $Email
                $IssueKey = [string]$Issue.key
                if ([string]::IsNullOrWhiteSpace($IssueKey)) { throw 'Jira created an issue but returned no issue key.' }
                $CreatedKeyByDocument[$Email.universalid] = $IssueKey
                Write-Host "Created Jira issue ${IssueKey}: $($Email.subject)" -ForegroundColor Green
                $ActionTaken = 'CreatedIssue'
                $RootId = $Email.universalid
            }
            else {
                $IssueKey = [string]$Item.IssueKey
                if ($IssueKey -like 'PENDING:*') {
                    $RootId = $IssueKey.Substring(8)
                    if (-not $CreatedKeyByDocument.ContainsKey($RootId)) {
                        throw "Could not resolve newly created Jira issue for Notes root '$RootId'."
                    }
                    $IssueKey = $CreatedKeyByDocument[$RootId]
                }
                Add-JiraEmailComment -Client $JiraClient -IssueKey $IssueKey -Email $Email | Out-Null
                Write-Host "Added Notes reply as a comment to ${IssueKey}: $($Email.subject)" -ForegroundColor Green
                $ActionTaken = 'AddedComment'
                $RootId = if ($Item.ConversationMarker -like 'PENDING:*') {
                    $Item.ConversationMarker.Substring(8)
                }
                elseif ($Email.parent) { [string]$Email.parent }
                else { $Email.universalid }
            }

            $StateRow = [pscustomobject]@{
                universalid = $Email.universalid
                parent = $Email.parent
                rootUniversalId = $RootId
                jiraIssueKey = $IssueKey
                jiraAction = $ActionTaken
                categoryPath = $Email.categoryPath
                subject = $Email.subject
                creationtime = $Email.creationtime
                notesurl = $Email.notesurl
                emailBody = $Email.emailBody
                authors = $Email.authors
                processedAt = (Get-Date).ToString('s')
            }
            $State.Add($StateRow)
            $null = $Processed.Add($Email.universalid)
            $IssueByDocument[$Email.universalid] = $IssueKey
            Save-State -State $State -Path $StatePath
        }
        Write-Host "Completed '$MailboxName'. State saved to $StatePath" -ForegroundColor Green
    }
}
finally {
    if ($null -ne $JiraClient) { Remove-JiraClient -Client $JiraClient }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
