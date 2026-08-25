Set-StrictMode -Version Latest

<#
.SYNOPSIS
    Jira Data Center REST API functions for Jira Email Importer setup.

.DESCRIPTION
    Provides the Jira REST functions required by Setup-JiraEmailImporter.ps1:
      - authenticated Jira client creation
      - resilient REST requests
      - connection testing
      - Jira field and Epic Link discovery
      - project issue-type discovery
      - project component discovery
      - Epic validation

    Import Modules\Security.psm1 before calling New-JiraClient.
    The module never writes the Jira PAT or Authorization header to output.
#>

function Test-JiraBlankString {
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

function ConvertTo-JiraUrlEncodedValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Value
    )

    $Encoding = New-Object System.Text.UTF8Encoding
    $Bytes = $Encoding.GetBytes($Value)
    $Builder = New-Object System.Text.StringBuilder

    foreach ($Byte in $Bytes) {
        $Character = [char]$Byte
        $IsUnreserved = (
            ($Byte -ge 65 -and $Byte -le 90) -or
            ($Byte -ge 97 -and $Byte -le 122) -or
            ($Byte -ge 48 -and $Byte -le 57) -or
            $Character -eq '-' -or
            $Character -eq '_' -or
            $Character -eq '.' -or
            $Character -eq '~'
        )

        if ($IsUnreserved) {
            $null = $Builder.Append($Character)
        }
        else {
            $null = $Builder.Append(('%{0:X2}' -f $Byte))
        }
    }

    return $Builder.ToString()
}

function Get-JiraObjectPropertyValue {
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

function Get-JiraErrorDetail {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    $Messages = New-Object 'System.Collections.Generic.List[string]'
    $RawDetail = $null

    if ($null -ne $ErrorRecord.ErrorDetails) {
        $RawDetail = [string]$ErrorRecord.ErrorDetails.Message
    }

    if (-not (Test-JiraBlankString -Value $RawDetail)) {
        try {
            $ParsedDetail = $RawDetail | ConvertFrom-Json -ErrorAction Stop
            $ErrorMessages = Get-JiraObjectPropertyValue -InputObject $ParsedDetail -Name 'errorMessages'
            $Errors = Get-JiraObjectPropertyValue -InputObject $ParsedDetail -Name 'errors'

            foreach ($Item in @($ErrorMessages)) {
                if (-not (Test-JiraBlankString -Value ([string]$Item))) {
                    $Messages.Add([string]$Item)
                }
            }

            if ($null -ne $Errors) {
                if ($Errors -is [System.Collections.IDictionary]) {
                    foreach ($Key in $Errors.Keys) {
                        $Messages.Add(('{0}: {1}' -f $Key, $Errors[$Key]))
                    }
                }
                else {
                    foreach ($Property in $Errors.PSObject.Properties) {
                        $Messages.Add(('{0}: {1}' -f $Property.Name, $Property.Value))
                    }
                }
            }
        }
        catch {
            $Messages.Add($RawDetail)
        }
    }

    if ($Messages.Count -eq 0 -and $null -ne $ErrorRecord.Exception) {
        $ExceptionMessage = [string]$ErrorRecord.Exception.Message
        if (-not (Test-JiraBlankString -Value $ExceptionMessage)) {
            $Messages.Add($ExceptionMessage)
        }
    }

    if ($Messages.Count -eq 0) {
        $Messages.Add('An unknown Jira REST API error occurred.')
    }

    return ($Messages -join '; ')
}

function Get-JiraHttpStatusCode {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    try {
        $Response = Get-JiraObjectPropertyValue -InputObject $ErrorRecord.Exception -Name 'Response'
        $StatusCode = Get-JiraObjectPropertyValue -InputObject $Response -Name 'StatusCode'

        if ($null -ne $StatusCode) {
            return [int]$StatusCode
        }
    }
    catch {
        return $null
    }

    return $null
}

function New-JiraClient {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$BaseUrl,

        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [System.Security.SecureString]$SecurePat,

        [Parameter()]
        [ValidateRange(1, 600)]
        [int]$TimeoutSeconds = 60,

        [Parameter()]
        [ValidateRange(0, 10)]
        [int]$MaximumRetryCount = 2
    )

    try {
        $JiraUri = New-Object System.Uri -ArgumentList $BaseUrl
    }
    catch {
        throw 'The Jira BaseUrl must be a valid HTTPS URL.'
    }

    if (-not $JiraUri.IsAbsoluteUri -or $JiraUri.Scheme -ne 'https') {
        throw 'The Jira BaseUrl must be a valid HTTPS URL.'
    }

    $ConvertCommand = Get-Command -Name ConvertFrom-SecureStringPlainText -ErrorAction SilentlyContinue
    if ($null -eq $ConvertCommand) {
        throw (
            'ConvertFrom-SecureStringPlainText is unavailable. ' +
            'Import Modules\Security.psm1 before creating the Jira client.'
        )
    }

    $PlainTextPat = $null

    try {
        $PlainTextPat = ConvertFrom-SecureStringPlainText -SecureString $SecurePat
        if (Test-JiraBlankString -Value $PlainTextPat) {
            throw 'The Jira PAT is empty.'
        }

        $Headers = @{
            Authorization = 'Bearer {0}' -f $PlainTextPat
            Accept        = 'application/json'
        }

        return [pscustomobject]@{
            PSTypeName        = 'JiraEmailImporter.JiraClient'
            BaseUrl           = $BaseUrl.TrimEnd('/')
            Headers           = $Headers
            TimeoutSeconds    = $TimeoutSeconds
            MaximumRetryCount = $MaximumRetryCount
        }
    }
    finally {
        $PlainTextPat = $null
    }
}
function Remove-JiraClient {
[CmdletBinding()]
param(
[Parameter(Mandatory)]
[ValidateNotNull()]
[psobject]$Client
)

$Headers = Get-JiraObjectPropertyValue `
-InputObject $Client `
-Name 'Headers'

if ($null -ne $Headers) {
if (
$Headers -is [System.Collections.IDictionary] -and
$Headers.Contains('Authorization')
) {
$Headers['Authorization'] = $null
$Headers.Remove('Authorization')
}
}

$HeadersProperty = $Client.PSObject.Properties['Headers']

if ($null -ne $HeadersProperty) {
$HeadersProperty.Value = $null
}
}

function Invoke-JiraRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [psobject]$Client,

        [Parameter(Mandatory)]
        [ValidateSet('Get', 'Post', 'Put', 'Delete')]
        [string]$Method,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter()]
        [AllowNull()]
        [object]$Body,

        [Parameter()]
        [switch]$RawBody,

        [Parameter()]
        [switch]$NoRetry
    )

    $BaseUrl = [string](Get-JiraObjectPropertyValue -InputObject $Client -Name 'BaseUrl')
    $Headers = Get-JiraObjectPropertyValue -InputObject $Client -Name 'Headers'
    $TimeoutSeconds = Get-JiraObjectPropertyValue -InputObject $Client -Name 'TimeoutSeconds'
    $MaximumRetryCount = Get-JiraObjectPropertyValue -InputObject $Client -Name 'MaximumRetryCount'

    if (Test-JiraBlankString -Value $BaseUrl) {
        throw 'The Jira client does not contain a BaseUrl.'
    }

    if ($null -eq $Headers -or -not $Headers.Contains('Authorization')) {
        throw 'The Jira client does not contain authentication headers.'
    }

    if ($null -eq $TimeoutSeconds) {
        $TimeoutSeconds = 60
    }

    if ($null -eq $MaximumRetryCount) {
        $MaximumRetryCount = 0
    }

    $RelativePath = $Path.TrimStart('/')
    $Uri = '{0}/{1}' -f $BaseUrl.TrimEnd('/'), $RelativePath
    $MaximumAttempts = 1

    if (-not $NoRetry) {
        $MaximumAttempts += [int]$MaximumRetryCount
    }

    for ($Attempt = 1; $Attempt -le $MaximumAttempts; $Attempt++) {
        try {
            $Parameters = @{
                Uri         = $Uri
                Method      = $Method
                Headers     = $Headers
                TimeoutSec  = [int]$TimeoutSeconds
                ErrorAction = 'Stop'
            }

            if ($PSBoundParameters.ContainsKey('Body') -and $null -ne $Body) {
                if ($RawBody) {
                    $Parameters.Body = $Body
                }
                else {
                    $Parameters.Body = $Body | ConvertTo-Json -Depth 20
                    $Parameters.ContentType = 'application/json; charset=utf-8'
                }
            }

            return Invoke-RestMethod @Parameters
        }
        catch {
            $StatusCode = Get-JiraHttpStatusCode -ErrorRecord $_
            $Retryable = (
                $null -eq $StatusCode -or
                $StatusCode -eq 408 -or
                $StatusCode -eq 429 -or
                ($StatusCode -ge 500 -and $StatusCode -le 599)
            )

            if ($Attempt -lt $MaximumAttempts -and $Retryable) {
                $DelaySeconds = [int](1 -shl $Attempt)
                if ($DelaySeconds -gt 10) {
                    $DelaySeconds = 10
                }

                Start-Sleep -Seconds $DelaySeconds
                continue
            }

            $Detail = Get-JiraErrorDetail -ErrorRecord $_
            $StatusText = ''
            if ($null -ne $StatusCode) {
                $StatusText = ' HTTP {0}.' -f $StatusCode
            }

            throw ('Jira request failed:{0} {1} {2}. {3}' -f $StatusText, $Method, $RelativePath, $Detail)
        }
    }
}

function Test-JiraConnection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [psobject]$Client
    )

    $Stopwatch = New-Object System.Diagnostics.Stopwatch
    $Stopwatch.Start()

    try {
        $CurrentUser = Invoke-JiraRequest -Client $Client -Method Get -Path 'rest/api/2/myself'
        $Stopwatch.Stop()

        return [pscustomobject]@{
            Success      = $true
            DisplayName  = [string](Get-JiraObjectPropertyValue -InputObject $CurrentUser -Name 'displayName')
            UserName     = [string](Get-JiraObjectPropertyValue -InputObject $CurrentUser -Name 'name')
            EmailAddress = [string](Get-JiraObjectPropertyValue -InputObject $CurrentUser -Name 'emailAddress')
            Active       = [bool](Get-JiraObjectPropertyValue -InputObject $CurrentUser -Name 'active')
            DurationMs   = [int]$Stopwatch.ElapsedMilliseconds
            JiraUrl      = [string](Get-JiraObjectPropertyValue -InputObject $Client -Name 'BaseUrl')
        }
    }
    catch {
        $Stopwatch.Stop()
        throw ('Jira connection test failed: {0}' -f $_.Exception.Message)
    }
}

function Get-JiraFields {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [psobject]$Client
    )

    $Response = Invoke-JiraRequest `
        -Client $Client `
        -Method Get `
        -Path 'rest/api/2/field'

    if ($null -eq $Response) {
        return
    }

    # Jira normally returns a top-level JSON array. Windows PowerShell 5.1
    # may preserve that response as one array object, so explicitly emit
    # each field.
    if ($Response -is [System.Array]) {
        for ($Index = 0; $Index -lt $Response.Count; $Index++) {
            $Item = $Response[$Index]

            if ($Item -is [System.Array]) {
                foreach ($NestedItem in $Item) {
                    if ($null -ne $NestedItem) {
                        Write-Output $NestedItem
                    }
                }
            }
            elseif ($null -ne $Item) {
                Write-Output $Item
            }
        }

        return
    }

    # Defensive support for a response wrapped in a values property.
    $Values = Get-JiraObjectPropertyValue `
        -InputObject $Response `
        -Name 'values'

    if ($null -ne $Values) {
        foreach ($Field in @($Values)) {
            if ($null -ne $Field) {
                Write-Output $Field
            }
        }

        return
    }

    Write-Output $Response
}

function Find-JiraField {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [object[]]$Fields,

        [Parameter()]
        [AllowNull()]
        [string]$Name,

        [Parameter()]
        [AllowNull()]
        [string]$SchemaCustom
    )

    if (
        (Test-JiraBlankString -Value $Name) -and
        (Test-JiraBlankString -Value $SchemaCustom)
    ) {
        throw 'Specify Name, SchemaCustom, or both.'
    }

    $Matches = New-Object 'System.Collections.Generic.List[object]'

    foreach ($Field in @($Fields)) {
        if ($null -eq $Field) {
            continue
        }

        # Windows PowerShell 5.1 can retain a top-level JSON array as a
        # nested array. Recursively inspect each contained Jira field.
        if ($Field -is [System.Array]) {
            $NestedParameters = @{
                Fields = @($Field)
            }

            if (-not (Test-JiraBlankString -Value $Name)) {
                $NestedParameters.Name = $Name
            }

            if (-not (Test-JiraBlankString -Value $SchemaCustom)) {
                $NestedParameters.SchemaCustom = $SchemaCustom
            }

            foreach ($NestedMatch in @(Find-JiraField @NestedParameters)) {
                if ($null -ne $NestedMatch) {
                    $Matches.Add($NestedMatch)
                }
            }

            continue
        }

        $IsMatch = $true

        if (-not (Test-JiraBlankString -Value $Name)) {
            $FieldName = Get-JiraObjectPropertyValue `
                -InputObject $Field `
                -Name 'name'

            # -ieq is case-insensitive and safe when $FieldName is null.
            $IsMatch = ([string]$FieldName -ieq $Name)
        }

        if (
            $IsMatch -and
            -not (Test-JiraBlankString -Value $SchemaCustom)
        ) {
            $Schema = Get-JiraObjectPropertyValue `
                -InputObject $Field `
                -Name 'schema'

            $Custom = Get-JiraObjectPropertyValue `
                -InputObject $Schema `
                -Name 'custom'

            # Many Jira system fields do not have schema.custom.
            # A missing value is simply not a match.
            $IsMatch = (
                -not (Test-JiraBlankString -Value $Custom) -and
                [string]$Custom -ieq $SchemaCustom
            )
        }

        if ($IsMatch) {
            $Matches.Add($Field)
        }
    }

    return $Matches.ToArray()
}

function Get-JiraEpicLinkField {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [psobject]$Client,

        [Parameter()]
        [object[]]$Fields
    )

    if ($null -eq $Fields -or @($Fields).Count -eq 0) {
        $Fields = @(Get-JiraFields -Client $Client)
    }
    else {
        # Flatten nested arrays supplied by the caller.
        $FlattenedFields = New-Object `
            'System.Collections.Generic.List[object]'

        foreach ($Field in @($Fields)) {
            if ($Field -is [System.Array]) {
                foreach ($NestedField in $Field) {
                    if ($null -ne $NestedField) {
                        $FlattenedFields.Add($NestedField)
                    }
                }
            }
            elseif ($null -ne $Field) {
                $FlattenedFields.Add($Field)
            }
        }

        $Fields = $FlattenedFields.ToArray()
    }

    if (@($Fields).Count -eq 0) {
        throw 'Jira returned no fields from rest/api/2/field.'
    }

    # Primary match: Jira Software Epic Link schema identifier.
    $EpicLinkMatches = @(
        Find-JiraField `
            -Fields $Fields `
            -SchemaCustom 'com.pyxis.greenhopper.jira:gh-epic-link'
    )

    # Fallback: standard display name.
    if ($EpicLinkMatches.Count -eq 0) {
        $EpicLinkMatches = @(
            Find-JiraField `
                -Fields $Fields `
                -Name 'Epic Link'
        )
    }

    # Additional fallback for renamed or differently represented fields.
    if ($EpicLinkMatches.Count -eq 0) {
        $EpicLinkMatches = @(
            $Fields | Where-Object {
                $FieldName = Get-JiraObjectPropertyValue `
                        -InputObject $_ `
                        -Name 'name'
                

                $ClauseNames = @(
                    Get-JiraObjectPropertyValue `
                        -InputObject $_ `
                        -Name 'clauseNames'
                )

                $FieldName -match '(?i)^epic\s*link$' -or
                @($ClauseNames | Where-Object {
                    [string]$_ -match '(?i)epic[ -]?link'
                }).Count -gt 0
            }
        )
    }

    if ($EpicLinkMatches.Count -gt 1) {
        $Ids = @(
            $EpicLinkMatches | ForEach-Object {
                Get-JiraObjectPropertyValue `
                    -InputObject $_ `
                    -Name 'id'
            }
        ) -join ', '

        throw "Multiple Epic Link fields were found: $Ids"
    }

    if ($EpicLinkMatches.Count -eq 1) {
        return $EpicLinkMatches[0]
    }

    # Provide diagnostics without exposing authentication data.
    $Candidates = @(
        $Fields | Where-Object {
            $FieldName = Get-JiraObjectPropertyValue `
                    -InputObject $_ `
                    -Name 'name'
            

            $Schema = Get-JiraObjectPropertyValue `
                -InputObject $_ `
                -Name 'schema'

            $SchemaCustom = Get-JiraObjectPropertyValue `
                    -InputObject $Schema `
                    -Name 'custom'
            

            $FieldName -match '(?i)epic|parent' -or
            $SchemaCustom -match '(?i)epic|greenhopper|parent'
        } | ForEach-Object {
            $Schema = Get-JiraObjectPropertyValue `
                -InputObject $_ `
                -Name 'schema'

            [pscustomobject]@{
                Id = Get-JiraObjectPropertyValue `
                        -InputObject $_ `
                        -Name 'id'
                
                Name = Get-JiraObjectPropertyValue `
                        -InputObject $_ `
                        -Name 'name'
                
                SchemaCustom = Get-JiraObjectPropertyValue `
                        -InputObject $Schema `
                        -Name 'custom'
                
            }
        }
    )

    if ($Candidates.Count -gt 0) {
        $CandidateText = @(
            $Candidates | ForEach-Object {
                "'{0}' (ID {1}; schema {2})" -f
                    $_.Name,
                    $_.Id,
                    $_.SchemaCustom
            }
        ) -join '; '

        throw (
            'The Jira Epic Link custom field could not be identified. ' +
            "Related Jira fields returned: $CandidateText"
        )
    }

    throw (
        'The Jira Epic Link custom field could not be found. ' +
        "Jira returned $(@($Fields).Count) fields, but none had an " +
        "Epic Link name, clause name, or Jira Software Epic Link schema."
    )
}

function Get-JiraProjectIssueTypes {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [psobject]$Client,

        [Parameter(Mandatory)]
        [ValidatePattern('^[A-Za-z][A-Za-z0-9_]*$')]
        [string]$ProjectKey
    )

    $EncodedProjectKey = ConvertTo-JiraUrlEncodedValue -Value $ProjectKey
    $Response = Invoke-JiraRequest `
        -Client $Client `
        -Method Get `
        -Path ('rest/api/2/issue/createmeta/{0}/issuetypes' -f $EncodedProjectKey)

    if ($Response -is [System.Array]) {
        return @($Response)
    }

    $IssueTypes = Get-JiraObjectPropertyValue -InputObject $Response -Name 'issueTypes'
    if ($null -ne $IssueTypes) {
        return @($IssueTypes)
    }

    $Values = Get-JiraObjectPropertyValue -InputObject $Response -Name 'values'
    if ($null -ne $Values) {
        return @($Values)
    }

    return @($Response)
}

function Get-JiraProjectComponents {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [psobject]$Client,

        [Parameter(Mandatory)]
        [ValidatePattern('^[A-Za-z][A-Za-z0-9_]*$')]
        [string]$ProjectKey
    )

    $EncodedProjectKey = ConvertTo-JiraUrlEncodedValue -Value $ProjectKey
    $Response = Invoke-JiraRequest `
        -Client $Client `
        -Method Get `
        -Path ('rest/api/2/project/{0}/components' -f $EncodedProjectKey)

    if ($Response -is [System.Array]) {
        for ($Index = 0; $Index -lt $Response.Count; $Index++) {
            Write-Output $Response[$Index]
        }

        return
    }

    $Values = Get-JiraObjectPropertyValue -InputObject $Response -Name 'values'
    if ($null -ne $Values) {
        foreach ($Component in @($Values)) {
            Write-Output $Component
        }

        return
    }

    if ($null -ne $Response) {
        Write-Output $Response
    }
}

function Get-JiraIssue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [psobject]$Client,

        [Parameter(Mandatory)]
        [ValidatePattern('^[A-Za-z][A-Za-z0-9_]*-[0-9]+$')]
        [string]$IssueKey,

        [Parameter()]
        [string[]]$Fields = @('summary', 'issuetype', 'status')
    )

    $EncodedIssueKey = ConvertTo-JiraUrlEncodedValue -Value $IssueKey
    $Path = 'rest/api/2/issue/{0}' -f $EncodedIssueKey

    if ($null -ne $Fields -and $Fields.Count -gt 0) {
        $EncodedFields = ConvertTo-JiraUrlEncodedValue -Value ($Fields -join ',')
        $Path += '?fields={0}' -f $EncodedFields
    }

    return Invoke-JiraRequest -Client $Client -Method Get -Path $Path
}

function Test-JiraEpic {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [psobject]$Client,

        [Parameter(Mandatory)]
        [ValidatePattern('^[A-Za-z][A-Za-z0-9_]*-[0-9]+$')]
        [string]$EpicKey
    )

    $Issue = Get-JiraIssue -Client $Client -IssueKey $EpicKey -Fields @('summary', 'issuetype', 'status')
    $IssueFields = Get-JiraObjectPropertyValue -InputObject $Issue -Name 'fields'
    $IssueType = Get-JiraObjectPropertyValue -InputObject $IssueFields -Name 'issuetype'
    $Status = Get-JiraObjectPropertyValue -InputObject $IssueFields -Name 'status'
    $IssueTypeName = [string](Get-JiraObjectPropertyValue -InputObject $IssueType -Name 'name')
    $IssueKey = [string](Get-JiraObjectPropertyValue -InputObject $Issue -Name 'key')
    $BaseUrl = [string](Get-JiraObjectPropertyValue -InputObject $Client -Name 'BaseUrl')

    return [pscustomobject]@{
        IsValid       = ($IssueTypeName -eq 'Epic')
        Key           = $IssueKey
        Summary       = [string](Get-JiraObjectPropertyValue -InputObject $IssueFields -Name 'summary')
        IssueTypeName = $IssueTypeName
        Status        = [string](Get-JiraObjectPropertyValue -InputObject $Status -Name 'name')
        Url           = '{0}/browse/{1}' -f $BaseUrl.TrimEnd('/'), $IssueKey
    }
}

Export-ModuleMember -Function @(
    'Find-JiraField',
    'Get-JiraEpicLinkField',
    'Get-JiraFields',
    'Get-JiraIssue',
    'Get-JiraProjectComponents',
    'Get-JiraProjectIssueTypes',
    'Invoke-JiraRequest',
    'New-JiraClient',
    'Remove-JiraClient',
    'Test-JiraConnection',
    'Test-JiraEpic'
)
