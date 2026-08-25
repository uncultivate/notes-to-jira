Set-StrictMode -Version Latest

<#
.SYNOPSIS
    Jira Data Center REST API functions for Jira Email Importer setup.

.DESCRIPTION
    Provides the Jira REST functions required by Setup-JiraEmailImporter.ps1
    and Task-NotesQuery.ps1:
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

function ConvertTo-JiraUserFacingError {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.Management.Automation.ErrorRecord]$ErrorRecord,

        [Parameter()]
        [string]$Method,

        [Parameter()]
        [string]$RelativePath,

        [Parameter()]
        [string]$Uri
    )

    $Existing = [string]$ErrorRecord.Exception.Message
    if ($Existing -match '(?s)^(Could not (reach|connect to|complete)|Jira (rejected|denied|could not|is unavailable)|Authentication to Jira)') {
        return $Existing
    }

    $StatusCode = Get-JiraHttpStatusCode -ErrorRecord $ErrorRecord
    $Detail = Get-JiraErrorDetail -ErrorRecord $ErrorRecord
    $InnerMessage = ''
    if ($null -ne $ErrorRecord.Exception.InnerException) {
        $InnerMessage = [string]$ErrorRecord.Exception.InnerException.Message
    }

    $Combined = @($Existing, $Detail, $InnerMessage) -join ' '
    $StatusNames = @{
        400 = 'Bad Request'
        401 = 'Unauthorized'
        403 = 'Forbidden'
        404 = 'Not Found'
        408 = 'Request Timeout'
        429 = 'Too Many Requests'
        500 = 'Internal Server Error'
        502 = 'Bad Gateway'
        503 = 'Service Unavailable'
        504 = 'Gateway Timeout'
    }

    $StatusName = $null
    if ($null -ne $StatusCode -and $StatusNames.ContainsKey([int]$StatusCode)) {
        $StatusName = $StatusNames[[int]$StatusCode]
    }

    $HostName = $null
    if ($Combined -match "remote name could not be resolved[:\s]+'?([^'\r\n]+)'?") {
        $HostName = $Matches[1].Trim()
    }
    elseif ($Combined -match 'No such host is known') {
        if (-not (Test-JiraBlankString -Value $Uri)) {
            try {
                $HostName = ([Uri]$Uri).Host
            }
            catch {
                $HostName = $null
            }
        }
    }

    $DisplayUrl = $Uri
    if (-not (Test-JiraBlankString -Value $Uri)) {
        try {
            $ParsedUri = New-Object System.Uri -ArgumentList $Uri
            if ($ParsedUri.IsAbsoluteUri) {
                $DisplayUrl = '{0}://{1}' -f $ParsedUri.Scheme, $ParsedUri.Authority
            }
        }
        catch {
            $DisplayUrl = $Uri
        }
    }

    $MethodLabel = $Method
    if (-not (Test-JiraBlankString -Value $Method)) {
        $MethodLabel = $Method.ToUpperInvariant()
    }

    $TechnicalParts = New-Object 'System.Collections.Generic.List[string]'
    $RequestText = ('{0} {1}' -f $MethodLabel, $RelativePath).Trim()
    if (-not (Test-JiraBlankString -Value $RequestText)) {
        $TechnicalParts.Add($RequestText)
    }
    if ($null -ne $StatusCode) {
        if (-not (Test-JiraBlankString -Value $StatusName)) {
            $TechnicalParts.Add(('HTTP {0} {1}' -f $StatusCode, $StatusName))
        }
        else {
            $TechnicalParts.Add(('HTTP {0}' -f $StatusCode))
        }
    }
    if (-not (Test-JiraBlankString -Value $Detail)) {
        $TechnicalParts.Add($Detail)
    }

    $Technical = ($TechnicalParts.ToArray() -join ' | ')
    $Lines = New-Object 'System.Collections.Generic.List[string]'

    if ($Combined -match 'remote name could not be resolved|No such host is known') {
        $Where = ''
        if (-not (Test-JiraBlankString -Value $DisplayUrl)) {
            $Where = " at $DisplayUrl"
        }
        elseif (-not (Test-JiraBlankString -Value $HostName)) {
            $Where = " at $HostName"
        }

        $Lines.Add("Could not reach Jira$Where.")
        $Lines.Add('')
        if (-not (Test-JiraBlankString -Value $HostName)) {
            $Lines.Add("The hostname '$HostName' could not be resolved.")
        }
        else {
            $Lines.Add('The Jira hostname could not be resolved.')
        }
        $Lines.Add('Typical causes:')
        $Lines.Add('  - you are not connected to the corporate network or VPN')
        $Lines.Add('  - jira.url in config.json is incorrect')
        $Lines.Add('')
        $Lines.Add('Open the Jira URL in a browser to confirm it is reachable, then retry.')
    }
    elseif ($Combined -match 'The operation has timed out|The request was aborted|timed out') {
        $Where = if (-not (Test-JiraBlankString -Value $DisplayUrl)) { " at $DisplayUrl" } else { '' }
        $Lines.Add("The Jira request timed out$Where.")
        $Lines.Add('')
        $Lines.Add('Typical causes:')
        $Lines.Add('  - Jira is slow or unreachable on the current network')
        $Lines.Add('  - jira.requestTimeoutSeconds in config.json is too low')
        $Lines.Add('')
        $Lines.Add('Check that Jira opens in a browser, then retry.')
    }
    elseif ($Combined -match 'Could not establish trust relationship|The underlying connection was closed|SSL/TLS|secure channel|certificate') {
        $Where = if (-not (Test-JiraBlankString -Value $DisplayUrl)) { " at $DisplayUrl" } else { '' }
        $Lines.Add("Could not establish a trusted HTTPS connection to Jira$Where.")
        $Lines.Add('')
        $Lines.Add('Typical causes:')
        $Lines.Add('  - corporate TLS inspection or an untrusted certificate')
        $Lines.Add('  - jira.url does not use a valid HTTPS address')
        $Lines.Add('')
        $Lines.Add('Open the Jira URL in a browser and confirm the certificate is trusted.')
    }
    elseif ($Combined -match 'Unable to connect to the remote server|actively refused|No connection could be made') {
        $Where = if (-not (Test-JiraBlankString -Value $DisplayUrl)) { " at $DisplayUrl" } else { '' }
        $Lines.Add("Could not connect to Jira$Where.")
        $Lines.Add('')
        $Lines.Add('Typical causes:')
        $Lines.Add('  - you are not connected to the corporate network or VPN')
        $Lines.Add('  - the Jira host is down or the URL port is wrong')
        $Lines.Add('')
        $Lines.Add('Open the Jira URL in a browser to confirm it is reachable, then retry.')
    }
    elseif ($StatusCode -eq 401) {
        $Lines.Add('Jira rejected the Personal Access Token (HTTP 401 Unauthorized).')
        $Lines.Add('')
        $Lines.Add('The stored token is missing, expired, or invalid.')
        $Lines.Add('Fix: re-run this script with -ReplacePat and paste a new PAT from Jira')
        $Lines.Add('     (your Jira profile > Personal Access Tokens).')
    }
    elseif ($StatusCode -eq 403) {
        $Lines.Add('Jira denied access (HTTP 403 Forbidden).')
        $Lines.Add('')
        $Lines.Add('The token was accepted, but this account cannot perform the request.')
        $Lines.Add('Check Browse Projects and Create Issues permissions for the target project.')
        if (-not (Test-JiraBlankString -Value $Detail)) {
            $Lines.Add('')
            $Lines.Add($Detail)
        }
    }
    elseif ($StatusCode -eq 404) {
        $Lines.Add('Jira could not find the requested resource (HTTP 404 Not Found).')
        $Lines.Add('')
        if ($RelativePath -match 'project/') {
            $Lines.Add('The project key may be wrong, or the PAT user cannot see that project.')
            $Lines.Add('Check mailboxes[].jira.projectKey in config.json.')
        }
        elseif ($RelativePath -match 'issue/') {
            $Lines.Add('The issue key may be wrong, or the PAT user cannot see that issue.')
            $Lines.Add('Check jira.defaultEpicKey or mailboxes[].jira.epicKey in config.json.')
        }
        else {
            $Lines.Add('Check the Jira URL, project key, and issue key in config.json.')
        }
        if (-not (Test-JiraBlankString -Value $Detail)) {
            $Lines.Add('')
            $Lines.Add($Detail)
        }
    }
    elseif ($StatusCode -eq 400) {
        $Lines.Add('Jira rejected the request (HTTP 400 Bad Request).')
        $Lines.Add('')
        if (-not (Test-JiraBlankString -Value $Detail)) {
            $Lines.Add($Detail)
            $Lines.Add('')
        }
        $Lines.Add('No Jira item was created or updated. Check required fields for the project')
        $Lines.Add('and issue type (summary, component, Epic Link, and any custom fields).')
    }
    elseif ($StatusCode -eq 429) {
        $Lines.Add('Jira rate-limited the request (HTTP 429 Too Many Requests).')
        $Lines.Add('')
        $Lines.Add('Wait a short time and retry. If this continues, import fewer items or')
        $Lines.Add('increase jira.requestTimeoutSeconds / retry settings in config.json.')
    }
    elseif ($null -ne $StatusCode -and $StatusCode -ge 500 -and $StatusCode -le 599) {
        $StatusLabel = if (-not (Test-JiraBlankString -Value $StatusName)) {
            "HTTP $StatusCode $StatusName"
        }
        else {
            "HTTP $StatusCode"
        }
        $Lines.Add("Jira is unavailable or returned an error ($StatusLabel).")
        $Lines.Add('')
        $Lines.Add('This is usually a temporary Jira server problem. Retry in a few minutes.')
        if (-not (Test-JiraBlankString -Value $Detail)) {
            $Lines.Add('')
            $Lines.Add($Detail)
        }
    }
    else {
        $Lines.Add('Could not complete the Jira request.')
        $Lines.Add('')
        if (-not (Test-JiraBlankString -Value $Detail)) {
            $Lines.Add($Detail)
        }
        else {
            $Lines.Add($Existing)
        }
        $Lines.Add('')
        $Lines.Add('Check the Jira URL, PAT, and network connection, then retry.')
    }

    if (-not (Test-JiraBlankString -Value $Technical)) {
        $Lines.Add('')
        $Lines.Add("Technical detail: $Technical")
    }

    return ($Lines.ToArray() -join [Environment]::NewLine)
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
        throw (
            "The Jira URL is not a valid HTTPS address: '$BaseUrl'.`n`n" +
            'Set jira.url in config.json to a full HTTPS URL, for example https://jira.example.gov.au'
        )
    }

    if (-not $JiraUri.IsAbsoluteUri -or $JiraUri.Scheme -ne 'https') {
        throw (
            "The Jira URL must be an absolute HTTPS address: '$BaseUrl'.`n`n" +
            'Set jira.url in config.json to a full HTTPS URL, for example https://jira.example.gov.au'
        )
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
            throw (
                "The stored Jira PAT is empty.`n`n" +
                'Fix: re-run this script with -ReplacePat and paste a new token from Jira.'
            )
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

            throw (
                ConvertTo-JiraUserFacingError `
                    -ErrorRecord $_ `
                    -Method $Method `
                    -RelativePath $RelativePath `
                    -Uri $Uri
            )
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
        throw
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

    $FieldMatches = New-Object 'System.Collections.Generic.List[object]'

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
                    $FieldMatches.Add($NestedMatch)
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
            $FieldMatches.Add($Field)
        }
    }

    return $FieldMatches.ToArray()
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

        throw (
            "Multiple Jira Epic Link fields were found: $Ids.`n`n" +
            'The importer cannot choose between them automatically. Ask a Jira admin ' +
            'which field is the Epic Link custom field, then confirm Jira Software is configured as expected.'
        )
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
            "The Jira Epic Link custom field could not be identified.`n`n" +
            'Jira Software may use a different field name in this instance. Related fields: ' +
            "$CandidateText`n`n" +
            'Ask a Jira admin to confirm the Epic Link field is available to the PAT user.'
        )
    }

    throw (
        "The Jira Epic Link custom field could not be found.`n`n" +
        "Jira returned $(@($Fields).Count) fields, but none matched an Epic Link name, " +
        "clause name, or Jira Software Epic Link schema.`n`n" +
        'Confirm this Jira instance has Jira Software (epics) enabled, and that the PAT user can browse issues.'
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
    'Get-JiraEpicLinkField',
    'Get-JiraProjectComponents',
    'Get-JiraProjectIssueTypes',
    'Invoke-JiraRequest',
    'New-JiraClient',
    'Remove-JiraClient',
    'Test-JiraConnection',
    'Test-JiraEpic'
)
