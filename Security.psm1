Set-StrictMode -Version Latest

<#
.SYNOPSIS
    Stores and retrieves the Jira Personal Access Token using
    Windows Credential Manager.

.DESCRIPTION
    This module stores the Jira PAT as a Generic Credential in the Windows
    Credential Manager credential set belonging to the current Windows user.

    No PAT file is written under the application Data directory.

    The Windows account that runs the Jira importer must be the same account
    that stores the credential.

    The PAT must never be written to config.json or application logs.
#>

function Assert-WindowsPlatform {
    [CmdletBinding()]
    param()

    $RunningOnWindows = $false

    if ($PSVersionTable.PSVersion.Major -lt 6) {
        $RunningOnWindows = $true
    }
    elseif (
        Get-Variable `
            -Name IsWindows `
            -Scope Global `
            -ErrorAction SilentlyContinue
    ) {
        $RunningOnWindows = [bool]$Global:IsWindows
    }

    if (-not $RunningOnWindows) {
        throw 'Jira PAT storage requires Windows Credential Manager.'
    }
}

Assert-WindowsPlatform

if (-not ('JiraCredentialManager' -as [type])) {
    Add-Type -Language CSharp -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Security;

public static class JiraCredentialManager
{
    private const int CRED_TYPE_GENERIC = 1;
    private const int CRED_PERSIST_LOCAL_MACHINE = 2;
    private const int ERROR_NOT_FOUND = 1168;

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct NativeCredential
    {
        public int Flags;
        public int Type;

        [MarshalAs(UnmanagedType.LPWStr)]
        public string TargetName;

        [MarshalAs(UnmanagedType.LPWStr)]
        public string Comment;

        public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;
        public int CredentialBlobSize;
        public IntPtr CredentialBlob;
        public int Persist;
        public int AttributeCount;
        public IntPtr Attributes;

        [MarshalAs(UnmanagedType.LPWStr)]
        public string TargetAlias;

        [MarshalAs(UnmanagedType.LPWStr)]
        public string UserName;
    }

    [DllImport(
        "advapi32.dll",
        EntryPoint = "CredWriteW",
        CharSet = CharSet.Unicode,
        SetLastError = true)]
    private static extern bool CredWrite(
        ref NativeCredential credential,
        int flags);

    [DllImport(
        "advapi32.dll",
        EntryPoint = "CredReadW",
        CharSet = CharSet.Unicode,
        SetLastError = true)]
    private static extern bool CredRead(
        string targetName,
        int type,
        int reservedFlag,
        out IntPtr credentialPointer);

    [DllImport(
        "advapi32.dll",
        EntryPoint = "CredDeleteW",
        CharSet = CharSet.Unicode,
        SetLastError = true)]
    private static extern bool CredDelete(
        string targetName,
        int type,
        int flags);

    [DllImport(
        "advapi32.dll",
        EntryPoint = "CredFree",
        SetLastError = false)]
    private static extern void CredFree(
        IntPtr credentialPointer);

    public static bool Exists(string targetName)
    {
        IntPtr credentialPointer = IntPtr.Zero;

        try
        {
            if (CredRead(
                targetName,
                CRED_TYPE_GENERIC,
                0,
                out credentialPointer))
            {
                return true;
            }

            int error = Marshal.GetLastWin32Error();

            if (error == ERROR_NOT_FOUND)
            {
                return false;
            }

            throw new Win32Exception(
                error,
                "Windows Credential Manager could not read credential '" +
                targetName + "'.");
        }
        finally
        {
            if (credentialPointer != IntPtr.Zero)
            {
                CredFree(credentialPointer);
            }
        }
    }

    public static void Write(
        string targetName,
        string userName,
        SecureString secret)
    {
        if (String.IsNullOrWhiteSpace(targetName))
        {
            throw new ArgumentException(
                "The credential target cannot be empty.",
                "targetName");
        }

        if (secret == null)
        {
            throw new ArgumentNullException("secret");
        }

        if (secret.Length == 0)
        {
            throw new ArgumentException(
                "The credential secret cannot be empty.",
                "secret");
        }

        IntPtr secretPointer = IntPtr.Zero;

        try
        {
            secretPointer =
                Marshal.SecureStringToGlobalAllocUnicode(secret);

            NativeCredential credential = new NativeCredential();

            credential.Flags = 0;
            credential.Type = CRED_TYPE_GENERIC;
            credential.TargetName = targetName;
            credential.Comment =
                "Jira Email Importer personal access token";
            credential.LastWritten =
                new System.Runtime.InteropServices.ComTypes.FILETIME();
            credential.CredentialBlobSize =
                checked(secret.Length * sizeof(char));
            credential.CredentialBlob = secretPointer;
            credential.Persist = CRED_PERSIST_LOCAL_MACHINE;
            credential.AttributeCount = 0;
            credential.Attributes = IntPtr.Zero;
            credential.TargetAlias = null;
            credential.UserName =
                String.IsNullOrWhiteSpace(userName)
                    ? "Jira PAT"
                    : userName;

            if (!CredWrite(ref credential, 0))
            {
                int error = Marshal.GetLastWin32Error();

                throw new Win32Exception(
                    error,
                    "Windows Credential Manager could not store credential '" +
                    targetName + "'.");
            }
        }
        finally
        {
            if (secretPointer != IntPtr.Zero)
            {
                Marshal.ZeroFreeGlobalAllocUnicode(secretPointer);
            }
        }
    }

    public static SecureString Read(string targetName)
    {
        IntPtr credentialPointer = IntPtr.Zero;

        try
        {
            if (!CredRead(
                targetName,
                CRED_TYPE_GENERIC,
                0,
                out credentialPointer))
            {
                int error = Marshal.GetLastWin32Error();

                if (error == ERROR_NOT_FOUND)
                {
                    throw new InvalidOperationException(
                        "Credential '" + targetName +
                        "' does not exist in Windows Credential Manager.");
                }

                throw new Win32Exception(
                    error,
                    "Windows Credential Manager could not read credential '" +
                    targetName + "'.");
            }

            NativeCredential credential =
                (NativeCredential)Marshal.PtrToStructure(
                    credentialPointer,
                    typeof(NativeCredential));

            if (credential.CredentialBlob == IntPtr.Zero ||
                credential.CredentialBlobSize <= 0)
            {
                throw new InvalidOperationException(
                    "Credential '" + targetName +
                    "' contains an empty secret.");
            }

            if ((credential.CredentialBlobSize % sizeof(char)) != 0)
            {
                throw new InvalidOperationException(
                    "Credential '" + targetName +
                    "' contains an invalid secret.");
            }

            int characterCount =
                credential.CredentialBlobSize / sizeof(char);

            SecureString result = new SecureString();

            for (int index = 0; index < characterCount; index++)
            {
                short characterValue = Marshal.ReadInt16(
                    credential.CredentialBlob,
                    index * sizeof(char));

                result.AppendChar((char)characterValue);
            }

            result.MakeReadOnly();
            return result;
        }
        finally
        {
            if (credentialPointer != IntPtr.Zero)
            {
                CredFree(credentialPointer);
            }
        }
    }

    public static bool Delete(string targetName)
    {
        if (CredDelete(targetName, CRED_TYPE_GENERIC, 0))
        {
            return true;
        }

        int error = Marshal.GetLastWin32Error();

        if (error == ERROR_NOT_FOUND)
        {
            return false;
        }

        throw new Win32Exception(
            error,
            "Windows Credential Manager could not delete credential '" +
            targetName + "'.");
    }
}
'@ -ErrorAction Stop
}

function Test-JiraPatSecret {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Target
    )

    Assert-WindowsPlatform

    try {
        return :Exists($Target)
    }
    catch {
        throw (
            "Could not check Jira PAT credential '$Target'. " +
            "Details: $($_.Exception.Message)"
        )
    }
}

function Set-JiraPatSecret {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Target,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$UserName = 'Jira PAT',

        [Parameter()]
        [System.Security.SecureString]$SecurePat,

        [Parameter()]
        [switch]$Force
    )

    Assert-WindowsPlatform

    $CredentialExists = Test-JiraPatSecret -Target $Target

    if ($CredentialExists -and -not $Force) {
        throw (
            "A Jira PAT credential already exists with target '$Target'. " +
            'Use -Force to replace it.'
        )
    }

    $PromptedForPat = $false

    if ($null -eq $SecurePat) {
        $SecurePat = Read-Host `
            -Prompt 'Enter the Jira Personal Access Token' `
            -AsSecureString

        $PromptedForPat = $true
    }

    try {
        if ($SecurePat.Length -lt 1) {
            throw 'The Jira PAT cannot be empty.'
        }

        if ($PSCmdlet.ShouldProcess(
            $Target,
            'Store Jira PAT in Windows Credential Manager'
        )) {
            try {
                :Write(
                    $Target,
                    $UserName,
                    $SecurePat
                )

                # Read the credential back to confirm that it was stored.
                $ValidationValue =
                    :Read($Target)

                try {
                    if (
                        $null -eq $ValidationValue -or
                        $ValidationValue.Length -lt 1
                    ) {
                        throw (
                            'Windows Credential Manager returned an ' +
                            'empty credential after storage.'
                        )
                    }
                }
                finally {
                    $ValidationValue = $null
                }

                return [pscustomobject]@{
                    Target    = $Target
                    CreatedAt = Get-Date
                    UserName  = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
                    Machine   = $env:COMPUTERNAME
                    Store     = 'Windows Credential Manager'
                }
            }
            catch {
                throw (
                    'Could not store the Jira PAT in Windows ' +
                    "Credential Manager. Details: $($_.Exception.Message)"
                )
            }
        }
    }
    finally {
        if ($PromptedForPat) {
            $SecurePat = $null
        }
    }
}

function Get-JiraPatSecret {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Target
    )

    Assert-WindowsPlatform

    try {
        $SecurePat = :Read($Target)

        if ($null -eq $SecurePat) {
            throw 'Windows Credential Manager returned no value.'
        }

        if ($SecurePat.Length -lt 1) {
            throw 'The stored Jira PAT is empty.'
        }

        return $SecurePat
    }
    catch {
        throw (
            "Could not retrieve Jira PAT credential '$Target'. " +
            'Ensure that the credential was created by the Windows ' +
            'account currently running this application. ' +
            "Details: $($_.Exception.Message)"
        )
    }
}

function ConvertFrom-SecureStringPlainText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [System.Security.SecureString]$SecureString
    )

    $Pointer = [System.IntPtr]::Zero

    try {
        $Pointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR(
            $SecureString
        )

        return [Runtime.InteropServices.Marshal]::PtrToStringBSTR(
            $Pointer
        )
    }
    finally {
        if ($Pointer -ne [System.IntPtr]::Zero) {
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR(
                $Pointer
            )
        }
    }
}

function Remove-JiraPatSecret {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Target
    )

    Assert-WindowsPlatform

    if (-not (Test-JiraPatSecret -Target $Target)) {
        return $false
    }

    if ($PSCmdlet.ShouldProcess(
        $Target,
        'Remove Jira PAT from Windows Credential Manager'
    )) {
        try {
            return :Delete($Target)
        }
        catch {
            throw (
                "Could not remove Jira PAT credential '$Target'. " +
                "Details: $($_.Exception.Message)"
            )
        }
    }

    return $false
}

Export-ModuleMember -Function @(
    'ConvertFrom-SecureStringPlainText',
    'Get-JiraPatSecret',
    'Remove-JiraPatSecret',
    'Set-JiraPatSecret',
    'Test-JiraPatSecret'
)
