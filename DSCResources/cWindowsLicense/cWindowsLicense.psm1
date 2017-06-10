DATA Const{
    ConvertFrom-StringData -stringdata @'
        ServiceClass = SoftwareLicensingService
        ProductClass = SoftwareLicensingProduct
        WindowsAppId = 55c92734-d682-4d71-983e-d6ec3f16059f
'@
}

# ////////////////////////////////////////////////////////////////////////////////////////
# ////////////////////////////////////////////////////////////////////////////////////////
#https://technet.microsoft.com/en-US/library/dn502536(v=ws.11).aspx
Enum WinLicenseStatus
{
    Unlicensed       = 0
    Licensed         = 1
    OOBGrace         = 2
    OOTGrace         = 3
    NonGenuineGrace  = 4
    Notification     = 5
    ExtendedGrace    = 6
}

# ////////////////////////////////////////////////////////////////////////////////////////
# ////////////////////////////////////////////////////////////////////////////////////////
function Get-TargetResource {
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param
    (
        [ValidateSet("Present", "Absent")]
        [string]
        $Ensure = 'Present',

        [parameter(Mandatory = $true)]
        [string]
        $ProductKey,

        [parameter()]
        [bool]
        $Activate = $true,

        [parameter()]
        [bool]
        $Force = $false

        # [parameter()]
        # [bool]
        # $IsKmsClient = $false,

        # [parameter()]
        # [string]
        # $KmsServer
    )

    $GetRes = @{
        Ensure            = $Ensure
        ProductKey        = ''
        Activate          = $false
        #IsKmsClient = $false
        #KmsServer = $KmsServer
        PartialProductKey = ''
        LicenseStatus     = ''
        OS                = ''
    }

    # Check the Product key was installed or not
    $Product = Get-PrimaryWindowsSKU -ea SilentlyContinue
    if (-not $Product) {
        #Write-Verbose ('Product key is not installed on this machine.')
        $GetRes.Ensure = 'Absent'
    }
    else {
        $GetRes.Ensure = 'Present'
        $GetRes.OS = $Product.Name

        #Check activation status
        $GetRes.LicenseStatus = [string]([WinLicenseStatus]$Product.LicenseStatus)
        if ($Product.LicenseStatus -ne [WinLicenseStatus]::Licensed) {
            #Write-Verbose ('Windows is NOT Licensed. (License status: "{0}")' -f [WinLicenseStatus]$Product.LicenseStatus)
            $GetRes.Activate = $false
        }
        else {
            #Write-Verbose 'Windows is Licensed.'
            $GetRes.Activate = $true
        }

        #Get partial product key (Last 5 chars)
        $GetRes.PartialProductKey = $Product.PartialProductKey
    }
    $GetRes
} # end of Get-TargetResource

# ////////////////////////////////////////////////////////////////////////////////////////
# ////////////////////////////////////////////////////////////////////////////////////////
function Set-TargetResource {
    [CmdletBinding()]
    param
    (
        [ValidateSet("Present", "Absent")]
        [string]
        $Ensure = 'Present',

        [parameter(Mandatory = $true)]
        [string]
        $ProductKey,

        [parameter()]
        [bool]
        $Activate = $true,

        [parameter()]
        [bool]
        $Force = $false
    )
    # $ErrorActionPreference = 'Stop'

    # Test the product key format (XXXXX-XXXXX-XXXXX-XXXXX-XXXXX)
    if ($ProductKey) {
        if ($ProductKey -notmatch '^[0-9A-Z]{5}(-[0-9A-Z]{5}){4}$') {
            Write-Error ('The product key is invalid')
            return
        }
    }

    $Slmgr = 'C:\Windows\System32\slmgr.vbs'
    $Cscript = 'C:\Windows\System32\cscript.exe'
    @($Slmgr, $Cscript).ForEach( {
            if (-not (Test-Path $_)) {
                Write-Error ('"{0}" not found' -f $_)
            }
        })

    $cState = (Get-TargetResource @PSBoundParameters)

    if ($Ensure -eq 'Absent') {
        # Remove Product Key
        $ExitCode = (Start-Command -FilePath $Cscript -ArgumentList ($Slmgr, '-upk')).ExitCode
        if ($ExitCode -ne 0) {Write-Error ('Error happend when removing the product key')}
        else {Write-Verbose ('Remove the product key succeeded')}

        # Remove Product Key from registry
        $ExitCode = (Start-Command -FilePath $Cscript -ArgumentList ($Slmgr, '-cpky')).ExitCode
        if ($ExitCode -ne 0) {Write-Error ('Error happend when removing the product key from registry')}
        else {Write-Verbose ('Remove the product key from registry succeeded')}
    }
    else {
        $isKeyChenged = $false

        $local:tmpPPKey = $ProductKey.Substring($ProductKey.Length - 5, 5)  # Last 5 chars
        if (($cState.Ensure -eq 'Absent') -or ($cState.PartialProductKey -ne $tmpPPKey) -or $Force) {
            # Install Product Key
            $ExitCode = (Start-Command -FilePath $Cscript -ArgumentList ($Slmgr, '-ipk', $ProductKey)).ExitCode
            if ($ExitCode -ne 0) {Write-Error ('Error happend when installing the product key')}
            else {Write-Verbose ('Install the product key succeeded')}
            $isKeyChenged = $true
        }

        if ($Activate) {
            if ($isKeyChenged -or (!$cState.Activate)) {
                #Activation
                $ExitCode = (Start-Command -FilePath $Cscript -ArgumentList ($Slmgr, '-ato')).ExitCode
                if ($ExitCode -ne 0) {Write-Error ('Error happend when try to activation')}
                else {Write-Verbose ('Activation succeeded')}
            }
        }
    }
} # end of Set-TargetResource

# ////////////////////////////////////////////////////////////////////////////////////////
# ////////////////////////////////////////////////////////////////////////////////////////
function Test-TargetResource {
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
        [ValidateSet("Present", "Absent")]
        [string]
        $Ensure = 'Present',

        [parameter(Mandatory = $true)]
        [string]
        $ProductKey,

        [parameter()]
        [bool]
        $Activate = $true,

        [parameter()]
        [bool]
        $Force = $false
    )

    # Test the product key format (XXXXX-XXXXX-XXXXX-XXXXX-XXXXX)
    if ($ProductKey) {
        if ($ProductKey -notmatch '^[0-9A-Z]{5}(-[0-9A-Z]{5}){4}$') {
            Write-Error ('The product key is invalid')
            return $true
        }
    }

    $cState = (Get-TargetResource @PSBoundParameters)

    if ($Ensure -eq 'Absent') {
        if ($cState.Ensure -eq 'Absent') {
            Write-Verbose ('Product key is not installed on this machine. It is your desired state.')
            return $true
        }
        else {
            Write-Verbose ('Product key is installed. It is NOT your desired state.')
            return $false
        }
    }
    elseif ($Ensure -eq 'Present') {
        if ($Force) {
            Write-Verbose ("Specified 'Force' switch. Skip test and return False.")
            return $false
        }

        if ($cState.Ensure -eq 'Absent') {
            Write-Verbose ('Product key is not installed on this machine. It is NOT your desired state.')
            return $false
        }

        $local:tmpPPKey = $ProductKey.Substring($ProductKey.Length - 5, 5)  # Last 5 chars
        if ($cState.PartialProductKey -ne $tmpPPKey) {
            Write-Verbose ('Partial product key is not matched. (current:"{0}" / desired: "{1}")' -f $cState.PartialProductKey, $tmpPPKey)
            return $false
        }

        if ($Activate) {
            if (!$cState.Activate) {
                Write-Verbose ('Windows is NOT Licensed. (License status: "{0}")' -f [WinLicenseStatus]$cState.LicenseStatus)
                return $false
            }
            else {
                Write-Verbose 'Windows is Licensed.'
            }
        }
    }

    return $true
} # end of Test-TargetResource

# ////////////////////////////////////////////////////////////////////////////////////////
# ////////////////////////////////////////////////////////////////////////////////////////
function Get-PrimaryWindowsSKU {
    [CmdletBinding()]
    Param()

    # This logic is ported from slmgr.vbs
    $QueryStr = ('SELECT * FROM {0} WHERE PartialProductKey IS NOT NULL' -f $Const.ProductClass)
    $object = Get-WmiObject -Query $QueryStr -ErrorAction SilentlyContinue |
        where {($_.ApplicationId -eq $Const.WindowsAppId) -and (!$_.LicenseIsAddon)}

    if (@($object).Count -ne 1) {
        return $null
    }
    else {
        return $object
    }
}

# ////////////////////////////////////////////////////////////////////////////////////////
# ////////////////////////////////////////////////////////////////////////////////////////
function Start-Command {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true, Position = 0)]
        [string] $FilePath,
        [Parameter(Mandatory = $false, Position = 1)]
        [string[]]$ArgumentList,
        [int]$Timeout = [int]::MaxValue
    )
    $ProcessInfo = New-Object System.Diagnostics.ProcessStartInfo
    $ProcessInfo.FileName = $FilePath
    $ProcessInfo.UseShellExecute = $false
    $ProcessInfo.Arguments = [string]$ArgumentList
    $Process = New-Object System.Diagnostics.Process
    $Process.StartInfo = $ProcessInfo
    $Process.Start() | Out-Null
    if (!$Process.WaitForExit($Timeout)) {
        $Process.Kill()
        Write-Error ('Process timeout. Terminated. (Timeout:{0}s, Process:{1})' -f ($Timeout * 0.001), $FilePath)
    }
    $Process
}

Export-ModuleMember -Function *-TargetResource
