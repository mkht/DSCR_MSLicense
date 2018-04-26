DATA Const{
    ConvertFrom-StringData -stringdata @'
        ProductClass = SoftwareLicensingProduct
        ProductClassWin7 = OfficeSoftwareProtectionProduct
        OfficeAppId = 0ff1ce15-a989-479d-af46-f275c6370663
        OfficeAppId2010 = 59a52881-a989-479d-af46-f275c6370663
'@
}

# ////////////////////////////////////////////////////////////////////////////////////////
# ////////////////////////////////////////////////////////////////////////////////////////
Enum OfficeLicenseStatus {
    Unlicensed = 0
    Licensed = 1
    OOBGrace = 2
    OOTGrace = 3
    NonGenuineGrace = 4
    Notification = 5
    ExtendedGrace = 6
}

# ////////////////////////////////////////////////////////////////////////////////////////
# ////////////////////////////////////////////////////////////////////////////////////////
Enum OSVersion {
    Win7 = 7     # Win7 or Server2008R2
    Win8 = 8     # Win8 or Server2012
    Win81 = 81    # Win8.1 or Server2012R2
    Win10 = 10    # Win10 or Server2016
    Unknown = 0     # Unknown or ~Vista
}

# ////////////////////////////////////////////////////////////////////////////////////////
# ////////////////////////////////////////////////////////////////////////////////////////
Enum OfficeVersion {
    Office2010 = 14
    Office2013 = 15
    Office2016 = 16
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
        $Force = $false,

        [parameter(Mandatory = $true)]
        #[ValidateSet("Office2010","Office2013","Office2016")]
        [ValidateSet("Office2013", "Office2016")]
        [string]
        $OfficeVersion
    )
    #$ErrorActionPreference = 'Stop'
    $GetRes = @{
        Ensure            = $Ensure
        ProductKey        = ''
        Activate          = $false
        OfficeVersion     = $OfficeVersion
        PartialProductKey = ''
        LicenseStatus     = ''
        Product           = ''
    }

    # Check specific version of Office is installed or not
    if (-not (Get-OSPPVBS $OfficeVersion)) {
        # Write-Verbose ('{0} is not installed.' -f $OfficeVersion)
        $GetRes.Ensure = 'Absent'
        return $GetRes
    }

    # Check the Product key was installed or not
    $Product = Get-PrimaryOfficeSKU $OfficeVersion -ea SilentlyContinue
    if (-not $Product) {
        #Write-Verbose ('Product key is not installed on this machine.')
        $GetRes.Ensure = 'Absent'
    }
    else {
        $GetRes.Ensure = 'Present'
        $GetRes.Product = $Product.Name

        #Check activation status
        $GetRes.LicenseStatus = [string]([OfficeLicenseStatus]$Product.LicenseStatus)
        if ($Product.LicenseStatus -ne [OfficeLicenseStatus]::Licensed) {
            #Write-Verbose ('Office is NOT Licensed. (License status: "{0}")' -f [OfficeLicenseStatus]$Product.LicenseStatus)
            $GetRes.Activate = $false
        }
        else {
            #Write-Verbose 'Office is Licensed.'
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
        $Force = $false,

        [parameter(Mandatory = $true)]
        #[ValidateSet("Office2010","Office2013","Office2016")]
        [ValidateSet("Office2013", "Office2016")]
        [string]
        $OfficeVersion
    )
    # $ErrorActionPreference = 'Stop'

    # Test the product key format (XXXXX-XXXXX-XXXXX-XXXXX-XXXXX)
    if ($ProductKey) {
        if ($ProductKey -notmatch '^[0-9A-Z]{5}(-[0-9A-Z]{5}){4}$') {
            Write-Error ('The product key is invalid')
            return
        }
    }

    $Ospp = Get-OSPPVBS $OfficeVersion
    if (-not $Ospp) {
        Write-Error ('OSPP.VBS not found')
    }
    $Ospp = ('"{0}"' -f $Ospp)

    $Cscript = 'C:\Windows\System32\cscript.exe'
    if (-not (Test-Path $Cscript)) {
        Write-Error ('"{0}" not found' -f $Cscript)
    }

    $cState = (Get-TargetResource @PSBoundParameters)

    if ($Ensure -eq 'Absent') {
        # Remove Product Key
        if ($cState.PartialProductKey.length -ne 5) {
            Write-Error 'Error happened when removing the product key (PartialProductKey not found)'
        }
        else {
            $Exec = (Start-Command -FilePath $Cscript -ArgumentList ($Ospp, ('/unpkey:{0}' -f $cState.PartialProductKey)))
            $Err = Assert-Error $Exec.StdOut
            if ($Err.ErrorCode) {Write-Error ('Error: {0} ({1})' -f $Err.ErrorMsg, $Err.ErrorCode)}
            else {Write-Verbose ('Remove the product key succeeded')}
        }
    }
    else {
        $isKeyChanged = $false

        $local:tmpPPKey = $ProductKey.Substring($ProductKey.Length - 5, 5)  # Last 5 chars
        if (($cState.Ensure -eq 'Absent') -or ($cState.PartialProductKey -ne $tmpPPKey) -or $Force) {
            # Install Product Key
            $Exec = (Start-Command -FilePath $Cscript -ArgumentList ($Ospp, ('/inpkey:{0}' -f $ProductKey)))
            $Err = Assert-Error $Exec.StdOut
            if ($Err.ErrorCode) {Write-Error ('Error: {0} ({1})' -f $Err.ErrorMsg, $Err.ErrorCode)}
            else {Write-Verbose ('Install the product key succeeded')}
            $isKeyChanged = $true
        }

        if ($Activate) {
            if ($isKeyChanged -or (!$cState.Activate)) {
                #Activation
                $Exec = (Start-Command -FilePath $Cscript -ArgumentList ($Ospp, '/act'))
                $Err = Assert-Error $Exec.StdOut
                if ($Err.ErrorCode) {Write-Error ('Error: {0} ({1})' -f $Err.ErrorMsg, $Err.ErrorCode)}
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
        $Force = $false,

        [parameter(Mandatory = $true)]
        #[ValidateSet("Office2010","Office2013","Office2016")]
        [ValidateSet("Office2013", "Office2016")]
        [string]
        $OfficeVersion
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
            Write-Verbose ('Product key not installed. It is your desired state.')
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
            Write-Verbose ('Product key not installed. It is NOT your desired state.')
            return $false
        }

        $local:tmpPPKey = $ProductKey.Substring($ProductKey.Length - 5, 5)  # Last 5 chars
        if ($cState.PartialProductKey -ne $tmpPPKey) {
            Write-Verbose ('Partial product key is not matched. (current:"{0}" / desired: "{1}")' -f $cState.PartialProductKey, $tmpPPKey)
            return $false
        }

        if ($Activate) {
            if (!$cState.Activate) {
                Write-Verbose ('Office is NOT Licensed. (License status: "{0}")' -f [OfficeLicenseStatus]$cState.LicenseStatus)
                return $false
            }
            else {
                Write-Verbose 'Office is Licensed.'
            }
        }
    }

    return $true
} # end of Test-TargetResource

# ////////////////////////////////////////////////////////////////////////////////////////
# ////////////////////////////////////////////////////////////////////////////////////////
function Get-PrimaryOfficeSKU {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory, Position = 0)]
        [OfficeVersion]$OfficeVersion
    )

    if ($OfficeVersion -eq [OfficeVersion]::Office2010) {
        $ApplicationId = $Const.OfficeAppId2010
    }
    else {
        $ApplicationId = $Const.OfficeAppId
    }

    if ((Get-OSVersion) -eq [OSVersion]::Win7) {
        $ProductClass = $Const.ProductClassWin7
    }
    else {
        $ProductClass = $Const.ProductClass
    }

    # This logic is ported from OSPP.VBS
    $QueryStr = ('SELECT * FROM {0} WHERE PartialProductKey IS NOT NULL' -f $ProductClass)
    $object = Get-WmiObject -Query $QueryStr -ErrorAction SilentlyContinue |
        where {($_.ApplicationId -eq $ApplicationId) -and (!$_.LicenseIsAddon)}

    if (@($object).Count -ne 1) {
        return $null
    }
    else {
        return $object
    }
}

# ////////////////////////////////////////////////////////////////////////////////////////
# ////////////////////////////////////////////////////////////////////////////////////////
function Get-OSVersion {
    [CmdletBinding()]
    Param()
    # https://msdn.microsoft.com/ja-jp/library/windows/desktop/ms724832(v=vs.85).aspx
    $Ver = ([System.Environment]::OSVersion).Version
    switch (('{0}.{1}' -f $Ver.Major, $Ver.Minor)) {
        '10.0' { return [OSVersion]::Win10 }
        '6.3' { return [OSVersion]::Win81 }
        '6.2' { return [OSVersion]::Win8  }
        '6.1' { return [OSVersion]::Win7  }
    }
    return [OSVersion]::Unknown
}

# ////////////////////////////////////////////////////////////////////////////////////////
# ////////////////////////////////////////////////////////////////////////////////////////
function Get-OSPPVBS {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory, Position = 0)]
        [OfficeVersion]$OfficeVersion
    )
    $ErrorActionPreference = 'SilentlyContinue'
    $ProgramFiles = $env:ProgramFiles
    $ProgramFilesX86 = ${env:ProgramFiles(x86)}

    $Path = Join-Path $ProgramFilesX86 ('Microsoft Office\Office{0}\OSPP.VBS' -f [int]$OfficeVersion) -ErrorAction SilentlyContinue
    if (Test-Path $Path -PathType Leaf) {
        return $Path
    }

    $Path = Join-Path $ProgramFiles ('Microsoft Office\Office{0}\OSPP.VBS' -f [int]$OfficeVersion) -ErrorAction SilentlyContinue
    if (Test-Path $Path -PathType Leaf) {
        return $Path
    }
}

# ////////////////////////////////////////////////////////////////////////////////////////
# ////////////////////////////////////////////////////////////////////////////////////////
function Assert-Error {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory, Position = 0)]
        [string]$OsppOutput
    )

    if ($OsppOutput -match 'ERROR CODE: (0x.+)') {
        $ErrorCode = $Matches[1]
        if ($OsppOutput -match 'ERROR DESCRIPTION: (.+)') {
            $ErrorMsg = $Matches[1]
        }
    }

    return [pscustomobject]@{
        ErrorCode = $ErrorCode
        ErrorMsg  = $ErrorMsg
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
    $ProcessInfo.RedirectStandardError = $true
    $ProcessInfo.RedirectStandardOutput = $true
    $Process = New-Object System.Diagnostics.Process
    $Process.StartInfo = $ProcessInfo
    $Process.Start() | Out-Null
    if (!$Process.WaitForExit($Timeout)) {
        $Process.Kill()
        Write-Error ('Process timeout. Terminated. (Timeout:{0}s, Process:{1})' -f ($Timeout * 0.001), $FilePath)
    }
    $stdout = $Process.StandardOutput.ReadToEnd()
    $stderr = $Process.StandardError.ReadToEnd()
    [pscustomobject]@{
        Process  = $Process
        StdOut   = $stdout
        StdErr   = $stderr
        ExitCode = $Process.ExitCode
    }
}

Export-ModuleMember -Function *-TargetResource
