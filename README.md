DSCR_MSLicense
====

PowerShell DSC Resource to set Windows / Office product key.

## Install
You can install Resource through [PowerShell Gallery](https://www.powershellgallery.com/packages/DSCR_MSLicense/).
```Powershell
Install-Module -Name DSCR_MSLicense
```

## Resources
* **cWindowsLicense**
    + For Windows operating system.
    + Support Ver. : Windows 7 and later

* **cOfficeLicense**
    + For Microsoft Office suite.
    + Support Ver. : Office 2013 & 2016
    + Office 365 is NOT supported

## Properties
### cWindowsLicense
+ [string] **Ensure** (Write):
    + Specify installation state of the product key.
    + The default value is Present. { Present | Absent }

+ [string] **ProductKey** (Key):
    + Product key for install.
    + The format is "XXXXX-XXXXX-XXXXX-XXXXX-XXXXX"

+ [Boolean] **Activate** (Write):
    + If `$true`, this resource check not only whether product key is installed, but also whether it is activated.
    + If the system hasn't activate yet, will try to that.
    + The default value is `$false`.

+ [Boolean] **Force** (Write):
    + If `$true`, Test-TargetResource always return `$false`.
    + In short, regardless of the current state, the installation of the product key will be executed.

### cOfficeLicense
+ [string] **Ensure** (Write):
    + Specify installation state of the product key.
    + The default value is Present. { Present | Absent }

+ [string] **ProductKey** (Key):
    + Product key for install.
    + The format is "XXXXX-XXXXX-XXXXX-XXXXX-XXXXX"

+ [string] **OfficeVersion** (Required):
    + Specify the version of the Office. { Office2013 | Office2016 }

+ [Boolean] **Activate** (Write):
    + If `$true`, this resource check not only whether product key is installed, but also whether it is activated.
    + If the system hasn't activate yet, will try to that.
    + The default value is `$false`.

+ [Boolean] **Force** (Write):
    + If `$true`, Test-TargetResource always return `$false`.
    + In short, regardless of the current state, the installation of the product key will be executed.

## Examples
+ **Example 1**: Install product key and try to activate
```Powershell
Configuration Example1
{
    Import-DscResource -ModuleName DSCR_MSLicense
    cWindowsLicense Win10_Pro
    {
        ProductKey = "W269N-WFGWX-YVC9B-4J6C9-T83GX"
        Activate   = $true
    }

    cOfficeLicense Office_ProPlus_2016
    {
        ProductKey    = "XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99"
        OfficeVersion = "Office2016"
        Activate      = $true
    }
}
```

## ChangeLog
