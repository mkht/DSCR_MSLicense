$output = 'C:\MOF'

Configuration DSCR_MSLicense_Sample
{
    Import-DscResource -ModuleName DSCR_MSLicense
    Node localhost
    {
        cWindowsLicense WinLicense_Sample
        {
            Ensure = "Present"
            ProductKey = "FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4"
            Activate = $false
        }

        cOfficeLicense OfficeLicense_Sample
        {
            Ensure = "Present"
            ProductKey = "YC7DK-G2NP3-2QQC3-J6H88-GVGXT"
            Activate = $false
            OfficeVersion = 'Office2013'
        }
    }
}

DSCR_MSLicense_Sample -OutputPath $output
Start-DscConfiguration -Path  $output -Verbose -wait -force
