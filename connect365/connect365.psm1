function connect365($variable1){
    $global:session365 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell-liveid?DelegatedOrg=$variable1 -Credential $UserCredential -Authentication Basic -AllowRedirection
    Import-Module (Import-PSSession $global:session365 -AllowClobber) -Global
}