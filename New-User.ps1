param(
    [Parameter(
        Mandatory=$True,
        Position=0
    )][string]$GivenName,
    [Parameter(
        Mandatory=$True,
        Position=1
    )][string]$SurName,
    [string]$Title,
    [string]$SamAccountName
    )

. .\menu.ps1
. .\New-SWRandomPassword.ps1

function Select-Department {
    $Department = fShowMenu "Select Department" @{
        "Compounding"="Compounding";
        "Customer Service"="Customer Service";
        "Finance"="Finance";
        "Human Resources"="Human Resources";
        "Information Systems"="Information Systems";
        "Maintenance"="Maintenance";
        "Production"="Production";
        "QA"="QA";
        "QC"="QC";
        "Quality"="Quality";
        "R&D"="R&D";
        "Supply Chain"="Supply Chain";
        "Warehouse"="Warehouse"
    }
    return $Department
}

$OUPath = "OU=Users,OU=USNP,OU=KDC Companies,DC=kdc-companies,DC=com"
$RandPass = New-SWRandomPassword
$AccountPassword = $RandPass | ConvertTo-SecureString -AsPlainText -Force
# $AccountPassword = "Costech1" | ConvertTo-SecureString -AsPlainText -Force
$EmailSuffix = "@kdc-thibiantnaturals.com"
$Company = "Thibiant Naturals"
$Office = "USNP"
$City = "Newbury Park"
$State = "CA"
$PostalCode = "91320"
$Country = "US"

if (-Not ($GivenName)) {
    $GivenName = Read-Host "First Name"
}
if (-Not ($SurName)) {
    $SurName = Read-Host "Last Name"
}
if (-Not ($Title)) {
    $Title = Read-Host "Title"
}

# Force Department string into USNP-DPT naming scheme. If not, then show a menu. Oh well, we tried.
if ($Department) {
    if ($Department -match "^USNP-DPT-") {
        continue
    } else {
        $Department = "USNP-DPT-$Department"
    }
    try {
        Get-ADGroup $Department
    } catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        $Department = Select-Department
    }
}
if (-Not ($Department)) {
    $Department = Select-Department
}

$ADGroup = "CN=USNP-DPT-$Department,OU=Departments,OU=Groups,OU=USNP,OU=KDC Companies,DC=kdc-companies,DC=com"

# Optional - specify SamAccountName (for example, middle name conflicts or capitalization fixes)
if (-Not ($SamAccountName)) {
    $UserName = "$($GivenName.Substring(0,1).ToUpper())$($SurName.Substring(0,1).ToUpper()+$SurName.Substring(1))"
} else {
    $UserName = $SamAccountName
}

$EmailAddress = "$UserName$EmailSuffix"
$FullName = "$GivenName $SurName"
$param = [ordered]@{
    Name                    = $FullName
    Enabled                 = $true
    DisplayName             = $FullName
    GivenName               = $GivenName
    SurName                 = $SurName
    SamAccountName          = $UserName
    EmailAddress            = $EmailAddress
    UserPrincipalName       = $EmailAddress
    Company                 = $Company
    Office                  = $Office
    City                    = $City
    State                   = $State
    PostalCode              = $PostalCode
    Country                 = $Country
    AccountPassword         = $AccountPassword
    ChangePasswordAtLogon   = $True
    Path                    = $OUPath
    OtherAttributes         = @{proxyAddresses = "SMTP:$EMailAddress"}
}
if ($Title) { $param["Title"] = $Title }
if ($Title) { $param["Description"] = $Title }
if ($Department) { $param["Department"] = $DepartmentSelect }
if ($OtherName) { $param["OtherName"] = $OtherName }
if ($StreetAddress) { $param["StreetAddress"] = $StreetAddress }
if ($OfficePhone) { $param["OfficePhone"] = $OfficePhone }
if ($MobilePhone) { $param["MobilePhone"] = $MobilePhone }

New-ADUser @param -Verbose
Add-ADGroupMember -Identity $ADGroup -Members $UserName -Verbose

$MailMessage = New-Object System.Collections.ArrayList
$MailMessage.Add("Access credentials for $FullName")
$MailMessage.Add("<strong>")
$MailMessage.Add("Microsoft")
$MailMessage.Add("</strong>")
$MailMessage.Add("Username: $Username")
$MailMessage.Add("Temporary password: (must change on first logon): $RandPass")

if ($QADUser) {
    $MailMessage.Add("<strong>")
    $MailMessage.Add("QAD / Barcode")
    $MailMessage.Add("</strong>")
    $MailMessage.Add("Username: $QADUsername")
    $MailMessage.Add("Temporary password: (must change on first logon): $QADPass")
    $MailMessage.Add("Please review the attached documents for new user onboarding.")
}
.\New-UserMessage -lines $MailMessage -OutFile "$PSScriptRoot\history\$UserName.docx"  
Write-Output "Username: $Username"
Write-Output "Temporary password (must change on first logon): $RandPass"