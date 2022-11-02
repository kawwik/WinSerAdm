Import-Module ActiveDirectory

$csvPath = $args[0]
$csv = Import-Csv -Path $csvPath

$groups = [System.Collections.ArrayList]@()
$containers = [System.Collections.ArrayList]@()
$users = [System.Collections.ArrayList]@()

$domainBios= (Get-ADDomain -Current LoggedOnUser).NetBIOSNAME
$domainName= (Get-ADDomain -Current LoggedOnUser).Forest

$ou = $csv.Container
$bios = (Get-ADDomain -Current LoggedOnUser).NetBIOSNAME

$checkOu = [adsi]::Exists("LDAP://OU=$ou, DC=$bios, DC=local")
if (!$checkOu) {
     New-ADOrganizationalUnit -Name $ou -Path "DC=$bios, DC=local"
     $containers.Add($ou)
}

$group = $csv.Group
$checkGroup = Get-ADGroup -Filter {SamAccountName -like $group}
if (!$checkGroup) {
    New-ADGroup -Name $group -SamAccountName $group -GroupScope DomainLocal -Path "OU=$ou, DC=$bios, DC=local"
    $groups.Add($group)
}

$homeDir = $csv.HomeDirectory
$checkDirectory = Test-Path -Path $homeDir
if (!$checkDirectory) {
    $shareName = $homeDir.Split("\")[-1]
    New-Item $homeDir -ItemType Directory
    New-SmbShare -Name $shareName -Path $homeDir
}

$name = $csv.Name.Split()[1]
$login = $csv.Login
try {
    Get-ADUser -Identity $name
    $userExists = $true
}
catch [Microsoft.ActiveDirectory.Management.ADIdentityResolutionException] {
    $userExists = $false
}

if (!$userExists) {
    New-ADUser -Server $domainName -Name "$name" -HomeDrive "X:" -HomeDirectory $csv.Profile -Department $csv.Department -EmailAddress $csv.Email -Title $csv.Position -HomePhone $csv.Phone -SamAccountName $csv.Login -ChangePasswordAtLogon $false -AccountPassword (ConvertTo-SecureString "$csv.Pass" -AsPlainText -Force) -Enabled $true -Path "OU=$ou, DC=$bios, DC=local"
    New-Item $homeDir\$login -ItemType Directory
    $users.Add($name)
}

Add-ADGroupMember -Identity $group -Members $login

class Report {
    [int]$GroupsCount
    [PSCustomobject]$GroupsCreated
    [int]$UsersCount
    [PSCustomobject]$UsersCreated
    [int]$ContainersCount
    [PSCustomobject]$ContainersCreated
}

$report = [Report]::new()
$report.ContainersCreated = $containers
$report.ContainersCount = $containers.Count

$report.GroupsCreated = $groups
$report.GroupsCount = $groups.Count

$report.UsersCreated = $users
$report.GroupsCount = $users.Count

ConvertTo-Html -InputObject $report >> report.html
