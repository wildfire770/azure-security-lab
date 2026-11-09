# New-LabUsers.ps1
# Creates 10 Marvel-themed test users in the Colberg Cloud lab tenant,
# assigns Microsoft 365 E5, and puts a few in a CA pilot group.
# Safe to rerun: existing users are skipped.
# Run as admin@lab.colberg.cloud (Cloud Shell or a non-managed PC).

$ErrorActionPreference = "Stop"
$Domain  = "lab.colberg.cloud"
$OutFile = "./lab-user-passwords.csv"   # keep this OUT of GitHub

# Load (or install) only the Graph modules this script needs
$Modules = "Microsoft.Graph.Authentication","Microsoft.Graph.Users","Microsoft.Graph.Users.Actions",
           "Microsoft.Graph.Groups","Microsoft.Graph.Identity.DirectoryManagement"
foreach ($m in $Modules) {
    if (-not (Get-Module -ListAvailable -Name $m)) {
        Write-Host "Installing $m ..."
        Install-Module $m -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $m
}

Connect-MgGraph -Scopes "User.ReadWrite.All","Group.ReadWrite.All","Organization.Read.All" -UseDeviceCode -NoWelcome

$Users = @(
    @{ First="Tony";    Last="Stark";    Title="CTO";                  Dept="IT";          Pilot=$true  }
    @{ First="Thor";    Last="Odinson";  Title="CEO";                  Dept="Executive";   Pilot=$false }
    @{ First="Carol";   Last="Danvers";  Title="Director of Security"; Dept="Security";    Pilot=$false }
    @{ First="Natasha"; Last="Romanoff"; Title="Security Analyst";     Dept="Security";    Pilot=$true  }
    @{ First="Steve";   Last="Rogers";   Title="Operations Manager";   Dept="Operations";  Pilot=$false }
    @{ First="Bruce";   Last="Banner";   Title="Research Scientist";   Dept="R&D";         Pilot=$false }
    @{ First="Wanda";   Last="Maximoff"; Title="HR Generalist";        Dept="HR";          Pilot=$false }
    @{ First="Sam";     Last="Wilson";   Title="Account Executive";    Dept="Sales";       Pilot=$false }
    @{ First="Shuri";   Last="Udaku";    Title="Software Engineer";    Dept="Engineering"; Pilot=$true  }
    @{ First="Peter";   Last="Parker";   Title="IT Intern";            Dept="IT";          Pilot=$false }
)

function New-RandomPassword {
    $chars = "ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789!@#%*"
    -join (1..16 | ForEach-Object { $chars[(Get-Random -Maximum $chars.Length)] })
}

# Microsoft 365 E5 SKU
$E5 = Get-MgSubscribedSku | Where-Object SkuPartNumber -eq "SPE_E5"
if (-not $E5) { Write-Warning "M365 E5 SKU not found - users will be created unlicensed." }

# Pilot group for Conditional Access testing
$Group = Get-MgGroup -Filter "displayName eq 'CA-Pilot'" | Select-Object -First 1
if (-not $Group) {
    $Group = New-MgGroup -DisplayName "CA-Pilot" -MailEnabled:$false -MailNickname "ca-pilot" `
        -SecurityEnabled -Description "Pilot group for Conditional Access testing"
}

$Results = foreach ($u in $Users) {
    $nick = "$($u.First.ToLower()).$($u.Last.ToLower())"
    $upn  = "$nick@$Domain"

    try {
        $existing = Get-MgUser -Filter "userPrincipalName eq '$upn'"
        if ($existing) { Write-Host "Skipped $upn (already exists)"; continue }

        $pw  = New-RandomPassword
        $new = New-MgUser -AccountEnabled `
            -DisplayName "$($u.First) $($u.Last)" -GivenName $u.First -Surname $u.Last `
            -UserPrincipalName $upn -MailNickname $nick `
            -JobTitle $u.Title -Department $u.Dept -UsageLocation "US" `
            -PasswordProfile @{ Password = $pw; ForceChangePasswordNextSignIn = $false }

        if ($E5) {
            Set-MgUserLicense -UserId $new.Id -AddLicenses @(@{ SkuId = $E5.SkuId }) -RemoveLicenses @() | Out-Null
        }
        if ($u.Pilot) {
            New-MgGroupMember -GroupId $Group.Id -DirectoryObjectId $new.Id
        }

        Write-Host "Created $upn" -ForegroundColor Green
        [pscustomobject]@{ UPN = $upn; Password = $pw; Department = $u.Dept; Pilot = $u.Pilot }
    }
    catch {
        Write-Host "FAILED $upn : $($_.Exception.Message)" -ForegroundColor Red
    }
}

if ($Results) {
    $Results | Export-Csv $OutFile -NoTypeInformation
    Write-Host "`nPasswords for new users saved to $OutFile - download it, then don't commit it."
} else {
    Write-Host "`nNo new users created."
}
