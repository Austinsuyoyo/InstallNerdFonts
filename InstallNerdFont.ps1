
#######################################################################################
# Import Module
Import-Module $PSScriptRoot\utils\Write-Menu.psm1
Import-Module $PSScriptRoot\utils\Show-Pause.psm1

#######################################################################################
$title = "1.[Option] All User or Current User"
$message = "Select 'All User' will installed in C:\Windows\Fonts\ `nSelect 'Current User' will installed in $($env:LOCALAPPDATA)\Microsoft\Windows\Fonts\"

$all = New-Object System.Management.Automation.Host.ChoiceDescription "&All User", "Installed in C:\Windows\Fonts\"
$cur = New-Object System.Management.Automation.Host.ChoiceDescription "&Current User", "Installed in $($env:LOCALAPPDATA)\Microsoft\Windows\Fonts\"

$choices = [System.Management.Automation.Host.ChoiceDescription[]]($all, $cur)
$AllUserFlag = $host.UI.PromptForChoice($title, $message, $choices, 0)

if($AllUserFlag -eq 0){
    #https://superuser.com/questions/749243/detect-if-powershell-is-running-as-administrator
    if(-not([bool](([System.Security.Principal.WindowsIdentity]::GetCurrent()).groups -match "S-1-5-32-544"))){
        Show-Pause "Please run as Admininstrator"
        Exit
    }
}

#######################################################################################
# Select Font
# ref:https://stackoverflow.com/questions/58855377/add-menu-options-in-a-running-powershell-script
Write-Host "Show the list of Nerd-Font latest version"
$ReleasePage = "https://api.github.com/repos/ryanoasis/nerd-fonts/releases/latest"
$Json = Invoke-WebRequest $ReleasePage | ConvertFrom-Json
$LatestVersion = $Json[0].tag_name

$Font_Name_Extend = Write-Menu -Title '2.[Option] Nerd-Font Menu' -Entries @($Json[0].assets.name)
if ($null -eq $Font_Name_Extend) {
    Show-Pause "Invalid choice `nPress any key to exit."
    Exit
}
$Font_Name = $Font_Name_Extend.Split(".")[0]



#######################################################################################
# Download Sources from github
$DownloadUrl = "https://github.com/ryanoasis/nerd-fonts/releases/download/$LatestVersion/$Font_Name_Extend"

#check whether *.zip file exist
if (-not(Test-Path $PSScriptRoot\$Font_Name_Extend)) {
    Write-Host Downloading latest release
    Invoke-WebRequest $DownloadUrl -Out $PSScriptRoot\$Font_Name_Extend
}
else {
    Write-Host Already Downloads the Font
}
#check whether font.zip already expan or not
if (-not(Test-Path $PSScriptRoot\$Font_Name\)) {
    Write-Host Extracting release files
    Expand-Archive -LiteralPath $PSScriptRoot\$Font_Name_Extend -DestinationPath $PSScriptRoot\$Font_Name
}
else {
    Write-Host Already expand archive
}
########################################################################################
## TODO:
## If downaloaded font have two types(otf & ttf)
## If just one type no need ask
$title = "3.[Option] Font Type"
$message = "Select Font Type to Installed: .otf (Recommend) or .ttf"

$otf = New-Object System.Management.Automation.Host.ChoiceDescription "&otf", "OpenType Font"
$ttf = New-Object System.Management.Automation.Host.ChoiceDescription "&ttf", "TrueType Font"

$choices = [System.Management.Automation.Host.ChoiceDescription[]]($otf, $ttf)
$returnValue = $host.UI.PromptForChoice($title, $message, $choices, 0)
$FontType = "*." + $choices[$returnValue].Label.Trim("&")


########################################################################################
# Insatll Font 
#Reference:
#https://github.com/mikeTWC1984/pwshise/blob/ee3209eac079e4b0adc40c569f25ec77a75a6ad3/fonts/FontInstaller.ps1
#https://gist.github.com/anthonyeden/0088b07de8951403a643a8485af2709b
#https://richardspowershellblog.wordpress.com/2008/03/20/special-folders/
#https://jordanmalcolm.com/deploying-windows-10-fonts-at-scale/
if ($IsWindows) {
    if ($AllUserFlag -eq 0) {
        # Check if already installed 
        # '*.ttf', '*.ttc', '*.otf'
        Get-ChildItem -Path $PSScriptRoot\$Font_Name -Include $FontType -Recurse | ForEach-Object {
            if (-not(Test-Path "C:\Windows\Fonts\$($_.Name)")) {
                Write-Host Installing font  $($_.BaseName) For All User
            
                # Install for all user
                Copy-Item $_.FullName "C:\Windows\Fonts"
                New-ItemProperty -Name $_.BaseName -Path "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Fonts" -PropertyType string -Value $_.name         
            }
            else {
                Write-Host $($_.Name) already installed
            }
        }
    }
    else {
        $Destination = (New-Object -ComObject Shell.Application).Namespace(0x14)
        Get-ChildItem -Path $PSScriptRoot\$Font_Name -Include $FontType -Recurse | ForEach-Object {
            if (-not(Test-Path "$env:LOCALAPPDATA\Microsoft\Windows\Fonts\$($_.Name)")) {   
                Write-Host Installing font  $($_.BaseName) For Current User

                # Install font for current user
                $Destination.CopyHere($_.FullName, 0x14)
            }
            else {
                Write-Host $($_.Name) already installed
            }
        }
    }
}
#if($IsLinux){}
#if($IsMacOS){}

#
# Clean
Write-Host Cleanup $Font_Name folder and zip
Remove-Item $PSScriptRoot\$Font_Name -Recurse -ErrorAction SilentlyContinue 
Remove-Item $PSScriptRoot\$Font_Name_Extend