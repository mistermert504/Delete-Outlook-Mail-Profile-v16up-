###############################################################################
# Info: Delete all mail profiles and create a new empty one                   #
# Compatibility: From Outlook 2016 upwards                                    #
# Author: Martin Eberle                                                       #
# Version: v1.0, 26.03.2021                                                   #
# Source: https://github.com/mistermert504/Delete-Outlook-Mail-Profile-v16up- #                                                          # 
###############################################################################

#Disable Outlook Save Mode
Write-Output "Disable Outlook Save Mode"
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office\16.0\Outlook\Security\" -Name "DisableSafeMode" -Type DWord -Value 1

#Kill Outlook if activ
if($proc=(get-process 'outlook' -ErrorAction SilentlyContinue))
{
    Write-Output "Outlook is running so close it.."
    Kill($proc)
    Write-Output "Outlook is stopped"
}

#Delete old Outlook Profile
$reg="HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles"
$child=(Get-ChildItem -Path $reg).name
foreach($item in $child)
{
    Remove-item -Path registry::$item -Recurse
}
Write-Output "All profiles removed successfully"

#Create new Outlook Profile
Write-Output "Now create new profile for outlook"
New-Item -Name "outlook" -Path $reg -Force -Verbose
Write-Output "New profile outlook created"

#Launch Outlook with new profile
Write-Output "Launch outlook with newly created profile"
Start-Process 'outlook' -ErrorAction SilentlyContinue -ArgumentList '/profile "outlook"'
