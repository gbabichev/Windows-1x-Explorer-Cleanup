<#
Hides the folders in This PC
https://learn.microsoft.com/en-us/dotnet/desktop/winforms/controls/known-folder-guids-for-file-dialog-custom-places?view=netframeworkdesktop-4.8
#>

$registryKeys = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions\{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}\PropertyBag',
'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions\{f42ee2d3-909f-4907-8871-4c22fc0bf756}\PropertyBag',
'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions\{7d83ee9b-2244-4e70-b1f5-5393042af1e4}\PropertyBag',
'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions\{a0c69a99-21c8-4671-8703-7934162fcf1d}\PropertyBag',
'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions\{0ddd015d-b06c-45d5-8c4c-f59713854639}\PropertyBag',
'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions\{35286a68-3c57-41a1-bbb1-0eae73d76c95}\PropertyBag',
'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions\{31C0DD25-9439-4F12-BF41-7FF4EDA38722}\PropertyBag',
'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions\{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}\PropertyBag',
'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions\{f42ee2d3-909f-4907-8871-4c22fc0bf756}\PropertyBag',
'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions\{7d83ee9b-2244-4e70-b1f5-5393042af1e4}\PropertyBag',
'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions\{a0c69a99-21c8-4671-8703-7934162fcf1d}\PropertyBag',
'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions\{0ddd015d-b06c-45d5-8c4c-f59713854639}\PropertyBag',
'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions\{35286a68-3c57-41a1-bbb1-0eae73d76c95}\PropertyBag',
'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Explorer\FolderDescriptions\{31C0DD25-9439-4F12-BF41-7FF4EDA38722}\PropertyBag'
)
$folderNames = @('Desktop64',
'Documents64',
'Downloads64',
'Music64',
'Pictures64',
'Videos64',
'3D Objects64',
'Desktop32',
'Documents32',
'Downloads32',
'Music32',
'Pictures32',
'Videos32',
'3D Objects32')

Function GetCurrentFolderStatus
{
    $table = New-Object System.Data.Datatable
    [void]$table.Columns.Add("Folder")
    [void]$table.Columns.Add("Status")

    $starter = 0
    foreach ($i in $registryKeys)
    {
        $a = Get-ItemProperty -Path $i -Name "ThisPCPolicy" -ErrorAction SilentlyContinue | select -ExpandProperty "ThisPCPolicy" 
        if ($a -eq "Show")
        {
            #Write-Host $a
            [void]$table.Rows.Add($folderNames[$starter],$a)
        }
        elseif ($a -eq "Hide"){
            [void]$table.Rows.Add($folderNames[$starter],$a)
        }
        elseif ($a -eq "")
        {
            [void]$table.Rows.Add($folderNames[$starter],"Show")
        }
        else {
            # PS threw an error. Check for key.
            if (-not (Get-Item -Path $i -ErrorAction SilentlyContinue)){
                # Path does not exist, create it.
                New-Item -Path $i | Out-Null
                New-ItemProperty -Path $i -Name "ThisPCPolicy" -PropertyType "string" | Out-Null
                [void]$table.Rows.Add($folderNames[$starter],"Show")
            }
            else {
                try
                {
                    New-ItemProperty -Path $i -Name "ThisPCPolicy" -PropertyType "string" | Out-Null
                    [void]$table.Rows.Add($folderNames[$starter],"Show")
                }
                catch
                {
                    [void]$table.Rows.Add($folderNames[$starter],"Error")
                }
            }
        }
        $starter += 1
    }
    $table | out-host
}

Function BonusExplorerEdits
{
    # Sets Explorer to launch to "This PC" by default instead of  "Quick Access".
    try
    {
        New-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "LaunchTo" -Value "1" -PropertyType "dword" | Out-Null 
    }
    catch
    {
        Write-Host "Error setting Explorer to open 'This PC' instead of 'Quick Access'"
    }
    # Modifies explorer, unchecks the "Show frequently used folders" box.
    try 
    {
        New-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer" -Name "ShowFrequent" -Value "0" -PropertyType "dword" | Out-Null
    }
    catch 
    {
        Write-Host "Error setting 'Show frequently used folders' setting."
    }
    # Modifies explorer, unchecks the "Show recently used files" box.
	try 
    {
        New-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer" -Name "ShowRecent" -Value "0" -PropertyType "dword" | Out-Null
    }
    catch
    {
        Write-Host "Error setting 'Show recently used folders' setting."
    }
    
    Write-Host "Done applying bonus Explorer changes."
}

# MAIN 
GetCurrentFolderStatus

$userInput = Read-Host "Would you like to hide the folders? (y/n)"
$starter = 0
if ($userInput -eq "y")
{
    foreach ($i in $registryKeys)
    {
        #Write-Host "Set" $folderNames[$starter] "to hide"
        try
        {
            Set-ItemProperty -Path $i -Name "ThisPCPolicy" -Value "Hide" 
            $starter += 1
        }
        catch 
        {
            Write-Host "An error occured setting $i"
        }
    }
}
else {
    Write-Host "Cancelling..."
    exit
}

GetCurrentFolderStatus

$userInput = Read-Host "Would you like to apply bonus Explorer changes? 
- Open explorer to 'This PC' instead of 'Quick Access'
- Do not 'show frequently used folders' under Home
- Do not 'show recently used files' under Home
...(y/n)"
if ($userInput -eq "y")
{
    BonusExplorerEdits
}

Write-Host "Done...
You may need to restart the 'explorer.exe' Process manually for all changes to take effect."
Read-Host "Press enter to quit"

