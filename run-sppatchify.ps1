# step 1 - change to SPPatchify Directory
cd D:\DEPL\Software\sppatchify
cd C:\DEPL\Software\sppatchify
cd D:\Artifacts\Software\sppatchify

# step 2 - Download Media
.\sppatchify.ps1 -downloadMedia -downloadVersion 2019
.\sppatchify.ps1 -downloadMedia -downloadVersion 2016
.\sppatchify.ps1 -downloadMedia -downloadVersion 2013

# step 3 - Copy Media to Servers
.\sppatchify.ps1 -CopyMedia

# step 4 - Stop Search (Option for SP2016/2019)
# We are no longer doing this for SP2016 and SP2019, as install not taking 5 hours
# some say this speeds up install of the CU# 
#.\sppatchify.ps1 -PauseSharePointSearch

# step 5 - Install CU and Run PSConfig afrer automatic reboot
.\sppatchify.ps1 -RunAndInstallCU # run parellel


# other One Off Commands

.\sppatchify.ps1 -RunConfigWizard
.\sppatchify.ps1 -StartSharePointSearch
.\sppatchify.ps1 -ClearCacheIni
.\sppatchify.ps1 -DismountContentDatabase
.\sppatchify.ps1 -MountContentDatabase #mount and update
.\sppatchify.ps1 -RebootServer
.\sppatchify.ps1 -DismountContentDatabase
.\sppatchify.ps1 -MountContentDatabase
.\sppatchify.ps1 -showVersionExit
.\sppatchify.ps1 -testRemotePSExit
.\sppatchify.ps1 -productlocalExit
.\sppatchify.ps1 -EnablePSRemoting
.\sppatchify.ps1 -IISStart




#future
.\sppatchify.ps1 -DismountContentDatabase -UpgradeNeeded #these block psconfig

<# 
verify mounted databases

$files = Get-ChildItem "D:\DEPL\Software\sppatchify\log\contentdbs-*.csv" | Sort-Object LastAccessTime -Desc
$dbs = Import-Csv $files.Fullname  
        Write-Host "Content DB - create script blocks" -Fore Yellow      
        foreach ($db in $dbs) {  
            
             $db.name
                     
                        Add-PSSnapIn Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue | Out-Null   
                        Get-SPContentDatabase | Where-Object {$_.Name -eq $db.name} 
                        }
                        #>

Get-SPContentDatabase | select NormalizedDataSource, name, needsupgrade | ft -AutoSize


