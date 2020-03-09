cd D:\DEPL\Software\sppatchify
cd C:\DEPL\Software\sppatchify
cd D:\Artifacts\Software\sppatchify
.\sppatchify.ps1 -downloadMedia -downloadVersion 2019
.\sppatchify.ps1 -downloadMedia -downloadVersion 2016
.\sppatchify.ps1 -downloadMedia -downloadVersion 2013
.\sppatchify.ps1 -Standard #install CU, run psconfig, open CA

.\sppatchify.ps1 -CopyMedia
.\sppatchify.ps1 -PauseSharePointSearch
.\sppatchify.ps1 -RunAndInstallCU # run parellel
.\sppatchify.ps1 -DismountContentDatabase
.\sppatchify.ps1 -RunConfigWizard
.\sppatchify.ps1 -MountContentDatabase #mount and update
.\sppatchify.ps1 -StartSharePointSearch
.\sppatchify.ps1 -RebootServer

.\sppatchify.ps1 -DismountContentDatabase
.\sppatchify.ps1 -MountContentDatabase
.\sppatchify.ps1 -showVersionExit
.\sppatchify.ps1 -testRemotePSExit
.\sppatchify.ps1 -productlocalExit
.\sppatchify.ps1 -EnablePSRemoting
.\sppatchify.ps1 -IISStart
.\sppatchify.ps1 -ClearCacheIni
.\sppatchify.ps1 -RunConfigWizard
.\sppatchify.ps1 -Advanced #dismount and mount

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


