cd D:\Artifacts\Software\sppatchify
.\sppatchify.ps1 -downloadMedia -downloadVersion 2019
.\sppatchify.ps1 -downloadMedia -downloadVersion 2016
.\sppatchify.ps1 -downloadMedia -downloadVersion 2013

.\sppatchify.ps1 -CopyMedia
.\sppatchify.ps1 -PauseSharePointSearch
.\sppatchify.ps1 -RunAndInstallCU


.\sppatchify.ps1 -DismountContentDatabase
.\sppatchify.ps1 -MountContentDatabase
.\sppatchify.ps1 -showVersionExit
.\sppatchify.ps1 -testRemotePSExit
.\sppatchify.ps1 -productlocalExit
.\sppatchify.ps1 -EnablePSRemoting
.\sppatchify.ps1 -reportContentDatabasesExit
.\sppatchify.ps1 -ClearCacheIni
.\sppatchify.ps1 -RunConfigWizard
.\sppatchify.ps1 -Advanced #dismount and mount



