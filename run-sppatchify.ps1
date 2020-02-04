cd D:\Artifacts\Software\sppatchify
.\sppatchify.ps1 -downloadMedia -downloadVersion 2019
.\sppatchify.ps1 -CopyMedia
.\sppatchify.ps1 -showVersionExit
.\sppatchify.ps1 -testRemotePSExit
.\sppatchify.ps1 -productlocalExit
.\sppatchify.ps1 -saveServiceInstanceExit
.\sppatchify.ps1 -reportContentDatabasesExit
.\sppatchify.ps1 -ClearCacheIni
.\sppatchify.ps1 -Standard
.\sppatchify.ps1 -RunAndInstallCU
<#
 -- AutoSPSourceBuilder SharePoint Update Download/Integration Utility --
Start-BitsTransfer : The operation being requested was not performed because the user has not logged on to the network. The specified service does not exist. (Exception from HRESULT: 0x800704DD)
At D:\Artifacts\Software\sppatchify\SPPatchify.ps1:2442 char:9
+         Start-BitsTransfer -DisplayName "Downloading AutoSPSourceBuil ...
+         ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : NotSpecified: (:) [Start-BitsTransfer], COMException
    + FullyQualifiedErrorId : System.Runtime.InteropServices.COMException,Microsoft.BackgroundIntelligentTransfer.Management.NewBitsTransferCommand
 
Could not download AutoSPSourceBuilder.xml file!
At D:\Artifacts\Software\sppatchify\SPPatchify.ps1:2444 char:13
+             throw "Could not download AutoSPSourceBuilder.xml file!"
+             ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : OperationStopped: (Could not downl...ilder.xml file!:String) [], RuntimeException
    + FullyQualifiedErrorId : Could not download AutoSPSourceBuilder.xml file!
    #>
.\sppatchify.ps1 -RunAndInstallCU

