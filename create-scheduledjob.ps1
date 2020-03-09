$parmater="RunConfigWizard"
$root ="C:\DEPL\Software\sppatchify"
$ScheduledJobOption  = New-ScheduledJobOption -RunElevated

$JobTrigge = New-JobTrigger -AtStartup -RandomDelay 00:01:00

 import-Module WebAdministration
$account=(Get-SPFarm).DefaultServiceAccount.Name
$password = (Get-ChildItem IIS:\AppPools | where { $_.processModel.userName -eq $account })[0].processModel.password  
$passwordSecure = ConvertTo-SecureString ($password) -AsPlainText -Force
$Cred = New-Object System.Management.Automation.PSCredential ($account, $passwordSecure)

Register-ScheduledJob -ScheduledJobOption $ScheduledJobOption -Trigger $JobTrigge -Credential $Cred   -Name "SSPparmater" -FilePath $root\sppatchify.ps1 -ArgumentList $parmater


#     Get-ScheduledJob -Name "SSPparmater" | Unregister-ScheduledJob -Force