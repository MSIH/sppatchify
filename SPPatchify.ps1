<#
.SYNOPSIS
	SharePoint Central Admin - View active services across entire farm. No more select machine drop down dance!
.DESCRIPTION
	Apply CU patch to entire farm from one PowerShell console.
    NOTE - must run local to a SharePoint server under account with farm admin rights.
	Comments and suggestions always welcome!  spjeff@spjeff.com or @spjeff
.NOTES
	File Namespace	: SPPatchify.ps1
	Author			: Jeff Jones - @spjeff
	Version			: 0.144
    Last Modified	: 10-04-2019
    
.LINK
	Source Code
	http://www.github.com/spjeff/sppatchify
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $False, ValueFromPipeline = $false, HelpMessage = 'Use -d -downloadMediaOnly to execute Media Download only.  No farm changes.  Prep step for real patching later.')]
    [Alias("d")]
    [switch]$downloadMediaOnly,
    [string]$downloadVersion,

    [Parameter(Mandatory = $False, ValueFromPipeline = $false, HelpMessage = 'Use -rebootFarmOnly.')] 
    [switch]$rebootFarmOnly,    

    [Parameter(Mandatory = $False, ValueFromPipeline = $false, HelpMessage = 'Use -StartSharePointSearchOnly.')] 
    [switch]$StartSharePointSearchOnly,
    
    [Parameter(Mandatory = $False, ValueFromPipeline = $false, HelpMessage = 'Use -PauseSharePointSearchOnly.')] 
    [switch]$PauseSharePointSearchOnly,  

    [Parameter(Mandatory = $False, ValueFromPipeline = $false, HelpMessage = 'Use -c -CopyMedia to copy \media\ across all peer machines.  No farm changes.  Prep step for real patching later.')]
    [switch]$CopyMedia,

    [Parameter(Mandatory = $False, ValueFromPipeline = $false, HelpMessage = 'Use -v -showVersionOnly to show farm version info.  READ ONLY, NO SYSTEM CHANGES.')]

    [switch]$showVersionOnly, 

    [Parameter(Mandatory = $False, ValueFromPipeline = $false, HelpMessage = 'Use -phaseTwo to execute Phase Two after local reboot.')]
    [switch]$phaseTwo,

    [Parameter(Mandatory = $False, ValueFromPipeline = $false, HelpMessage = 'Use -phaseThree to execute Phase Three attach and upgrade content.')]
    [switch]$phaseThree,
	
    [Parameter(Mandatory = $False, ValueFromPipeline = $false, HelpMessage = 'Use -remoteSessionPort to open PSSession (remoting) with custom port number.')]
    [string]$remoteSessionPort,

    [Parameter(Mandatory = $False, ValueFromPipeline = $false, HelpMessage = 'Use -remoteSessionSSL to open PSSession (remoting) with SSL encryption.')]
    [switch]$remoteSessionSSL,

    [Parameter(Mandatory = $False, ValueFromPipeline = $false, HelpMessage = 'Use -test to open Remote PS Session and verify connectivity all farm members.')]
    [switch]$testRemotePSOnly,

    [Parameter(Mandatory = $False, ValueFromPipeline = $false, HelpMessage = 'Use -targetServers to run for specific machines only.  Applicable to PhaseOne and PhaseTwo.')]
    [string[]]$targetServers,

    [Parameter(Mandatory = $False, ValueFromPipeline = $false, HelpMessage = 'Use -productlocalOnly to execute remote cmdlet [Get-SPProduct -Local] on all servers in farm, or target/wave servers only if given.')]
    [switch]$productlocalOnly,

    [Parameter(Mandatory = $False, ValueFromPipeline = $false, HelpMessage = 'Use -mount to execute Mount-SPContentDatabase to load CSV and attach content databases to web applications.')]
    [string]$mount,

    [Parameter(Mandatory = $False, ValueFromPipeline = $false, HelpMessage = 'Use -appOffline TRUE/FALSE to COPY app_offline.htm] file to all servers and all IIS websites (except Default Website).')]
    [string]$appOffline,

    [Parameter(Mandatory = $False, ValueFromPipeline = $false, HelpMessage = 'Use -bypass to run with PACKAGE.BYPASS.DETECTION.CHECK=1')]
    [switch]$bypass, 

    [Parameter(Mandatory = $False, ValueFromPipeline = $false, HelpMessage = 'Use -saveServiceInstanceOnly to snapshot CSV with current Service Instances running.')]
    [switch]$saveServiceInstanceOnly,

    [Parameter(Mandatory = $False, ValueFromPipeline = $false, HelpMessage = 'Use -reportContentDatabasesOnly to snapshot CSV with Content Databases.')]
    [switch]$reportContentDatabasesOnly,

    [Parameter(Mandatory = $False, ValueFromPipeline = $false, HelpMessage = 'Use -startSharePointRelatedServicesOnly to start sharepoint and iis services.')]
    [switch]$startSharePointRelatedServicesOnly,

    [Parameter(Mandatory = $False, ValueFromPipeline = $false, HelpMessage = 'Use -stopSharePointRelatedServicesOnly to stop sharepoint and iis services.')]
    [switch]$stopSharePointRelatedServicesOnly,

    [Parameter(Mandatory = $False, ValueFromPipeline = $false, HelpMessage = 'Use -EnablePSRemotingOnly to enable CredSSP.')]
    [switch]$EnablePSRemotingOnly,

    [Parameter(Mandatory = $False, ValueFromPipeline = $false, HelpMessage = 'Use -ClearCacheIni to clear chache ini folder.')]
    [switch]$ClearCacheIni,

    [Parameter(Mandatory = $False, ValueFromPipeline = $false, HelpMessage = 'Use -StopSPDistributedCache to stop sharepoint and iis services.')]
    [switch]$StopSPDistributedCache,

    [Parameter(Mandatory = $False, ValueFromPipeline = $false, HelpMessage = 'Use -IISStart to stop sharepoint and iis services.')]
    [switch]$IISStart,

    [Parameter(Mandatory = $False, ValueFromPipeline = $false, HelpMessage = 'Use -RunAndInstallCU to stop sharepoint and iis services.')]
    [switch]$RunAndInstallCU,

    [Parameter(Mandatory = $False, ValueFromPipeline = $false, HelpMessage = 'Use -DismountContentDatabase to stop sharepoint and iis services.')]
    [switch]$DismountContentDatabase,
    [bool] $needsUpdateOnly = $false,

    [Parameter(Mandatory = $False, ValueFromPipeline = $false, HelpMessage = 'Use -RunConfigWizard to stop sharepoint and iis services.')]
    [switch]$RunConfigWizard,

    [Parameter(Mandatory = $False, ValueFromPipeline = $false, HelpMessage = 'Use -MountContentDatabase to stop sharepoint and iis services.')]
    [switch]$MountContentDatabase,

    [Parameter(Mandatory = $False, ValueFromPipeline = $false, HelpMessage = 'Use -UpgradeContent to stop sharepoint and iis services.')]
    [switch]$UpgradeContent,

    [Parameter(Mandatory = $False, ValueFromPipeline = $false, HelpMessage = 'Use -Standard download and copy CU, pause search, install CU, run psconfig and then start search.')]
    [switch]$Standard,

    [Parameter(Mandatory = $False, ValueFromPipeline = $false, HelpMessage = 'Use -Advanced download and copy CU, pause search, install CU, dismount content databases, run psconfig, mount content databases, and then start search.')]
    [switch]$Advanced,

    [Parameter(Mandatory = $False, ValueFromPipeline = $false, HelpMessage = 'Use -Complete download and copy CU, clear ini cache, restart IIS,save status of services, pause search, install CU, dismount content databases, run psconfig, mount content databases, start services, and then start search.')]
    [switch]$Complete,

    [Parameter(Mandatory = $False, ValueFromPipeline = $false, HelpMessage = 'Use -AfterReboot to tell command starting after reboot.')]
    [switch]$AfterReboot
)

# Plugin
Add-PSSnapIn Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue | Out-Null
Import-Module WebAdministration -ErrorAction SilentlyContinue | Out-Null

$host.ui.RawUI.WindowTitle = "SPPatchify"
$rootCmd = $MyInvocation.MyCommand.Definition
$root = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$maxattempt = 3
$maxrebootminutes = 120
$logFolder = "$root\log"

function Main() {
    # Clean up
    Get-PSSession | Remove-PSSession -Confirm:$false

    $mainArgs = $args

    if (-not $mainArgs) {
        return
    }

    # what does this do
    If (!(Test-Path -Path $root -PathType Container)) {
        $remoteRoot = MakeRemote $root
    }

    $remoteRoot = MakeRemote $root

    # Local farm servers - note removed global variables
    # $global:servers = Get-SPServer | Where-Object { $_.Role -ne "Invalid" } | Sort-Object Address
    # List - Target servers
    <#
    if ($targetServers) {
        $global:servers = Get-SPServer | Where-Object { $targetServers -contains $_.Name } | Sort-Object Address
    }
    #>    
    
    #Create Local Log Folders
    If (!(Test-Path -Path $logFolder -PathType Container)) {
        mkdir "$logFolder" -ErrorAction SilentlyContinue | Out-Null
    }
    If (!(Test-Path -Path $logFolder\msp -PathType Container)) {
        mkdir "$logFolder\msp" -ErrorAction SilentlyContinue | Out-Null
    }

    # Start Logging
    $start = Get-Date
    $when = $start.ToString("yyyy-MM-dd-hh-mm-ss")
    $logFile = "$logFolder\SPPatchify-$when.txt"
    Start-Transcript $logFile          

    # Download media
    if ($downloadMediaOnly) {
        PatchRemoval
        PatchMenu 
        Stop-Transcript; Exit       
    }    

    # Halt if no servers detected
    if ((getFarmServers).Count -eq 0) {
        Write-Host "HALT - POWERSHELL ERROR - No SharePoint servers detected.  Close this window and run from new window." -Fore Red
        Stop-Transcript; Exit        
    }
    else {
        Write-Host "Servers Online: $((getFarmServers).Count)"   
    }   

    # Create remote log folders
    LoopRemoteCmd "Create log directory on" "mkdir '$logFolder' -ErrorAction SilentlyContinue | Out-Null"
    LoopRemoteCmd "Create log directory on" "mkdir '$logFolder\msp' -ErrorAction SilentlyContinue | Out-Null"  

    if ($rebootFarmOnly) {
        rebootFarm
        Stop-Transcript; Exit       
    }

    if ($EnablePSRemotingOnly) {  
        # Enable CredSSP remoting      
        EnablePSRemoting
        Stop-Transcript; Exit
    }

    # Enable CredSSP remoting      
    EnablePSRemoting

    # Verify Remote PowerShell
    if (-not (VerifyRemotePS)) {
        return
    }
    
    # Save Service Instance
    if ($saveServiceInstanceOnly) {
        SaveServiceInst
        # display file
        Stop-Transcript; Exit
    }

    # Run SPPL to detect new binary patches
    if ($productlocalOnly) {
        TestRemotePS
        ProductLocal
        Stop-Transcript; Exit
    }
        
    # Test PowerShell
    if ($testRemotePSOnly) {
        TestRemotePS
        Stop-Transcript; Exit
    }
    
	
    # Display version
    if ($showVersionOnly) {
        ShowVersion
        Stop-Transcript; Exit
    }
 
    # Mount Databases
    if ($reportContentDatabasesOnly) {
        ReportContentDatabases
        # display file
        Stop-Transcript; Exit
    }

    #  search
    if ($PauseSharePointSearchOnly) {
        PauseSharePointSearch
        # display file
        Stop-Transcript; Exit
    }

    if ($StartSharePointSearchOnly) {
        StartSharePointSearch
        # display file
        Stop-Transcript; Exit
    }



    <# Change Services
    if ($changeServices.ToUpper() -eq "TRUE") {
        changeServices $true
        Exit
    }
    if ($changeServices.ToUpper() -eq "FALSE") {
        changeServices $false
        Exit
    }
    #>

    # Change Services
    if ($startSharePointRelatedServicesOnly) {
        changeServices $true
        Stop-Transcript; Exit
    }
    if ($stopSharePointRelatedServicesOnly) {
        changeServices $false
        Stop-Transcript; Exit
    }

    # Install App_Offline
    if ($appOffline.ToUpper() -eq "TRUE") {
        AppOffline $true
        Stop-Transcript; Exit
    }
    if ($appOffline.ToUpper() -eq "FALSE") {
        AppOffline $false
        Stop-Transcript; Exit
    }	

    function StartSharePointRelatedServices() {
        changeServices $true
       
    }

    function StopSharePointRelatedServices() {
        changeServices $false
    }

    if ($CopyMedia) {   
        # Copy media only (switch -C)  
        # does not require remoting, use unc path    
        CopyMedia "Copy"
    }  
    
    if ($ClearCacheIni) {  
        # Cler the Cache INI Folder 
        # does not require remoting, use unc path     
        ClearCacheIni
    }   
    
    if ($SaveServiceInst) { 
        # Save SP Service Instances that are online to CVS file  
        # does not require remoting    
        SaveServiceInst
    }
    
    if ($StopSPDistributedCache) { 
        # Stop StopSPDistributedCache on all farm servers  
        # uses CredSSP remoting    
        StopSPDistributedCache
    }
    
    if ($IISStart) {   
        # Start IIS App Pools, Sites, IIS Admin service, and Web service
        # uses CredSSP remoting    
        IISStart
    }


    # Run CU, wait for servers to reboot, verify installed on all servers
    if ($RunAndInstallCU) {  
        RunAndInstallCU
    } 


    if ($DismountContentDatabase) {   
        # Run PSconfigure on all servers
        # does not require remoting     
        DismountContentDatabase $needsUpdateOnly
    } 

    if ($RunConfigWizard) {   
        # Run PSconfigure on all servers
        # uses CredSSP remoting    
        RunPSconfig
        VerifyCUInstalledOnAllServers
        DisplayCA
    } 

    if ($MountContentDatabase) {  
        # Run PSconfigure on all servers
        # does not require remoting      
        MountContentDatabase
    } 

    if ($UpgradeContent) {  
        # Run PSconfigure on all servers
        # does not require remoting      
        UpgradeContent
    } 

    if ($Standard) {      

        RunAndInstallCU($mainArgs)        
        VerifyCUInstalledOnAllServers  
        RunPSconfig
        StartSharePointSearch
        DisplayCA
    }

    if ($Advanced) {
        # dismount databases before running psconfig
        # mount databases distributed  
        # PauseSharePointSearch - this can take up to an hour, so do before
        RunAndInstallCU($mainArgs)   
        VerifyCUInstalledOnAllServers 
        DismountContentDatabase 
        RunPSconfig
        MountContentDatabase 
        StartSharePointSearch
        DisplayCA
    }

    if ($Complete) {
        # dismount databases before running psconfig
        # mount databases distributed           
        PatchRemoval
        PatchMenu    
        CopyMedia "Copy"
        ClearCacheIni #Complete
        IISStart #Complete
        SaveServiceInst  
        PauseSharePointSearch      
        RunAndInstallCU($mainArgs)           
        VerifyCUInstalledOnAllServers 
        DismountContentDatabase #Advanced
        RunPSconfig
        MountContentDatabase #Advanced
        StartSharePointRelatedServices #Complete
        UpgradeContent
        StartSharePointSearch
        DisplayCA
    }  

    #remove all scheduled tasks
    $taskName = "SSP_*"
    foreach ($server in getFarmServers) {       
        $addr = $server.Address
        Write-Host "Unregister task $taskName from - $addr" -Fore Green
        if ($addr -eq $env:computername) {              
            # Remove SCHTASK if found
            $found = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue 
            if ($found) {
                $found | Unregister-ScheduledTask -Confirm:$false 
            }   
        }
        else {
            # Remove SCHTASK if found
            $found = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue -CimSession $addr
            if ($found) {
                $found | Unregister-ScheduledTask -Confirm:$false -CimSession $addr
            }
        } 
    }
            
    # remove all pssession
    Get-PSSession | Remove-PSSession -Confirm:$false

    # Calculate Duration and Run Cleanup    
    CalcDuration
  
    Write-Host "DONE ===$(Get-Date)"
    Stop-Transcript
}

#region binary EXE
function MakeRemote($path) {
    # Remote UNC
    $char = $path.ToCharArray()
    if ($char[1] -eq ':') {
        $char[1] = '$'
    }
    return (-join $char)
}

function CopyMedia($action = "Copy") {
    Write-Host "===== $action EXE ===== $(Get-Date)" -Fore "Yellow"

    # Clear old session
    Get-Job | Remove-Job -Force
    Get-PSSession | Remove-PSSession -Confirm:$false

    # Start Jobs
    foreach ($server in getRemoteServers) {
        $addr = $server.Address
        if ($addr -ne $env:computername) {
            Write-Host "===== ROBOCOPY ===== $(Get-Date)" -Fore "Yellow"
            # Dynamic command
            $dest = "\\$addr\$remoteRoot\media"
            mkdir $dest -Force -ErrorAction SilentlyContinue | Out-Null;
            ROBOCOPY ""$root\media"" ""$dest"" /Z /MIR /W:0 /R:0
        }
    }

    # Watch Jobs
    Start-Sleep 5

    $counter = 0
    do {
        foreach ($server in getRemoteServers) {
            # Progress
            if (Get-Job) {
                $prct = [Math]::Round(($counter / (Get-Job).Count) * 100)
                if ($prct) {
                    Write-Progress -Activity "Copy EXE ($prct %) $(Get-Date)" -Status $addr -PercentComplete $prct -ErrorAction SilentlyContinue
                }
            }

            # Check Job Status
            Get-Job | Format-Table -AutoSize
        }
        Start-Sleep 5
        $pending = Get-Job | Where-Object { $_.State -eq "Running" -or $_.State -eq "NotStarted" }
        $counter = (Get-Job).Count - $pending.Count
    }
    while ($pending)

    # Complete
    Get-Job | Format-Table -a
    Write-Progress -Activity "Completed $(Get-Date)" -Completed
}

function VerifyCUInstalledOnAllServers() {
    # Display server upgrade
    Write-Host "Farm Servers - Upgrade Status $(Get-Date)" -Fore "Yellow"
    (Get-SPProduct).Servers | Select-Object Servername, InstallStatus | Sort-Object Servername | Format-Table -AutoSize

    $halt = (Get-SPProduct).Servers | Where-Object { $_.InstallStatus -eq "InstallRequired" }
    if ($halt) {
        $halt | Format-Table -AutoSize
        Write-Host "HALT - MEDIA ERROR - Install on servers" -Fore Red
        Stop-Transcript; Exit
    }
}

function SafetyEXE() {
    # notused
    Write-Host "===== SafetyEXE ===== $(Get-Date)" -Fore "Yellow"

    # Count number of files.   Must be 3 for SP2013 (major ver 15)

    # Build CMD
    $ver = (Get-SPFarm).BuildVersion.Major
    if ($ver -eq 15) {
        foreach ($server in getFarmServers) {
            $addr = $server.Address
            $c = (Get-ChildItem "\\$addr\$remoteRoot\media").Count
            if ($c -ne 3) {
                $halt = $true
                Write-Host "HALT - MEDIA ERROR - Expected 3 files on \\$addr\$remoteRoot\media" -Fore Red
            }
        }

        # Halt
        if ($halt) {
            Stop-Transcript; Exit
        }
    }
}


function RunAndInstallCU($mainArgs) {
    Write-Host "===== RunAndInstallCU ===== $(Get-Date)" -Fore "Yellow"

    # Remove MSPLOG
    Write-Host "===== Remove MSPLOG on ===== $(Get-Date)" -Fore "Yellow"
    LoopRemoteCmd "Remove MSPLOG on " "Remove-Item '$logfolder\msp\*' -Confirm:`$false -ErrorAction SilentlyContinue" -isJob $true

    # Remove MSPLOG
    Write-Host "===== Unblock EXE on ===== $(Get-Date)" -Fore "Yellow"
    LoopRemoteCmd "Unblock EXE on " "gci '$root\media\*' | Unblock-File -Confirm:`$false -ErrorAction SilentlyContinue" -isJob $true

    # Build CMD
    $files = Get-ChildItem "$root\media\*.exe" -Recurse | Sort-Object Name
    If ($files) {
        foreach ($f in $files) {
            # Display patch name
            $name = $f.Name
            Write-Host $name -Fore Yellow
            $patchName = $name.replace(".exe", "")
            $cmd = $f.FullName
            Write-Host "$cmd ===== $(Get-Date)" -Fore "Yellow"
            $params = "/passive /forcerestart /log:""$root\log\msp\$name.log"""
            if ($bypass) {
                $params += " PACKAGE.BYPASS.DETECTION.CHECK=1"
            }
            $taskName = "SPP_InstallCU"

            # Loop - Run Task Scheduler
            foreach ($server in getFarmServers) {
                # Local PC - No reboot
                $addr = $server.Address
                Write-Host $addr 
                if ($addr -eq $env:computername) {
                    $params = $params.Replace("forcerestart", "norestart")
                    # Remove SCHTASK if found
                    $found = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue 
                    if ($found) {
                        $found | Unregister-ScheduledTask -Confirm:$false 
                    }

                    # New SCHTASK parameters
                    $user = "System"
                    $folder = Split-Path $f
                    $a = New-ScheduledTaskAction -Execute $cmd -Argument $params -WorkingDirectory $folder 
                    $p = New-ScheduledTaskPrincipal -RunLevel Highest -UserId $user -LogonType S4U

                    # Create and Start SCHTASK
                
                    Write-Host "Register and start SCHTASK - $addr - $cmd" -Fore Green
                    Register-ScheduledTask -TaskName $taskName -Action $a -Principal $p -Description "Install SharePoint CU created by SPPatchify Tool" 
                    Start-ScheduledTask -TaskName $taskName 
                    Write-Host "Start SCHTASK $addr ===== $(Get-Date)" -Fore "Yellow"

                    # Event log START
                    New-EventLog -LogName "Application" -Source "SPPatchify" -ComputerName $addr -ErrorAction SilentlyContinue | Out-Null
                    Write-EventLog -LogName "Application" -Source "SPPatchify" -EntryType Information -Category 1000 -EventId 1000 -Message "START" -ComputerName $addr              
                }
                else {

                    if ($addr -eq $env:computername) { 
                        Start-ScheduledTask -TaskName $taskName 
                    }

                    # New SCHTASK parameters
                    $user = "System"
                    $folder = Split-Path $f
                    $a = New-ScheduledTaskAction -Execute $cmd -Argument $params -WorkingDirectory $folder -CimSession $addr
                    $p = New-ScheduledTaskPrincipal -RunLevel Highest -UserId $user -LogonType S4U

                    # Create and start SCHTASK
                    Write-Host "Register and start SCHTASK - $addr - $cmd" -Fore Green
                    Register-ScheduledTask -TaskName $taskName -Action $a -Principal $p -CimSession $addr -Description "Install SharePoint CU created by SPPatchify Tool" 
                    Start-ScheduledTask -TaskName $taskName -CimSession $addr
                    Write-Host "Start SCHTASK $addr ===== $(Get-Date)" -Fore "Yellow"

                    # Event log START
                    New-EventLog -LogName "Application" -Source "SPPatchify" -ComputerName $addr -ErrorAction SilentlyContinue | Out-Null
                    Write-EventLog -LogName "Application" -Source "SPPatchify" -EntryType Information -Category 1000 -EventId 1000 -Message "START" -ComputerName $addr            
                }
            }

            # WaitEXE Watch EXE binary complete
            WaitEXE $patchName      
        }

        #delete scheduled task
        # Loop - Run Task Scheduler
        foreach ($server in getFarmServers) {       
            $addr = $server.Address
            Write-Host "Unregister task $taskName from - $addr" -Fore Green
            if ($addr -eq $env:computername) {              
                # Remove SCHTASK if found
                $found = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue 
                if ($found) {
                    $found | Unregister-ScheduledTask -Confirm:$false 
                }   
            }
            else {
                # Remove SCHTASK if found
                $found = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue -CimSession $addr
                if ($found) {
                    $found | Unregister-ScheduledTask -Confirm:$false -CimSession $addr
                }
            }
        }
	
        # SharePoint 2016 Force Reboot
        $ver = (Get-SPFarm).BuildVersion.Major
        if ($ver -eq 16) {
            Write-Host "Force Reboot ===== $(Get-Date)" -Fore "Yellow"
            foreach ($server in getRemoteServers) {
                $addr = $server.Address
                if ($addr -ne $env:computername) {
                    Write-Host "Reboot $($addr)" -Fore Yellow
                    Restart-Computer -ComputerName $addr -Force
                }
            }        
            #LocalReboot RunConfigWizard 
            <# 
            $rebootArgs = "-RunConfigWizard"
            $taskName = "SPP_RunPSconfigAfterReboot"
            $cmd = "%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe"
            $params = "-ExecutionPolicy Bypass -File '$root\sppatchify.ps1' $rebootArgs"
            #>

            $taskName = "SPP_RunPSconfigAfterReboot"
            $cmd = "%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe"
            $params = "-ExecutionPolicy Bypass -File ""$root\sppatchify.ps1"" -RunConfigWizard"

            $found = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue 
            if ($found) {
                $found | Unregister-ScheduledTask -Confirm:$false 
            }

            # New SCHTASK parameters
            $user = GetFarmAccount 
            $a = New-ScheduledTaskAction -Execute $cmd -Argument $params 
            $p = New-ScheduledTaskPrincipal -RunLevel Highest -UserId $user -LogonType S4U
            $t = New-ScheduledTaskTrigger -AtStartup -RandomDelay (New-TimeSpan -Minutes 2)
            $task = New-ScheduledTask -Action $a -Principal $p -Trigger $t -Description "Run SPPatchify $rebootArgs after reboot" 
            # Create SCHTASK                
            Write-Host "Register and start SCHTASK - $addr - $cmd" -Fore Green
            $password = (GetFarmAccountPassword)
            Register-ScheduledTask -InputObject  $task -User $user -Password $password -TaskName $taskName           

            Write-Host "Reboot $($env:computername) ===== $(Get-Date)" -Fore Yellow
            Stop-Transcript
            Restart-Computer  -Force
        } 
    }
    else {
        write-host "No Install Files Found. Plaese run .\sppatchify.ps1 -downloadMedia"
        Stop-Transcript
        exit
    }

}

function Sendmail($from = "SharePointPatching@nih.gov", $to = "ContentDeploymentMonitoring@woodbournesolutions.com", $subject = "PS Patchify Notice", $body) {

    $MailMessage = New-Object system.net.mail.mailmessage
    $MailMessage.From = $from
    $MailMessage.To.Add($to)       
    $MailMessage.Subject = $subject 
    $MailMessage.Body = $body
    $MailMessage.IsBodyHtml = $true
    $smtp = New-Object Net.Mail.SmtpClient
    $smtp.Host = "mailfwd.nih.gov"
    $smtp.Port = 25
    $smtp.Send($MailMessage)
}


function RunPSconfig() {
    # not needed. will start with Central Admin server and will run for a while. Should be long enogh for other servers to start.
    #Write-Host " Waiting for Servers to reboot  ===== $(Get-Date)" -Fore "Yellow"
    #WaitReboot
    
    $taskName = "SPP_RunPSconfigAfterReboot"
    Write-Host " Remove Task after reboot  ===== $(Get-Date)" -Fore "Yellow"
   
    $found = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue 
    if ($found) {
        $found | Unregister-ScheduledTask -Confirm:$false 
    }
 

    Write-Host " RunPSconfig ===== $(Get-Date)" -Fore "Yellow"
       
    $taskName = "SPP_RunPSconfig"
    $cmd = "powershell.exe"
    $params = "-ExecutionPolicy Bypass -Command & { Add-PsSnapin Microsoft.SharePoint.PowerShell; & PSConfig.exe -cmd upgrade -inplace b2b -wait -cmd applicationcontent -install -cmd installfeatures -cmd secureresources -cmd services -install }"
    # Loop - Run Task Scheduler
    foreach ($server in getFarmServers) {
        # Local PC - No reboot
        $addr = $server.Address
        Write-Host $addr 
        if ($addr -eq $env:computername) {
            $found = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue 
            if ($found) {
                $found | Unregister-ScheduledTask -Confirm:$false 
            }

            # New SCHTASK parameters
            $user = GetFarmAccount 
            $a = New-ScheduledTaskAction -Execute $cmd -Argument $params 
            $p = New-ScheduledTaskPrincipal -RunLevel Highest -UserId $user -LogonType S4U
            $task = New-ScheduledTask -Action $a -Principal $p -Description "Run PSconfig created by SPPatchify Tool" 
            # Create SCHTASK                
            Write-Host "Register and start SCHTASK - $addr - $cmd" -Fore Green
            $password = (GetFarmAccountPassword)
            Register-ScheduledTask -InputObject  $task  -User $user -Password $password -TaskName $taskName 
            
            # Event log START                
            Start-ScheduledTask -TaskName $taskName 
            Write-Host "Start SCHTASK $addr ===== $(Get-Date)" -Fore "Yellow"
        }    
        else {

            # Remove SCHTASK if found
            $found = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue -CimSession $addr
            if ($found) {
                $found | Unregister-ScheduledTask -Confirm:$false -CimSession $addr
            }

            # New SCHTASK parameters
            $user = GetFarmAccount                
            $a = New-ScheduledTaskAction -Execute $cmd -Argument $params -CimSession $addr
            $p = New-ScheduledTaskPrincipal -RunLevel Highest -UserId $user -LogonType S4U
            $task = New-ScheduledTask -Action $a -Principal $p -Description "Run PSconfig created by SPPatchify Tool"  
            
            # Create SCHTASK
            Write-Host "Register and start SCHTASK - $addr - $cmd" -Fore Green
            $password = (GetFarmAccountPassword)      
            Register-ScheduledTask -InputObject $task -TaskName $taskName -CimSession $addr -User $user -Password $password 

            # Event log START            
            Start-ScheduledTask -TaskName $taskName -CimSession $addr
            Write-Host "Start SCHTASK $addr ===== $(Get-Date)" -Fore "Yellow"
        }
        do {
  
            if ($addr -eq $env:computername) {
                $taskStatus = (Get-ScheduledTask -TaskName $taskName).State -ne 'Ready'
            }
            else {
                $taskStatus = (Get-ScheduledTask -TaskName $taskName -CimSession $addr).State -ne 'Ready'
            }

            Write-Host "." -NoNewLine
            Start-Sleep 600
        }  
        while ($taskStatus)
    }  
    
    #delete scheduled tasks
    $taskName = "SPP_RunPSconfig"
    foreach ($server in getFarmServers) {
        # Local PC - No reboot        
        $addr = $server.Address
        Write-Host "Unregister task $taskName from $addr" -Fore Green
        if ($addr -eq $env:computername) {
            $found = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue 
            if ($found) {
                $found | Unregister-ScheduledTask -Confirm:$false 
            }
        }    
        else {
            # Remove SCHTASK if found
            $found = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue -CimSession $addr
            if ($found) {
                $found | Unregister-ScheduledTask -Confirm:$false -CimSession $addr
            }           
        }
    }
    # restart IIS and app pools on all servers
    IISStart     
}
    
function CreateScheduleTask($addr, $cmd, $params, $taskName, $user, $password, $descirption, $wait = $false) {
    # NotUsed at this time
    # TODO chck if paramser are missing
    if ($addr -eq $env:computername) {
        $found = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue 
        if ($found) {
            $found | Unregister-ScheduledTask -Confirm:$false 
        }

        # New SCHTASK parameters           
        $a = New-ScheduledTaskAction -Execute $cmd -Argument $params 
        $p = New-ScheduledTaskPrincipal -RunLevel Highest -UserId $user -LogonType S4U
        $task = New-ScheduledTask -Action $a -Principal $p -Description $descirption 
            
        # Create SCHTASK                
        Write-Host "Register and start SCHTASK - $addr - $cmd" -Fore Green            
        Register-ScheduledTask -InputObject  $task -User $user -Password $password -TaskName $taskName 
            
        # Start SCHTASK                               
        Start-ScheduledTask -TaskName $taskName 
        Write-Host "Start SCHTASK $addr ===== $(Get-Date)" -Fore "Yellow"
    }    
    else {

        # Remove SCHTASK if found
        $found = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue -CimSession $addr
        if ($found) {
            $found | Unregister-ScheduledTask -Confirm:$false -CimSession $addr
        }

        # New SCHTASK parameters                          
        $a = New-ScheduledTaskAction -Execute $cmd -Argument $params -CimSession $addr
        $p = New-ScheduledTaskPrincipal -RunLevel Highest -UserId $user -LogonType S4U
        $task = New-ScheduledTask -Action $a -Principal $p -Description $descirption 
            
        # Create SCHTASK
        Write-Host "Register and start SCHTASK - $addr - $cmd" -Fore Green                  
        Register-ScheduledTask -InputObject $task -TaskName $taskName -CimSession $addr -User $user -Password $password 
            
        # Start SCHTASK         
        Start-ScheduledTask -TaskName $taskName -CimSession $addr
        Write-Host "Start SCHTASK $addr ===== $(Get-Date)" -Fore "Yellow"
    }
        
    # Wait for scheduled task to complete; return to ready state
    if ($wait) {
        do {
  
            if ($addr -eq $env:computername) {
                $taskStatus = (Get-ScheduledTask -TaskName $taskName).State -ne 'Ready'
            }
            else {
                $taskStatus = (Get-ScheduledTask -TaskName $taskName -CimSession $addr).State -ne 'Ready'
            }

            Write-Host "." -NoNewLine
            Start-Sleep 60
        }  
        while ($taskStatus)
    }
}

function WaitEXE($patchName) {
    Write-Host "===== WaitEXE ===== $(Get-Date)" -Fore "Yellow"
	
    # Wait for EXE intialize
    Write-Host "Wait 60 sec..."
    Start-Sleep 60

    # Watch binary complete
    $counter = 0
    if (getFarmServers) {
        foreach ($server in getFarmServers) {	
            # Progress
            $addr = $server.Address
            $prct = [Math]::Round(($counter / (getFarmServers).Count) * 100)
            if ($prct) {
                Write-Progress -Activity "Wait EXE ($prct % ) $(Get-Date)" -Status $addr -PercentComplete $prct
            }
            $counter++

            # Remote Posh
            $attempt = 0
            Write-Host "`nEXE monitor started on $addr at $(Get-Date) " -NoNewLine
            do {
                # Monitor EXE process
                $proc = Get-Process -Name $patchName -Computer $addr -ErrorAction SilentlyContinue
                Write-Host "." -NoNewLine
                Start-Sleep 60

                # Priority (High) from https://gallery.technet.microsoft.com/scriptcenter/Set-the-process-priority-9826a55f
                $cmd = "`$proc = Get-Process -Name ""$patchName"" -ErrorAction SilentlyContinue; if (`$proc) { if (`$proc.PriorityClass.ToString() -ne ""High"") { `$proc.PriorityClass = [System.Diagnostics.ProcessPriorityClass]::HIGH } }"
                $sb = [Scriptblock]::Create($cmd)        
                InvokeCommand  -ScriptBlock $sb  -server $addr 
                # Measure EXE
                $proc | Select-Object Id, HandleCount, WorkingSet, PrivateMemorySize

                # Count MSPLOG files
                $cmd = "`$f=Get-ChildItem ""$logFolder\*MSPLOG*""; `$c=`$f.count; `$l=(`$f | sort last -desc | select -first 1).LastWriteTime; `$s=`$env:computername; New-Object -TypeName PSObject -Prop (@{""Server"" = `$s; ""Count"" = `$c; ""LastWriteTime"" = `$l })"
                $sb = [Scriptblock]::Create($cmd)
                InvokeCommand  -ScriptBlock $sb  -server $addr  
              
                $progress = "Server: $($result.Server)  /  MSP Count: $($result.Count)  /  Last Write: $($result.LastWriteTime)"
                Write-Progress $progress
            }
            while ($proc)
            Write-Host $progress
			
            # Check Schtask Exit Code
            if ($addr -eq $env:computername) { 
                $task = Get-ScheduledTask -TaskName $taskName 
            }
            else { 
                $task = Get-ScheduledTask -TaskName $taskName -CimSession $addr
            }
           
            $info = $task | Get-ScheduledTaskInfo
            $exit = $info.LastTaskResult
            if ($exit -eq 0) {
                Write-Host "EXIT CODE $exit - $taskName $(Get-Date)" -Fore White -Backgroundcolor Green
            }
            else {
                Write-Host "EXIT CODE $exit - $taskName $(Get-Date)" -Fore White -Backgroundcolor Red
            }
			
            # Event Log
            New-EventLog -LogName "Application" -Source "SPPatchify" -ComputerName $addr -ErrorAction SilentlyContinue | Out-Null
            Write-EventLog -LogName "Application" -Source "SPPatchify" -EntryType Information -Category 1000 -EventId 1000 -Message "DONE - Exit Code $exit" -ComputerName $addr

            # Retry Attempt
            if ($exit -gt 0) {
                # Retry
                $attempt++
                if ($attempt -lt $maxattempt) {
                    # Event log START
                    New-EventLog -LogName "Application" -Source "SPPatchify" -ComputerName $addr -ErrorAction SilentlyContinue | Out-Null
                    Write-EventLog -LogName "Application" -Source "SPPatchify" -EntryType Information -Category 1000 -EventId 1000 -Message "RETRY ATTEMPT # $attempt" -ComputerName $addr

                    # Run
                    Write-Host "RETRY ATTEMPT  # $attempt of $maxattempt" -Fore White -Backgroundcolor Red

                    if ($addr -eq $env:computername) { 
                        Start-ScheduledTask -TaskName $taskName -
                    }
                    else { 
                        Start-ScheduledTask -TaskName $taskName -CimSession $addr
                    }                    
                }
            }
        }
    }
}

function WaitReboot() {
    
    Write-Host "`n===== WaitReboot ===== $(Get-Date)" -Fore "Yellow"
	
    # Wait for farm peer machines to reboot
    Write-Host "Wait 60 sec..."
    Start-Sleep 60
	
    # Clean up
    Get-PSSession | Remove-PSSession -Confirm:$false
	
    # Verify machines online
    $counter = 0
    foreach ($server in getFarmServers) {
        # Progress
        $addr = $server.Address
        Write-Host $addr -Fore Yellow
        if ($addr -ne $env:COMPUTERNAME) {
            $prct = [Math]::Round(($counter / (getFarmServers).Count) * 100)
            if ($prct) {
                Write-Progress -Activity "Waiting for machine ($prct %) $(Get-Date)" -Status $addr -PercentComplete $prct
            }
            $counter++
		
            # Remote PowerShell session
            do {
                # Dynamic open PSSession                
                $remote = GetRemotePSSession $addr  (GetFarmAccountCredentials) 
                # Display
                Write-Host "."  -NoNewLine
                Start-Sleep 5
            }
            while (!$remote)
        }
    }
    Write-Host "`n===== All server online ===== $(Get-Date)" -Fore "Yellow"
	
    # Clean up
    Get-PSSession | Remove-PSSession -Confirm:$false
}

function LocalReboot($parmater) {
    # NotUsed at this time
    # Create Schedued Job
    Get-ScheduledJob -Name "SSPparmater" -ErrorAction SilentlyContinue | Unregister-ScheduledJob -Force
    #To run as highest level add schedulejoboption
    if ($parmater) {
        $ScheduledJobOption = New-ScheduledJobOption -RunElevated
        $JobTrigge = New-JobTrigger -AtStartup -RandomDelay 00:01:00
        $Cred = GetFarmAccountCredentials
        Register-ScheduledJob -ScheduledJobOption $ScheduledJobOption -Trigger $JobTrigge -Credential $Cred  -Name "SSPparmater" -FilePath $root\sppatchify.ps1 -ArgumentList $parmater
    
        # Reboot
        Write-Host "`n ===== REBOOT LOCAL ===== $(Get-Date)"
        $th = [Math]::Round(((Get-Date) - $start).TotalHours, 2)
        Write-Host "Duration Total Hours: $th" -Fore "Yellow"
        Stop-Transcript
        Restart-Computer -Force
        Exit
    }
}
function LaunchPhaseThree() {
    # NotUsed at this time
    # Launch script in new windows for Phase Three - Add Content
    Start-Process "powershell.exe" -ArgumentList "$root\SPPatchify.ps1 -phaseThree"
}

<# function LaunchPhaseThree() {
    Start-Process "powershell.exe" -ArgumentList "$scriptFile -phaseThree"
} #>

function CalcDuration() {
    Write-Host "===== DONE ===== $(Get-Date)" -Fore "Yellow"
    $totalHours = [Math]::Round(((Get-Date) - $start).TotalHours, 2)
    Write-Host "Duration Hours: $totalHours" -Fore "Yellow"
    $c = (Get-SPContentDatabase).Count
    Write-Host "Content Databases Online: $c"
	
    # Add both Phase one and two
    $regHive = "HKCU:\Software"
    $regKey = "SPPatchify"
    if (!$phaseTwo) {
        # Create Regkey
        New-Item -Path $regHive -Name "$regKey" -ErrorAction SilentlyContinue | Out-Null
        New-ItemProperty -Path "$regHive\$regKey" -Name "PhaseOneTotalHours" -Value $totalHours -ErrorAction SilentlyContinue | Out-Null
    }
    else {
        # Read Regkey
        $key = Get-ItemProperty -Path "$regHive\PhaseOneTotalHours" -ErrorAction SilentlyContinue
        if ($key) {
            $totalHours += [double]($key."PhaseOneTotalHours")
        }
        Write-Host "TOTAL Hours (Phase One and Two): $totalHours" -Fore "Yellow"
        Remove-Item -Path "$regHive\$regKey" -ErrorAction SilentlyContinue | Out-Null
    }
}
function FinalCleanUp() {
    # NotUsed at this time
    # Close sessions
    Get-PSSession | Remove-PSSession -Confirm:$false
    Stop-Transcript
}
#endregion

#region SP Config Wizard
function LoopRemotePatch($msg, $cmd, $params) {
    if (!$cmd) {
        return
    }

    # Clean up
    Get-PSSession | Remove-PSSession -Confirm:$false

    # Loop servers
    $counter = 0
    foreach ($server in getFarmServers) {
        
        # Overwrite restart parameter
        $ver = (Get-SPFarm).BuildVersion.Major
        $addr = $server.Address
        if ($ver -eq 16 -or $env:computername -eq $addr) {
            $cmd = $cmd.replace("forcerestart", "norestart")
        }

        # Script block
        if ($cmd.GetType().Name -eq "String") {
            $sb = [ScriptBlock]::Create($cmd)
        }
        else {
            $sb = $cmd
        }
	
        # Progress
        $prct = [Math]::Round(($counter / (getFarmServers).Count) * 100)
        if ($prct) {
            Write-Progress -Activity $msg -Status "$addr ($prct %) $(Get-Date)" -PercentComplete $prct
        }
        $counter++
		
        # Remote Posh
        Write-Host ">> invoke on $addr $(Get-Date)" -Fore "Green"
		
        # Dynamic open PSSession        
        $remote = GetRemotePSSession $addr (GetFarmAccountCredentials)        

        # Invoke
        foreach ($s in $sb) {
            Write-Host $s.ToString()
            if ($remote) {
                Invoke-Command -Session $remote -ScriptBlock $s
            }
        }
        Write-Host "<< complete on $addr $(Get-Date)" -Fore "Green"
    }
    Write-Progress -Activity "Completed $(Get-Date)" -Completed	
    # Clean up
    Get-PSSession | Remove-PSSession -Confirm:$false
}

function GetRemotePSSession([string]$server, [System.Management.Automation.PSCredential]$credentials = [System.Management.Automation.PSCredential]::Empty ) {
    $session = Get-PSSession | Where-Object { $_.ComputerName -eq $server }
    if (!$session) {   
        if ($env:computername -eq $server) {
            $session = New-PSSession -Credential $credentials
        }
        else {     
            if ($remoteSessionPort -and $remoteSessionSSL) {
                if ($credentials -eq [System.Management.Automation.PSCredential]::Empty) {
                    $session = New-PSSession -ComputerName $server -Port $remoteSessionPort -UseSSL
                }
                else {
                    $session = New-PSSession -ComputerName $server -Credential $credentials -Authentication Kerberos  -Port $remoteSessionPort -UseSSL 
                }

            }
            elseif ($remoteSessionPort) {
                if ($credentials -eq [System.Management.Automation.PSCredential]::Empty) {
                    $session = New-PSSession -ComputerName $server -Port $remoteSessionPort 
                }
                else {
                    $session = New-PSSession -ComputerName $server -Credential $credentials -Authentication Kerberos  -Port $remoteSessionPort 
                }
            }
            elseif ($remoteSessionSSL) {
      
                if ($credentials -eq [System.Management.Automation.PSCredential]::Empty) {
                    $session = New-PSSession -ComputerName $server  -UseSSL
                }
                else {
                    $session = New-PSSession -ComputerName $server -Credential $credentials -Authentication Kerberos  -UseSSL 
                }        
            }
            else {
                if ($credentials -eq [System.Management.Automation.PSCredential]::Empty) {
                    $session = New-PSSession -ComputerName $server 
                }
                else {
                    $session = New-PSSession -ComputerName $server -Credential $credentials -Authentication Kerberos  
                }
            }
        }
    }
    return $session
}

function LoopRemoteCmd($msg, $cmd, $isJob = $false) {
    if (!$cmd) {
        return
    }

    # Clean up
    Get-PSSession | Remove-PSSession -Confirm:$false
	
    # Loop servers
    $counter = 0
    foreach ($server in getFarmServers) {
        Write-Host $server.Address -Fore Yellow

        # Script block
        if ($cmd.GetType().Name -eq "String") {
            $sb = [ScriptBlock]::Create($cmd)
        }
        else {
            $sb = $cmd
        }
	
        # Progress
        $addr = $server.Address
        $prct = [Math]::Round(($counter / (getFarmServers).Count) * 100)
        if ($prct) {
            Write-Progress -Activity $msg -Status "$addr ($prct %) $(Get-Date)" -PercentComplete $prct
        }
        $counter++        

        # Merge script block array
        $mergeSb = $sb
        $mergeCmd = ""
        if ($sb -is [array]) {
            foreach ($s in $sb) {
                $mergeCmd += $s.ToString() + "`n"
            }
            $mergeSb = [Scriptblock]::Create($mergeCmd)
        }

        # Remote Posh
        Write-Host ">> invoke on $addr $(Get-Date)" -Fore "Green"
        Write-Host "mergeSb $mergeSb" -Fore "Cyan"
        InvokeCommand -ScriptBlock $mergeSb -server $addr -isJob $isJob    
        Write-Host "<< complete on $addr $(Get-Date)" -Fore "Green"
    }
    Write-Progress -Activity "Completed $(Get-Date)" -Completed
}

function InvokeCommand($server, $ScriptBlock, $isJob = $false) {
    # InvokeCommand -server -command -isJob
    # if local server
    ## invoke command
    # if remote server
    ## get sesscion
    ## invoke command

    # Script block
    # write-host "sb: $ScriptBlock"
    # write-host "sb.GetType().Name: $($ScriptBlock.GetType().Name)"

    if ([string]::IsNullOrEmpty($ScriptBlock)) {
        write-host "Scriptblock is empty"
        return
    }
    if ($ScriptBlock.GetType().Name -eq "String") {
        $ScriptBlock = [ScriptBlock]::Create($ScriptBlock)
    }
    #write-host "sb.GetType().Name: $($ScriptBlock.GetType().Name)"
    # write-host "ScriptBlock: $ScriptBlock"
    $session = GetRemotePSSession $server (GetFarmAccountCredentials)
    if ($env:computername -eq $server -or $server -eq "localhost") {

        if ($isJob) {
            Start-Job -ScriptBlock $ScriptBlock -Credential (GetFarmAccountCredentials)
        }
        else {
            #Start-Job -ScriptBlock $ScriptBlock -Credential (GetFarmAccountCredentials)
            Invoke-Command -Session $Session -ScriptBlock $ScriptBlock 
            #Start-Process -FilePath Powershell -Credential (GetFarmAccountCredentials) -Wait -ArgumentList '-Command', $ScriptBlock 
        }
    }
    else {
        
        if ($session) {
            if ($isJob) {
                Invoke-Command  -ScriptBlock $ScriptBlock -AsJob -Session $session
            }
            else {
                Invoke-Command  -ScriptBlock $ScriptBlock -Session $session
            }
        }
        else {
            Write-Host "could not invoke, no remote session"
        }
    }
}



function StopSPDistributedCache() {
    Write-Host "===== StopSPDistributedCache OFF ===== $(Get-Date)" -Fore "Yellow"

    # Distributed Cache
    $sb = {
        try {
            Use-CacheCluster
            Get-AFCacheClusterHealth -ErrorAction SilentlyContinue
            $computer = [System.Net.Dns]::GetHostByName($env:computername).HostName
            $counter = 0
            $maxLoops = 60

            $cache = Get-CacheHost | Where-Object { $_.HostName -eq $computer }
            if ($cache) {
                do {
                    try {
                        # Wait for graceful stop
                        $hostInfo = Stop-CacheHost -Graceful -CachePort 22233 -HostName $computer -ErrorAction SilentlyContinue
                        Write-Host $computer $hostInfo.Status
                        Start-Sleep 5
                        $counter++
                    }
                    catch {
                        break
                    }
                }
                while ($hostInfo -and $hostInfo.Status -ne "Down" -and $counter -lt $maxLoops)

                # Force stop
                Add-PSSnapIn Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue | Out-Null
                Stop-SPDistributedCacheServiceInstance
            }
        }
        catch {
        }
    }
    LoopRemoteCmd "Stop Distributed Cache on " $sb
}

function ChangeServices($state) {
    Write-Host "===== ChangeServices $state ===== $(Get-Date)" -Fore "Yellow"
    $ver = (Get-SPFarm).BuildVersion.Major

    # Logic core
    if ($state) {
        $action = "START"
        $sb = {
            @("IISADMIN", "W3SVC", "SPAdminV4", "SPTimerV4", "SQLBrowser", "Schedule", "SPInsights", "DocAve 6 Agent Service") | ForEach-Object {
                if (Get-Service $_ -ErrorAction SilentlyContinue) {
                    Set-Service -Name $_ -StartupType Automatic -ErrorAction SilentlyContinue
                    Start-Service $_ -ErrorAction SilentlyContinue
                }
            }
            @("OSearch$ver", "SPSearchHostController") | ForEach-Object {
                Start-Service $_ -ErrorAction SilentlyContinue
            }
            Start-Process 'iisreset.exe' -ArgumentList '/start' -Wait -PassThru -NoNewWindow | Out-Null
        }
    }
    else {
        $action = "STOP"
        $sb = {
            Start-Process 'iisreset.exe' -ArgumentList '/stop' -Wait -PassThru -NoNewWindow | Out-Null
            @("IISADMIN", "W3SVC", "SPAdminV4", "SPTimerV4", "SQLBrowser", "Schedule", "SPInsights", "DocAve 6 Agent Service") | ForEach-Object {
                if (Get-Service $_ -ErrorAction SilentlyContinue) {
                    Set-Service -Name $_ -StartupType Disabled -ErrorAction SilentlyContinue
                    Stop-Service $_ -ErrorAction SilentlyContinue
                }
            }
            @("OSearch$ver", "SPSearchHostController") | ForEach-Object {
                Stop-Service $_ -ErrorAction SilentlyContinue
            }
        }
    }

    # Search Crawler
    Write-Host "$action search crawler ..."
    try {
        $ssa = Get-SPEenterpriseSearchServiceApplication 
        if ($state) {
            $ssa.resume()
        }
        else {
            $ssa.pause()
        }
    }
    catch {
    }

    LoopRemoteCmd "$action services on " $sb
}

function PauseSharePointSearch() {

    Write-Host "Start pausing search crawler ... ===== $(Get-Date)" -Fore "Yellow" 
    $ssa = Get-SPEnterpriseSearchServiceApplication  
    $ssa.pause()   
    Write-Host "search crawler paused... ===== $(Get-Date)" -Fore "Yellow"  
}

function StartSharePointSearch() {
    Write-Host "Starting search crawler ... ===== $(Get-Date)" -Fore "Yellow"   
    $ssa = Get-SPEnterpriseSearchServiceApplication         
    $ssa.resume()    
    Write-Host "Started search crawler ... ===== $(Get-Date)" -Fore "Yellow"     
}

function DismountContentDatabase($needUpgradeOnly = $False) {
    #ChangeContent $false
 
    Write-Host "===== ContentDB $state ===== $(Get-Date)" -Fore "Yellow"
    # Display
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint") | Out-Null
    $dbs = Get-SPContentDatabase
    if ($needUpgradeOnly) {
        $dbs = $dbs | where ($_.NeedsUpgrade)
    }
    $c = $dbs.Count
    Write-Host "Content Databases Online: $c"
    $sb = @()  
    # Remove content database
       
    if ($dbs) {
        $dbs | ForEach-Object { $wa = $_.WebApplication.Url; $_ | Select-Object Name, NormalizedDataSource, @{n = "WebApp"; e = { $wa } } } | Export-Csv "$logFolder\contentdbs-$when.csv" -NoTypeInformation
        $dbs | ForEach-Object {
            "$($_.Name),$($_.NormalizedDataSource)"
            Dismount-SPContentDatabase $_ -Confirm:$false
        }
    }
   
}

function MountContentDatabase() {
    # ChangeContent $true
    # create script block based on saved content database
    $files = Get-ChildItem "$logFolder\contentdbs-*.csv" | Sort-Object LastAccessTime -Desc
    if ($files -is [Array]) {
        $files = $files[0]
    }
    $sb = @()
    # Loop databases and create script block
    if ($files) {
        Write-Host "Content DB - from CSV $($files.Fullname)" -Fore Yellow
        $dbs = Import-Csv $files.Fullname  
        Write-Host "Content DB - create script blocks" -Fore Yellow      
        foreach ($db in $dbs) {  
            $wa = [Microsoft.SharePoint.Administration.SPWebApplication]::Lookup($db.WebApp)
            if ($wa) {   
                                           
                $sb2 = '
                        Add-PSSnapIn Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue | Out-Null   
                        $db = Get-SPContentDatabase | Where-Object {$_.Name -eq "' + $($db.name) + '"}  
                        If(-not $db) {Mount-SPContentDatabase -AssignNewDatabaseId -WebApplication "' + $($wa.url) + '" -Name "' + $($db.name) + '" -DatabaseServer "' + $($db.NormalizedDataSource) + '" | Out-Null}
                        $NeedsUpgrade = (Get-SPContentDatabase | Where-Object {$_.Name -eq "' + $($db.name) + '"}).NeedsUpgrade
                        if($NeedsUpgrade){Upgrade-SPContentDatabase -Name "' + $($db.name) + '" -WebApplication "' + $($wa.url) + '" -ErrorAction SilentlyContinue -Confirm:$false | Out-Null}
                '
                $sb += $sb2           
            }
        }
        $serverAddress = getFarmServers | ForEach-Object { $_.Address }
        DistributedJobs -scriptBlocks $sb -servers $serverAddress -credentials (GetFarmAccountCredentials)
            
    }
    else {
        Write-Host "Content DB - CSV not found" -Fore Yellow
    }
}


function ReportContentDatabases() {
    $dbs = Get-SPContentDatabase
    if ($dbs) {
        $dbs | ForEach-Object { $wa = $_.WebApplication.Url; $_ | Select-Object Name, NormalizedDataSource, @{n = "WebApp"; e = { $wa } } } | Export-Csv "$logFolder\contentdbs-$when.csv" -NoTypeInformation
    }
}


function DistributedJobs2($scriptBlocks, [string[]]$servers, [int]$maxJobs = 1, [System.Management.Automation.PSCredential]$credentials = [System.Management.Automation.PSCredential]::Empty) {
    
    if (!$servers -or !$scriptBlocks) {
        return
    }   
    Get-Job | Remove-Job 
    Get-PSSession | Remove-PSSession

    $remoteServers = $servers | Where-Object { $_ -ne $env:computername }
    $remoteServers += "localhost"
    $servers = $remoteServers
    $servers
    
    $counter = 0
    foreach ($scriptBlock in $scriptBlocks) {
        $avaialableServer = $null
        $activeJobs = $null
        $wait = $true

        $ActiveJobs = @(Get-Job | Where-Object { $_.State -eq "Running" }) -or $_.State -eq "NotStarted"
        
        foreach ($server in $servers) {
            $server
            ($activeJobs | Where-Object { $_.Location -eq $server }).Count
            # if a server has less than maxJobs running, then do not wait
            if ($maxJobs -gt ($activeJobs | Where-Object { $_.Location -eq $server }).Count) { 
                $wait = $False 
                $avaialableServer = $server 
                break
            }
        }

        # wait while servers have maxJobs running
        if ($wait) {
            $activeJobs | Wait-Job -Any | Out-Null
        }  

        if (!$avaialableServer) {
            $activeJobs
            Write-Host "---"
            $server 
        }

        Write-Host "Starting job for $avaialableServer"
        InvokeCommand -server $avaialableServer -ScriptBlock $scriptBlock -isJob $true                
        
        # Progress
        $prct = [Math]::Round(($counter / $scriptBlocks.Count) * 100)
        if ($prct) {
            Write-Progress -Activity "Jobs" -Status "($prct %) $(Get-Date)" -PercentComplete $prct
        }
        $counter++
    }

    # Wait for all jobs to complete and results ready to be received
    Wait-Job * | Out-Null

    
    # Process the results
    foreach ($job in Get-Job -IncludeChildJob) {
        $result = Receive-Job $job 
        Write-Host $result
    }
    

    Get-Job | Remove-Job 
    Get-PSSession | Remove-PSSession
     
}


function ChangeContent($state) {
    Write-Host "===== ContentDB $state ===== $(Get-Date)" -Fore "Yellow"
    # Display
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint") | Out-Null
    $c = (Get-SPContentDatabase).Count
    Write-Host "Content Databases Online: $c"
    $sb = @()
    if (!$state) {
        # Remove content database
        $dbs = Get-SPContentDatabase
        if ($dbs) {
            $dbs | ForEach-Object { $wa = $_.WebApplication.Url; $_ | Select-Object Name, NormalizedDataSource, @{n = "WebApp"; e = { $wa } } } | Export-Csv "$logFolder\contentdbs-$when.csv" -NoTypeInformation
            $dbs | ForEach-Object {
                "$($_.Name),$($_.NormalizedDataSource)"
                Dismount-SPContentDatabase $_ -Confirm:$false
            }
        }
    }
    else {
        # create script block based on saved content database
        $files = Get-ChildItem "$logFolder\contentdbs-*.csv" | Sort-Object LastAccessTime -Desc
        if ($files -is [Array]) {
            $files = $files[0]
        }

        # Loop databases and create script block
        if ($files) {
            Write-Host "Content DB - from CSV $($files.Fullname)" -Fore Yellow
            $dbs = Import-Csv $files.Fullname  
            Write-Host "Content DB - create script blocks" -Fore Yellow      
            foreach ($db in $dbs) {  
                $wa = [Microsoft.SharePoint.Administration.SPWebApplication]::Lookup($db.WebApp)
                if ($wa) {   
                                           
                    $sb2 = '
                        Add-PSSnapIn Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue | Out-Null   

                        $db = Get-SPContentDatabase | Where-Object {$_.Name -eq "' + $($db.name) + '"}  

                        If(-not $db) {Mount-SPContentDatabase -AssignNewDatabaseId -WebApplication "' + $($wa.url) + '" -Name "' + $($db.name) + '" -DatabaseServer ' + $($db.NormalizedDataSource) + ' | Out-Null}
                        $NeedsUpgrade = (Get-SPContentDatabase | Where-Object {$_.Name -eq "' + $($db.name) + '"}).NeedsUpgrade
                        if($NeedsUpgrade){Upgrade-SPContentDatabase -Name "' + $($db.name) + '" -WebApplication "' + $($wa.url) + '" -ErrorAction SilentlyContinue -Confirm:$false | Out-Null}
                        '
                    $sb += $sb2           
                }
            }
            $serverAddress = getFarmServers | ForEach-Object { $_.Address }
            DistributedJobs -scriptBlocks $sb -servers $serverAddress -credentials (GetFarmAccountCredentials)
            
        }
        else {
            Write-Host "Content DB - CSV not found" -Fore Yellow
        }
    }
}
#endregion

#region general
function getRemoteServers() {
    return getFarmServers | Where-Object { $_.Address -ne $env:computername }
}

function getFarmServers() {
    return Get-SPServer | Where-Object { $_.Role -ne "Invalid" } | Sort-Object Address
}

function EnablePSRemoting() {
    $ssp = Get-WSManCredSSP
    if ($ssp[0] -match "not configured to allow delegating") {
        # Enable remote PowerShell over CredSSP authentication
        Enable-WSManCredSSP -DelegateComputer * -Role Client -Force
        Restart-Service WinRM
    }
    $remoteServers = getRemoteServers
    foreach ($remoteServer in $remoteServers ) {
        invoke-command -computername $remoteServer.Address -scriptblock {
            $isCredSSPServer = ((Get-Item WSMan:\LocalHost\Service\Auth\CredSSP).Value -eq "true")
            if (-not $isCredSSPServer) {
                Enable-WsManCredSSP -Role Server -Force
                Restart-Service WinRM
            }
        }
    }

}

function GetFarmAccount() {
    return (Get-SPFarm).DefaultServiceAccount.Name
}
function GetFarmAccountCredentials() {
    $farmAccount = GetFarmAccount
    $farmAccountPassword = GetFarmAccountPassword

    if ($farmAccount -and $farmAccountPassword ) {
        $securePassword = $farmAccountPassword | ConvertTo-SecureString -AsPlainText -Force
        $PSCredential = New-Object System.Management.Automation.PSCredential -ArgumentList $farmAccount, $securePassword
    }    
    
    return $PSCredential      
}

Function GetFarmAccountPassword() {    
    $farmAccount = GetFarmAccount
    import-Module WebAdministration
    $password = (Get-ChildItem IIS:\AppPools | Where-Object { $_.processModel.userName -eq $farmAccount })[0].processModel.password 
    return $password
}

function ReadIISPW {
    #SecurityTokenServiceApplicationPool SharePoint Web Services System
    Write-Host "===== Read IIS PW ===== $(Get-Date)" -Fore "Yellow"

    # Current user (ex: Farm Account)
    $domain = $env:userdomain
    $user = $env:username
    Write-Host "Logged in as $domain\$user"
	
    # Start IISAdm` if needed
    $iisadmin = Get-Service IISADMIN -ErrorAction SilentlyContinue | Out-Null
    if ($iisadmin.Status -ne "Running") {
        # Set Automatic and Start
        Set-Service -Name IISADMIN -StartupType Automatic -ErrorAction SilentlyContinue
        Start-Service IISADMIN -ErrorAction SilentlyContinue
    }
	
    # Attempt to detect password from IIS Pool (if current user is local admin and farm account)
    Import-Module WebAdministration -ErrorAction SilentlyContinue | Out-Null
    $m = Get-Module WebAdministration
    if ($m) {
        # PowerShell ver 2.0+ IIS technique
        $appPools = Get-ChildItem "IIS:\AppPools\"
        foreach ($pool in $appPools) {	
            if ($pool.processModel.userName -like "*$user") {
                Write-Host "Found - "$pool.processModel.userName
                $pass = $pool.processModel.password
                if ($pass) {
                    break
                }
            }
        }
    }
    else {
        # PowerShell ver 3.0+ WMI technique
        $appPools = Get-CimInstance -Namespace "root/MicrosoftIISv2" -ClassName "IIsApplicationPoolSetting" -Property Name, WAMUserName, WAMUserPass | Select-Object WAMUserName, WAMUserPass
        foreach ($pool in $appPools) {	
            if ($pool.WAMUserName -like "*$user") {
                Write-Host "Found - "$pool.WAMUserName
                $pass = $pool.WAMUserPass
                if ($pass) {
                    break
                }
            }
        }
    }

    # Prompt for password
    if (!$pass) {
        $sec = Read-Host "Enter password " -AsSecureString
    }
    else {
        $sec = $pass | ConvertTo-SecureString -AsPlainText -Force
    }

    # Save global
    $global:cred = New-Object System.Management.Automation.PSCredential -ArgumentList "$domain\$user", $sec
}

function DisplayCA() {
    # Version DLL File
    $sb = {
        ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        Add-PSSnapIn Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue | Out-Null;
        $ver = (Get-SPFarm).BuildVersion.Major;
        [System.Diagnostics.FileVersionInfo]::GetVersionInfo("C:\Program Files\Common Files\microsoft shared\Web Server Extensions\$ver\ISAPI\Microsoft.SharePoint.dll") | Select-Object FileVersion, @{N = 'PC'; E = { $env:computername } }
    }
    LoopRemoteCmd "Get file version on " $sb
	
    # Display Version
    ShowVersion
	
    # Open Central Admin
    $ca = (Get-SPWebApplication -IncludeCentralAdministration) | Where-Object { $_.IsAdministrationWebApplication -eq $true }
    $pages = @("PatchStatus.aspx", "UpgradeStatus.aspx", "FarmServers.aspx")
    $pages | ForEach-Object { Start-Process ($ca.Url + "_admin/" + $_) }
}
function ShowVersion() {
    # Version Max Patch
    $maxv = 0
    $f = Get-SPFarm
    $p = Get-SPProduct
    foreach ($u in $p.PatchableUnitDisplayNames) {
        $n = $u
        $v = ($p.GetPatchableUnitInfoByDisplayName($n).patches | Sort-Object version -desc)[0].version
        if (!$maxv) {
            $maxv = $v
        }
        if ($v -gt $maxv) {
            $maxv = $v
        }
    }

    # Control Panel Add/Remove Programs
    
	
    # IIS UP/DOWN Load Balancer
    Write-Host "IIS UP/DOWN Load Balancer"
    $coll = @()
    getFarmServers | ForEach-Object {
        try {
            $addr = $_.Address;
            $root = (Get-Website "Default Web Site").PhysicalPath.ToLower().Replace("%systemdrive%", $env:SystemDrive)
            $remoteRoot = "\\$addr\"
            $remoteRoot += MakeRemote $root
            $status = (Get-Content "$remoteRoot\status.html" -ErrorAction SilentlyContinue)[1];
            $coll += @{"Server" = $addr; "Status" = $status }
        }
        catch {
            # Suppress any error
        }
    }
    $coll | Format-Table -AutoSize

    # Database table
    $d = Get-SPWebapplication -IncludeCentralAdministration | Get-SPContentDatabase 
    $d | Sort-Object NeedsUpgrade, Name | Select-Object NeedsUpgrade, Name | Format-Table -AutoSize

    # Database summary
    $d | Group-Object NeedsUpgrade | Format-Table -AutoSize
    "---"
	
    # Server status table
    (Get-SPProduct).Servers | Select-Object Servername, InstallStatus -Unique | Group-Object InstallStatus, Servername | Sort-Object Name | Format-Table -AutoSize
	
    # Server status summary
    (Get-SPProduct).Servers | Select-Object Servername, InstallStatus -Unique | Group-Object InstallStatus | Sort-Object Name | Format-Table -AutoSize

    # Display data
    if ($maxv -eq $f.BuildVersion) {
        Write-Host "Max Product = $maxv" -Fore Green
        Write-Host "Farm Build  = $($f.BuildVersion)" -Fore Green
    }
    else {
        Write-Host "Max Product = $maxv" -Fore Yellow
        Write-Host "Farm Build  = $($f.BuildVersion)" -Fore Yellow
    }
}
function IISStart() {
    # Start IIS pools and sites
    $sb = {
        Import-Module WebAdministration

        # IISAdmin
        $iisadmin = Get-Service "IISADMIN"
        if ($iisadmin) {
            Set-Service -Name $iisadmin -StartupType Automatic -ErrorAction SilentlyContinue
            Start-Service $iisadmin -ErrorAction SilentlyContinue
        }

        # W3WP
        Start-Service w3svc | Out-Null
        Get-ChildItem "IIS:\AppPools\" | ForEach-Object { $n = $_.Name; Start-WebAppPool $n | Out-Null }
        Get-WebSite | Start-WebSite | Out-Null
    }
    LoopRemoteCmd "Start IIS on " $sb
}

function ProductLocal() {
    # Sync local SKU binary to config DB
    $sb = {
        Add-PSSnapIn Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue | Out-Null
        Get-SPProduct -Local
    }
    LoopRemoteCmd "Product local SKU on " $sb
	
    # Display server upgrade
    Write-Host "Farm Servers - Upgrade Status " -Fore "Yellow"
    (Get-SPProduct).Servers | Select-Object Servername, InstallStatus | Sort-Object Servername | Format-Table -AutoSize
}

function UpgradeContent() {
    Write-Host "===== Upgrade Content Databases ===== $(Get-Date)" -Fore "Yellow"
	
    # Tracking table - assign DB to server
    $maxWorkmaxWorkersers = 4
    $track = @()
    $dbs = Get-SPContentDatabase
    $i = 0
    foreach ($db in $dbs) {
        # Assign to SPServer
        $mod = $i % (getFarmServers).count
        $pc = getFarmServers[$mod].Address
		
        # Collect
        $obj = New-Object -TypeName PSObject -Prop (@{"Name" = $db.Name; "Id" = $db.Id; "UpgradePC" = $pc; "JID" = 0; "Status" = "New" })
        $track += $obj
        $i++
    }
    $track | Format-Table -Auto
	

    # Clean up
    Get-PSSession | Remove-PSSession -Confirm:$false
    Get-Job | Remove-Job
	
    # Open sessions ?
    foreach ($server in getRemoteServers) {
        $addr = $server.Address
        
        # Dynamic open PSSesion
        GetRemotePSSession $addr (GetFarmAccountCredentials) 
    }

    # Monitor and Run loop
    do {
        # Get latest PID status
        $active = $track | Where-Object { $_.Status -eq "InProgress" }
        foreach ($db in $active) {
            # Monitor remote server job
            if ($db.JID) {
                $job = Get-Job $db.JID
                if ($job.State -eq "Completed") {
                    # Update DB tracking
                    $db.Status = "Completed"
                }
                elseif ($job.State -eq "Failed") {
                    # Update DB tracking
                    $db.Status = "Failed"
                }
                else {
                    Write-host "-" -NoNewline
                }
            }
        }
		
        # Ensure workers are active
        foreach ($server in getFarmServers) {
            # Count active workers per server
            $active = $track | Where-Object { $_.Status -eq "InProgress" -and $_.UpgradePC -eq $server.Address }
            if ($active.count -lt $maxWorkers) {
			
                # Choose next available DB
                $avail = $track | Where-Object { $_.Status -eq "New" -and $_.UpgradePC -eq $server.Address }
                if ($avail) {
                    if ($avail -is [array]) {
                        $row = $avail[0]
                    }
                    else {
                        $row = $avail
                    }
				
                    # Kick off new worker
                    $id = $row.Id
                    $name = $row.Name
                    $remoteStr = "`$cmd = New-Object System.Diagnostics.ProcessStartInfo; " + 
                    "`$cmd.FileName = 'powershell.exe'; " + 
                    "`$internal = Add-PSSnapin Microsoft.SharePoint.Powershell -ErrorAction SilentlyContinue | Out-Null; Upgrade-SPContentDatabase -Id $id -Confirm:`$false; " + 
                    "`$cmd.Arguments = '-NoProfile -Command ""$internal""'; " + 
                    "[System.Diagnostics.Process]::Start(`$cmd);"
					
                    # Run on remote server
                    $remoteCmd = [Scriptblock]::Create($remoteStr) 
                    $pc = $server.Address
                    Write-Host $pc -Fore "Green"
                    Get-PSSession | Format-Table -AutoSize
                    $session = Get-PSSession | Where-Object { $_.ComputerName -like "$pc*" }
                    if (!$session) {
                        # Dynamic open PSSession
                        $session = GetRemotePSSession $addr (GetFarmAccountCredentials)                       
                    }
                    $result = Invoke-Command $remoteCmd -Session $session -AsJob
					
                    # Update DB tracking
                    $row.JID = $result.Id
                    $row.Status = "InProgress"
                }
				
                # Progress
                $counter = ($track | Where-Object { $_.Status -eq "Completed" }).Count
                $prct = 0
                if ($track) {
                    $prct = [Math]::Round(($counter / $track.Count) * 100)
                }
                if ($prct) {
                    Write-Progress -Activity "Upgrade database" -Status "$name ($prct %) $(Get-Date)" -PercentComplete $prct
                }
                $track | Format-Table -AutoSize
            }
        }

        # Latest counter
        $remain = $track | Where-Object { $_.status -ne "Completed" -and $_.status -ne "Failed" }
    }
    while ($remain)
    Write-Host "===== Upgrade Content Databases DONE ===== $(Get-Date)"
    $track | Group-Object status | Format-Table -AutoSize
    $track | Format-Table -AutoSize
	
    # GUI
    $msg = "Upgrade Content DB Complete (100 %)"
	
    # Clean up
    Get-PSSession | Remove-PSSession -Confirm:$false
    Get-Job | Remove-Job
}

function ShowMenu($prod) {
    # Choices
    $csv = Import-Csv "$root\SPPatchify-Download-CU.csv" | Select-Object -Property @{n = 'MonthInt'; e = { [int]$_.Month } }, *
    $choices = $csv | Where-Object { $_.Product -eq $prod } | Sort-Object Year, MonthInt -Desc | Select-Object Year, Month -Unique

    # Menu
    Write-Host "Download CU Media to \media\ - $prod" -Fore "Yellow"
    Write-Host "---------"
    $menu = @()
    $i = 0
    $choices | ForEach-Object {
        $n = (getMonth($_.Month)) + " " + ($_.Year)
        $menu += $n
        if ($i -eq 0) {
            $default = $n
            $n += "[default] <=="
            Write-Host "$i $n" -Fore "Green"
        }
        else {
            Write-Host "$i $n"
        }
        $i++
    }

    # Return
    $sel = Read-Host "Select month. Press [enter] for default"
    if (!$sel) {
        $sel = $default
    }
    else {
        $sel = $menu[$sel]
    }
    $global:selmonth = $sel
} 

function GetMonth($mo) {
    # Convert integer to three letter month name
    try {
        $mo = (Get-Culture).DateTimeFormat.GetAbbreviatedMonthName($mo)
    }
    catch {
        return $mo
    }
    return $mo
}

function GetMonthInt($name) {
    # Convert three letter month name to integer
    $found = $false
    1 .. 12 | ForEach-Object {
        if ($name -eq (Get-Culture).DateTimeFormat.GetAbbreviatedMonthName($_)) {
            $found = $true
            return $_
        }
    }
    if (!$found) { return $name }
}
function PatchRemoval() {
    # Remove patch media
    $files = Get-ChildItem "$root\media\*" -Recurse -ErrorAction SilentlyContinue #| Out-Null
    $files | Format-Table -AutoSize
    $files | Remove-Item -Confirm:$false -Force
}
function PatchMenu() {
    # Ensure folder
    mkdir "$root\media" -ErrorAction SilentlyContinue | Out-Null
    PatchRemoval
    
    If ($downloadVersion) {
        $sharePointVersion = $downloadVersion
    }
    else {
        $spYear = $null
        if (Test-Path "C:\Program Files\Common Files\Shared Tools\Web Server Extensions") {
            $spYear = (Get-SPFarm).BuildVersion
        }
       
        $spAvailableVersionNumbers = @{
            19 = "2019"
            16 = "2016"
            15 = "2013"
        }

        if ($spYear.Major -eq 15) {
            $spYear = 15
        }
        elseif ($spYear.Major -eq 16  ) {
            if ($spYear.build -ge 10337) {
                $spYear = 19
            }
            else {
                $spYear = 16
            }
        }
        while ([string]::IsNullOrEmpty($spYear)) {

            $spYear = ($spAvailableVersionNumbers | Out-GridView -Title "Please select the version of SharePoint to download updates for:" -PassThru).Name
            if ($spYear.Count -gt 1) {
                Write-Warning "Please only select ONE version. Re-prompting..."
                Remove-Variable -Name spYear -Force -ErrorAction SilentlyContinue
            }
        }
        $sharePointVersion = $spAvailableVersionNumbers[$spYear]     
    }
    
    Write-Host " - SharePoint $sharePointVersion selected."
    AutoSPSourceBuilder -UpdateLocation "$root\media" -SharePointVersion $sharePointVersion -Destination "$root\media"
    #$Destination 
    #Get-ChildItem -Path $Destination -Recurse -File $Destination | Copy-Item -Destination $root\media 
    #Get-ChildItem -Path $root\media -Recurse -Directory | Remove-Item 
    <#

    # Skip if we already have media
    $files = Get-ChildItem "$root\media\*.exe"
    if ($files) {
        Write-Host "Using EXE found in \media\.`nTo trigger download GUI first delete \media\ folder and run script again."		
        $files | Format-Table -Auto
        Return
    }

    # Download CSV of patch URLs
    $source = "https://raw.githubusercontent.com/spjeff/sppatchify/master/SPPatchify-Download-CU.csv"
    $local = "$root\SPPatchify-Download-CU.csv"
    $wc = New-Object System.Net.Webclient
    $dest = $local.Replace(".csv", "-temp.csv")
    $wc.DownloadFile($source, $dest)
	
    # Overwrite if downloaded OK
    if (Test-Path $dest) {
        Copy-Item $dest $local -Force
        Remove-Item $dest
    }
    $csv = Import-Csv $local
	
    # SKU - SharePoint or Project?
    $sku = "SP"
    $ver = "15"
    if ($downloadVersion) {
        $ver = $downloadVersion
    }
    if (Get-Command Get-SPFarm -ErrorAction SilentlyContinue) {
        # Local farm
        $farm = Get-SPFarm -ErrorAction SilentlyContinue
        if ($farm) {
            $ver = $farm.BuildVersion.Major
            $sppl = (Get-SPProduct -Local) | Where-Object { $_.ProductName -like "*Microsoft Project*" }
            if ($sppl) {
                if ($ver -ne 16) {
                    $sku = "PROJ"
                }
            }
        }
        else {
            # Detect binary folder - fallback if not joined to farm
            $detect16 = Get-ChildItem "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16"
            if ($detect16) {
                $ver = "16"
            }
        }
    }

    # Product and menu
    $prod = "$sku$ver"
    Write-Host "Product = $prod"
    ShowMenu $prod
	
    # Filter CSV for selected CU month
    Write-Host "SELECTED = $($global:selmonth)" -Fore "Yellow"
    $year = $global:selmonth.Split(" ")[1]
    $month = GetMonthInt $global:selmonth.Split(" ")[0]
    Write-Host "$year-$month-$sku$ver"
    $patchFiles = $csv | Where-Object { $_.Year -eq $year -and $_.Month -eq $month -and $_.Product -eq "$sku$ver" }
    $patchFiles | Format-Table -Auto
	
    # Download patch files
    $bits = (Get-Command Start-BitsTransfer -ErrorAction SilentlyContinue)
    foreach ($file in $patchFiles) {
        # Parameters
        $splits = $file.URL.Split("/")
        $name = $splits[$splits.Count - 1]
        $dest = "$root\media\$name"

        # Download file if missing
        if (Test-Path $dest) {
            Write-Host "Found $name"
        }
        else {
            Write-Host "Downloading $name"
            if ($bits) {
                # pefer BITS
                Write-Host "BITS $dest"
                Start-BitsTransfer -Source $file.URL -Destination $dest
            }
            else {
                # Dot Net
                Write-Host "WebClient $dest"
                (New-Object System.Net.WebClient).DownloadFile($file.URL, $dest)
            }
        }
    }

    # Halt if Farm is PROJ and media is not
    $files = Get-ChildItem "$root\media\*prj*.exe"
    if ($sku -eq "PROJ" -and !$files) {
        Write-Host "HALT - have Project Server farm and \media\ folder missing PRJ.  Download correct media and try again." -Fore Red
        Stop-Transcript
        Exit
    }
	
    # Halt if have multiple EXE and not SP2016
    $files = Get-ChildItem "$root\media\*.exe"
    if ($files -is [System.Array] -and $ver -ne 16) {
        # HALT - multiple EXE found - require clean up before continuing
        $files | Format-Table -AutoSize
        Write-Host "HALT - Multiple EXEs found. Clean up \media\ folder and try again." -Fore Red
        Stop-Transcript
        Exit
    }
    #>
}

function DetectAdmin() {
    # Are we running as local Administrator
    $wid = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $prp = New-Object System.Security.Principal.WindowsPrincipal($wid)
    $adm = [System.Security.Principal.WindowsBuiltInRole]::Administrator
    $IsAdmin = $prp.IsInRole($adm)
    if (!$IsAdmin) {
        (Get-Host).UI.RawUI.Backgroundcolor = "DarkRed"
        Clear-Host
        Write-Host "===== SPPatchify - Not Running as Administrator =====`nStarting an elevated PowerShell window...`n"
        $arguments = "& '" + $rootCmd + "' -phaseTwo"
        $arguments
        Start-Process powershell -Verb runAs -ArgumentList $arguments
        Break
    }
}

function SaveServiceInst() {
    # Save config to CSV
    $sos = Get-SPServiceInstance | Where-Object { $_.Status -eq "Online" } | Select-Object Id, TypeName, @{n = "Server"; e = { $_.Server.Address } }
    $sos | Export-Csv "$logFolder\sos-before-$when.csv" -Force -NoTypeInformation
}

function StartServiceInst() {
    # Restore config from CSV
    $files = Get-ChildItem "$logFolder\sos-before-*.csv" | Sort-Object LastWriteTime -Descending
    $sos = Import-Csv $files[0].FullName
    if ($sos) {
        foreach ($row in $sos) {
            $si = Get-SPServiceInstance $row.Id
            if ($si) {
                if ($si.Status -ne "Online") {
                    $row | Format-Table -AutoSize
                    Write-Host "Starting ... " -Fore Green
                    if ($si.TypeName -ne "User Profile Synchronization Service") {
                        # UPS needs password input to start via Central Admin GUI
                        $si.Provision()
                    }
                    if ($si.TypeName -eq "Distributed Cache") {
                        # Special command to initialize
                        Add-SPDistributedCacheServiceInstance
                        $si.Provision()
                    }
                    Write-Host "OK"
                }
            }
        }
    }
}
#endregion

function IsLocalServer($serverName) {
    if ($serverName.ToLower() -eq ($env:computername).ToLower()) {
        return Test
    }
    else {
        return false
    }
}

function VerifyRemotePS() {
    try {
        Write-Host "Test Remote PowerShell " -Fore Green
        # Loop servers
        foreach ($server in getFarmServers) {
            $addr = $server.Address
            if ($addr -ne $env:computername) {
                # Dynamic open PSSession               
                $remote = GetRemotePSSession $addr  (GetFarmAccountCredentials)                 
            }
        }
        Write-Host "Succeess" -Fore Green
        return $true
    }
    catch {
        throw 'ERROR - Not able to connect to one or more computers in the farm. Please make sure you have run [Enable-PSRemoting] and [Enable-WSManCredSSP -Role Server]'
    }
}

function ClearCacheIni() {  
    Write-Host "Clear CACHE.INI " -Fore Green

    # Stop the SharePoint Timer Service on each server in the farm
    Write-Host "Change SPTimer to OFF" -Fore Green
    ChangeSPTimer $false

    # Delete all xml files from cache config folder on each server in the farm
    Write-Host "Clear XML cache"
    DeleteXmlCache

    # Start the SharePoint Timer Service on each server in the farm
    Write-Host "Change SPTimer to ON" -Fore Red
    ChangeSPTimer $true
    Write-Host "Succeess" -Fore Green
}



# Stops the SharePoint Timer Service on each server in the SharePoint Farm.
function ChangeSPTimer($state) {
    # Constants
    $timer = "SPTimerV4"
    $timerInstance = "Microsoft SharePoint Foundation Timer"  

    # Iterate through each server in the farm, and each service in each server
    foreach ($server in getFarmServers) {
        foreach ($instance in $server.ServiceInstances) {
            # If the server has the timer service then stop the service
            if ($instance.TypeName -eq $timerInstance) {
                # Display
                $addr = $server.Address
                if ($state) {
                    $change = "Running"
                }
                else {
                    $change = "Stopped"
                }

                Write-Host -Foregroundcolor DarkGray -NoNewline "$timer service on server: "
                Write-Host -Foregroundcolor Gray $addr

                # Change
                $svc = Get-Service -ComputerName $addr -Name $timer
                $svc | Set-Service -StartupType Automatic
                if ($state) {
                    $svc | Start-Service
                }
                else {
                    $svc | Stop-Service
                }

                # Wait for service stop/start
                WaitSPTimer $addr $timer $change $state
                break;
            }
        }
    }
}


# Waits for the service on the server to reach the required service state.
function WaitSPTimer($addr, $service, $change, $state) {
    Write-Host -foregroundcolor DarkGray -NoNewLine "Waiting for $service to change to $change on server $addr"

    do {
        # Display
        Write-Host -Foregroundcolor DarkGray -NoNewLine "."

        # Get Service
        $svc = Get-Service -ComputerName $addr -Name $timer

        # Modify Service
        $svc | Set-Service -StartupType Automatic
        if ($state) {
            $svc | Start-Service
        }
        else {
            $svc | Stop-Service
        }
    }
    while ($svc.Status -ne $change)
    Write-Host -Foregroundcolor DarkGray -NoNewLine " Service is "
    Write-Host -Foregroundcolor Gray $change
}


# Removes all xml files recursive on an UNC path
function DeleteXmlCache() {
    Write-Host -foregroundcolor DarkGray "Delete xml files"

    # Iterate through each server in the farm, and each service in each server
    foreach ($server in getFarmServers) {
        foreach ($instance in $server.ServiceInstances) {
            # If the server has the timer service delete the XML files from the config cache
            if ($instance.TypeName -eq $timerServiceInstanceName) {
                [string]$serverName = $server.Name

                Write-Host -foregroundcolor DarkGray -NoNewline "Deleting xml files from config cache on server: $serverName"
                Write-Host -foregroundcolor Gray $serverName

                # Remove all xml files recursive on an UNC path
                $path = "\\" + $serverName + "\c$\ProgramData\Microsoft\SharePoint\Config\*-*\*.xml"
                Remove-Item -path $path -Force

                # 1 = refresh all cache settings
                $path = "\\" + $serverName + "\c$\ProgramData\Microsoft\SharePoint\Config\*-*\cache.ini"
                Set-Content -path $path -Value "1"

                break
            }
        }
    }
}

function TestRemotePS() {
    # Prepare
    Get-PSSession | Remove-PSSession -Confirm:$false
    #ReadIISPW

    # Connect
    foreach ($f in getRemoteServers) {
        New-PSSession -ComputerName $f.Address -Authentication Credssp -Credential (GetFarmAccountCredentials)
    }

    # WMI Uptime
    $sb = {
        $wmi = Get-WmiObject -Class Win32_OperatingSystem;
        $t = $wmi.ConvertToDateTime($wmi.LocalDateTime) - $wmi.ConvertToDateTime($wmi.LastBootUpTime);
        $t | Select-Object Days, Hours, Minutes
    }
    Invoke-Command -Session (Get-PSSession) -ScriptBlock $sb | Format-Table -AutoSize

    # Display
    Get-PSSession | Format-Table -AutoSize
    if ((getRemoteServers).Count -eq (Get-PSSession).Count) {
        $color = "Green"
    }
    else {
        $color = "Red"
    }
    Write-Host "Farm Servers : $((getFarmServers).Count)" -Fore $color
    Write-Host "Sessions     : $((Get-PSSession).Count)" -Fore $color
}

function VerifyWMIUptime() {
    # WMI Uptime
    $sb = {
        $wmi = Get-WmiObject -Class Win32_OperatingSystem;
        $t = $wmi.ConvertToDateTime($wmi.LocalDateTime) - $wmi.ConvertToDateTime($wmi.LastBootUpTime);
        $t;
    }
    $result = Invoke-Command -Session (Get-PSSession) -ScriptBlock $sb 

    # Compare threshold and suggest reboot
    $warn = 0
    foreach ($r in $result) {
        $TotalMinutes = [int]$r.TotalMinutes
        if ($TotalMinutes -gt $maxrebootminutes) {
            Write-Host "WARNING - Last reboot was $TotalMinutes minutes ago for $($r.PSComputerName)" -Fore Black -Backgroundcolor Yellow
            $warn++
        }
    }

    # Suggest reboot
    if ($warn) {
        # Prompt user
        $Readhost = Read-Host "Do you want to reboot above servers?  [Type R to Reboot.  Anything else to continue.]" 
        if ($ReadHost -like 'R*') { 
            # Reboot all
            Get-PSSession | Format-Table -Auto
            Write-Host "Rebooting above servers ... "
            $sb = { Restart-Computer -Force }
            Invoke-Command -ScriptBlock $sb -Session (Get-PSSession)
        }
    }
}


function AppOffline ($state) {
    # Deploy App_Offline.ht to peer IIS instances across the farm
    $ao = "app_offline.htm"
    $folders = Get-SPWebApplication | ForEach-Object { $_.IIsSettings[0].Path.FullName }
    # Start Jobs
    foreach ($server in getFarmServers) {
        $addr = $server.Address
        if ($addr -ne $env:computername) {
            foreach ($f in $folders) {
                # IIS Home Folders
                $remoteRoot = MakeRemote $f
                if ($state) {
                    # Install by HTM file copy
                    # Dynamic command
                    $dest = "\\$addr\$remoteroot\app_offline.htm"
                    Write-Host "Copying $ao to $dest" -Fore Yellow
                    ROBOCOPY $ao $dest /Z /MIR /W:0 /R:0
                }
                else {
                    # Uinstall by HTM file delete
                    # Dynamic command
                    $dest = "\\$addr\$remoteroot\app_offline.htm"
                    Write-Host "Deleting $ao to $dest" -Fore Yellow
                    Remove-ChildItem $dest -Confirm:$false
                }
            }
        }
    }
}


function rebootFarm() {
    foreach ($server in getRemoteServers) {
            
        Write-Host "Reboot $($server)" -Fore Yellow
        Restart-Computer -ComputerName $server.Name -Force            
    }
    Stop-Transcript
    Restart-Computer -Force
}

function AutoSPSourceBuilder() {
    <#PSScriptInfo
.VERSION 2.0.1.1
.GUID 6ba84db4-f1a9-4079-bd19-39cd044c6b11
.AUTHOR Brian Lalancette (@brianlala)
.COMPANYNAME
.COPYRIGHT 2019 Brian Lalancette
.TAGS SharePoint
.LICENSEURI
.PROJECTURI https://github.com/brianlala/AutoSPSourceBuilder
.ICONURI
.EXTERNALMODULEDEPENDENCIES BitsTransfer
.REQUIREDSCRIPTS
.EXTERNALSCRIPTDEPENDENCIES
.RELEASENOTES
.PRIVATEDATA
#> 

    <# 
.SYNOPSIS
    Builds a SharePoint 2010/2013/2016/2019 Service Pack + Cumulative/Public Update (and optionally slipstreamed) installation source.
.DESCRIPTION 
    Builds a SharePoint 2010/2013/2016/2019 Service Pack + Cumulative/Public Update (and optionally, slipstreamed) installation source. 
    Starting from existing (user-provided) SharePoint 2010/2013/2016/2019 installation media/files (and optional Office Web Apps / Online Server media/files),
    the script can download prerequisites, the specified Service Pack, and CU/PU packages for SharePoint/WAC, along with specified (optional) language packs, then extract them to a destination path structure.
    By default, automatically downloads the latest AutoSPSourceBuilder.xml inventory file as the source of product information (URLs, builds, naming, etc.) to the same local path as the AutoSPSourceBuilder.ps1 script.
.EXAMPLE
    AutoSPSourceBuilder.ps1 -UpdateLocation "C:\Users\brianl\Downloads\SP" -Destination "D:\SP\2010"
.EXAMPLE
    AutoSPSourceBuilder.ps1 -SourceLocation E: -Destination "C:\Source\SP\2010" -CumulativeUpdate "December 2011" -Languages fr-fr,es-es
.EXAMPLE
    AutoSPSourceBuilder.ps1 -SharePointVersion 2013 -Destination "C:\SP" -Verbose -UseExistingLocalXML
.PARAMETER SharePointVersion
    The version of SharePoint for which to download updates. Valid options are 2010, 2013, 2016 and 2019.
    This parameter is mandatory.
.PARAMETER SourceLocation
    The location (path, drive letter, etc.) where the SharePoint binary files are located.
    You can specify a UNC path (\\server\share\SP\2010), a drive letter (E:) or a local/mapped folder (Z:\SP\2010).
    If you don't provide a value, the script will not attempt to do any slipstreaming and will only download updates (and build out some of the folder structure).
.PARAMETER Destination
    The file path for the final slipstreamed SP2010/SP2013/2016/2019 installation files.
    The default value is $env:SystemDrive\SP\201x (where 201x is the version of SharePoint we want to download/integrate updates from).
.PARAMETER UpdateLocation
    The optional file path where the downloaded service pack and cumulative update files are located, or where they should be placed in case they need to be downloaded.
    It's recommended to omit this parameter in order to use the default value <Destination>\Updates (so, typically C:\SP\201x\Updates).
.PARAMETER GetPrerequisites
    Switch that specifies whether to attempt to download all prerequisite files for the selected product, which can be subsequently used to perform an offline installation.
    By default prerequisites are not downloaded.
.PARAMETER CumulativeUpdate
    The name of the cumulative update (CU) you'd like to integrate.
    The format should be e.g. "December 2011".
    If no value is provided, the script will prompt for an available CU name.
.PARAMETER WACSourceLocation
    The location (path, drive letter, etc.) where the Office Web Apps / Online Server binary files are located.
    You can specify a UNC path (\\server\share\SP\2010), a drive letter (E:) or a local/mapped folder (Z:\WAC).
    If no value is provided, the script will simply skip the WAC integration altogether.
.PARAMETER Languages
    A comma-separated list of languages (in the culture ID format, e.g. de-de,fr-fr) used to specify which language packs to download.
    If no languages are provided, and PromptForLanguages isn't specified, the script will simply skip language pack/update integration altogether.
.PARAMETER PromptForLanguages
    This switch indicates that the script should prompt the user for which languages to download updates (and language packs) for.
    If this switch is omitted, and no languages have been specified on the command line, languages are skipped entirely.
.PARAMETER UseExistingLocalXML
    If you want to use a custom or pre-existing AutoSPSourceBuilder.XML file, or simply want to skip downloading the official one on-the-fly, use this switch parameter.
.LINK
    https://github.com/brianlala/autospsourcebuilder
    https://github.com/brianlala/autospinstaller
    http://www.toddklindt.com/sp2010builds
.NOTES
    Created & maintained by Brian Lalancette (@brianlala), 2012-2018.
#>
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)][ValidateSet("2010", "2013", "2016", "2019")]
        [String]$SharePointVersion,
        [Parameter(Mandatory = $false)][ValidateNotNullOrEmpty()]
        [String]$SourceLocation,
        [Parameter(Mandatory = $false)][ValidateNotNullOrEmpty()]
        [String]$Destination = $env:SystemDrive + "\SP\$SharePointVersion",
        [Parameter(Mandatory = $false)][ValidateNotNullOrEmpty()]
        [String]$UpdateLocation,
        [Parameter(Mandatory = $false)]
        [Switch]$GetPrerequisites,
        [Parameter(Mandatory = $false)][ValidateNotNullOrEmpty()]
        [String]$CumulativeUpdate,
        [Parameter(Mandatory = $false)][ValidateNotNullOrEmpty()]
        [String]$WACSourceLocation,
        [Parameter(Mandatory = $false)][ValidateNotNullOrEmpty()]
        [Array]$Languages,
        [Parameter(Mandatory = $false)]
        [switch]$PromptForLanguages,
        [Parameter(Mandatory = $false)]
        [switch]$UseExistingLocalXML = $false
    )

    #region Functions
    # ===================================================================================
    # Func: Pause
    # Desc: Wait for user to press a key - normally used after an error has occured or input is required
    # ===================================================================================
    Function Pause($action, $key) {
        # From http://www.microsoft.com/technet/scriptcenter/resources/pstips/jan08/pstip0118.mspx
        if ($key -eq "any" -or ([string]::IsNullOrEmpty($key))) {
            $actionString = "Press any key to $action..."
            Write-Host $actionString
            $null = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
        else {
            $actionString = "Enter `"$key`" to $action"
            $continue = Read-Host -Prompt $actionString
            if ($continue -ne $key) { pause $action $key }
        }
    }

    Function WriteLine {
        Write-Host -ForegroundColor White "--------------------------------------------------------------"
    }

    Function DownloadPackage {
        Param
        (
            [Parameter()][string]$url,
            [Parameter()][string]$ExpandedFile,
            [Parameter()][string]$DestinationFolder,
            [Parameter()][string]$destinationFile
        )
        $file = $url.Split('/')[-1]
        If (!$destinationFile) { $destinationFile = $file }
        If (!$expandedFile) { $expandedFile = $file }
        Try {
            # Check if destination file or its expanded version already exists
            If (Test-Path "$DestinationFolder\$expandedFile") {
                # Check if the expanded file is already there
                Write-Host -ForegroundColor DarkGray "  - File $expandedFile exists, skipping download."
            }
            ElseIf (Test-Path "$DestinationFolder\$destinationFile") {
                # Check if the packed downloaded file is already there (in case of a CU)
                Write-Host -ForegroundColor DarkGray "  - File $destinationFile exists, skipping download."
            }
            ElseIf ((($file -eq $destinationFile) -or ("$file.zip" -eq $destinationFile)) -and ((Test-Path "$DestinationFolder\$file") -or (Test-Path "$DestinationFolder\$file.zip")) -and !((Get-Item $file -ErrorAction SilentlyContinue).Mode -eq "d----")) {
                # Check if the packed downloaded file is already there (in case of a CU or Prerequisite)
                Write-Host -ForegroundColor DarkGray "  - File $file exists, skipping download."
                If (!($file -like "*.zip")) {
                    # Give the CU package a .zip extension so we can work with it like a compressed folder
                    Rename-Item -Path "$DestinationFolder\$file" -NewName ($file + ".zip") -Force -ErrorAction SilentlyContinue
                }
            }
            Else {
                # Go ahead and download the missing package
                # Begin download
                Write-Verbose -Message " - Attempting to download from $url..."
                Import-Module BitsTransfer
                $job = Start-BitsTransfer -Asynchronous -Source $url -Destination "$DestinationFolder\$destinationFile" -DisplayName "Downloading `'$file`' to $DestinationFolder\$destinationFile" -Priority Foreground -Description "From $url..." -RetryInterval 60 -RetryTimeout 3600 -ErrorVariable err
                # When proxy is enabled
                # $job = Start-BitsTransfer -Asynchronous -Source $url -Destination "$DestinationFolder\$destinationFile" -DisplayName "Downloading `'$file`' to $DestinationFolder\$destinationFile" -Priority Foreground -Description "From $url..." -RetryInterval 60 -RetryTimeout 3600 -ProxyList canatsrv06:80 -ProxyUsage Override -ProxyAuthentication Ntlm -ProxyCredential $proxyCredentials -ErrorVariable err
                Write-Host "  - Connecting..." -NoNewline
                while ($job.JobState -eq "Connecting") {
                    Write-Host "." -NoNewline
                    Start-Sleep -Milliseconds 500
                }
                Write-Host "."
                If ($err) { Throw }
                Write-Host "  - Downloading $file..."
                while ($job.JobState -ne "Transferred") {
                    $percentDone = "{0:N2}" -f $($job.BytesTransferred / $job.BytesTotal * 100) + "% - $($job.JobState)"
                    Write-Host $percentDone -NoNewline
                    Start-Sleep -Milliseconds 500
                    $backspaceCount = (($percentDone).ToString()).Length
                    for ($count = 1; $count -le $backspaceCount; $count++) { Write-Host "`b `b" -NoNewline }
                    if ($job.JobState -like "*Error") {
                        Write-Host -ForegroundColor Yellow "  - An error occurred downloading $file, retrying..."
                        Resume-BitsTransfer -BitsJob $job -Asynchronous | Out-Null
                    }
                }
                Write-Host "  - Completing transfer..."
                Complete-BitsTransfer -BitsJob $job
                Write-Host " - Done!"
            }
        }
        Catch {
            Write-Output $err
            Write-Debug $_
            Write-Warning " - An error occurred downloading `'$file`'"
            $global:errorWarning = $true
            break
        }
    }

    Function Expand-Zip ($InputFile, $DestinationFolder) {
        $Shell = New-Object -ComObject Shell.Application
        $fileZip = $Shell.Namespace($InputFile)
        $Location = $Shell.Namespace($DestinationFolder)
        $Location.Copyhere($fileZip.items())
    }


    Function Read-Log() {
        $log = Get-ChildItem -Path (Get-Item $env:TEMP).FullName | Where-Object { $_.Name -like "opatchinstall*.log" } | Sort-Object -Descending -Property "LastWriteTime" | Select-Object -first 1
        If ($null -eq $log) {
            Write-Host `n
            Throw " - Could not find extraction log file!"
        }
        # Get error(s) from log
        $lastError = $log | select-string -SimpleMatch -Pattern "OPatchInstall: The extraction of the files failed" | Select-Object -Last 1
        If ($lastError) {
            Write-Host `n
            Write-Warning $lastError.Line
            $global:errorWarning = $true
            Invoke-Item $log.FullName
            Throw " - Review the log file and try to correct any error conditions."
        }
        Remove-Variable -Name log -ErrorAction SilentlyContinue
    }

    Function Remove-ReadOnlyAttribute ($Path) {
        ForEach ($item in (Get-ChildItem -File -Path $Path -Recurse -ErrorAction SilentlyContinue)) {
            $attributes = @((Get-ItemProperty -Path $item.FullName).Attributes)
            If ($attributes -match "ReadOnly") {
                # Set the file to just have the 'Archive' attribute
                Write-Host "  - Removing Read-Only attribute from file: $item"
                Set-ItemProperty -Path $item.FullName -Name Attributes -Value "Archive"
            }
        }
    }

    # ====================================================================================
    # Func: EnsureFolder
    # Desc: Checks for the existence and validity of a given path, and attempts to create if it doesn't exist.
    # From: Modified from patch 9833 at http://autospinstaller.codeplex.com/SourceControl/list/patches by user timiun
    # ====================================================================================
    Function EnsureFolder ($Path) {
        If (!(Test-Path -Path $Path -PathType Container)) {
            Write-Host -ForegroundColor White " - $Path doesn't exist; creating..."
            Try {
                New-Item -Path $Path -ItemType Directory | Out-Null
            }
            Catch {
                Write-Warning " - $($_.Exception.Message)"
                Throw " - Could not create folder $Path!"
                $global:errorWarning = $true
            }
        }
    }
    #endregion

    #region Admin Check
    # First check if we are running this under an elevated session. Pulled from the script at http://gallery.technet.microsoft.com/scriptcenter/1b5df952-9e10-470f-ad7c-dc2bdc2ac946
    If (!([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Warning " - You should run this script under an elevated PowerShell prompt. Launch an elevated PowerShell prompt by right-clicking the PowerShell shortcut and selecting `"Run as Administrator`"."
        Write-Warning " - Running without elevation may cause certain things to fail, e.g. file extraction."
        Pause -action "proceed if you are sure this is OK, or Ctrl-C to exit" -key "y"
    }
    #endregion

    #region OS Check
    # Then check if we are running Server 2012, Windows 8 or newer (e.g. Windows 10)
    $windowsMajorVersion, $windowsMinorVersion, $null = (Get-WmiObject Win32_OperatingSystem).Version -split "\."
    if (($windowsMajorVersion -lt 6 -or (($windowsMajorVersion -eq 6) -and ($windowsMinorVersion -lt 2)) -and $windowsMajorVersion -ne 10) -and ($Languages.Count -gt 0)) {
        Write-Warning "You should be running Windows Server 2012 or Windows 8 (minimum) to get the full functionality of this script."
        Write-Host -ForegroundColor Yellow " - Some features (e.g. image extraction) may not work otherwise."
        Pause "proceed if you are sure this is OK, or Ctrl-C to exit" "y"
    }
    #endregion

    #region Start
    $oldTitle = $Host.UI.RawUI.WindowTitle
    $Host.UI.RawUI.WindowTitle = "--AutoSPSourceBuilder--"
    $dp0 = $PSScriptRoot

    # Only needed if proxy is enabled
    # $proxyCredentials = (Get-Credential -Message "Enter credentials for proxy server:" -UserName "$env:USERDOMAIN\$env:USERNAME")


    Write-Host -ForegroundColor Green " -- AutoSPSourceBuilder SharePoint Update Download/Integration Utility --"
    <##>
    #$UseExistingLocalXML = $true
    if ($UseExistingLocalXML) {
        Write-Warning "'UseExistingLocalXML' specified; skipping download of AutoSPSourceBuilder.xml, and attempting to use local copy at '$dp0\AutoSPSourceBuilder.xml'."
        Write-Warning "This could mean you won't have the latest updates in your local copy."
        Write-Warning "To use the latest online AutoSPSourceBuilder.xml inventory file, omit the -UseExistingLocalXML switch."
    }
    else {
        # Get latest official XML file from the master GitHub repo
        Write-Output " - Grabbing latest AutoSPSourceBuilder.xml patch inventory file..."
        Start-BitsTransfer -DisplayName "Downloading AutoSPSourceBuilder.xml patch inventory file" -Description "From 'https://raw.githubusercontent.com/brianlala/AutoSPSourceBuilder/master/Scripts'..." -Destination "$dp0\AutoSPSourceBuilder.xml" -Priority Foreground -Source "https://raw.githubusercontent.com/brianlala/AutoSPSourceBuilder/master/Scripts/AutoSPSourceBuilder.xml" -RetryInterval 60 -RetryTimeout 3600
        if (!$?) {
            throw "Could not download AutoSPSourceBuilder.xml file!"
        }
    }
    [xml]$xml = (Get-Content -Path "$dp0\AutoSPSourceBuilder.xml" -ErrorAction SilentlyContinue)
    if ([string]::IsNullOrEmpty($xml)) {
        throw "Required AutoSPSourceBuilder.xml file not found! Please ensure it's located in the same folder as the AutoSPSourceBuilder.ps1 script ('$dp0')."
    }
    $spAvailableVersions = $xml.Products.Product.Name | Where-Object { $_ -like "SP*" }
    $spAvailableVersionNumbers = $spAvailableVersions -replace "SP", ""
    #endregion

    #region Determine Product Version and Languages Requested

    if (!([string]::IsNullOrEmpty($SharePointVersion))) {
        Write-Host " - SharePoint $SharePointVersion specified on the command line."
        $spYear = $SharePointVersion
    }

    # is is a hack to remove error from download media
    if ($UpdateLocation) {
        $Destination = $UpdateLocation
    }
    $Destination = $Destination.TrimEnd("\")
    # Ensure the Destination has the year at the end of the path, in case we forgot to type it in when/if prompted
    <#if (!($Destination -like "*$spYear")) {
        $Destination = $Destination + "\" + $spYear
    }
    #>
    Write-Verbose -Message "Destination is `"$Destination`""
    if ([string]::IsNullOrEmpty($UpdateLocation)) {
        $UpdateLocation = $Destination + "\Updates"
    }

    Write-Verbose -Message "Update location is `"$UpdateLocation`""

    if ($SourceLocation) {
        $sourceDir = $SourceLocation.TrimEnd("\")
        Write-Host " - Checking for $sourceDir\Setup.exe and $sourceDir\PrerequisiteInstaller.exe..."
        $sourceFound = ((Test-Path -Path "$sourceDir\Setup.exe") -and (Test-Path -Path "$sourceDir\PrerequisiteInstaller.exe"))
        # Inspired by http://vnucleus.com/2011/08/alphabet-range-sequences-in-powershell-and-a-usage-example/
        while (!$sourceFound) {
            foreach ($driveLetter in 68..90) {
                # Letters from D-Z
                # Check for the SharePoint DVD in all possible drive letters
                $sourceDir = "$([char]$driveLetter):"
                Write-Host " - Checking for $sourceDir\Setup.exe and $sourceDir\PrerequisiteInstaller.exe..."
                $sourceFound = ((Test-Path -Path "$sourceDir\Setup.exe") -and (Test-Path -Path "$sourceDir\PrerequisiteInstaller.exe"))
                If ($sourceFound -or $driveLetter -ge 90) { break }
            }
            break
        }
        if (!$sourceFound) {
            Write-Warning " - The correct SharePoint source files/media were not found!"
            Write-Warning " - Please insert/mount the correct media, or specify a valid path."
            $global:errorWarning = $true
            break
            Pause "exit"
            Stop-Transcript; Exit
        }
        else {
            Write-Host " - Source found in $sourceDir."
            $spVer, $null, $spBuild, $null = (Get-Item -Path "$sourceDir\setup.exe").VersionInfo.ProductVersion -split "\."
            # Create a hash table with 'wave' to product year mappings
            $spYears = @{"14" = "2010"; "15" = "2013"; "16" = "2016" } # Can't use this hashtable to map SharePoint 2019 versions because it uses version 16 as well
            $spYear = $spYears.$spVer
            # Accomodate SharePoint 2019 (uses the same major version number, but 5-digit build numbers)
            if ($spBuild.Length -eq 5) {
                $spYear = "2019"
            }
            If (!$sourceDir -and ([string]::IsNullOrEmpty($SharePointVersion))) {
                Write-Warning " - Cannot determine version of SharePoint setup binaries, and SharePointVersion was not specified."
                $global:errorWarning = $true
                break
                Pause "exit"
                Stop-Transcript; Exit
            }
            Write-Host " - SharePoint $spYear detected."
            if ($spYear -eq "2013") {
                $installerVer = (Get-Command "$sourceDir\setup.dll").FileVersionInfo.ProductVersion
                $null, $null, [int]$build, $null = $installerVer -split "\."
                If ($build -ge 4569) {
                    # SP2013 SP1
                    $sp2013SP1 = $true
                    Write-Host "  - Service Pack 1 detected."
                }
            }
        }
        if (!($sourceDir -eq "$Destination\SharePoint")) {
            WriteLine
            Write-Host " - (Robo-)copying files from $sourceDir to $Destination\SharePoint..."
            Start-Process -FilePath robocopy.exe -ArgumentList "`"$sourceDir`" `"$Destination\SharePoint`" /E /Z /ETA /NDL /NFL /NJH /XO /A-:R" -Wait -NoNewWindow
            Write-Host " - Done copying original files to $Destination\SharePoint."
            WriteLine
        }
    }
    else {
        Write-Host " - No source location specified; updates will be downloaded but slipstreaming will be skipped."
        if ($SharePointVersion) {
            Write-Host " - SharePoint $SharePointVersion specified on the command line."
            $spYear = $SharePointVersion
        }
        else {
            Write-Verbose -Message "`$SharePointVersion not specified; prompting..."
            Write-Host -ForegroundColor Cyan " - Please select the version of SharePoint from the list that appears..."
            Start-Sleep -Seconds 1
            while ([string]::IsNullOrEmpty($spYear)) {
                Start-Sleep -Seconds 1
                $spYear = $spAvailableVersionNumbers | Sort-Object | Out-GridView -Title "Please select the version of SharePoint to download updates for:" -OutputMode Single -PassThru
  
                if ($spYear.Count -gt 1) {
                    Write-Warning "Please only select ONE version. Re-prompting..."
                    Remove-Variable -Name spYear -Force -ErrorAction SilentlyContinue
                }
            }
            Write-Host " - SharePoint $spYear selected."
        }
    }
    if ($Languages.Count -lt 1) {
        Write-Host " - No languages specified on the command line.." -NoNewline
        if ($PromptForLanguages) {
            Start-Sleep -Seconds 1
            Write-Host "Will prompt for language(s)" -NoNewline
        }
        Write-Host "."
    }
    #endregion

    #region Determine Update Requested
    $spNode = $xml.Products.Product | Where-Object { $_.Name -eq "SP$spYear" }
    # Figure out which CU we want, but only if there are any available
    [array]$spCuNodes = $spNode.CumulativeUpdates.ChildNodes | Where-Object { $_.NodeType -ne "Comment" }
    if ((!([string]::IsNullOrEmpty($CumulativeUpdate))) -and !($spNode.CumulativeUpdates.CumulativeUpdate | Where-Object { $_.Name -eq $CumulativeUpdate })) {
        Write-Warning " - Invalid entry for update: `"$CumulativeUpdate`""
        Remove-Variable -Name CumulativeUpdate -ErrorAction SilentlyContinue
    }
    # Only prompt for an update if there are actually any to choose from, and if we haven't already specified one on the command line
    if (($spCuNodes).Count -ge 1 -and ([string]::IsNullOrEmpty($CumulativeUpdate))) {
        Start-Sleep -Seconds 1
        Write-Host -ForegroundColor Cyan " - Please select ONE available $(if ($spYear -eq "2016") {"Public"} else {"Cumulative"}) Update from the list that appears..."
        while ([string]::IsNullOrEmpty($selectedCumulativeUpdate)) {
            Start-Sleep -Seconds 1
            $selectedCumulativeUpdate = $spNode.CumulativeUpdates.CumulativeUpdate.Name | Select-Object -Unique | Out-GridView -Title "Please select an available $(if ($spYear -ge 2016) {"Public"} else {"Cumulative"}) Update for SharePoint $spYear`:" -PassThru
            if ($selectedCumulativeUpdate.Count -gt 1) {
                Write-Warning "Please only select ONE update. Re-prompting..."
                Remove-Variable -Name selectedCumulativeUpdate -Force -ErrorAction SilentlyContinue
            }
            if ([string]::IsNullOrEmpty($selectedCumulativeUpdate)) {
                Stop-Transcript; Exit
            }
        }
        $CumulativeUpdate = $selectedCumulativeUpdate
        Write-Host " - SharePoint $spYear $CumulativeUpdate $(if ($spYear -ge 2016) {"Public"} else {"Cumulative"}) Update selected."
    }
    [array]$spCU = $spNode.CumulativeUpdates.CumulativeUpdate | Where-Object { $_.Name -eq $CumulativeUpdate }
    if ($spCU.Count -ge 1) {
        # Only do this stuff if we actually have requested a CU
        $spCUName = $spCU[0].Name
        $spCUBuild = $spCU[0].Build
        if ($spYear -eq "2010") {
            # For SP2010 service packs
            $null, $null, $updateSubBuild, $null = $spCU[0].Build -split "\."
            # Get the service pack required, based on the sp* value in the CU URL - the URL will refer to the *upcoming* service pack and not the service pack required to apply the CU...
            if ($spCU[0].Url -like "*sp2*" -and $CumulativeUpdate -ne "August 2013" -and $CumulativeUpdate -ne "October 2013") {
                # As we would probably want at least SP1 if we are installing a CU prior to the August 2013 CU for SP2010
                $spServicePack = $spNode.ServicePacks.ServicePack | Where-Object { $_.Name -eq "SP1" }

            }
            elseif ($spCU[0].Url -like "*sp3*" -or $CumulativeUpdate -eq "August 2013" -or $updateSubBuild -gt 7140) {
                # We probably want SP2 if we are installing the August 2013 CU for SP2010, or a version newer than 14.0.7140.5000
                $spServicePack = $spNode.ServicePacks.ServicePack | Where-Object { $_.Name -eq "SP2" }
            }
        }
        elseif ($spYear -eq "2013") {
            # For SP2013 service packs
            if ($sp2013SP1) {
                $spServicePack = $spNode.ServicePacks.ServicePack | Where-Object { $_.Name -eq "SP1" }
            }
        }

        # Check if we are requesting the August 2014 CU, which sure enough, isn't cumulative and requires SP1 + July 2014 CU
        if ($CumulativeUpdate -eq "August 2014") {
            Write-Host " - The $CumulativeUpdate CU requires the July 2014 CU to be present first; will now attempt to integrate both."
            [array]$spCU = ($spNode.CumulativeUpdates.CumulativeUpdate | Where-Object { $_.Name -eq "July 2014" }), $spCU
        }
    }
    #endregion

    #region SharePoint Prerequisites
    if ($GetPrerequisites) {
        ## -and (!([string]::IsNullOrEmpty($SourceLocation))))
        WriteLine
        $spPrerequisiteNode = $spNode.Prerequisites
        EnsureFolder -Path "$Destination\SharePoint\PrerequisiteInstallerFiles"
        foreach ($prerequisite in $spPrerequisiteNode.Prerequisite) {
            Write-Host " - Getting prerequisite `"$($prerequisite.Name)`"..."
            # Because MS added a newer WcfDataServices.exe (yes, with the same filename) to the prerequisites list with SP2013 SP1, we need a special case here to ensure it's downloaded with a different name
            if ($prerequisite.Name -eq "Microsoft WCF Data Services 5.6" -and $spYear -eq "2013") {
                DownloadPackage -Url $($prerequisite.Url) -ExpandedFile "WcfDataServices56.exe" -DestinationFolder "$Destination\SharePoint\PrerequisiteInstallerFiles" -DestinationFile "WcfDataServices56.exe"
            }
            else {
                DownloadPackage -Url $($prerequisite.Url) -DestinationFolder "$Destination\SharePoint\PrerequisiteInstallerFiles"
            }
        }
        # Apply KB3087184 for SharePoint 2013 installer
        # KB3087184 can be considered a "prerequisite" for successful installation of SharePoint 2013 on a server that already has the .Net Framework 4.6 installed
        # Per https://support.microsoft.com/en-us/help/3087184/sharepoint-2013-or-project-server-2013-setup-error-if-the-net-framewor
        If ($spYear -eq "2013") {
            Write-Host " - Checking version of '$Destination\SharePoint\updates\svrsetup.dll'..."
            # Check to see if we have already patched/replaced svrsetup.dll
            $svrSetupDll = Get-Item -Path "$Destination\SharePoint\updates\svrsetup.dll" -ErrorAction SilentlyContinue
            # Check for presence and version of svrsetup.dll
            if ($svrSetupDll) {
                $null, $null, [int]$svrSetupDllBuild, $null = $svrSetupDll.VersionInfo.ProductVersion -split "\."
            }
            else {
                Write-Host " - '$Destination\SharePoint\updates\svrsetup.dll' was not found (or version could not be determined); attempting to update..."
            }
            if ($null -eq $svrSetupDllBuild -or ($svrSetupDllBuild -lt 4709)) {
                # 4709 is the version substring/build of the patched svrsetup.dll
                if (Test-Path -Path "$Destination\SharePoint\PrerequisiteInstallerFiles\svrsetup_15-0-4709-1000_x64.zip" -ErrorAction SilentlyContinue) {
                    Write-Host "  - Attempting to patch SharePoint 2013 installation source with updated svrsetup.dll from KB3087184..."
                    Write-Host "  - Per https://support.microsoft.com/en-us/help/3087184/sharepoint-2013-or-project-server-2013-setup-error-if-the-net-framewor"
                    # Rename the original file
                    Write-Host "   - Copying patched version of svrsetup.dll to '$Destination\SharePoint\updates'..."
                    Expand-Zip -InputFile "$Destination\SharePoint\PrerequisiteInstallerFiles\svrsetup_15-0-4709-1000_x64.zip" -DestinationFolder "$Destination\SharePoint\updates"
                    Write-Host " - Done."
                    $patchedForKB3087184 = $true
                }
                else {
                    Write-Host -ForegroundColor Yellow "  - Package for KB3087184 was not found; skipping patching of SharePoint $spYear installation source."
                }
            }
            else {
                Write-Host -ForegroundColor DarkGray "  - `"$Destination\SharePoint\updates\svrsetup.dll`" already exists and is already updated ($svrSetupDllBuild)."
                $patchedForKB3087184 = $true
            }
        }
        WriteLine
    }
    else {
        Write-Verbose -Message "Skipping prerequisites since GetPrerequisites or SourceLocation were not specified."
    }
    #endregion

    #region Prompt for Language Packs
    if (($PromptForLanguages)) {
        $lpNode = $spNode.LanguagePacks
        # Prompt for an available language pack
        $availableLanguageNames = $lpNode.LanguagePack | Where-Object { $null -ne $_.Url } | Select-Object Name | Sort-Object Name
        Write-Host -ForegroundColor Cyan " - Please select one or more available language pack(s) from the list that appears..."
        Start-Sleep -Seconds 2
        [array]$Languages = $availableLanguageNames.Name | Out-GridView -Title "Please select one or more available language pack(s). Hold down Ctrl to select multiple, or click Cancel to skip:" -PassThru
        if ($Languages.Count -eq 0) {
            Write-Host " - No languages selected."
        }
    }
    #endregion

    #region SharePoint Service Pack
    If ($spServicePack -and ($spYear -ne "2013") -and (!([string]::IsNullOrEmpty($SourceLocation)))) {
        # Exclude SharePoint 2013 service packs as slipstreaming support has changed
        if ($spServicePack.Name -eq "SP1" -and $spYear -eq "2010") { $spMspCount = 40 } # Service Pack 1 should have 40 .msp files
        if ($spServicePack.Name -eq "SP2" -and $spYear -eq "2010") { $spMspCount = 47 } # Service Pack 2 should have 47 .msp files
        else { $spMspCount = 0 }
        WriteLine
        # Check if a SharePoint service pack already appears to be included in the source
        If ((Get-ChildItem "$sourceDir\Updates" -Filter *.msp).Count -lt $spMspCount) {
            # Checking for specific number of MSP patch files in the \Updates folder
            Write-Host " - $($spServicePack.Name) seems to be missing, or incomplete in $sourceDir\; downloading..."
            # Set the subfolder name for easy update build & name identification, for example, "15.0.4481.1005 (March 2013)"
            $spServicePackSubfolder = $spServicePack.Build + " (" + $spServicePack.Name + ")"
            EnsureFolder -Path "$UpdateLocation\$spServicePackSubfolder"
            DownloadPackage -Url $($spServicePack.Url) -DestinationFolder "$UpdateLocation\$spServicePackSubfolder"
            Remove-ReadOnlyAttribute -Path "$Destination\SharePoint\Updates"
            # Extract SharePoint service pack patch files
            $spServicePackExpandedFile = $($spServicePack.Url).Split('/')[-1]
            Write-Verbose -Message " - Extracting from '$UpdateLocation\$spServicePackSubfolder\$spServicePackExpandedFile'"
            Write-Host " - Extracting SharePoint $($spServicePack.Name) patch files..." -NoNewline
            Start-Process -FilePath "$UpdateLocation\$spServicePackSubfolder\$spServicePackExpandedFile" -ArgumentList "/extract:`"$Destination\SharePoint\Updates`" /passive" -Wait -NoNewWindow
            Read-Log
            Write-Host "done!"
        }
        Else { Write-Host " - $($spServicePack.Name) appears to be already slipstreamed into the SharePoint binary source location." }

        ## Extract SharePoint w/SP1 files (future functionality?)
        ## Start-Process -FilePath "$UpdateLocation\en_sharepoint_server_2010_with_service_pack_1_x64_759775.exe" -ArgumentList "/extract:$Destination\SharePoint /passive" -NoNewWindow -Wait -NoNewWindow
        WriteLine
    }
    else {
        Write-Verbose -Message "Not processing service pack."
    }
    #endregion

    #region March PU for SharePoint 2013
    # Since the March 2013 PU for SharePoint 2013 is considered the baseline build for all patches going forward (prior to SP1), we need to download and extract it if we are looking for a SP2013 CU dated March 2013 or later
    If ($spCU.Count -ge 1 -and $spCU[0].Name -ne "December 2012" -and $spYear -eq "2013" -and !$sp2013SP1 -and !([string]::IsNullOrEmpty($SourceLocation))) {
        WriteLine
        $march2013PU = $spNode.CumulativeUpdates.CumulativeUpdate | Where-Object { $_.Name -eq "March 2013" }
        Write-Host " - Getting SharePoint $spYear baseline update $($march2013PU.Name) PU:"
        $march2013PUFile = $($march2013PU.Url).Split('/')[-1]
        if ($march2013PU.Url -like "*zip.exe") {
            $march2013PUFileIsZip = $true
            $march2013PUFile += ".zip"
        }
        # Set the subfolder name for easy update build & name identification, for example, "15.0.4481.1005 (March 2013)"
        $updateSubfolder = $march2013PU.Build + " (" + $march2013PU.Name + ")"
        EnsureFolder -Path "$UpdateLocation\$updateSubfolder"
        DownloadPackage -Url $($march2013PU.Url) -ExpandedFile $($march2013PU.ExpandedFile) -DestinationFolder "$UpdateLocation\$updateSubfolder" -destinationFile $march2013PUFile
        # Expand PU executable to $UpdateLocation\$updateSubfolder
        If (!(Test-Path "$UpdateLocation\$updateSubfolder\$($march2013PU.ExpandedFile)") -and $march2013PUFileIsZip) {
            # Ensure the expanded file isn't already there, and the PU is a zip
            $march2013PUFileZipPath = Join-Path -Path "$UpdateLocation\$updateSubfolder" -ChildPath $march2013PUFile
            Write-Host " - Expanding $($march2013PU.Name) Public Update (single file)..."
            # Remove any pre-existing hotfix.txt file so we aren't prompted to replace it by Expand-Zip and cause our script to pause
            if (Test-Path -Path "$UpdateLocation\$updateSubfolder\hotfix.txt" -ErrorAction SilentlyContinue) {
                Remove-Item -Path "$UpdateLocation\$updateSubfolder\hotfix.txt" -Confirm:$false -ErrorAction SilentlyContinue
            }
            Expand-Zip -InputFile $march2013PUFileZipPath -DestinationFolder "$UpdateLocation\$updateSubfolder"
        }
        Remove-ReadOnlyAttribute -Path "$Destination\SharePoint\Updates"
        $march2013PUTempFolder = "$Destination\SharePoint\Updates\March2013PU_TEMP"
        # Remove any existing .xml or .msp files
        foreach ($existingItem in (Get-ChildItem -Path $march2013PUTempFolder -ErrorAction SilentlyContinue)) {
            $existingItem | Remove-Item -Force -Confirm:$false
        }
        # Extract SharePoint PU files to $march2013PUTempFolder
        Write-Verbose -Message " - Extracting from '$UpdateLocation\$updateSubfolder\$($march2013PU.ExpandedFile)'"
        Write-Host " - Extracting $($march2013PU.Name) Public Update patch files..." -NoNewline
        Start-Process -FilePath "$UpdateLocation\$updateSubfolder\$($march2013PU.ExpandedFile)" -ArgumentList "/extract:`"$march2013PUTempFolder`" /passive" -Wait -NoNewWindow
        Read-Log
        Write-Host "done!"
        # Now that we have a supported way to slispstream BOTH the March 2013 PU as well as a subsequent CU (per http://blogs.technet.com/b/acasilla/archive/2014/03/09/slipstream-sharepoint-2013-with-march-pu-cu.aspx), let's make it happen.
        Write-Host " - Processing $($march2013PU.Name) Public Update patch files (to allow slipstreaming with a later CU)..." -NoNewline
        # Grab every file except for the eula.txt (or any other text files) and any pre-existing renamed files
        foreach ($item in (Get-ChildItem -Path "$march2013PUTempFolder" | Where-Object { $_.Name -notlike "*.txt" -and $_.Name -notlike "_*SP0" })) {
            $prefix, $extension = $item -split "\."
            $newName = "_$prefix-SP0.$extension"
            if (Test-Path -Path "$march2013PUTempFolder\$newName") {
                Remove-Item -Path "$march2013PUTempFolder\$newName" -Force -Confirm:$false
            }
            Rename-Item -Path "$($item.FullName)" -NewName $newName -ErrorAction Inquire
        }
        # Move March 2013 PU files up into \Updates folder
        foreach ($item in (Get-ChildItem -Path "$march2013PUTempFolder")) {
            $item | Move-Item -Destination "$Destination\SharePoint\Updates" -Force
        }
        Remove-Item -Path $march2013PUTempFolder -Force -Confirm:$false
        Write-Host "done!"
        WriteLine
    }
    #endregion

    #region SharePoint Updates
    If ($spCU.Count -ge 1) {
        $null, $null, [int]$spServicePackBuildNumber, $null = $spServicePack.Build -split "\."
        $null, $null, [int]$spCUBuildNumber, $null = $spCUBuild -split "\."
        if (($spCU.Url[0] -like "*`/$($spServicePack.Name)`/*") -and ($spServicePackBuildNumber -gt $spCUBuildNumber)) {
            # New; only get the CU if its URL doesn't contain the service pack we already have and if the build is older, as it will likely be older
            Write-Host -ForegroundColor DarkGray " - The $($spCU.Name[0]) update appears to be older than the SharePoint $spYear service pack or binaries, skipping."
            # Mark that the CU, although requested, has been skipped for the reason above. Used so that the output .txt file report remains accurate.
            $spCUSkipped = $true
        }
        else {
            WriteLine
            foreach ($spCUpackage in $spCU) {
                $spCUpackageIndex ++
                $spCUPackageName = $spCUpackage.Name
                $spCUPackageBuild = $spCUpackage.Build
                $spCUFile = $($spCUPackage.Url).Split('/')[-1]
                Write-Host " - Getting SharePoint $spYear $($spCUPackageName) update file ($spCUFile):"
                if ($spCUPackage.Url -like "*zip.exe") {
                    $spCuFileIsZip = $true
                    $spCuFile += ".zip"
                }
                # Set the subfolder name for easy update build & name identification, for example, "15.0.4481.1005 (March 2013)"
                $updateSubfolder = $spCUPackageBuild + " (" + $spCUPackageName + ")"
                EnsureFolder -Path "$UpdateLocation\$updateSubfolder"
                DownloadPackage -Url $($spCUPackage.Url) -ExpandedFile $($spCUPackage.ExpandedFile) -DestinationFolder "$UpdateLocation\$updateSubfolder" -destinationFile $spCuFile
                # Only do this if we are interested in slipstreaming, e.g. if we have a $SourceLocation
                if (!([string]::IsNullOrEmpty($SourceLocation))) {
                    # Expand CU executable to $UpdateLocation\$updateSubfolder
                    if (!(Test-Path "$UpdateLocation\$updateSubfolder\$($spCUPackage.ExpandedFile)") -and $spCuFileIsZip) {
                        # Ensure the expanded file isn't already there, and the CU is a zip
                        $spCuFileZipPath = Join-Path -Path "$UpdateLocation\$updateSubfolder" -ChildPath $spCuFile
                        Write-Host " - Expanding $spCuFile $(if ($spYear -ge 2016) {"Public"} else {"Cumulative"}) Update (single file)..."
                        # Remove any pre-existing hotfix.txt file so we aren't prompted to replace it by Expand-Zip and cause our script to pause
                        if (Test-Path -Path "$UpdateLocation\$updateSubfolder\hotfix.txt" -ErrorAction SilentlyContinue) {
                            Remove-Item -Path "$UpdateLocation\$updateSubfolder\hotfix.txt" -Confirm:$false -ErrorAction SilentlyContinue
                        }
                        Expand-Zip -InputFile $spCuFileZipPath -DestinationFolder "$UpdateLocation\$updateSubfolder"
                    }
                    Remove-ReadOnlyAttribute -Path "$Destination\SharePoint\Updates"
                    # Extract SharePoint CU files to $Destination\SharePoint\Updates (but only if the source file is an .exe)
                    if ($spCUPackage.ExpandedFile -like "*.exe") {
                        # Assuming this is the the "launcher" package and the only one with an .exe extension. This is to differentiate from the ubersrv*.cab files included recently as part of CUs
                        [array]$spCULaunchers += $spCUPackage.ExpandedFile
                    }
                    if ($spCUpackageIndex -eq $spCU.Count) {
                        # Now that all packages have been downloaded we can call the launcher .exe to extract the CU
                        if ($spYear -eq "2013") {
                            Write-Host -ForegroundColor Cyan " - NOTE: SP2013 updates can take a VERY long time to extract, so please be patient!"
                        }
                        Write-Host " - Extracting $($spCUName) $(if ($spYear -ge 2016) {"Public"} else {"Cumulative"}) Update patch files..."
                        foreach ($spCULauncher in $spCULaunchers) {
                            Write-Verbose -Message "  - Extracting from '$UpdateLocation\$updateSubfolder\$spCULauncher'"
                            Write-Host "  - $spCULauncher..." -NoNewline
                            Start-Process -FilePath "$UpdateLocation\$updateSubfolder\$spCULauncher" -ArgumentList "/extract:`"$Destination\SharePoint\Updates`" /passive" -Wait -NoNewWindow
                            Write-Host "done!"
                            Read-Log
                        }
                        Write-Host " - Extracting update patch files done!"
                    }
                }
                else {
                    Write-Verbose -Message "Skipping slipstreaming since no source location was specified."
                }
                WriteLine
            }
        }
    }
    else {
        Write-Verbose -Message "Skipping SharePoint updates since no SharePoint $spYear updates were found."
    }
    #endregion

    #region Office Web Apps / Online Server
    if ($spYear -le 2013) { $wacProductName = "Office Web Apps"; $wacNodeName = "OfficeWebApps" }
    elseif ($spYear -ge 2016) { $wacProductName = "Office Online Server"; $wacNodeName = "OfficeOnlineServer" }
    if ($WACSourceLocation) {
        $wacNode = $xml.Products.Product | Where-Object { $_.Name -eq "$wacNodeName$spYear" }
        $wacServicePack = $wacNode.ServicePacks.ServicePack | Where-Object { $_.Name -eq $spServicePack.Name } # To match the chosen SharePoint service pack
        if ($wacServicePack.Name -eq "SP1" -and $spYear -eq "2010") { $wacMspCount = 16 }
        if ($wacServicePack.Name -eq "SP2" -and $spYear -eq "2010") { $wacMspCount = 32 }
        else { $wacMspCount = 0 }
        # Create a hash table with some directories to look for to confirm the valid presence of the WAC binaries. Not perfect.
        $wacTestDirs = @{"2010" = "XLSERVERWAC.en-us"; "2013" = "wacservermui.en-us"; "2016" = "wacserver.ww" }
        ##if ($spYear -eq "2010") {$wacTestDir = "XLSERVERWAC.en-us"}
        ##elseif ($spYear -eq "2013") {$wacTestDir = "wacservermui.en-us"}
        # Try to find a OWA/OOS update that matches the current month for the SharePoint update
        [array]$wacCU = $wacNode.CumulativeUpdates.CumulativeUpdate | Where-Object { $_.Name -eq $spCUName }
        [array]$wacCUNodes = $wacNode.CumulativeUpdates.ChildNodes | Where-Object { $_.NodeType -ne "Comment" }
        if ([string]::IsNullOrEmpty($wacCU)) {
            Write-Host " - There is no $($spCUName) update for $wacProductName available."
            Start-Sleep -Seconds 1
            while ([string]::IsNullOrEmpty($wacCUName) -and (($wacCUNodes).Count -ge 1)) {
                Write-Host -ForegroundColor Cyan " - Please select another available $wacProductName update..."
                Start-Sleep -Seconds 1
                $wacCUName = $wacNode.CumulativeUpdates.CumulativeUpdate.Name | Select-Object -Unique | Out-GridView -Title "Please select another available $wacProductName update:" -PassThru
                if ($wacCUName.Count -gt 1) {
                    Write-Warning "Please only select ONE update. Re-prompting..."
                    Remove-Variable -Name wacCUName -Force -ErrorAction SilentlyContinue
                }
            }
            [array]$wacCU = $wacNode.CumulativeUpdates.CumulativeUpdate | Where-Object { $_.Name -eq $wacCUName }
        }
        else {
            Write-Host " - $($wacCU[0].Name) update found for $wacProductName."
        }
        if ($wacCU.Count -ge 1) {
            $wacCUName = $wacCU[0].Name
            $wacCUBuild = $wacCU[0].Build
        }
        WriteLine
        # Download Office Web Apps / Online Server?

        # Download Office Web Apps / Online Server 2013 Prerequisites
        If ($GetPrerequisites -and $spYear -ge 2013) {
            WriteLine
            $wacPrerequisiteNode = $wacNode.Prerequisites
            EnsureFolder -Path "$Destination\$wacNodeName\PrerequisiteInstallerFiles"
            foreach ($prerequisite in $wacPrerequisiteNode.Prerequisite) {
                Write-Host " - Getting $wacProductName prerequisite `"$($prerequisite.Name)`"..."
                DownloadPackage -Url $($prerequisite.Url) -DestinationFolder "$Destination\$wacNodeName\PrerequisiteInstallerFiles"
            }
            WriteLine
        }
        {
            Write-Verbose -Message "Skipping $wacProductName updates since GetPrerequisites wasn't specified or $wacProductName is older than 2013."
        }
        # Extract Office Web Apps / Online Server files to $Destination\$wacNodeName
        $sourceDirWAC = $WACSourceLocation.TrimEnd("\")
        Write-Host " - Checking for $sourceDirWAC\$($wacTestDirs.$spYear)\..."
        $sourceFoundWAC = (Test-Path -Path "$sourceDirWAC\$($wacTestDirs.$spYear)" -ErrorAction SilentlyContinue)
        if (!$sourceFoundWAC) {
            Write-Warning " - The correct $wacProductName source files/media were not found!"
            Write-Warning " - Please specify a valid path."
            $global:errorWarning = $true
            break
            Pause "exit"
            Stop-Transcript; Exit
        }
        else {
            Write-Host " - Source found in $sourceDirWAC."
        }
        if (!($sourceDirWAC -eq "$Destination\$wacNodeName")) {
            Write-Host " - (Robo-)copying files from $sourceDirWAC to $Destination\$wacNodeName..."
            Start-Process -FilePath robocopy.exe -ArgumentList "`"$sourceDirWAC`" `"$Destination\$wacNodeName`" /E /Z /ETA /NDL /NFL /NJH /XO /A-:R" -Wait -NoNewWindow
            Write-Host " - Done copying original files to $Destination\$wacNodeName."
        }

        if (!([string]::IsNullOrEmpty($wacServicePack.Name))) {
            # Check if WAC SP already appears to be included in the source
            if ((Get-ChildItem "$sourceDirWAC\Updates" -Filter *.msp).Count -lt $wacMspCount) {
                # Checking for ($wacMspCount) MSP patch files in the \Updates folder
                Write-Host " - WAC $($wacServicePack.Name) seems to be missing or incomplete in $sourceDirWAC; downloading..."
                # Download Office Web Apps / Online Server service pack
                Write-Host " - Getting $wacProductName $($wacServicePack.Name):"
                # Set the subfolder name for easy update build & name identification, for example, "15.0.4481.1005 (March 2013)"
                $wacServicePackSubfolder = $wacServicePack.Build + " (" + $wacServicePack.Name + ")"
                EnsureFolder -Path "$UpdateLocation\$wacServicePackSubfolder"
                DownloadPackage -Url $($wacServicePack.Url) -DestinationFolder "$UpdateLocation\$wacServicePackSubfolder"
                Remove-ReadOnlyAttribute -Path "$Destination\$wacNodeName\Updates"
                # Extract Office Web Apps / Online Server service pack files to $Destination\$wacNodeName\Updates
                $wacServicePackExpandedFile = $($wacServicePack.Url).Split('/')[-1]
                Write-Verbose -Message " - Extracting from '$$UpdateLocation\$wacServicePackSubfolder\$wacServicePackExpandedFile'"
                Write-Host " - Extracting $wacProductName $($wacServicePack.Name) patch files..." -NoNewline
                Start-Process -FilePath "$UpdateLocation\$wacServicePackSubfolder\$wacServicePackExpandedFile" -ArgumentList "/extract:`"$Destination\$wacNodeName\Updates`" /passive" -Wait -NoNewWindow
                Read-Log
                Write-Host "done!"
            }
            else { Write-Host " - WAC $($wacServicePack.Name) appears to be already slipstreamed into the SharePoint binary source location." }
        }
        else { Write-Host " - No WAC service packs are available or applicable for this version." }

        if ($spCU.Count -ge 1 -and [string]::IsNullOrEmpty($wacCU)) {
        }
        if (!([string]::IsNullOrEmpty($wacCU))) {
            # Only attempt this if we actually have a CU for WAC that matches the SP revision
            # Download Office Web Apps / Online Server CU
            foreach ($wacCUPackage in $wacCU) {
                Write-Host " - Getting $wacProductName $wacCUName update:"
                $wacCuFileZip = $($wacCUPackage.Url).Split('/')[-1] + ".zip"
                # Set the subfolder name for easy update build & name identification, for example, "15.0.4481.1005 (March 2013)"
                $wacUpdateSubfolder = $wacCUBuild + " (" + $wacCUName + ")"
                EnsureFolder -Path "$UpdateLocation\$wacUpdateSubfolder"
                DownloadPackage -Url $($wacCUPackage.Url) -ExpandedFile $($wacCUPackage.ExpandedFile) -DestinationFolder "$UpdateLocation\$wacUpdateSubfolder" -destinationFile $wacCuFileZip

                # Expand Office Web Apps / Online Server CU executable to $UpdateLocation\$wacUpdateSubfolder
                If (!(Test-Path "$UpdateLocation\$wacUpdateSubfolder\$($wacCUPackage.ExpandedFile)")) {
                    # Check if the expanded file is already there
                    $wacCuFileZipPath = Join-Path -Path "$UpdateLocation\$wacUpdateSubfolder" -ChildPath $wacCuFileZip
                    Write-Host " - Expanding $wacProductName $(if ($spYear -ge 2016) {"Public"} else {"Cumulative"}) Update (single file)..."
                    EnsureFolder -Path "$UpdateLocation\$wacUpdateSubfolder"
                    # Remove any pre-existing hotfix.txt file so we aren't prompted to replace it by Expand-Zip and cause our script to pause
                    if (Test-Path -Path "$UpdateLocation\$wacUpdateSubfolder\hotfix.txt" -ErrorAction SilentlyContinue) {
                        Remove-Item -Path "$UpdateLocation\$wacUpdateSubfolder\hotfix.txt" -Confirm:$false -ErrorAction SilentlyContinue
                    }
                    Expand-Zip -InputFile $wacCuFileZipPath -DestinationFolder "$UpdateLocation\$wacUpdateSubfolder"
                }
                Remove-ReadOnlyAttribute -Path "$Destination\$wacNodeName\Updates"
                # Extract Office Web Apps / Online Server CU files to $Destination\$wacNodeName\Updates
                Write-Verbose -Message " - Extracting from '$UpdateLocation\$wacUpdateSubfolder\$($wacCUPackage.ExpandedFile)'"
                Write-Host " - Extracting $wacProductName $(if ($spYear -ge 2016) {"Public"} else {"Cumulative"}) Update patch files..." -NoNewline
                Start-Process -FilePath "$UpdateLocation\$wacUpdateSubfolder\$($wacCUPackage.ExpandedFile)" -ArgumentList "/extract:`"$Destination\$wacNodeName\Updates`" /passive" -Wait -NoNewWindow
                Write-Host "done!"
            }
        }
        else { Write-Host " - No $wacProductName updates are available or applicable for this version." }
        WriteLine
    }
    else {
        Write-Verbose -Message "Skipping $wacProductName updates since no $wacProductName location was specified."
    }
    #endregion

    #region Language Packs
    if ($Languages.Count -gt 0) {
        # Remove any spaces or quotes and ensure each one is split out
        [array]$languages = $Languages -replace ' ', '' -split ","
        Write-Host " - Languages requested:"
        foreach ($language in $Languages) {
            Write-Host "  - $language"
        }
        $lpNode = $spNode.LanguagePacks
        ForEach ($language in $Languages) {
            WriteLine
            $spLanguagePack = $lpNode.LanguagePack | Where-Object { $_.Name -eq $language }
            If (!$spLanguagePack) {
                Write-Warning " - Language Pack `"$language`" invalid, or not found - skipping."
            }
            if ([string]::IsNullOrEmpty($spLanguagePack.Url)) {
                Write-Warning " - There is no download URL for Language Pack `"$language`" yet - skipping. You may need to download it manually from MSDN/Technet."
            }
            Else {
                # Download the language pack
                [array]$validLanguages += $language
                $lpDestinationFile = $($spLanguagePack.Url).Split('/')[-1]
                # Give it a more descriptive name if the language sub-string is not already present
                if (!($lpDestinationFile -like "*$language*")) {
                    if ($spver -eq "14") {
                        $lpDestinationFile = $lpDestinationFile -replace ".exe", "_$language.exe"
                    }
                    else {
                        $lpDestinationFile = $lpDestinationFile -replace ".img", "_$language.img"
                    }
                }
                Write-Host " - Getting SharePoint $spYear Language Pack ($language):"
                # Set the subfolder name for easy update build & name identification, for example, "15.0.4481.1005 (March 2013)"
                $spLanguagePackSubfolder = $spLanguagePack.Name
                EnsureFolder -Path "$UpdateLocation\$spLanguagePackSubfolder"
                DownloadPackage -Url $($spLanguagePack.Url) -DestinationFolder "$UpdateLocation\$spLanguagePackSubfolder" -DestinationFile $lpDestinationFile
                Remove-ReadOnlyAttribute -Path "$Destination\LanguagePacks\$language"
                # Extract the language pack to $Destination\LanguagePacks\xx-xx (where xx-xx is the culture ID of the language pack, for example fr-fr)
                if ($lpDestinationFile -match ".img$" -or $lpDestinationFile -match ".iso$") {
                    # Mount the ISO/IMG file ($UpdateLocation\$spLanguagePackSubfolder\$lpDestinationFile) and robo-copy the files to $Destination\LanguagePacks\$language
                    Write-Host " - Mounting language pack disk image..." -NoNewline
                    Mount-DiskImage -ImagePath "$UpdateLocation\$spLanguagePackSubfolder\$lpDestinationFile" -StorageType ISO
                    $isoDrive = (Get-DiskImage -ImagePath "$UpdateLocation\$spLanguagePackSubfolder\$lpDestinationFile" | Get-Volume).DriveLetter + ":"
                    Write-Host "Done."

                    # Copy files
                    Write-Host " - (Robo-)copying language pack files from $isoDrive to $Destination\LanguagePacks\$language"
                    Start-Process -FilePath robocopy.exe -ArgumentList "`"$isoDrive`" `"$Destination\LanguagePacks\$language`" /E /Z /ETA /NDL /NFL /NJH /XO /A-:R" -Wait -NoNewWindow
                    Write-Host " - Done copying language pack files to $Destination\LanguagePacks\$language."
                    # Dismount the ISO/IMG
                    Dismount-DiskImage -ImagePath "$UpdateLocation\$spLanguagePackSubfolder\$lpDestinationFile"
                }
                else {
                    Write-Verbose -Message " - Extracting from '$UpdateLocation\$spLanguagePackSubfolder\$lpDestinationFile'"
                    Write-Host " - Extracting Language Pack files ($language)..." -NoNewline
                    Start-Process -FilePath "$UpdateLocation\$spLanguagePackSubfolder\$lpDestinationFile" -ArgumentList "/extract:`"$Destination\LanguagePacks\$language`" /quiet" -Wait -NoNewWindow
                    Write-Host "done!"
                }
                [array]$lpSpNodes = $splanguagePack.ServicePacks.ChildNodes | Where-Object { $_.NodeType -ne "Comment" }
                if (($lpSpNodes).Count -ge 1 -and $spServicePack) {
                    # Download service pack for the language pack
                    $lpServicePack = $spLanguagePack.ServicePacks.ServicePack | Where-Object { $_.Name -eq $spServicePack.Name } # To match the chosen SharePoint service pack
                    $lpServicePackDestinationFile = $($lpServicePack.Url).Split('/')[-1]
                    Write-Host " - Getting SharePoint $spYear Language Pack $($lpServicePack.Name) ($language):"
                    EnsureFolder -Path "$UpdateLocation\$spLanguagePackSubfolder"
                    DownloadPackage -Url $($lpServicePack.Url) -DestinationFolder "$UpdateLocation\$spLanguagePackSubfolder" -DestinationFile $lpServicePackDestinationFile
                    if (Test-Path -Path "$Destination\LanguagePacks\$language\Updates") { Remove-ReadOnlyAttribute -Path "$Destination\LanguagePacks\$language\Updates" }
                    # Extract each language pack to $Destination\LanguagePacks\xx-xx (where xx-xx is the culture ID of the language pack, for example fr-fr)
                    if ($lpServicePackDestinationFile -match ".img$") {
                        # Mount the ISO/IMG file ($UpdateLocation\$spLanguagePackSubfolder\$lpDestinationFile) and robo-copy the files to $Destination\LanguagePacks\$language
                        Write-Host " - Mounting language pack service pack disk image..." -NoNewline
                        Mount-DiskImage -ImagePath "$UpdateLocation\$spLanguagePackSubfolder\$lpServicePackDestinationFile" -StorageType ISO
                        $isoDrive = (Get-DiskImage -ImagePath "$UpdateLocation\$spLanguagePackSubfolder\$lpServicePackDestinationFile" | Get-Volume).DriveLetter + ":"
                        Write-Host "Done."

                        # Copy files
                        Write-Host " - (Robo-)copying language pack service pack files from $isoDrive to $Destination\LanguagePacks\$language"
                        Start-Process -FilePath robocopy.exe -ArgumentList "`"$isoDrive`" `"$Destination\LanguagePacks\$language`" /E /Z /ETA /NDL /NFL /NJH /XO /A-:R" -Wait -NoNewWindow
                        Write-Host " - Done copying language pack service pack files to $Destination\LanguagePacks\$language."
                        # Dismount the ISO/IMG
                        Dismount-DiskImage -ImagePath "$UpdateLocation\$spLanguagePackSubfolder\$lpServicePackDestinationFile"
                    }
                    else {
                        Write-Verbose -Message " - Extracting from '$UpdateLocation\$spLanguagePackSubfolder\$lpServicePackDestinationFile'"
                        Write-Host " - Extracting Language Pack $($lpServicePack.Name) files ($language)..." -NoNewline
                        Start-Process -FilePath "$UpdateLocation\$spLanguagePackSubfolder\$lpServicePackDestinationFile" -ArgumentList "/extract:`"$Destination\LanguagePacks\$language\Updates`" /quiet" -Wait -NoNewWindow
                        Write-Host "done!"
                    }
                }
            }
            If ($spCU.Count -ge 1 -and (Test-Path -Path "$Destination\LanguagePacks\$language\Updates") -and (!([string]::IsNullOrEmpty($SourceLocation)))) {
                # Copy matching culture files from $Destination\SharePoint\Updates folder (e.g. spsmui-fr-fr.msp) to $Destination\LanguagePacks\$language\Updates
                Write-Host " - Updating $Destination\LanguagePacks\$language with the $($spCUName) SharePoint update..."
                ForEach ($patch in (Get-ChildItem -Path $Destination\SharePoint\Updates -Filter *$language*)) {
                    Copy-Item -Path $patch.FullName -Destination "$Destination\LanguagePacks\$language\Updates" -Force
                }
            }
            else {
                Write-Verbose -Message "Skipping language pack slipstreaming since no source location was specified."
            }
            WriteLine
        }
    }
    else {
        Write-Verbose -Message "Skipping language packs & language updates since no language was specified."
    }

    #endregion
    #$Destination = $UpdateLocation
    #region Wrap Up
    WriteLine
    if (!([string]::IsNullOrEmpty($SourceLocation))) {
        $textFileName = "_SLIPSTREAM_HISTORY.txt"
    }
    else {
        $textFileName = "_UPDATE_HISTORY.txt"
    }
    # Append the history file if it already exists
    if (Get-item -Path "$Destination\$textFileName" -ErrorAction SilentlyContinue) {
        Write-Output " - Appending version history file `"$textFileName`"..."
    }
    else {
        Write-Output " - Adding a version history file `"$textFileName`"..."
        Set-Content -Path "$Destination\$textFileName" -Value "This media source directory has been prepared with:" -Force
    }
    Add-Content -Path "$Destination\$textFileName" -Value "-------------------------------------------------------------------------------------------------------------------------------------" -Force
    Add-Content -Path "$Destination\$textFileName" -Value "$(Get-Date):" -Force
    if (!([string]::IsNullOrEmpty($SourceLocation))) {
        Add-Content -Path "$Destination\$textFileName" -Value "- SharePoint $spYear" -Force
    }
    If (!([string]::IsNullOrEmpty($spServicePack))) {
        Add-Content -Path "$Destination\$textFileName" -Value " - $($spServicePack.Name) for SharePoint $spYear" -Force
    }
    If (!([string]::IsNullOrEmpty($march2013PU))) {
        Add-Content -Path "$Destination\$textFileName" -Value " - $($march2013PU.Name) Public Update for SharePoint $spYear" -Force
    }
    If (($spCU.Count -ge 1) -and !$spCUSkipped) {
        Add-Content -Path "$Destination\$textFileName" -Value " - $($spCUName) $(if ($spYear -ge 2016) {"Public"} else {"Cumulative"}) Update for SharePoint $spYear" -Force
    }
    If ($GetPrerequisites -and !([string]::IsNullOrEmpty($SourceLocation))) {
        Add-Content -Path "$Destination\$textFileName" -Value " - Prerequisite software for SharePoint $spYear" -Force
    }
    if ($patchedForKB3087184) {
        Add-Content -Path "$Destination\$textFileName" -Value " - .Net Framework 4.6 installation compatibility update (KB3087184) for SharePoint $spYear" -Force
    }
    If ($validLanguages.Count -gt 0) {
        # Add the language packs to the txt file only if they were actually valid
        Add-Content -Path "$Destination\$textFileName" -Value " - Language Packs:" -Force
        ForEach ($language in $validLanguages) {
            Add-Content -Path "$Destination\$textFileName" -Value "  - $language" -Force
        }
    }
    If (!([string]::IsNullOrEmpty($WACSourceLocation))) {
        Add-Content -Path "$Destination\$textFileName" -Value "- $wacProductName $spYear" -Force
        if (!([string]::IsNullOrEmpty($wacPrerequisiteNode))) {
            Add-Content -Path "$Destination\$textFileName" -Value " - Prerequisite software for $wacProductName $spYear" -Force
        }
        if (!([string]::IsNullOrEmpty($wacServicePack))) {
            Add-Content -Path "$Destination\$textFileName" -Value " - $($wacServicePack.Name) for $wacProductName $spYear" -Force
        }
        if (!([string]::IsNullOrEmpty($wacCU))) {
            Add-Content -Path "$Destination\$textFileName" -Value " - $($wacCUName) $(if ($spYear -ge 2016) {"Public"} else {"Cumulative"}) Update for $wacProductName $spYear" -Force
        }
    }
    Add-Content -Path "$Destination\$textFileName" -Value "Using AutoSPSourceBuilder (https://github.com/brianlala/autospsourcebuilder)." -Force
    Add-Content -Path "$Destination\$textFileName" -Value "-------------------------------------------------------------------------------------------------------------------------------------" -Force
    If ($errorWarning) {
        Write-Host -ForegroundColor Yellow " - At least one non-trivial error was encountered."
        if (!([string]::IsNullOrEmpty($SourceLocation))) {
            Write-Host -ForegroundColor Yellow " - Your SharePoint installation source could therefore be incomplete."
        }
        Write-Host -ForegroundColor Yellow " - You should re-run this script until there are no more errors."
    }
    Write-Output " - Done!"
    Write-Output " - Review the output and check your source/update file integrity carefully."
    Start-Sleep -Seconds 3
    Invoke-Item -Path $Destination


    WriteLine
    $Host.UI.RawUI.WindowTitle = $oldTitle
    #Pause "exit"
    #endregion
    Write-Output "$UpdateLocation\$spServicePackSubfolder"
    Return "$UpdateLocation\$spServicePackSubfolder"
}


# function distribute($servers, $jobs) {
function DistributedJobs($scriptBlocks, [string[]]$servers, [System.Management.Automation.PSCredential]$credentials = (GetFarmAccountCredentials)) {
    
    if (!$servers -or !$scriptBlocks) {
        return
    }  
    if ($servers.count -eq 1) {
        $servers = $env:computername
    }
    else { 
        $servers = $servers | Where-Object { $_ -ne $env:computername }
    }
    $scriptBlocks 
    Get-Job | Stop-job | Remove-Job 
    Get-PSSession | Remove-PSSession
    # $servers = "WBSSP201902", "WBSSP201901"
    # $data = $databases = Get-SPContentDatabase | select name
    $data = [System.Collections.ArrayList]$scriptBlocks
    # $data = New-Object System.Collections.ArrayList
    # $data.AddRange(@("param1", "param2", "param3", "param4", "param5", "param6", "param7", "param8", "param9", "param10", "para11"))
    $jobs = New-Object System.Collections.ArrayList

    do {
        Write-Host "Checking job states." -ForegroundColor Yellow
        $toremove = @()
        foreach ($job in $jobs) {
            Write-Verbose $job.State
            if ($job.State -ne "Running") {
                $result = Receive-Job $job
                Write-Host "result $result"
                <# if ($result[0] -ne "ScriptRan") {
                    Write-Host "  Adding data back to que >> $($job.InData)" -ForegroundColor Green
                    $data.Add($job.InData) | Out-Null
                }#>

                $toremove += $job
            }
        }

        Write-Host "Removing completed/failed jobs" -ForegroundColor Yellow
        foreach ($job in $toremove) {
            Write-Host "  Removing job >> $($job.Location)" -ForegroundColor Green
            $jobs.Remove($job) | Out-Null
        }

        # Check if there is room to start another job
        Write-Host "$($jobs.Count) -lt $($servers.Count) -and $($data.Count)"
        if ($jobs.Count -lt $servers.Count -and $data.Count -gt 0) {
            Write-Host "Checking servers if they can start a new job." -ForegroundColor Yellow
            foreach ($server in $servers ) {  
                write-host "server: $server "    
                $job = $jobs | ? Location -eq $server
                if ($job -eq $null) {
                    Write-Host "  Adding job for $server >> $($data[0].Name)" -ForegroundColor Green
                    write-host $data[0]
                    # No active job was found for the server, so add new job
                    # $job = Invoke-Command -ScriptBlock $data[0] -Session $session -AsJob 
                    $job = InvokeCommand -server $server -ScriptBlock $data[0] -isJob $true

                    <#{
                        param($data, $hostname)
                        if ((Get-PSSnapin | Where-Object { $_.Name -eq "Microsoft.SharePoint.PowerShell" }) -eq $null) { 
                            Add-PSSnapIn "Microsoft.SharePoint.Powershell" 
                        }
                        $results = Upgrade-SPContentDatabase $data -Confirm:$false -WarningVariable warn -ErrorVariable erro                 
                        @("ScriptRan", $warn, $erro)                   
                    } #>
                    
                    $job | Add-Member -MemberType NoteProperty -Name InData -Value $data[0]
                    $jobs.Add($job) | Out-Null
                    $data.Remove($data[0])
                }
            }
        }
        # Just a manual check of $jobs
        Write-Output $jobs
        # Wait a bit before checking again
        Start-Sleep -Seconds 60
    } while ($data.Count -gt 0)

}


Main


