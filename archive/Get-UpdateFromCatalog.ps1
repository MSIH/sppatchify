function Get-UpdateFromCatalog { 
    <# 
        .SYNOPSIS 
            Get the specified Windows update details. 
            For that it generates COM IE windows, and parses the source code of the generated web pages. 
             
            Requirements: 
            The 'Microsoft Update Catalog' ActiveX (http://www.catalog.update.microsoft.com/Search.aspx) must be already installed. 
             
            Notes: 
            Script should not be called twice at the same server due it will cause issues identifying child IE opened sessions. 
            Script cannot handle more than 25 results (1 search page output) 
             
        .PARAMETER Name 
            (Mandatory) Name of the item to look for. It can be a product, KB title or number. 
            Multiple values can be provided 
 
        .PARAMETER DownloadFolder 
            If specified, the installers will be downloaded to '<specified folder>\<kb title>\<installer>' 
                 
        .EXAMPLE 
            PS> Get-UpdateFromCatalog -Name "3197867","KB3197876" -verbose | out-gridview 
             
        .EXAMPLE 
            PS> Get-UpdateFromCatalog -Name "3197867","KB3197876" -DownloadFolder ".\DownloadedPatches" -verbose  
    #> 
    [CmdletBinding()] 
    param ( 
        [parameter(Mandatory=$True)][ValidateNotNullOrEmpty()][string[]]$Name=$Null, 
        [parameter(Mandatory=$False)][ValidateNotNullOrEmpty()][string]$DownloadFolder=$Null 
    ) 
     
    ## Closes any previous IE popup 
    (New-Object -ComObject Shell.Application).Windows() | Where-Object {$_.LocationUrl -like "*catalog.update.microsoft.com/*DownloadDialog.aspx*"} | foreach-object { $_.Quit() } 
     
    ## Adds '*.update.microsoft.com to the list of Trusted sites to be able to allow popups from it and get the 'direct link' 
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\EscDomains\microsoft.com\*.update" -Name "https" -Value 2 -Type DWORD -Force | Out-null 
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Internet Explorer\New Windows\Allow" -Name '*.update.microsoft.com' -Value "0" -Type Binary -Force | Out-Null 
         
    [array]$Report=$Null 
    Foreach ($Item in $Name) { 
        $URL='http://www.catalog.update.microsoft.com/Search.aspx?q=' + $Item.replace(' ','%20') 
         
        write-verbose " * Get-UpdateFromCatalog: Search url: `n $URL" 
         
        $WebClient = New-Object System.Net.WebClient 
        $WebClient.Encoding = [System.Text.Encoding]::UTF8 
        do { 
            try {  
                $Content = $null 
                $Content = $WebClient.DownloadString($URL) 
                $Result=$? 
            } catch {} 
        } while (-not ($Result)) 
         
         
        if ($Content -match '(?s)(?<=>We did not find any results for)(.+?)(?=<\/span>)') {  
            write-warning "No results were found. Try using another search criteria..." 
            return $Null 
        } ## 25 is the max amount of rows in a page 
         
        try { 
            $Temp=([regex]::Match($Content, '(?s)(?<=searchDuration">)(.+?)(?=<\/span>)')).value.trim().tostring().split(' ') | where { $_ } 
            $Results=[int]"$($Temp[4])" 
        } catch {} 
        if ($Results -gt 25) { write-warning "Your search returned $($Results) results, but no more than 25 will be handled by the script. Try using another search criteria..."} ## 25 is the max amount of rows in a page 
         
        $Items=[regex]::Matches($Content, '(?s)<tr id=(.|\n)*?<\/tr>') | foreach-object { $_.Groups[0].Value } 
        foreach ($Item in $Items) { 
            $UpdateID=([regex]::Match($item, '(?s)(?<=<input id=")(.+?)(?=")')).value.trim() 
             
            ####### Get Update additional information 
            $URL='http://www.catalog.update.microsoft.com/ScopedViewInline.aspx?updateid=' + $UpdateID 
            write-verbose " * Get-UpdateFromCatalog: Getting additional update information from: `n $URL" 
             
            $ie = New-Object -ComObject InternetExplorer.Application  
            $ie.Navigate2($URL, "0x1000") 
            $ie.Visible = $False 
            while( $ie.busy){Start-Sleep 1} 
             
            try { 
                [array]$SupersedesBy=[regex]::Match($ie.document.documentElement.innerHTML, "(?s)(?<=supersededbyInfo\>)(.+?)(?=<SPAN id\=)").value.trim() | foreach-object { 
                    ([regex]::Matches($_, '(?s)(?<="\>)(.+?)(?=<\/DIV>)')).Groups[1].Value.trim() | foreach-object { 
                        ([regex]::Match($_, '(?s)(?<="\>)(.+?)(?=<)')).value.trim()  
                    } 
                } 
            } catch {} 
                     
            try { 
                [array]$Supersedes=[regex]::Match($ie.document.documentElement.innerHTML, "(?s)(?<=supersedesInfo\>)(.+?)(?=<DIV id\=)").value.trim() | foreach-object { 
                    [regex]::Matches($_, '(?s)(?<="\>)(.+?)(?=<\/DIV>)') | foreach-object {  
                        $_.Groups[0].Value.trim()  
                    } 
                } 
            } catch {} 
             
            $ie.Quit() 
             
             
            ####### Get Update direct download link 
            $URL='http://catalog.update.microsoft.com/?updateid=' + $UpdateID 
            write-verbose " * Get-UpdateFromCatalog: Getting direct download link from: `n $URL" 
             
            $ie = New-Object -ComObject InternetExplorer.Application  
            $ie.Navigate2($URL, "0x1000") 
            $ie.Visible = $False 
            while( $ie.busy){Start-Sleep 1}  
            $DownloadBTN = $ie.Document.getElementsByName("downloadButton") | Where-Object {$_.Type -eq 'button'}  
            $DownloadBTN.click() 
            $IEDownloadWindow = (New-Object -ComObject Shell.Application).Windows() | Where-Object {$_.LocationUrl -like "*catalog.update.microsoft.com/*DownloadDialog.aspx*"} 
            $IEDownloadWindow.Visible = $False 
            while( $IEDownloadWindow.busy){Start-Sleep 1}  
            $DirectDownloadLink=[regex]::Match($IEDownloadWindow.document.documentElement.innerHTML, "(?s)(?<=url \= ')(.+?)(?=';)").value.trim() 
            $ie.Quit() 
            $IEDownloadWindow.Quit() 
         
            $Obj = New-Object -TypeName PSObject -Property (@{ 
                'Title'=[regex]::Match($item, '(?s)(?<=<a id=)(.+?)(?=<\/a>)').value.split('>')[-1].trim() 
                'Products'=[regex]::Match($item, '(?s)(?<=_C2_)(.+?)(?=<\/td>)').value.split('>')[-1].trim() 
                'Classification'=[regex]::Match($item, '(?s)(?<=_C3_)(.+?)(?=<\/td>)').value.split('>')[-1].trim() 
                'LastUpdated'=[regex]::Match($item, '(?s)(?<=_C4_)(.+?)(?=<\/td>)').value.split('>')[-1].trim() 
                'Version'=[regex]::Match($item, '(?s)(?<=_C5_)(.+?)(?=<\/td>)').value.split('>')[-1].trim() 
                'Size'=[regex]::Match($item, '(?s)(?<=size">)(.+?)(?=<\/span>)').value.trim() 
                'CatalogURL'='http://catalog.update.microsoft.com/?updateid=' + $UpdateID 
                'DirectDownloadLink'=$DirectDownloadLink 
                'Supersedes'=$Supersedes 
                'SupersedesBy'=$SupersedesBy 
            }) 
             
            ## To have PS object sorted 
            $Report+=$Obj | select Title, Products, Classification, LastUpdated, Version, Size, CatalogURL, DirectDownloadLink, Supersedes, SupersedesBy 
        } 
    } 
     
    ##### Download section 
    if ($DownloadFolder) { 
        write-verbose " * Get-UpdateFromCatalog: Attemping to download files..." 
         
        if (-not (Test-Path $DownloadFolder)) { 
            write-verbose " * Get-UpdateFromCatalog: Creating folder ""$($DownloadFolder)""..." 
            $DownloadFolderPath=New-Item $DownloadFolder -Type "Directory" -Confirm:$False -Force | foreach-object { $_.Fullname } 
        } else { $DownloadFolderPath=get-item $DownloadFolder | foreach-object { $_.Fullname } } 
         
        foreach ($Item in $Report) 
        { 
            $SubFolder=New-Item "$($DownloadFolderPath)\$($Item.Title)" -Type "Directory" -Confirm:$False -Force | foreach-object { $_.Fullname } 
            [string]$Url=$Item.DirectDownloadLink 
            [string]$FileName="$($SubFolder)\$(($Url).split('/')[-1])" 
            write-verbose " * Get-UpdateFromCatalog: Downloading ""$Url"" to ""$FileName""..." 
            try { (new-object System.Net.WebClient).DownloadFile($Url,$FileName)  
            } catch { write-error "$(Get-Date -Format o) - Get-MBSAMissingUpdates: $_" } 
        } 
    } 
     
    return $Report 
}
