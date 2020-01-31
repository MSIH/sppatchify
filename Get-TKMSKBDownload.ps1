function Get-TKMSKBDownload

{
<#
.Synopsis
   Downloads file from Microsoft KB Download link.
.DESCRIPTION
   Takes the URL for a Microsoft patch download and downloads it to the current directory. Find out more at http://toddklindt.com/PSDownloadMSPatch 
   v1 Published 12/8/2017
.EXAMPLE
   Get-TKMSKBDownload -url https://www.microsoft.com/en-us/download/details.aspx?id=56230
   Download the patch linked from a Microsoft patch page

.EXAMPLE
   Get-TKMSKBDownload -url https://www.microsoft.com/en-us/download/confirmation.aspx?id=56230
   You can also use the confirmation URL
   
.EXAMPLE
   Get-TKMSKBDownload -url 56230
   You can pass just the patch ID as well
   
#>
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        # details.aspx or confirmation.aspx url for the patch you want
        [Parameter(Mandatory=$true,
                   Position=0)]
        $url
    )

    Process
    {
        
        # support for passing just the patch number, i.e. 56230
        If ($url -match '^[0-9]+$') {
            $url = "https://www.microsoft.com/en-us/download/confirmation.aspx?id=$url"

            }
        
        # details.aspx pages won't work but confirmation pages will. Swapping them out.
        $url = $url.replace("details.","confirmation.")


        # Get the URL for the binary download from the patch page
        try 
            {
                $downloadurl = ((Invoke-WebRequest -UseBasicParsing -Uri $url).links | Where-Object -Property data-bi-cN -Like -Value "click here to download manually" | select -First 1).href
            } 
        catch
            {
                $url = $url.replace("confirmation.","details.")
                Write-Host "$url could not be found"
                break
            }

        # Filename of the file
        $file = $downloadurl.Substring($downloadurl.LastIndexOf("/") + 1)

        # Putting it all together
        Write-Host "Downloading $file from $downloadurl"
        Invoke-WebRequest -UseBasicParsing -Uri $downloadurl -OutFile $file
    }
    
}
