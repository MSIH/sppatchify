$spAvailableVersionNumbers = @{
    2019 = "2019"
    2016 = "2016"
    2013 = "2013"
}
$spYear = $null

Write-Verbose -Message "`$SharePointVersion not specified; prompting..."
Write-Host -ForegroundColor Cyan " - Please select the version of SharePoint from the list that appears..."
Start-Sleep -Seconds 1
while ([string]::IsNullOrEmpty($spYear)) {
    Start-Sleep -Seconds 1
    $spYear = $spAvailableVersionNumbers | Sort-Object | Out-GridView -Title "Please select the version of SharePoint to download updates for:" -PassThru
    if ($spYear.Count -gt 1) {
        Write-Warning "Please only select ONE version. Re-prompting..."
        Remove-Variable -Name spYear -Force -ErrorAction SilentlyContinue
    }
}
$spYear
Write-Host " - SharePoint $spYear selected."

$spYear.Value

$URL = 'http://www.catalog.update.microsoft.com/Search.aspx?q=' + "SharePoint%20$($spYear.Key)"

$result = Invoke-WebRequest $url

$result.Links | Select href

$Menu = [ordered]@{

    1 = 'Do something'

    2 = 'Do this instead'

    3 = 'Do whatever you  want'

}
$selectedCumulativeUpdate = $null
while ([string]::IsNullOrEmpty($selectedCumulativeUpdate)) {
    Start-Sleep -Seconds 1
    $selectedCumulativeUpdate = $Menu | Select-Object -Unique | Out-GridView -Title "Please select an available $(if ($spYear -ge 2016) {"Public"} else {"Cumulative"}) Update for SharePoint $spYear`:" -PassThru
    if ($selectedCumulativeUpdate.Count -gt 1) {
        Write-Warning "Please only select ONE update. Re-prompting..."
        Remove-Variable -Name selectedCumulativeUpdate -Force -ErrorAction SilentlyContinue
    }
}
$CumulativeUpdate = $selectedCumulativeUpdate
Write-Host " - SharePoint $spYear $CumulativeUpdate $(if ($spYear -ge 2016) {"Public"} else {"Cumulative"}) Update selected."
