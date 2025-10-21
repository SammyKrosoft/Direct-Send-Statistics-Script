Start-DirectSendAudit {
    [CmdletBinding()]
    param()

    # ----------------------------
    # Clear any existing Exchange Online sessions (silent)
    # ----------------------------
    Get-PSSession | Where-Object { $_.ComputerName -like "*.outlook.com" -or $_.ComputerName -like "*.office365.com" } | Remove-PSSession -ErrorAction SilentlyContinue

    # ----------------------------
    # Prompt for output folder
    # ----------------------------
    $defaultFolder = "C:\Temp"
    $userFolder = Read-Host "Enter output folder path or press Enter to use default [$defaultFolder]"
    if ([string]::IsNullOrWhiteSpace($userFolder)) { $OutputFolder = $defaultFolder } else { $OutputFolder = $userFolder }
    if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder | Out-Null }

    # ----------------------------
    # Prompt for number of days
    # ----------------------------
    $defaultDays = 7
    $userDays = Read-Host "Enter number of days to generate the report for (1-90, default $defaultDays)"
    if ([string]::IsNullOrWhiteSpace($userDays)) { $Days = $defaultDays }
    elseif ([int]::TryParse($userDays, [ref]$null) -and $userDays -ge 1 -and $userDays -le 90) { $Days = [int]$userDays }
    else { Write-Host "Invalid input, using default $defaultDays days."; $Days = $defaultDays }

    $endDate = Get-Date
    $startDate = $endDate.AddDays(-$Days)
    Write-Host "`nGenerating report from $($startDate.ToShortDateString()) to $($endDate.ToShortDateString())" -ForegroundColor Cyan
    Write-Host "Output folder: $OutputFolder`n" -ForegroundColor Cyan

    # ----------------------------
    # Connect to Exchange Online
    # ----------------------------
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
    Connect-ExchangeOnline -ErrorAction Stop

    # ----------------------------
    # Check Direct Send status
    # ----------------------------
    $orgConfig = Get-OrganizationConfig | Select-Object Identity,RejectDirectSend
    if ($orgConfig.RejectDirectSend -eq $false) { Write-Host "Direct Send appears to be enabled on your tenant." -ForegroundColor Green }
    else { Write-Host "Direct Send is not enabled on your tenant." -ForegroundColor Green }

    # ----------------------------
    # Detect tenant domains
    # ----------------------------
    $tenantDomains = (Get-AcceptedDomain).DomainName
    $primaryDomain = (Get-AcceptedDomain | Where-Object { $_.Default -eq $true }).DomainName
    Write-Host "Detected tenant domains:" -ForegroundColor Cyan -NoNewLine
    Write-Host "$($tenantDomains -join ', ')`n" -ForegroundColor White

    # ----------------------------
    # Determine notification email
    # ----------------------------
    try {
        $currentUser = (Get-ConnectionInformation).UserPrincipalName
    } catch {
        $currentUser = $null
    }

    if (-not $currentUser) {
        $currentUser = (Get-ExoMailbox -ResultSize 1).PrimarySmtpAddress.ToString()
    }

    $defaultNotifyAddress = $currentUser
    $notifyAddress = Read-Host "Enter notification email address. This must exist on the tenant. (default: $defaultNotifyAddress)"
    if ([string]::IsNullOrWhiteSpace($notifyAddress)) { $notifyAddress = $defaultNotifyAddress }

    Write-Host "`nUsing notification address: $notifyAddress`n" -ForegroundColor Cyan

    # ----------------------------
    # Start Historical Search
    # ----------------------------
    $reportTitle = "NoConnectorMessages-" + ([Guid]::NewGuid().ToString())
    Start-HistoricalSearch -ReportTitle $reportTitle `
        -StartDate $startDate `
        -EndDate $endDate `
        -ReportType ConnectorReport `
        -ConnectorType NoConnector `
        -Direction Received `
        -NotifyAddress $notifyAddress

    Write-Host "Historical search started." -ForegroundColor Cyan -NoNewLine
    Write-Host " This can take 30 minutes or longer for large data sets." -ForegroundColor White
    Write-Host "Please keep this powershell tab open." -ForegroundColor Yellow
    Write-Host "Searching for internal messages sent with no connector, indicative of Direct Send messages." -ForegroundColor Cyan
    Write-Host "Please note this report only returns results for messages sent with no connector. Further investigation may be required to determine Direct Send was used."
    Write-Host "To review all inbound messages sent without a connector please see the original Historical Message Trace report in your chosen output folder."
    Write-Host "Waiting until report is ready..." -ForegroundColor Cyan

    # ----------------------------
    # Poll until report is ready
    # ----------------------------
    $report = $null
    do {
        Start-Sleep -Seconds 60
        $report = Get-HistoricalSearch | Where-Object { $_.ReportTitle -eq $reportTitle -and $_.Status -eq "Done" }
        Write-Host "Checking report status..."
    } until ($report)

    Write-Host "`nReport is ready." -ForegroundColor Green

    # ----------------------------
    # Purview login and CSV download
    # ----------------------------
    do {
        Write-Host "`nPress " -ForegroundColor White -NoNewLine
        Write-Host "Enter " -ForegroundColor Yellow -NoNewLine
        Write-Host "to open your default browser and log in to Purview..." -ForegroundColor White
        [void][System.Console]::ReadLine()
        Start-Process $report.FileUrl

        Write-Host "After logging in to Purview the report should automatically start downloading. This can take a while."
        Write-Host "Please note this may appear as a blank browser page with 'Working...' displayed in the tab. "
        Write-Host "Once the report has finished downloading press " -ForegroundColor White -NoNewLine
        Write-Host "Enter " -ForegroundColor Yellow -NoNewLine
        Write-Host "to continue." -ForegroundColor White
        [void][System.Console]::ReadLine()

        $downloads = [Environment]::GetFolderPath("UserProfile") + "\Downloads"
        $latestCsv = Get-ChildItem -Path $downloads -Filter "*ConnectorReport*.csv" | Sort-Object LastWriteTime -Descending | Select-Object -First 1

        if ($latestCsv -eq $null) {
            Write-Host "No ConnectorReport CSV found in Downloads folder. Please wait for the download to finish and press Enter again."
        }
    } until ($latestCsv -ne $null)

    # Keep original filename when moving
    $downloadPath = Join-Path $OutputFolder $latestCsv.Name
    Move-Item -Path $latestCsv.FullName -Destination $downloadPath -Force
    Write-Host "Report moved to: $downloadPath`n" -ForegroundColor Cyan

    # ----------------------------
    # CSV filtering function
    # ----------------------------
    function Filter-CsvBySender {
        param(
            [Parameter(Mandatory=$true)][string]$InputCsv,
            [Parameter(Mandatory=$true)][string]$OutputCsv,
            [Parameter(Mandatory=$true)][string[]]$Domains
        )

        $bytes = [System.IO.File]::ReadAllBytes($InputCsv)
        $csvText = -join ($bytes | ForEach-Object {
            if (($_ -ge 32 -and $_ -le 126) -or $_ -eq 10 -or $_ -eq 13) { [char]$_ } else { '' }
        })
        $lines = $csvText -split "`r?`n"
        $header = $lines[0] -replace '[<>"]', '' -replace '^\uFEFF', ''

        $matchingLines = $lines[1..($lines.Count-1)] | Where-Object {
            $fields = $_ -split ','
            if ($fields.Count -gt 3) {
                $sender = $fields[3] -replace '[<>"]', ''
                $Domains | ForEach-Object { $sender -match [regex]::Escape($_) } | Where-Object { $_ } | Measure-Object | Select-Object -ExpandProperty Count
            } else { $false }
        }

        if ($matchingLines.Count -gt 0) {
            $filteredCsv = @($header) + $matchingLines
            $cleanedCsv = $filteredCsv | ForEach-Object { $_ -replace '[<>"]', '' }
            Set-Content -Path $OutputCsv -Value $cleanedCsv -Encoding UTF8
            Write-Host "Filtered CSV saved with $($matchingLines.Count) matching rows."
        } else {
            Write-Host "No matching rows found where sender contains any of the specified domains."
        }
    }

    # Rename filtered file based on tenant domain + timestamp
    $timestamp = Get-Date -Format "yyyyMMdd-HHmm"
    $filteredPath = Join-Path $OutputFolder "$($primaryDomain)-DirectSendReport-$timestamp.csv"
    Filter-CsvBySender -InputCsv $downloadPath -OutputCsv $filteredPath -Domains $tenantDomains

    # ----------------------------
    # Summary and optional open
    # ----------------------------
    Write-Host "`nFiltered CSV is available at: $filteredPath"
    $openNow = Read-Host "Would you like to open the filtered CSV now? (Y/N)" -ForegroundColor Cyan
    if ($openNow -in @('Y','y')) { Invoke-Item $filteredPath }

    Write-Host "`n--- Summary of messages per tenant domain ---" -ForegroundColor Cyan
    $filteredCsv = Import-Csv $filteredPath
    foreach ($domain in $tenantDomains) {
        $count = ($filteredCsv | Where-Object { $_.sender_address -match [regex]::Escape($domain) }).Count
        Write-Host "$domain : $count messages"
    }
    $grandTotal = $filteredCsv.Count
    Write-Host "`nGrand Total: $grandTotal messages" -ForegroundColor Green

    # ----------------------------
    # Prompt to disconnect from Exchange Online
    # ----------------------------
    $logout = Read-Host "`nWould you like to log out of Exchange Online? (Y/N)"
    if ($logout -in @('Y','y')) { Disconnect-ExchangeOnline -Confirm:$false; Write-Host "Disconnected from Exchange Online." }
}
