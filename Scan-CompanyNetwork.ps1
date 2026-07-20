<#
    Scan-CompanyNetwork.ps1
    Công cụ rà soát bản quyền và crack trên toàn bộ máy tính công ty từ xa.
    Yêu cầu: Chạy dưới quyền Domain Admin hoặc tài khoản có quyền WinRM (PowerShell Remoting) trên các máy trạm.
#>

param (
    # Danh sách tên máy tính hoặc địa chỉ IP cần quét.
    # Mặc định quét máy hiện tại. Để quét toàn mạng, có thể truyền danh sách: @("PC01", "PC02", "192.168.1.50")
    # Hoặc tự động lấy từ Active Directory: (Get-ADComputer -Filter *).Name
    [string[]]$ComputerNames = @("localhost"),
    
    # Đường dẫn xuất tệp báo cáo tổng hợp
    [string]$OutputFile = "BaoCao_HeThong_CongTy.csv"
)

# Khối lệnh kiểm tra sẽ được đẩy sang thực thi trên các máy trạm
$ScriptBlock = {
    # 1. Quét bản quyền Windows
    $winStatus = "Không xác định"
    $winChannel = "Không xác định"
    $winKms = ""
    try {
        $win = Get-CimInstance -ClassName SoftwareLicensingProduct -Filter "ApplicationId = '55c92734-d682-4d71-983e-d6ec3f16059f' and PartialProductKey is not null" -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($win) {
            $winStatus = switch ($win.LicenseStatus) {
                0 { "Chưa kích hoạt (Unlicensed)" }
                1 { "Đã kích hoạt (Licensed)" }
                2 { "Ân hạn ban đầu" }
                3 { "Ân hạn hết hạn" }
                4 { "Ân hạn không chính hãng" }
                5 { "Chờ kích hoạt / Hết hạn" }
                6 { "Ân hạn mở rộng" }
                default { "Không xác định" }
            }
            if ($win.Description -like "*RETAIL*") { $winChannel = "Retail" }
            elseif ($win.Description -like "*OEM*") { $winChannel = "OEM" }
            elseif ($win.Description -like "*VOLUME_KMSCLIENT*") { $winChannel = "Volume KMS" }
            elseif ($win.Description -like "*VOLUME_MAK*") { $winChannel = "Volume MAK" }
            
            $licService = Get-CimInstance -ClassName SoftwareLicensingService -ErrorAction SilentlyContinue
            if ($licService -and $licService.KeyManagementServiceMachine) {
                $winKms = $licService.KeyManagementServiceMachine
            }
        }
    } catch {}

    # 2. Quét bản quyền Office
    $officeStatus = "Không phát hiện"
    try {
        $offices = Get-CimInstance -ClassName SoftwareLicensingProduct -Filter "Description like '%Office%' and PartialProductKey is not null" -ErrorAction SilentlyContinue
        if ($offices) {
            $officeStatus = ($offices | ForEach-Object { 
                $status = if ($_.LicenseStatus -eq 1) { "Đã kích hoạt" } else { "Chưa kích hoạt" }
                "$($_.Name): $status" 
            }) -join " | "
        }
    } catch {}

    # 3. Quét các dấu vết phần mềm bẻ khóa (Crack)
    $crackRisks = @()
    
    # A. Cổng KMS 1688
    try {
        if (Get-NetTCPConnection -LocalPort 1688 -State Listen -ErrorAction SilentlyContinue) {
            $crackRisks += "Cổng KMS 1688 đang mở"
        }
    } catch {
        if (netstat -ano | Select-String ":1688\s+.*LISTENING") { $crackRisks += "Cổng KMS 1688 đang mở" }
    }
    
    # B. Dấu vết KMS Tools
    $kmsFolders = @("${env:ProgramFiles}\KMSpico", "${env:ProgramData}\KMSAuto", "C:\Windows\Setup\Scripts")
    foreach ($f in $kmsFolders) { if (Test-Path $f) { $crackRisks += "Thư mục KMS lậu: $f" } }

    # C. File hosts chặn Microsoft hoặc các hãng phần mềm khác
    $hostsPath = "$env:windir\System32\drivers\etc\hosts"
    if (Test-Path $hostsPath) {
        $hostsLines = Get-Content $hostsPath -ErrorAction SilentlyContinue
        $blockedApps = @()
        $thirdPartyDomains = @(
            "minitool.com", "adobe.com", "autodesk.com", "corel.com", 
            "registeridm.com", "internetdownloadmanager.com", "sketchup.com"
        )
        foreach ($line in $hostsLines) {
            $clean = $line.Trim()
            if ($clean -and -not $clean.StartsWith("#")) {
                foreach ($domain in $thirdPartyDomains) {
                    if ($clean -like "*$domain*" -and $clean -match "(127\.0\.0\.1|0\.0\.0\.0)") {
                        if ($blockedApps -notcontains $domain) { $blockedApps += $domain }
                    }
                }
            }
        }
        if ($blockedApps.Count -gt 0) { $crackRisks += "Hosts chặn domain: $($blockedApps -join ', ')" }
    }
    
    # D. Kiểm tra Chữ ký số Authenticode (MiniTool, IDM...)
    $targetKeywords = @("MiniTool", "Adobe", "IDM", "Internet Download Manager")
    $targetFilesMap = @{
        "MiniTool" = @("partitionwizard.dll", "partitionwizard.exe")
        "IDM" = @("idm.core.dll", "IDMan.exe")
        "Internet Download Manager" = @("idm.core.dll", "IDMan.exe")
        "Adobe" = @("amtlib.dll")
    }
    $regPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    $apps = Get-ItemProperty -Path $regPaths -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName }
    foreach ($app in $apps) {
        foreach ($kw in $targetKeywords) {
            if ($app.DisplayName -like "*$kw*") {
                $installDir = $app.InstallLocation
                if (-not $installDir -and $app.UninstallString -match '^["]?([^"]+\.[eE][xX][eE])["]?') {
                    $installDir = Split-Path $matches[1]
                }
                if ($installDir -and (Test-Path $installDir)) {
                    foreach ($file in $targetFilesMap[$kw]) {
                        $found = Get-ChildItem -Path $installDir -Filter $file -Recurse -Depth 2 -ErrorAction SilentlyContinue
                        foreach ($f in $found) {
                            $sig = Get-AuthenticodeSignature -FilePath $f.FullName -ErrorAction SilentlyContinue
                            if ($sig -and $sig.Status -ne 'Valid') {
                                $crackRisks += "$($app.DisplayName) mất chữ ký số ($file)"
                            }
                        }
                    }
                }
            }
        }
    }

    # Trả về đối tượng báo cáo của máy này
    [PSCustomObject]@{
        ComputerName   = $env:COMPUTERNAME
        WindowsLicense = $winStatus
        WindowsChannel = $winChannel
        WindowsKMS     = $winKms
        OfficeStatus   = $officeStatus
        CrackRisks     = ($crackRisks -join "; ")
        ScanTime       = (Get-Date -Format 'dd/MM/yyyy HH:mm:ss')
    }
}

# Thử thực thi truy vấn
Write-Host "==========================================================================" -ForegroundColor Cyan
Write-Host "     BẮT ĐẦU RÀ SOÁT BẢN QUYỀN & CRACK TRÊN CÁC MÁY TRẠM NỘI BỘ MẠNG      " -ForegroundColor White -BackgroundColor DarkBlue
Write-Host "==========================================================================" -ForegroundColor Cyan
Write-Host "Đang quét các máy tính: $($ComputerNames -join ', ')..." -ForegroundColor Yellow

$Results = @()
if ($ComputerNames.Count -eq 1 -and $ComputerNames[0] -eq "localhost") {
    $Results += Invoke-Command -ScriptBlock $ScriptBlock
} else {
    $Results += Invoke-Command -ComputerName $ComputerNames -ScriptBlock $ScriptBlock -ErrorAction SilentlyContinue
}

# Xuất báo cáo CSV dạng UTF-8
if ($Results.Count -gt 0) {
    # Thêm BOM cho UTF-8 để Excel mở không lỗi font tiếng Việt
    $Utf8NoBOM = New-Object System.Text.UTF8Encoding $false
    $Results | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding utf8
    
    Write-Host "`nQuá trình rà soát hoàn tất!" -ForegroundColor Green
    Write-Host "Đã xuất báo cáo tổng hợp tại: $OutputFile" -ForegroundColor Green
    Write-Host "Mở file CSV bằng Excel để lọc nhanh các máy vi phạm (cột CrackRisks)." -ForegroundColor Gray
} else {
    Write-Host "Không thể kết nối hoặc quét bất kỳ máy trạm nào." -ForegroundColor Red
}
Write-Host "==========================================================================" -ForegroundColor Cyan
