# Scan-LicenseStatus.ps1 - Cong cu Quet Ban Quyen & Phat Hien Crack (Standalone Scanner)

# TU DONG NANG QUYEN ADMINISTRATOR
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "========================================================" -ForegroundColor Yellow
    Write-Host "CANH BAO: Cong cu yeu cau quyen Administrator!" -ForegroundColor Yellow
    Write-Host "Dang mo cua so PowerShell moi duoi quyen Administrator..." -ForegroundColor Yellow
    Write-Host "========================================================" -ForegroundColor Yellow
    try {
        Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
        exit
    } catch {
        Write-Host "Loi: Khong the khoi chay duoi quyen Administrator!" -ForegroundColor Red
        exit
    }
}

function Get-WindowsActivation {
    $result = [PSCustomObject]@{
        ProductName = "Chua xac dinh"
        StatusText  = "Chua kiem tra"
        Status      = "Unknown"
        RawStatus   = -1
        Channel     = "Chua xac dinh"
        KMSServer   = "Khong phat hien"
        Details     = ""
    }
    try {
        $win = Get-CimInstance -ClassName SoftwareLicensingProduct -Filter "ApplicationId = '55c92734-d682-4d71-983e-d6ec3f16059f' and PartialProductKey is not null" -ErrorAction SilentlyContinue
        if ($win) {
            $win = $win | Select-Object -First 1
            $result.ProductName = $win.Name
            $result.RawStatus = $win.LicenseStatus
            
            switch ($win.LicenseStatus) {
                0 { $result.StatusText = "Chua kich hoat (Unlicensed)"; $result.Status = "Unlicensed" }
                1 { $result.StatusText = "Da kich hoat (Licensed)"; $result.Status = "Licensed" }
                2 { $result.StatusText = "An han ban dau (OOB Grace)"; $result.Status = "Grace" }
                3 { $result.StatusText = "An han het han (OOT Grace)"; $result.Status = "Grace" }
                4 { $result.StatusText = "An han khong chinh hang (Non-genuine)"; $result.Status = "Warning" }
                5 { $result.StatusText = "Cho kich hoat / Het han (Notification)"; $result.Status = "Notification" }
                6 { $result.StatusText = "An han mo rong (Extended Grace)"; $result.Status = "Grace" }
                default { $result.StatusText = "Khong xac dinh"; $result.Status = "Unknown" }
            }
            
            if ($win.Description -like "*RETAIL*") { $result.Channel = "Retail (Ban le)" }
            elseif ($win.Description -like "*OEM*") { $result.Channel = "OEM (Theo may)" }
            elseif ($win.Description -like "*VOLUME_KMSCLIENT*") { $result.Channel = "Volume KMS (KMS Doanh nghiep)" }
            elseif ($win.Description -like "*VOLUME_MAK*") { $result.Channel = "Volume MAK" }
            else { $result.Channel = "Khac/Tu cau hinh" }
            
            $licService = Get-CimInstance -ClassName SoftwareLicensingService -ErrorAction SilentlyContinue
            if ($licService -and $licService.KeyManagementServiceMachine) {
                $result.KMSServer = "$($licService.KeyManagementServiceMachine):$($licService.KeyManagementServicePort)"
            }
            
            $result.Details = "San pham: $($win.Name)`nPhan nhom: $($result.Channel)`nTrang thai: $($result.StatusText)`nMay chu kich hoat KMS: $($result.KMSServer)"
        } else {
            $result.StatusText = "Khong co san pham hoat dong"
            $result.Details = "Khong the tim thay thong tin giay phep Windows hop le qua WMI."
        }
    } catch {
        $result.StatusText = "Loi truy van"
        $result.Details = "Loi khi lay thong tin ban quyen Windows: $($_.Exception.Message)"
    }
    return $result
}

function Get-OfficeActivation {
    $officeList = @()
    $offices = Get-CimInstance -ClassName SoftwareLicensingProduct -Filter "Description like '%Office%' and PartialProductKey is not null" -ErrorAction SilentlyContinue
    if ($offices) {
        foreach ($off in $offices) {
            $statusText = switch ($off.LicenseStatus) {
                0 { "Chua kich hoat" }
                1 { "Da kich hoat" }
                2 { "An han ban dau" }
                3 { "An han het han" }
                4 { "An han khong chinh hang" }
                5 { "Cho kich hoat / Het han" }
                6 { "An han mo rong" }
                default { "Khong xac dinh" }
            }
            
            $channel = "Khong xac dinh"
            if ($off.Description -like "*RETAIL*") { $channel = "Retail (Ban le)" }
            elseif ($off.Description -like "*VOLUME_KMS*") { $channel = "Volume KMS" }
            elseif ($off.Description -like "*VOLUME_MAK*") { $channel = "Volume MAK" }
            elseif ($off.Description -like "*SUBSCRIPTION*") { $channel = "Subscription (Office 365)" }
            
            $officeList += [PSCustomObject]@{
                ProductName = $off.Name
                Status = $statusText
                Channel = $channel
                KMSServer = $off.KeyManagementServiceMachine
                GraceRemaining = $off.GracePeriodRemaining
                Source = "WMI"
            }
        }
    }
    
    $officePaths = @(
        "${env:ProgramFiles}\Microsoft Office\Office16",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office16",
        "${env:ProgramFiles}\Microsoft Office\Office15",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office15"
    )
    foreach ($path in $officePaths) {
        $ospp = Join-Path $path "ospp.vbs"
        if (Test-Path $ospp) {
            $output = cscript.exe //NoLogo "$ospp" /dstatus 2>&1
            $currentProduct = ""
            $status = ""
            $kms = ""
            $licenseName = ""
            
            foreach ($line in $output) {
                if ($line -like "*PRODUCT ID:*") {
                    if ($currentProduct) {
                        if (-not ($officeList | Where-Object { $_.ProductName -like "*$currentProduct*" -or $_.ProductName -like "*$licenseName*" })) {
                            $officeList += [PSCustomObject]@{
                                ProductName = "$licenseName"
                                Status = $status
                                Channel = "Volume Client (OSPP)"
                                KMSServer = $kms
                                GraceRemaining = "N/A"
                                Source = "OSPP"
                            }
                        }
                    }
                    $currentProduct = ($line -split "PRODUCT ID:")[1].Trim()
                    $status = ""
                    $kms = ""
                    $licenseName = ""
                }
                elseif ($line -like "*LICENSE NAME:*") {
                    $licenseName = ($line -split "LICENSE NAME:")[1].Trim()
                }
                elseif ($line -like "*LICENSE STATUS:*") {
                    $statusLine = ($line -split "LICENSE STATUS:")[1].Trim()
                    if ($statusLine -like "*---LICENSED---*") { $status = "Da kich hoat" }
                    elseif ($statusLine -like "*---NOTIFIED---*") { $status = "Cho kich hoat / Het han" }
                    else { $status = "Chua kich hoat ($statusLine)" }
                }
                elseif ($line -like "*KMS MACHINE NAME:*") {
                    $kms = ($line -split "KMS MACHINE NAME:")[1].Trim()
                }
            }
            if ($currentProduct) {
                if (-not ($officeList | Where-Object { $_.ProductName -like "*$currentProduct*" -or $_.ProductName -like "*$licenseName*" })) {
                    $officeList += [PSCustomObject]@{
                        ProductName = "$licenseName"
                        Status = $status
                        Channel = "Volume Client (OSPP)"
                        KMSServer = $kms
                        GraceRemaining = "N/A"
                        Source = "OSPP"
                    }
                }
            }
        }
    }
    return $officeList
}

function Get-CrackDetection {
    $risks = @()
    $details = @()

    $crackTasks = Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object { 
        $_.TaskName -like "*AutoKMS*" -or 
        $_.TaskName -like "*KMSAuto*" -or 
        $_.TaskName -like "*KMSConnection*" -or 
        $_.TaskName -like "*AutoPico*" -or 
        $_.TaskName -like "*HEU_KMS*" -or
        $_.TaskName -like "*KMS-Activator*"
    }
    if ($crackTasks) {
        foreach ($t in $crackTasks) {
            $risks += "Phat hien Tac vu lap lich be khoa (Scheduled Task): $($t.TaskName)"
            $details += "Tac vu ngam '$($t.TaskName)' dang duoc len lich chay tai '$($t.TaskPath)'."
        }
    }

    $crackServices = Get-Service -ErrorAction SilentlyContinue | Where-Object { 
        $_.Name -like "*AutoKMS*" -or 
        $_.Name -like "*KMSpico*" -or 
        $_.Name -like "*KMSConnection*" -or 
        $_.Name -like "*Service_KMS*" 
    }
    if ($crackServices) {
        foreach ($s in $crackServices) {
            $risks += "Phat hien Dich vu be khoa chay ngam (Service): $($s.DisplayName)"
            $details += "Dich vu '$($s.DisplayName)' ($($s.Name)) dang o trang thai $($s.Status)."
        }
    }

    $hostsPath = "$env:windir\System32\drivers\etc\hosts"
    if (Test-Path $hostsPath) {
        $hostsContent = Get-Content $hostsPath -ErrorAction SilentlyContinue
        $kmsBlocks = $hostsContent | Where-Object { 
            $_ -and (-not $_.StartsWith("#")) -and (
                $_ -like "*kms*" -or 
                $_ -like "*microsoft.com*" -or 
                $_ -like "*activation*" -or
                $_ -like "*adobe*" -or
                $_ -like "*autodesk*"
            )
        }
        if ($kmsBlocks) {
            $risks += "File Hosts bi can thiep de chan xac minh ban quyen"
            $details += "Phat hien $($kmsBlocks.Count) dong cau hinh Redirect/Block trong file Hosts ($hostsPath)."
        }
    }

    $suspiciousPaths = @(
        "$env:windir\AutoKMS\AutoKMS.exe",
        "$env:ProgramFiles\KMSpico\KMSpico.exe",
        "${env:ProgramFiles(x86)}\KMSpico\KMSpico.exe",
        "$env:windir\KMSAuto Net.exe",
        "C:\ProgramData\KMSAutoS\KMSAuto X64.exe"
    )
    foreach ($path in $suspiciousPaths) {
        if (Test-Path $path) {
            $risks += "Phat hien tep tin cong cu be khoa tren o đia: $path"
            $details += "Tep tin thuc thi be khoa nguy hiem ton tai tai '$path'."
        }
    }

    $regUninstallPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    $installedApps = Get-ItemProperty -Path $regUninstallPaths -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName }
    $crackKeywords = @("AutoKMS", "KMSpico", "KMSAuto", "MiniTool", "CCleaner Pro Crack", "Adobe GenP", "Patch", "Crack", "Keygen")
    
    foreach ($app in $installedApps) {
        foreach ($kw in $crackKeywords) {
            if ($app.DisplayName -like "*$kw*") {
                $risks += "Phan mem thuong mai bi Crack/Be khoa: $($app.DisplayName)"
                $details += "Ung dung '$($app.DisplayName)' duoc ghi nhan trong Registry voi InstallLocation: $($app.InstallLocation)."
            }
        }
    }

    $officeRegKms = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform" -Name "KeyManagementServiceMachine" -ErrorAction SilentlyContinue
    if ($officeRegKms -and $officeRegKms.KeyManagementServiceMachine) {
        $kmsMachineOffice = $officeRegKms.KeyManagementServiceMachine
        if ($kmsMachineOffice -eq "127.0.0.1" -or $kmsMachineOffice -eq "localhost" -or $kmsMachineOffice -eq "::1") {
            $risks += "Office cau hinh may chu KMS noi bo (Localhost)"
            $details += "Ban quyen Office tro may chu kich hoat ve chinh may tinh thong qua Registry."
        }
    }

    try {
        $kmsPort = Get-NetTCPConnection -LocalPort 1688 -State Listen -ErrorAction SilentlyContinue
        if ($kmsPort) {
            $risks += "Phat hien cong dich vu KMS (1688) dang mo"
            $details += "Cong TCP 1688 dang lang nghe (Listen) cuc bo. Dau hieu cua KMS gia lap chay an."
        }
    } catch {
        $netstat = netstat -ano 2>&1 | Select-String ":1688\s+.*LISTENING"
        if ($netstat) {
            $risks += "Phat hien cong dich vu KMS (1688) dang mo (netstat)"
            $details += "Cong 1688 dang mo o trang thai LISTENING qua lenh netstat."
        }
    }

    $crackRegs = @("HKLM:\SOFTWARE\KMSAuto", "HKLM:\SOFTWARE\KMSAutoS", "HKLM:\SOFTWARE\KMSpico", "HKCU:\Software\KMSAuto", "HKCU:\Software\KMSAutoS")
    foreach ($reg in $crackRegs) {
        if (Test-Path $reg) {
            $risks += "Phat hien khoa Registry be khoa: $reg"
            $details += "Khoa cau hinh cua KMSAuto/KMSpico duoc phat hien tai '$reg'."
        }
    }
    
    return [PSCustomObject]@{
        Risks = $risks
        Details = $details
    }
}

# CHAY QUET VA IN BAO CAO
Write-Host "`n==========================================================================" -ForegroundColor Cyan
Write-Host "             CONG CU KIEM TRA BAN QUYEN WINDOWS, OFFICE & CRACK            " -ForegroundColor White -BackgroundColor DarkBlue
Write-Host "==========================================================================" -ForegroundColor Cyan
Write-Host "Thoi gian thuc hien : $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')" -ForegroundColor Gray
Write-Host "Ten may tinh        : $env:COMPUTERNAME" -ForegroundColor Gray
Write-Host "--------------------------------------------------------------------------" -ForegroundColor Cyan

Write-Host "`n[1] DANG QUET BAN QUYEN WINDOWS..." -ForegroundColor Yellow
$win = Get-WindowsActivation
Write-Host " - Phien ban    : $($win.ProductName)"
Write-Host " - Phan loai    : $($win.Channel)"
Write-Host " - Trang thai   : $($win.StatusText)"
if ($win.KMSServer -and $win.KMSServer -ne "Khong phat hien") {
    Write-Host " - May chu KMS  : $($win.KMSServer)" -ForegroundColor Red
}

Write-Host "`n[2] DANG QUET BAN QUYEN MICROSOFT OFFICE / PROJECT / VISIO..." -ForegroundColor Yellow
$officeList = Get-OfficeActivation
if ($officeList.Count -eq 0) {
    Write-Host " - Khong tim thay san pham Office / Project / Visio nao dang ky khoa ban quyen." -ForegroundColor Gray
} else {
    foreach ($off in $officeList) {
        Write-Host " - Ten san pham : $($off.ProductName)"
        Write-Host "   Trang thai   : $($off.Status)"
        Write-Host "   Phan nhom    : $($off.Channel)"
        if ($off.KMSServer) {
            Write-Host "   May chu KMS  : $($off.KMSServer)" -ForegroundColor Red
        }
        Write-Host "   ----------------------------------------------------------------------" -ForegroundColor DarkGray
    }
}

Write-Host "`n[3] DANG QUET CAC DAU VET CRACK / BE KHOA..." -ForegroundColor Yellow
$crack = Get-CrackDetection
if ($crack.Risks.Count -eq 0) {
    Write-Host " -> [AN TOAN] Khong phat hien thay bat ky dau hieu phan mem crack/be khoa nao." -ForegroundColor Green
} else {
    Write-Host " -> [CANH BAO] Phat hien $($crack.Risks.Count) dau hieu crack/be khoa he thong:" -ForegroundColor Red
    for ($i = 0; $i -lt $crack.Risks.Count; $i++) {
        Write-Host "   * Dau hieu: $($crack.Risks[$i])" -ForegroundColor Red
        Write-Host "     Mo ta   : $($crack.Details[$i])" -ForegroundColor Gray
    }
}
Write-Host "`n==========================================================================" -ForegroundColor Cyan
Write-Host "Hoan tat qua trinh quet ban quyen va kiem tra he thong!" -ForegroundColor Green
