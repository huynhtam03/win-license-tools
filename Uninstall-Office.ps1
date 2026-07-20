# Uninstall-Office.ps1 - Cong cu Go Ca i Dat & Lam Sach Toan Bo Microsoft Office / Project / Visio

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

function Start-OfficeUninstallProcess {
    Write-Host "`n==========================================================================" -ForegroundColor Cyan
    Write-Host "     BAT DAU QUA TRINH GO CAI DAT & LAM SACH OFFICE / PROJECT / VISIO     " -ForegroundColor White -BackgroundColor DarkRed
    Write-Host "==========================================================================" -ForegroundColor Cyan
    Write-Host "Thoi gian thuc hien : $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')" -ForegroundColor Gray
    Write-Host "Ten may tinh        : $env:COMPUTERNAME" -ForegroundColor Gray
    Write-Host "--------------------------------------------------------------------------" -ForegroundColor Cyan

    # 1. DUA VAO KET QUA QUET (Get-OfficeActivation) DE LIET KE SAN PHAM CAN GO
    Write-Host "`n[BUOC 1] Dang doc ket qua quet giay phep (Get-OfficeActivation)..." -ForegroundColor Yellow
    $scanList = Get-OfficeActivation
    
    if ($scanList.Count -gt 0) {
        Write-Host " -> Phat hien $($scanList.Count) phien ban/giay phep Office/Project/Visio tu ket qua quet:" -ForegroundColor Yellow
        foreach ($item in $scanList) {
            Write-Host "    * San pham : $($item.ProductName)" -ForegroundColor Red
            Write-Host "      Trang thai: $($item.Status)"
            Write-Host "      Kenh      : $($item.Channel)"
            Write-Host "      Nguon     : $($item.Source)"
            Write-Host "      --------------------------------------------------" -ForegroundColor DarkGray
        }
    } else {
        Write-Host " -> Khong tim thay giay phep Office/Project/Visio nao trong ket qua quet." -ForegroundColor Gray
    }

    # 2. GO PHAN MEM CAI DAT TRONG REGISTRY & CLICK-TO-RUN
    Write-Host "`n[BUOC 2] Dang tien hanh go bo phan mem cai dat tren he thong..." -ForegroundColor Yellow
    
    # 2.1 Click-To-Run Uninstaller
    $c2rPaths = @(
        "${env:ProgramFiles}\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe",
        "${env:ProgramFiles(x86)}\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe"
    )
    foreach ($c2rPath in $c2rPaths) {
        if (Test-Path $c2rPath) {
            Write-Host " -> Dang chay trinh go sach Click-To-Run ($c2rPath)..." -ForegroundColor Yellow
            try {
                $p = Start-Process -FilePath $c2rPath -ArgumentList "scenario=uninstall forceuninstall=True displaylevel=False" -PassThru -NoNewWindow -ErrorAction SilentlyContinue
                if ($p) {
                    $p.WaitForExit(300000)
                    Write-Host " -> Da thuc thi trinh go Click-To-Run. Ma thoat: $($p.ExitCode)" -ForegroundColor Green
                }
            } catch {
                Write-Host " -> Loi khi chay Click-To-Run: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }

    # 2.2 Registry Uninstall (MSI & Cac goi phan mem)
    $regUninstallPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    $installedApps = Get-ItemProperty -Path $regUninstallPaths -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName }
    $officeRegApps = @()
    foreach ($app in $installedApps) {
        $name = $app.DisplayName
        if ((($name -like "*Microsoft Office*") -or ($name -like "*Microsoft 365*") -or ($name -like "*Microsoft Project*") -or ($name -like "*Microsoft Visio*")) -and 
            ($name -notlike "*Proof*") -and 
            ($name -notlike "*Language*") -and 
            ($name -notlike "*Database Engine*") -and 
            ($name -notlike "*Visual Studio*") -and 
            ($name -notlike "*Viewer*") -and
            ($name -notlike "*Office Web Apps*") -and
            ($name -notlike "*OneDrive*") -and
            ($name -notlike "*Teams*") -and
            ($name -notlike "*Update*") -and
            ($name -notlike "*Desktop App Support*")) {
            $officeRegApps += $app
        }
    }

    foreach ($app in $officeRegApps) {
        $unCmd = $app.QuietUninstallString
        $isQuiet = $true
        if ([string]::IsNullOrWhiteSpace($unCmd)) {
            $unCmd = $app.UninstallString
            $isQuiet = $false
        }
        if (-not [string]::IsNullOrWhiteSpace($unCmd)) {
            Write-Host " -> Dang go goi Registry: $($app.DisplayName)..." -ForegroundColor Yellow
            try {
                $exe = ""
                $rawArgs = ""
                $cleanUnCmd = $unCmd.Trim()
                if ($cleanUnCmd -match '^\s*"([^"]+\.[eE][xX][eE])"\s*(.*)$') {
                    $exe = $matches[1]
                    $rawArgs = $matches[2]
                } elseif ($cleanUnCmd -match '^\s*([^\s]+\.[eE][xX][eE])\s*(.*)$') {
                    $exe = $matches[1]
                    $rawArgs = $matches[2]
                } else {
                    $exe = $cleanUnCmd
                    $rawArgs = ""
                }
                if ($exe -and (Test-Path $exe -ErrorAction SilentlyContinue)) {
                    $argList = if ([string]::IsNullOrWhiteSpace($rawArgs)) { "" } else { $rawArgs.Trim() }
                    if (-not $isQuiet -and $exe -like "*msiexec.exe*") {
                        if ($argList -notlike "*/qn*") { $argList = ($argList + " /qn /norestart").Trim() }
                    }
                    $p = if ([string]::IsNullOrWhiteSpace($argList)) {
                        Start-Process -FilePath $exe -PassThru -ErrorAction SilentlyContinue
                    } else {
                        Start-Process -FilePath $exe -ArgumentList $argList -PassThru -ErrorAction SilentlyContinue
                    }
                    if ($p) { $p.WaitForExit(180000) }
                } else {
                    $p = Start-Process cmd.exe -ArgumentList "/c `"$cleanUnCmd`"" -PassThru -ErrorAction SilentlyContinue
                    if ($p) { $p.WaitForExit(180000) }
                }
                Write-Host "    Da hoan tat go Registry: $($app.DisplayName)" -ForegroundColor Green
            } catch {}
        }
    }

    # 3. GO BO TOAN BO GIAY PHEP SAN PHAM VA MAY CHU KMS QUA OSPP VA WMI
    Write-Host "`n[BUOC 3] Dang xoa sach tat ca Product Key va thong tin KMS cua Office / Project / Visio..." -ForegroundColor Yellow
    
    # 3.1 Xoa qua OSPP
    $officePaths = @(
        "${env:ProgramFiles}\Microsoft Office\Office16",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office16",
        "${env:ProgramFiles}\Microsoft Office\Office15",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office15"
    )
    foreach ($path in $officePaths) {
        $ospp = Join-Path $path "ospp.vbs"
        if (Test-Path $ospp) {
            Write-Host " -> Dang thuc thi OSPP tool ($ospp)..." -ForegroundColor Gray
            try { cscript.exe //NoLogo "$ospp" /remkms 2>&1 | Out-Null } catch {}
            try {
                $dstatusOut = cscript.exe //NoLogo "$ospp" /dstatus 2>&1
                $keysToUninstall = @()
                foreach ($line in $dstatusOut) {
                    if ($line -match "Last 5 characters of installed product key:\s*([A-Za-z0-9]{5})") {
                        $key5 = $matches[1]
                        if ($key5 -and $keysToUninstall -notcontains $key5) {
                            $keysToUninstall += $key5
                        }
                    }
                }
                foreach ($k5 in $keysToUninstall) {
                    Write-Host " -> Dang go Product Key OSPP (5 ky tu cuoi: $k5)..." -ForegroundColor Green
                    cscript.exe //NoLogo "$ospp" /unpkey:$k5 2>&1 | Out-Null
                }
            } catch {}
        }
    }

    # 3.2 Xoa qua WMI SoftwareLicensingProduct (Cho tat ca cac san pham con lai)
    try {
        $wmiProducts = Get-CimInstance -ClassName SoftwareLicensingProduct -ErrorAction SilentlyContinue | Where-Object { 
            ($_.PartialProductKey -or $_.LicenseStatus -eq 1 -or $_.LicenseStatus -eq 5) -and 
            ($_.Name -like "*Office*" -or $_.Name -like "*Project*" -or $_.Name -like "*Visio*" -or $_.Description -like "*Office*" -or $_.Description -like "*Project*" -or $_.Description -like "*Visio*")
        }
        foreach ($p in $wmiProducts) {
            Write-Host " -> Dang go khoa WMI san pham: $($p.Name)..." -ForegroundColor Green
            try {
                Invoke-CimMethod -InputObject $p -MethodName "UninstallProductKey" -ErrorAction SilentlyContinue | Out-Null
            } catch {}
        }
    } catch {}

    Remove-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform" -Name "KeyManagementServiceMachine" -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\OfficeSoftwareProtectionPlatform" -Name "KeyManagementServiceMachine" -ErrorAction SilentlyContinue

    Write-Host "`n==========================================================================" -ForegroundColor Cyan
    Write-Host "HOAN TAT QUA TRINH GO CAI DAT VA LAM SACH TOAN BO OFFICE / PROJECT / VISIO!" -ForegroundColor Green
}

Start-OfficeUninstallProcess
