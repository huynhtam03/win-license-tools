# Remove-KMSActivation.ps1 - Cong cu Go Bo Ban Quyen KMS Lau & Lam Sach He Thong

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

function Start-KMSRemovalProcess {
    Write-Host "`n==========================================================================" -ForegroundColor Cyan
    Write-Host "        CONG CU GO BO BAN QUYEN KMS LAU VA LAM SACH HE THONG               " -ForegroundColor White -BackgroundColor DarkRed
    Write-Host "==========================================================================" -ForegroundColor Cyan
    Write-Host "Thoi gian thuc hien : $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')" -ForegroundColor Gray
    Write-Host "Ten may tinh        : $env:COMPUTERNAME" -ForegroundColor Gray
    Write-Host "--------------------------------------------------------------------------" -ForegroundColor Cyan

    # Buoc 1: Xoa cau hinh may chu KMS Windows (slmgr /ckms)
    Write-Host "`n[BUOC 1] Dang go bo cau hinh may chu KMS Windows (Clear KMS Server)..." -ForegroundColor Yellow
    try {
        $ckmsResult = cscript.exe //NoLogo "$env:windir\System32\slmgr.vbs" /ckms 2>&1
        Write-Host " -> Ket qua slmgr /ckms: $($ckmsResult -join ' ')" -ForegroundColor Green
    } catch {
        Write-Host " -> Loi khi chay slmgr /ckms: $($_.Exception.Message)" -ForegroundColor Red
    }

    # Buoc 2: Go bo kho a san pham KMS Windows (slmgr /upk)
    Write-Host "`n[BUOC 2] Dang go bo Product Key KMS Windows (Uninstall Product Key)..." -ForegroundColor Yellow
    try {
        $upkResult = cscript.exe //NoLogo "$env:windir\System32\slmgr.vbs" /upk 2>&1
        Write-Host " -> Ket qua slmgr /upk: $($upkResult -join ' ')" -ForegroundColor Green
    } catch {
        Write-Host " -> Loi khi chay slmgr /upk: $($_.Exception.Message)" -ForegroundColor Red
    }

    # Buoc 3: Xoa thong tin Product Key khoi Registry (slmgr /cpky)
    Write-Host "`n[BUOC 3] Dang xoa thong tin Product Key Windows khoi Registry..." -ForegroundColor Yellow
    try {
        $cpkyResult = cscript.exe //NoLogo "$env:windir\System32\slmgr.vbs" /cpky 2>&1
        Write-Host " -> Ket qua slmgr /cpky: $($cpkyResult -join ' ')" -ForegroundColor Green
    } catch {
        Write-Host " -> Loi khi chay slmgr /cpky: $($_.Exception.Message)" -ForegroundColor Red
    }

    # Buoc 4: Lam sach tep tin Hosts
    Write-Host "`n[BUOC 4] Dang kiem tra va lam sach tep tin Hosts..." -ForegroundColor Yellow
    $hostsPath = "$env:windir\System32\drivers\etc\hosts"
    if (Test-Path $hostsPath) {
        try {
            $lines = Get-Content $hostsPath -ErrorAction SilentlyContinue
            $cleanLines = @()
            $removedCount = 0
            
            foreach ($line in $lines) {
                $clean = $line.Trim()
                if ($clean -and -not $clean.StartsWith("#") -and ($clean -like "*microsoft*" -or $clean -like "*activation*" -or $clean -like "*kms*")) {
                    Write-Host " -> Da go bo dong chan: $clean" -ForegroundColor Red
                    $removedCount++
                } else {
                    $cleanLines += $line
                }
            }
            if ($removedCount -gt 0) {
                $cleanLines | Set-Content $hostsPath -Encoding utf8 -Force
                Write-Host " -> Da lam sach $removedCount dong cau hinh chan trong file Hosts!" -ForegroundColor Green
            } else {
                Write-Host " -> Tep Hosts khong chua cau hinh chan may chu Microsoft." -ForegroundColor Gray
            }
        } catch {
            Write-Host " -> Loi khi chinh sua file Hosts: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # Buoc 5: Tim va go bo Tac vu chay ngam be khoa (Scheduled Tasks)
    Write-Host "`n[BUOC 5] Dang don dep cac tac vu be khoa tu dong (Scheduled Tasks)..." -ForegroundColor Yellow
    $suspiciousTaskNames = @("*AutoKMS*", "*KMSAuto*", "*KMSConnection*", "*AutoPico*", "*HEU_KMS*", "*KMS-Activator*")
    $removedTasks = 0
    foreach ($pattern in $suspiciousTaskNames) {
        $tasks = Get-ScheduledTask -TaskName $pattern -ErrorAction SilentlyContinue
        foreach ($t in $tasks) {
            try {
                Unregister-ScheduledTask -TaskName $t.TaskName -TaskPath $t.TaskPath -Confirm:$false -ErrorAction SilentlyContinue
                Write-Host " -> Da go bo tac vu be khoa: $($t.TaskName) ($($t.TaskPath))" -ForegroundColor Green
                $removedTasks++
            } catch {}
        }
    }
    if ($removedTasks -eq 0) {
        Write-Host " -> Khong phat hien tac vu chay ngam be khoa nao." -ForegroundColor Gray
    }

    # Buoc 6: Tim va dung/go bo Dich vu KMS lau (Services)
    Write-Host "`n[BUOC 6] Dang don dep cac dich vu KMS lau (Services)..." -ForegroundColor Yellow
    $suspiciousServices = @("Service_KMS", "KMSpico Service", "KMSConnectionMonitor", "AutoKMS")
    $removedSrvs = 0
    foreach ($srvName in $suspiciousServices) {
        $srv = Get-Service -Name $srvName -ErrorAction SilentlyContinue
        if ($srv) {
            try {
                Stop-Service -Name $srvName -Force -ErrorAction SilentlyContinue
                sc.exe delete $srvName | Out-Null
                Write-Host " -> Da dung va go bo dich vu KMS lau: $($srv.DisplayName) ($srvName)" -ForegroundColor Green
                $removedSrvs++
            } catch {}
        }
    }
    if ($removedSrvs -eq 0) {
        Write-Host " -> Khong phat hien dich vu KMS lau chay ngam nao." -ForegroundColor Gray
    }

    # Buoc 7: Go bo ban quyen Office / Project / Visio KMS qua ospp.vbs va WMI
    Write-Host "`n[BUOC 7] Dang go bo ban quyen KMS cua Office/Project/Visio qua ospp.vbs va WMI..." -ForegroundColor Yellow
    
    $officePaths = @(
        "${env:ProgramFiles}\Microsoft Office\Office16",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office16",
        "${env:ProgramFiles}\Microsoft Office\Office15",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office15"
    )
    
    $osppFound = $false
    foreach ($path in $officePaths) {
        $ospp = Join-Path $path "ospp.vbs"
        if (Test-Path $ospp) {
            $osppFound = $true
            Write-Host " -> Phat hien OSPP tool tai: $ospp" -ForegroundColor Gray
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
    
    if (-not $osppFound) {
        Write-Host " -> Khong tim thay tep tin ospp.vbs tren cac duong dan mac dinh." -ForegroundColor Gray
    }

    # Go khoa san pham qua WMI va xoa Registry KMS
    Write-Host " -> Dang quet va go toan bo khoa san pham Office/Project/Visio trong WMI..." -ForegroundColor Yellow
    try {
        $wmiProds = Get-CimInstance -ClassName SoftwareLicensingProduct -ErrorAction SilentlyContinue | Where-Object { 
            ($_.PartialProductKey -or $_.LicenseStatus -eq 1 -or $_.LicenseStatus -eq 5) -and 
            ($_.Name -like "*Office*" -or $_.Name -like "*Project*" -or $_.Name -like "*Visio*" -or $_.Description -like "*Office*" -or $_.Description -like "*Project*" -or $_.Description -like "*Visio*")
        }
        foreach ($wp in $wmiProds) {
            Write-Host " -> Dang go khoa WMI san pham: $($wp.Name)..." -ForegroundColor Green
            try { Invoke-CimMethod -InputObject $wp -MethodName "UninstallProductKey" -ErrorAction SilentlyContinue | Out-Null } catch {}
        }
    } catch {}

    Remove-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform" -Name "KeyManagementServiceMachine" -ErrorAction SilentlyContinue
    Remove-ItemProperty -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\OfficeSoftwareProtectionPlatform" -Name "KeyManagementServiceMachine" -ErrorAction SilentlyContinue

    Write-Host "`n==========================================================================" -ForegroundColor Cyan
    Write-Host "HOAN TAT QUA TRINH GO BO BAN QUYEN KMS LAU VA LAM SACH HE THONG!" -ForegroundColor Green
    Write-Host "Trang thai Windows/Office hien tai da duoc lam sach thong tin KMS lau." -ForegroundColor Gray
}

Start-KMSRemovalProcess
