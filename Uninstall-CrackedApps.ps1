# Uninstall-CrackedApps.ps1 - Cong cu Go Cài Dat Hang Loat CAC APP CRACK BEN THU BA

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

function Get-CrackDetection {
    $risks = @()
    $details = @()

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
    
    return [PSCustomObject]@{
        Risks = $risks
        Details = $details
    }
}

function Start-CrackedAppsUninstallProcess {
    Write-Host "`n==========================================================================" -ForegroundColor Cyan
    Write-Host "        CONG CU GO CAI DAT HANG LOAT (BULK UNINSTALL) CAC APP CRACK         " -ForegroundColor White -BackgroundColor DarkYellow
    Write-Host "==========================================================================" -ForegroundColor Cyan
    Write-Host "Thoi gian thuc hien : $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')" -ForegroundColor Gray
    Write-Host "Ten may tinh        : $env:COMPUTERNAME" -ForegroundColor Gray
    Write-Host "--------------------------------------------------------------------------" -ForegroundColor Cyan

    $crackResult = Get-CrackDetection
    $regUninstallPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    $installedApps = Get-ItemProperty -Path $regUninstallPaths -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName }
    
    $crackedAppsFound = @()
    foreach ($risk in $crackResult.Risks) {
        if ($risk -like "*bi Crack*") {
            foreach ($app in $installedApps) {
                if ($risk -like "*$($app.DisplayName)*") {
                    if (-not ($crackedAppsFound | Where-Object { $_.DisplayName -eq $app.DisplayName })) {
                        $crackedAppsFound += $app
                    }
                }
            }
        }
    }
    
    if ($crackedAppsFound.Count -eq 0) {
        Write-Host " -> [THONG BAO] Khong phat hien thay phan mem ben thu ba nao dang bi be khoa tren he thong." -ForegroundColor Green
        Write-Host "==========================================================================" -ForegroundColor Cyan
        return
    }

    Write-Host " -> Phat hien $($crackedAppsFound.Count) phan mem bi be khoa can go hang loat:" -ForegroundColor Yellow
    foreach ($a in $crackedAppsFound) {
        Write-Host "    * $($a.DisplayName)" -ForegroundColor Red
    }
    Write-Host ""

    $uninstalledCount = 0
    foreach ($app in $crackedAppsFound) {
        $unCmd = $app.QuietUninstallString
        $isQuiet = $true
        if ([string]::IsNullOrWhiteSpace($unCmd)) { 
            $unCmd = $app.UninstallString 
            $isQuiet = $false
        }
        
        if (-not [string]::IsNullOrWhiteSpace($unCmd)) {
            Write-Host "[+] Dang go tu dong ($($uninstalledCount + 1)/$($crackedAppsFound.Count)): $($app.DisplayName)..." -ForegroundColor Yellow
            
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
                    if (-not $isQuiet) {
                        if ($exe -like "*unins000.exe*" -or $exe -like "*uninstall.exe*") {
                            if ($argList -notlike "*/SILENT*") {
                                $argList = ($argList + " /SILENT /VERYSILENT /SUPPRESSMSGBOXES /NORESTART").Trim()
                            }
                        } elseif ($exe -like "*msiexec.exe*") {
                            if ($argList -notlike "*/qn*") {
                                $argList = ($argList + " /qn /norestart").Trim()
                            }
                        }
                    }
                    $argList = $argList.Trim()

                    $p = if ([string]::IsNullOrWhiteSpace($argList)) {
                        Write-Host "    Lenh thuc thi: `"$exe`"" -ForegroundColor Gray
                        Start-Process -FilePath $exe -PassThru -ErrorAction SilentlyContinue
                    } else {
                        Write-Host "    Lenh thuc thi: `"$exe`" $argList" -ForegroundColor Gray
                        Start-Process -FilePath $exe -ArgumentList $argList -PassThru -ErrorAction SilentlyContinue
                    }
                    
                    if ($p) { $p.WaitForExit(60000) }
                } else {
                    Write-Host "    Thuc thi lenh CMD: $cleanUnCmd" -ForegroundColor Gray
                    $p = Start-Process cmd.exe -ArgumentList "/c `"$cleanUnCmd`"" -PassThru -ErrorAction SilentlyContinue
                    if ($p) { $p.WaitForExit(60000) }
                }
                
                Write-Host " -> DA GO THANH CONG: $($app.DisplayName)" -ForegroundColor Green
                $uninstalledCount++
            } catch {
                Write-Host " -> Loi khi go $($app.DisplayName): $($_.Exception.Message)" -ForegroundColor Red
            }
        } else {
            Write-Host " -> Bo qua $($app.DisplayName): Khong tim thay UninstallString trong Registry." -ForegroundColor Gray
        }
        Write-Host "--------------------------------------------------------------------------" -ForegroundColor DarkGray
    }
    
    Write-Host "==========================================================================" -ForegroundColor Cyan
    Write-Host "       HOAN TAT GO CAI DAT HANG LOAT $uninstalledCount / $($crackedAppsFound.Count) PHAN MEM BE KHOA           " -ForegroundColor Green
    Write-Host "==========================================================================" -ForegroundColor Cyan
}

Start-CrackedAppsUninstallProcess
