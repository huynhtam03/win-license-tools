# Clean-SystemTemp.ps1 - Cong cu Don Dẹp He Thong, Xoa Cache Temp & Reset Mang cho IT Helpdesk

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

function Start-SystemAndNetworkCleanup {
    Write-Host "`n==========================================================================" -ForegroundColor Cyan
    Write-Host "     CONG CU DON DEP HE THONG, XOA TEMP CACHE & RESET MANG (IT HELPDESK)  " -ForegroundColor White -BackgroundColor DarkGreen
    Write-Host "==========================================================================" -ForegroundColor Cyan
    Write-Host "Thoi gian thuc hien : $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')" -ForegroundColor Gray
    Write-Host "Ten may tinh        : $env:COMPUTERNAME" -ForegroundColor Gray
    Write-Host "--------------------------------------------------------------------------" -ForegroundColor Cyan

    # 1. DON DEP FILE TEMP & PREFETCH
    Write-Host "`n[BUOC 1] Dang xoa cac tep tin rac va cache trong he thong..." -ForegroundColor Yellow
    $tempFolders = @(
        $env:TEMP,
        "$env:windir\Temp",
        "$env:windir\Prefetch",
        "$env:windir\SoftwareDistribution\Download"
    )

    $deletedFiles = 0
    $deletedBytes = 0

    foreach ($folder in $tempFolders) {
        if (Test-Path $folder) {
            Write-Host " -> Dang don dep thu muc: $folder" -ForegroundColor Gray
            $files = Get-ChildItem -Path $folder -Recurse -File -ErrorAction SilentlyContinue
            foreach ($f in $files) {
                try {
                    $size = $f.Length
                    Remove-Item -Path $f.FullName -Force -ErrorAction SilentlyContinue
                    $deletedFiles++
                    $deletedBytes += $size
                } catch {}
            }
        }
    }

    $freedMB = [math]::Round($deletedBytes / 1MB, 2)
    Write-Host " -> DA XOA THANH CONG: $deletedFiles tep tin rac (Giai phong ~$freedMB MB dung luong o C)." -ForegroundColor Green

    # 2. XOA CACHE DNS & RESET MANG
    Write-Host "`n[BUOC 2] Dang lam sach cache DNS va thiet lap lai he thong mang..." -ForegroundColor Yellow
    try {
        $flushRes = ipconfig /flushdns 2>&1
        Write-Host " -> Xoa Cache DNS (ipconfig /flushdns): $($flushRes -join ' ')" -ForegroundColor Green
    } catch {}

    try {
        Clear-DnsClientCache -ErrorAction SilentlyContinue
        Write-Host " -> Clear-DnsClientCache: Thanh cong." -ForegroundColor Green
    } catch {}

    try {
        $winsockRes = netsh winsock reset 2>&1
        Write-Host " -> Reset Winsock catalog (netsh winsock reset): Thanh cong." -ForegroundColor Green
    } catch {}

    try {
        $ipResetRes = netsh int ip reset 2>&1
        Write-Host " -> Reset IP stack (netsh int ip reset): Thanh cong." -ForegroundColor Green
    } catch {}

    Write-Host "`n==========================================================================" -ForegroundColor Cyan
    Write-Host "HOAN TAT QUA TRINH DON DEP HE THONG VA RE-SET CACHE MANG!" -ForegroundColor Green
    Write-Host "He thong da duoc toi uu hoa va lam sach cache." -ForegroundColor Gray
}

Start-SystemAndNetworkCleanup
