# Get-ComputerAssetInfo.ps1 - Cong cu Thu Thap Thong Tin Tai San IT (IT Asset Audit & Inventory)

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

function Get-FullComputerAssetInfo {
    $asset = [ordered]@{
        ComputerName      = $env:COMPUTERNAME
        Manufacturer      = "Khong xac dinh"
        Model             = "Khong xac dinh"
        SerialNumber      = "Khong xac dinh"
        MotherboardSerial = "Khong xac dinh"
        OSName            = "Khong xac dinh"
        OSVersion         = "Khong xac dinh"
        OSArchitecture    = "Khong xac dinh"
        OSInstallDate     = "Khong xac dinh"
        CurrentUser       = $env:USERNAME
        UserDomain        = $env:USERDOMAIN
        CPU               = "Khong xac dinh"
        CPUCoresThreads   = "Khong xac dinh"
        RAM_Total_GB      = 0
        RAM_Slots_Used    = "Khong xac dinh"
        Disk_Summary      = "Khong xac dinh"
        Network_Adapter   = "Khong xac dinh"
        IPAddress         = "Khong xac dinh"
        MACAddress        = "Khong xac dinh"
        AuditTimestamp    = (Get-Date -Format 'dd/MM/yyyy HH:mm:ss')
    }

    try {
        $cs = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue
        if ($cs) {
            $asset.Manufacturer = $cs.Manufacturer
            $asset.Model = $cs.Model
        }

        $bios = Get-CimInstance -ClassName Win32_BIOS -ErrorAction SilentlyContinue
        if ($bios) {
            $asset.SerialNumber = $bios.SerialNumber
        }

        $baseboard = Get-CimInstance -ClassName Win32_BaseBoard -ErrorAction SilentlyContinue
        if ($baseboard) {
            $asset.MotherboardSerial = $baseboard.SerialNumber
        }

        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue
        if ($os) {
            $asset.OSName = $os.Caption
            $asset.OSVersion = "$($os.Version) (Build $($os.BuildNumber))"
            $asset.OSArchitecture = $os.OSArchitecture
            $asset.OSInstallDate = $os.InstallDate.ToString('dd/MM/yyyy HH:mm:ss')
        }

        $cpu = Get-CimInstance -ClassName Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($cpu) {
            $asset.CPU = $cpu.Name.Trim()
            $asset.CPUCoresThreads = "$($cpu.NumberOfCores) Cores / $($cpu.NumberOfLogicalProcessors) Threads"
        }

        $ramChips = Get-CimInstance -ClassName Win32_PhysicalMemory -ErrorAction SilentlyContinue
        if ($ramChips) {
            $totalBytes = ($ramChips | Measure-Object -Property Capacity -Sum).Sum
            $asset.RAM_Total_GB = [math]::Round($totalBytes / 1GB, 2)
            $asset.RAM_Slots_Used = "$($ramChips.Count) Khe RAM"
        }

        $disks = Get-CimInstance -ClassName Win32_DiskDrive -ErrorAction SilentlyContinue
        if ($disks) {
            $diskSummaryList = @()
            foreach ($d in $disks) {
                $gb = [math]::Round($d.Size / 1GB, 2)
                $diskSummaryList += "$($d.Model) ($gb GB, SN: $($d.SerialNumber))"
            }
            $asset.Disk_Summary = $diskSummaryList -join " | "
        }

        $net = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object { $_.IPAddress -notlike "127.*" -and $_.IPAddress -notlike "169.254.*" } | Select-Object -First 1
        if ($net) {
            $asset.IPAddress = $net.IPAddress
            $nic = Get-NetAdapter -InterfaceIndex $net.InterfaceIndex -ErrorAction SilentlyContinue
            if ($nic) {
                $asset.Network_Adapter = $nic.InterfaceDescription
                $asset.MACAddress = $nic.MacAddress
            }
        }
    } catch {}

    return [PSCustomObject]$asset
}

# IN KET QUA RA CONSOLE & XUAT FILE
$assetObj = Get-FullComputerAssetInfo

Write-Host "`n==========================================================================" -ForegroundColor Cyan
Write-Host "         THU THAP THONG TIN TAI SAN IT (IT ASSET AUDIT & INVENTORY)       " -ForegroundColor White -BackgroundColor DarkBlue
Write-Host "==========================================================================" -ForegroundColor Cyan
Write-Host "Thoi gian thu thap : $($assetObj.AuditTimestamp)" -ForegroundColor Gray
Write-Host "Ten may tinh       : $($assetObj.ComputerName)" -ForegroundColor Gray
Write-Host "--------------------------------------------------------------------------" -ForegroundColor Cyan

Write-Host "`n[1] THONG TIN NHAN DANG THIET BI (DEVICE IDENTITY):" -ForegroundColor Yellow
Write-Host " - Ten may tinh    : $($assetObj.ComputerName)"
Write-Host " - Nha san xuat    : $($assetObj.Manufacturer)"
Write-Host " - Model thiet bi  : $($assetObj.Model)"
Write-Host " - Serial Number   : $($assetObj.SerialNumber)" -ForegroundColor Green
Write-Host " - Mainboard Serial: $($assetObj.MotherboardSerial)"

Write-Host "`n[2] HE DIEU HANH & NGUOI DUNG (OS & USER):" -ForegroundColor Yellow
Write-Host " - Ten he dieu hanh: $($assetObj.OSName) ($($assetObj.OSArchitecture))"
Write-Host " - Phien ban OS    : $($assetObj.OSVersion)"
Write-Host " - Ngay cai dat OS : $($assetObj.OSInstallDate)"
Write-Host " - Nguoi dung hien tai: $($assetObj.UserDomain)\$($assetObj.CurrentUser)"

Write-Host "`n[3] THONG TIN CAU HINH PHAN CUNG (HARDWARE):" -ForegroundColor Yellow
Write-Host " - Vi xu ly (CPU)  : $($assetObj.CPU) ($($assetObj.CPUCoresThreads))"
Write-Host " - Dung luong RAM  : $($assetObj.RAM_Total_GB) GB ($($assetObj.RAM_Slots_Used))"
Write-Host " - Danh sach O đia : $($assetObj.Disk_Summary)"

Write-Host "`n[4] THONG TIN MANG (NETWORK):" -ForegroundColor Yellow
Write-Host " - Card mang (NIC) : $($assetObj.Network_Adapter)"
Write-Host " - Dia chi IP      : $($assetObj.IPAddress)" -ForegroundColor Cyan
Write-Host " - Dia chi MAC     : $($assetObj.MACAddress)" -ForegroundColor Cyan

# XUAT FILE CSV VA JSON DE NHAP VAO HETHONG QUAN LY TAI SAN
$csvPath = Join-Path $env:USERPROFILE "Desktop\TaiSan_IT_$($assetObj.ComputerName).csv"
$jsonPath = Join-Path $env:USERPROFILE "Desktop\TaiSan_IT_$($assetObj.ComputerName).json"

try {
    $assetObj | Export-Csv -Path $csvPath -NoTypeInformation -Encoding utf8 -Force
    $assetObj | ConvertTo-Json | Set-Content -Path $jsonPath -Encoding utf8 -Force

    Write-Host "`n==========================================================================" -ForegroundColor Cyan
    Write-Host " [THANH CONG] Da xuat du lieu tai san IT ra Desktop:" -ForegroundColor Green
    Write-Host "  * CSV  : $csvPath" -ForegroundColor White
    Write-Host "  * JSON : $jsonPath" -ForegroundColor White
    Write-Host "==========================================================================" -ForegroundColor Cyan
} catch {
    Write-Host "`nLoi khi xuat file tai san: $($_.Exception.Message)" -ForegroundColor Red
}
