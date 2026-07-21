# Get-ComputerAssetInfo.ps1 - Cong cu Thu Thap Thong Tin Tai San IT Chi Tiet (Deep IT Asset Audit)

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
        ComputerName            = $env:COMPUTERNAME
        SystemManufacturer      = "Khong xac dinh"
        SystemModel             = "Khong xac dinh"
        SystemUUID              = "Khong xac dinh"
        BIOSSerialNumber        = "Khong xac dinh"
        BIOSVersion             = "Khong xac dinh"
        MotherboardManufacturer = "Khong xac dinh"
        MotherboardProduct      = "Khong xac dinh"
        MotherboardSerial       = "Khong xac dinh"
        
        OSName                  = "Khong xac dinh"
        OSVersion               = "Khong xac dinh"
        OSArchitecture          = "Khong xac dinh"
        OSInstallDate           = "Khong xac dinh"
        CurrentUser             = $env:USERNAME
        UserDomain              = $env:USERDOMAIN
        
        CPU                     = "Khong xac dinh"
        CPUCoresThreads         = "Khong xac dinh"
        GPU                     = "Khong xac dinh"
        GPUVRAM_GB              = "Khong xac dinh"
        
        TotalRAM_GB             = 0
        RAMSlotsCount           = 0
        RAMSticksDetail         = "Khong xac dinh"
        
        TotalStorageCapacityGB  = 0
        TotalFreeStorageGB      = 0
        PhysicalDisksDetail     = "Khong xac dinh"
        LogicalDrivesDetail     = "Khong xac dinh"
        
        NetworkAdapter          = "Khong xac dinh"
        IPAddress               = "Khong xac dinh"
        MACAddress              = "Khong xac dinh"
        AuditTimestamp          = (Get-Date -Format 'dd/MM/yyyy HH:mm:ss')
    }

    try {
        # 1. System & Motherboard
        $cs = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue
        if ($cs) {
            $asset.SystemManufacturer = $cs.Manufacturer
            $asset.SystemModel = $cs.Model
        }

        $csp = Get-CimInstance -ClassName Win32_ComputerSystemProduct -ErrorAction SilentlyContinue
        if ($csp) {
            $asset.SystemUUID = $csp.UUID
        }

        $bios = Get-CimInstance -ClassName Win32_BIOS -ErrorAction SilentlyContinue
        if ($bios) {
            $asset.BIOSSerialNumber = $bios.SerialNumber
            $asset.BIOSVersion = "$($bios.Manufacturer) $($bios.SMBIOSBIOSVersion)"
        }

        $baseboard = Get-CimInstance -ClassName Win32_BaseBoard -ErrorAction SilentlyContinue
        if ($baseboard) {
            $asset.MotherboardManufacturer = $baseboard.Manufacturer
            $asset.MotherboardProduct = $baseboard.Product
            $asset.MotherboardSerial = $baseboard.SerialNumber
        }

        # 2. Operating System
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction SilentlyContinue
        if ($os) {
            $asset.OSName = $os.Caption
            $asset.OSVersion = "$($os.Version) (Build $($os.BuildNumber))"
            $asset.OSArchitecture = $os.OSArchitecture
            $asset.OSInstallDate = $os.InstallDate.ToString('dd/MM/yyyy HH:mm:ss')
        }

        # 3. CPU & GPU
        $cpu = Get-CimInstance -ClassName Win32_Processor -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($cpu) {
            $asset.CPU = $cpu.Name.Trim()
            $asset.CPUCoresThreads = "$($cpu.NumberOfCores) Cores / $($cpu.NumberOfLogicalProcessors) Threads"
        }

        $gpus = Get-CimInstance -ClassName Win32_VideoController -ErrorAction SilentlyContinue
        if ($gpus) {
            $gpuNames = @()
            $vramGBs = @()
            foreach ($g in $gpus) {
                $gpuNames += $g.Name
                if ($g.AdapterRAM) {
                    $vramGBs += "$([math]::Round($g.AdapterRAM / 1GB, 2)) GB"
                }
            }
            $asset.GPU = $gpuNames -join " | "
            $asset.GPUVRAM_GB = if ($vramGBs.Count -gt 0) { $vramGBs -join " | " } else { "N/A" }
        }

        # 4. Memory RAM Details
        $ramChips = Get-CimInstance -ClassName Win32_PhysicalMemory -ErrorAction SilentlyContinue
        if ($ramChips) {
            $totalBytes = ($ramChips | Measure-Object -Property Capacity -Sum).Sum
            $asset.TotalRAM_GB = [math]::Round($totalBytes / 1GB, 2)
            $asset.RAMSlotsCount = $ramChips.Count
            
            $ramDetailsList = @()
            foreach ($r in $ramChips) {
                $chipGB = [math]::Round($r.Capacity / 1GB, 2)
                $speed = if ($r.Speed) { "$($r.Speed)MHz" } else { "" }
                $mfg = if ($r.Manufacturer) { $r.Manufacturer.Trim() } else { "" }
                $bank = if ($r.DeviceLocator) { $r.DeviceLocator } else { "Slot" }
                $ramDetailsList += "[$($bank): $chipGB GB $speed $mfg]"
            }
            $asset.RAMSticksDetail = $ramDetailsList -join " | "
        }

        # 5. Storage (Physical Disks & Logical Partitions)
        $physDisks = Get-CimInstance -ClassName Win32_DiskDrive -ErrorAction SilentlyContinue
        if ($physDisks) {
            $pDiskList = @()
            $totalDiskBytes = 0
            foreach ($d in $physDisks) {
                $gb = [math]::Round($d.Size / 1GB, 2)
                $totalDiskBytes += $d.Size
                $pDiskList += "$($d.Model) ($gb GB - SN: $($d.SerialNumber))"
            }
            $asset.TotalStorageCapacityGB = [math]::Round($totalDiskBytes / 1GB, 2)
            $asset.PhysicalDisksDetail = $pDiskList -join " | "
        }

        $logDrives = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DriveType = 3" -ErrorAction SilentlyContinue
        if ($logDrives) {
            $lDriveList = @()
            $totalFreeBytes = 0
            foreach ($ld in $logDrives) {
                $totalGB = [math]::Round($ld.Size / 1GB, 2)
                $freeGB = [math]::Round($ld.FreeSpace / 1GB, 2)
                $usedGB = [math]::Round(($ld.Size - $ld.FreeSpace) / 1GB, 2)
                $totalFreeBytes += $ld.FreeSpace
                $percentFree = [math]::Round(($ld.FreeSpace / $ld.Size) * 100, 1)
                $lDriveList += "[$($ld.DeviceID) Tong: $totalGB GB | Dung: $usedGB GB | Trong: $freeGB GB ($percentFree% trong)]"
            }
            $asset.TotalFreeStorageGB = [math]::Round($totalFreeBytes / 1GB, 2)
            $asset.LogicalDrivesDetail = $lDriveList -join " | "
        }

        # 6. Network Info
        $net = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue | Where-Object { $_.IPAddress -notlike "127.*" -and $_.IPAddress -notlike "169.254.*" } | Select-Object -First 1
        if ($net) {
            $asset.IPAddress = $net.IPAddress
            $nic = Get-NetAdapter -InterfaceIndex $net.InterfaceIndex -ErrorAction SilentlyContinue
            if ($nic) {
                $asset.NetworkAdapter = $nic.InterfaceDescription
                $asset.MACAddress = $nic.MacAddress
            }
        }
    } catch {}

    return [PSCustomObject]$asset
}

# IN KET QUA RA CONSOLE & XUAT FILE
$assetObj = Get-FullComputerAssetInfo

Write-Host "`n==========================================================================" -ForegroundColor Cyan
Write-Host "     THU THAP THONG TIN TAI SAN IT CHI TIET (DEEP IT ASSET AUDIT)         " -ForegroundColor White -BackgroundColor DarkBlue
Write-Host "==========================================================================" -ForegroundColor Cyan
Write-Host "Thoi gian thu thap : $($assetObj.AuditTimestamp)" -ForegroundColor Gray
Write-Host "Ten may tinh       : $($assetObj.ComputerName)" -ForegroundColor Gray
Write-Host "--------------------------------------------------------------------------" -ForegroundColor Cyan

Write-Host "`n[1] HE THONG & BO MACH CHU (SYSTEM & MAINBOARD):" -ForegroundColor Yellow
Write-Host " - Ten may tinh     : $($assetObj.ComputerName)"
Write-Host " - Nha san xuat     : $($assetObj.SystemManufacturer)"
Write-Host " - Model thiet bi   : $($assetObj.SystemModel)"
Write-Host " - Serial (Service): $($assetObj.BIOSSerialNumber)" -ForegroundColor Green
Write-Host " - System UUID      : $($assetObj.SystemUUID)"
Write-Host " - Mainboard Hang   : $($assetObj.MotherboardManufacturer)"
Write-Host " - Mainboard Model  : $($assetObj.MotherboardProduct)" -ForegroundColor Green
Write-Host " - Mainboard Serial : $($assetObj.MotherboardSerial)"
Write-Host " - BIOS Version     : $($assetObj.BIOSVersion)"

Write-Host "`n[2] HE DIEU HANH & NGUOI DUNG (OS & USER):" -ForegroundColor Yellow
Write-Host " - Ten he dieu hanh : $($assetObj.OSName) ($($assetObj.OSArchitecture))"
Write-Host " - Phien ban OS     : $($assetObj.OSVersion)"
Write-Host " - Ngay cai dat OS  : $($assetObj.OSInstallDate)"
Write-Host " - Nguoi dung       : $($assetObj.UserDomain)\$($assetObj.CurrentUser)"

Write-Host "`n[3] BO NHO RAM & CARD DO HOA (MEMORY & GPU):" -ForegroundColor Yellow
Write-Host " - Tong RAM he thong: $($assetObj.TotalRAM_GB) GB ($($assetObj.RAMSlotsCount) thanh RAM)" -ForegroundColor Green
Write-Host " - Chi tiet thanh RAM: $($assetObj.RAMSticksDetail)"
Write-Host " - Card do hoa (GPU): $($assetObj.GPU) (VRAM: $($assetObj.GPUVRAM_GB))"

Write-Host "`n[4] CHI TIET OU DIA & PHAN VUNG (STORAGE & DRIVES):" -ForegroundColor Yellow
Write-Host " - Tong dung luong o: $($assetObj.TotalStorageCapacityGB) GB (Trong: $($assetObj.TotalFreeStorageGB) GB)" -ForegroundColor Green
Write-Host " - O đia vat ly      : $($assetObj.PhysicalDisksDetail)"
Write-Host " - Phan vung o đia   : $($assetObj.LogicalDrivesDetail)" -ForegroundColor Cyan

Write-Host "`n[5] THONG TIN MANG (NETWORK):" -ForegroundColor Yellow
Write-Host " - Card mang (NIC)  : $($assetObj.NetworkAdapter)"
Write-Host " - Dia chi IP       : $($assetObj.IPAddress)" -ForegroundColor Cyan
Write-Host " - Dia chi MAC      : $($assetObj.MACAddress)" -ForegroundColor Cyan

# XUAT FILE CSV VA JSON
$csvPath = Join-Path $env:USERPROFILE "Desktop\TaiSan_IT_$($assetObj.ComputerName).csv"
$jsonPath = Join-Path $env:USERPROFILE "Desktop\TaiSan_IT_$($assetObj.ComputerName).json"

try {
    $assetObj | Export-Csv -Path $csvPath -NoTypeInformation -Encoding utf8 -Force
    $assetObj | ConvertTo-Json | Set-Content -Path $jsonPath -Encoding utf8 -Force

    Write-Host "`n==========================================================================" -ForegroundColor Cyan
    Write-Host " [THANH CONG] Da xuat du lieu tai san IT chi tiet ra Desktop:" -ForegroundColor Green
    Write-Host "  * CSV  : $csvPath" -ForegroundColor White
    Write-Host "  * JSON : $jsonPath" -ForegroundColor White
    Write-Host "==========================================================================" -ForegroundColor Cyan
} catch {
    Write-Host "`nLoi khi xuat file tai san: $($_.Exception.Message)" -ForegroundColor Red
}
