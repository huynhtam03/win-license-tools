# WIN LICENSE & IT HELPDESK TOOLS

Bo cong cu kiem tra ban quyen Windows, Office, Project, Visio, go bo KMS lau, gỡ ung dung crack va ho tro dac luc cho IT Helpdesk.

---

## 🚀 HUONG DAN SU DUNG (LENH CHAY CHUAN PASSED CACHE)

Mo **Windows PowerShell** duoi quyen **Administrator** va copy cac lenh tuong ung duoi day de chay:

### 1. Giao dien Do hoa Modern Dark Dashboard (Tat-ca-trong-mot)
Dung khi muon su dung giao dien do hoa WPF bam nut hien dai, day du tinh nang:

```powershell
irm "https://raw.githubusercontent.com/huynhtam03/win-license-tools/main/Check-LicenseStatus-Standalone.ps1" -Headers @{"Cache-Control"="no-cache"} | iex
```

---

### 2. Cac Cong cu Chuc nang Doc lap (Dinh dang Dong lenh CLI)

#### 📋 A. Chi Quet Ban Quyen va Phat Hien Crack
Quet va in bao cao chi tiet ve ban quyen Windows, Office, Project, Visio va phat hien cac dau hieu be khoa.
```powershell
irm "https://raw.githubusercontent.com/huynhtam03/win-license-tools/main/Scan-LicenseStatus.ps1" -Headers @{"Cache-Control"="no-cache"} | iex
```

#### 🛡️ B. Chi Go Bo Ban Quyen KMS Lau & Lam Sach He Thong
Xoa may chu KMS, go Product Key KMS Windows/Office/Project/Visio, lam sach Hosts, Tasks va Services be khoa.
```powershell
irm "https://raw.githubusercontent.com/huynhtam03/win-license-tools/main/Remove-KMSActivation.ps1" -Headers @{"Cache-Control"="no-cache"} | iex
```

#### 🗑️ C. Chi Go Ca i Dat Hang Loat Cac App Crack Ben Thu Ba
Quet registry va go cai dat an (Silent Uninstall) cac phan mem bi be khoa (MiniTool, Adobe, IDM, AutoCAD...).
```powershell
irm "https://raw.githubusercontent.com/huynhtam03/win-license-tools/main/Uninstall-CrackedApps.ps1" -Headers @{"Cache-Control"="no-cache"} | iex
```

#### ❌ D. Chi Go Sach Office / Project / Visio
Go bo bo cai Click-To-Run, Registry, OSPP keys va WMI keys cua Office, Project va Visio.
```powershell
irm "https://raw.githubusercontent.com/huynhtam03/win-license-tools/main/Uninstall-Office.ps1" -Headers @{"Cache-Control"="no-cache"} | iex
```

#### ⚡ E. Chi Don Dep Temp & Reset Cache Mang
Xoa file rac Windows Temp, User Temp, Prefetch, SoftwareDistribution, xoa DNS Cache (`ipconfig /flushdns`) va Reset Winsock catalog.
```powershell
irm "https://raw.githubusercontent.com/huynhtam03/win-license-tools/main/Clean-SystemTemp.ps1" -Headers @{"Cache-Control"="no-cache"} | iex
```

#### 📦 F. Chi Thu Thap Thong Tin Tai San IT & Xuat CSV/JSON
Thu thap Serial Number, Mainboard, CPU, RAM, Disk, IP, MAC va tu dong xuat file CSV & JSON len Desktop de nhap lieu vao phan mem Quan ly tai san (Snipe-IT, GLPI, Excel...).
```powershell
irm "https://raw.githubusercontent.com/huynhtam03/win-license-tools/main/Get-ComputerAssetInfo.ps1" -Headers @{"Cache-Control"="no-cache"} | iex
```

---

## ✨ TINH NANG NOI BAT

* **Tieu chuan ASCII Khong Dau:** Dam bao khong bao gio bi loi font hoac loi cu phap khi chay `irm | iex` tren Windows PowerShell 5.1 & 7.x.
* **Xua tan Cache Web:** Su dung cờ `Cache-Control: no-cache` dam bao luon luon tai phien ban code moi nhat tu GitHub.
* **WMI & OSPP Integration:** Can quet va xoa triet de cac khoa ban quyen KMS qua trinh quan ly OSPP va WMI.
* **Export HTML Audit Report:** Xuat bao cao kiem tra ban quyen va cau hinh he thong ra file HTML chuyen nghiep ngay tren Desktop.
