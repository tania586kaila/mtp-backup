# MTP Backup via WPD COM API (PowerShell)
# 直接调用 Windows Portable Device API，支持中文文件名
# 用法: .\mtp_backup.ps1

param(
    [string]$SourceDir   = "F:\share\sync",
    [string]$DeviceName  = "F50",
    [string]$BackupRoot  = "PC备份",
    [string]$DateFormat  = "yyyyMMdd",
    [string]$LogFile     = "$PSScriptRoot\backup.log"
)

# ── 日志 ──────────────────────────────────────
function Write-Log {
    param([string]$Level, [string]$Msg)
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts] $Level - $Msg"
    Write-Host $line
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
}
function Log-Info  { param($m) Write-Log "INFO " $m }
function Log-Debug { param($m) Write-Log "DEBUG" $m }
function Log-Error { param($m) Write-Log "ERROR" $m }
function Log-Ok    { param($m) Write-Log "OK   " $m }
function Log-Skip  { param($m) Write-Log "SKIP " $m }

# ── WPD 辅助 ─────────────────────────────────
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

[ComImport, Guid("A1567595-4C2F-4574-A6FA-ECEF917B9A40"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IPortableDeviceManager {
    void GetDevices([In, Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex=1)] string[] pPnPDeviceIDs, ref uint pcPnPDeviceIDs);
    void RefreshDeviceList();
    void GetDeviceFriendlyName([In, MarshalAs(UnmanagedType.LPWStr)] string pszPnPDeviceID,
        [In, Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex=2)] char[] pDeviceFriendlyName, ref uint pcchDeviceFriendlyName);
    void GetDeviceDescription([In, MarshalAs(UnmanagedType.LPWStr)] string pszPnPDeviceID,
        [In, Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex=2)] char[] pDeviceDescription, ref uint pcchDeviceDescription);
    void GetDeviceManufacturer([In, MarshalAs(UnmanagedType.LPWStr)] string pszPnPDeviceID,
        [In, Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex=2)] char[] pDeviceManufacturer, ref uint pcchDeviceManufacturer);
}
"@ -ErrorAction SilentlyContinue

function Get-WPDDeviceId {
    param([string]$Name)
    $mgr = New-Object -ComObject "PortableDeviceManager"
    $count = [uint32]0
    $mgr.GetDevices($null, [ref]$count)
    if ($count -eq 0) { Log-Error "未找到 WPD 设备"; return $null }
    $ids = New-Object string[] $count
    $mgr.GetDevices($ids, [ref]$count)
    foreach ($id in $ids) {
        $len = [uint32]256
        $buf = New-Object char[] $len
        try {
            $mgr.GetDeviceFriendlyName($id, $buf, [ref]$len)
            $devName = New-Object string($buf, 0, [int]$len - 1)
            Log-Debug "WPD 设备: $devName"
            if ($devName -like "*$Name*") {
                Log-Info "匹配设备: $devName ($id)"
                return $id
            }
        } catch {}
    }
    return $null
}

# ── Shell Namespace MTP 操作 ──────────────────
# 用 Shell.Application 做目录遍历和文件上传
# 关键：用 FolderItem.Path 拼接子路径，用 Folder.CopyHere 上传

function Get-StorageFolder {
    param([string]$DevName)
    $sh = New-Object -ComObject Shell.Application
    $pc = $sh.Namespace(0x11)
    foreach ($item in $pc.Items()) {
        if ($item.Name -like "*$DevName*") {
            $devNS = $sh.Namespace($item.Path)
            if ($devNS -and $devNS.Items().Count -gt 0) {
                $storage = $devNS.Items().Item(0)
                Log-Info "存储空间: $($storage.Name)"
                return @{ Shell=$sh; Folder=$storage.GetFolder; Path=$storage.Path }
            }
        }
    }
    return $null
}

function Find-SubFolder {
    param($ParentFolder, [string]$Name)
    $items = $ParentFolder.Items()
    for ($i = 0; $i -lt $items.Count; $i++) {
        $item = $items.Item($i)
        if ($item.Name -eq $Name -and $item.IsFolder) {
            return $item
        }
    }
    return $null
}

function Ensure-MTPFolder {
    param($Shell, $ParentFolder, $ParentPath, [string]$Name)
    $existing = Find-SubFolder $ParentFolder $Name
    if ($existing) {
        Log-Debug "已存在: $Name"
        return @{ Item=$existing; Folder=$existing.GetFolder; Path="$ParentPath\$Name" }
    }
    Log-Info "创建目录: $Name"
    $ParentFolder.NewFolder($Name)
    $deadline = (Get-Date).AddSeconds(10)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 500
        $found = Find-SubFolder $ParentFolder $Name
        if ($found) {
            return @{ Item=$found; Folder=$found.GetFolder; Path="$ParentPath\$Name" }
        }
    }
    Log-Error "创建目录超时: $Name"
    return $null
}

function Get-MTPFileNames {
    param($Folder)
    $names = @{}
    if ($Folder -eq $null) { return $names }
    $items = $Folder.Items()
    for ($i = 0; $i -lt $items.Count; $i++) {
        $names[$items.Item($i).Name] = $true
    }
    return $names
}

function Copy-FileToMTP {
    param($Shell, [string]$LocalFile, $MTPFolder, [string]$FileName)
    $srcDir  = Split-Path $LocalFile -Parent
    $srcNS   = $Shell.Namespace($srcDir)
    if ($srcNS -eq $null) { Log-Error "无法打开: $srcDir"; return $false }
    $fileItem = $srcNS.ParseName($FileName)
    if ($fileItem -eq $null) { Log-Error "无法获取: $FileName"; return $false }

    # CopyHere + 轮询等待
    $MTPFolder.CopyHere($fileItem, 0x14)
    $deadline = (Get-Date).AddSeconds(60)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds 1
        $names = Get-MTPFileNames $MTPFolder
        if ($names.ContainsKey($FileName)) {
            return $true
        }
    }
    return $false
}

# ── 主流程 ────────────────────────────────────
Log-Info ("=" * 60)
Log-Info "MTP 备份开始"
Log-Info "源目录: $SourceDir"
Log-Info ("=" * 60)

# 1. 获取存储
$storage = Get-StorageFolder $DeviceName
if ($storage -eq $null) { Log-Error "未找到 F50 存储"; exit 1 }

$sh = $storage.Shell

# 2. 创建备份目录结构
$backupRootInfo = Ensure-MTPFolder $sh $storage.Folder $storage.Path $BackupRoot
if ($backupRootInfo -eq $null) { exit 1 }

$dateStr = Get-Date -Format $DateFormat
$dateInfo = Ensure-MTPFolder $sh $backupRootInfo.Folder $backupRootInfo.Path $dateStr
if ($dateInfo -eq $null) { exit 1 }

$dirName = Split-Path $SourceDir -Leaf
$dirInfo  = Ensure-MTPFolder $sh $dateInfo.Folder $dateInfo.Path $dirName
if ($dirInfo -eq $null) { exit 1 }

Log-Info "备份目标: F50\内部共享存储空间\$BackupRoot\$dateStr\$dirName"

# 3. 遍历源目录所有文件
$allFiles = Get-ChildItem -Path $SourceDir -Recurse -File
Log-Info "发现 $($allFiles.Count) 个文件"

$copied = 0; $skipped = 0; $failed = 0

foreach ($file in $allFiles) {
    $relPath = $file.FullName.Substring($SourceDir.TrimEnd('\').Length + 1)
    $parts   = $relPath -split '\\'

    # 确保子目录结构
    $curInfo = $dirInfo
    $dirOk   = $true
    for ($i = 0; $i -lt $parts.Count - 1; $i++) {
        $curInfo = Ensure-MTPFolder $sh $curInfo.Folder $curInfo.Path $parts[$i]
        if ($curInfo -eq $null) { $failed++; $dirOk = $false; break }
    }
    if (-not $dirOk) { continue }

    $fname = $file.Name

    # 检查是否已存在
    $existing = Get-MTPFileNames $curInfo.Folder
    if ($existing.ContainsKey($fname)) {
        Log-Skip $relPath
        $skipped++
        continue
    }

    Log-Info "[COPY] $relPath"
    $ok = Copy-FileToMTP $sh $file.FullName $curInfo.Folder $fname
    if ($ok) {
        Log-Ok  "[OK]   $fname"
        $copied++
    } else {
        Log-Error "[FAIL] $fname"
        $failed++
    }
}

Log-Info ("=" * 60)
Log-Info "备份完成! [OK]$copied  [SKIP]$skipped  [FAIL]$failed"
Log-Info ("=" * 60)

# 4. 验证：列出 F50 目标目录内容
Write-Host ""
Write-Host "=== 验证：F50\内部共享存储空间\$BackupRoot\$dateStr\$dirName ==="
$verFolder = (Ensure-MTPFolder $sh $dateInfo.Folder $dateInfo.Path $dirName).Folder
if ($verFolder) {
    $verItems = $verFolder.Items()
    Write-Host "共 $($verItems.Count) 个文件:"
    for ($i = 0; $i -lt $verItems.Count; $i++) {
        $item = $verItems.Item($i)
        Write-Host "  [OK] $($item.Name)"
    }
}
