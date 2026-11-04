# ============================================================
# ELKA v8.2 — общие функции установщиков (Windows PowerShell 5.1+)
# ============================================================
$ErrorActionPreference = "Stop"

function Say([string]$t, [string]$color = "Gray") { Write-Host $t -ForegroundColor $color }
function Step([string]$t) { Write-Host ""; Write-Host ("=== " + $t + " ===") -ForegroundColor Cyan }
function Ok([string]$t)   { Write-Host ("  [OK] " + $t) -ForegroundColor Green }
function Warn([string]$t) { Write-Host ("  [!]  " + $t) -ForegroundColor Yellow }
function Bad([string]$t)  { Write-Host ("  [X]  " + $t) -ForegroundColor Red }

function Ask([string]$prompt, [string]$default) {
    $v = Read-Host ($prompt + " [" + $default + "]")
    if ([string]::IsNullOrWhiteSpace($v)) { return $default }
    return $v.Trim().Trim('"')
}
function AskYes([string]$prompt, [bool]$defaultYes = $true) {
    $d = "Н"; if ($defaultYes) { $d = "Д" }
    $v = Read-Host ($prompt + " (Д/Н) [" + $d + "]")
    if ([string]::IsNullOrWhiteSpace($v)) { return $defaultYes }
    return ($v.Trim().ToUpper() -in @("Д","ДА","Y","YES","L"))
}

# Версия Excel: "16.0" / "15.0" / "14.0"
function Get-ExcelVersion {
    $k = [Microsoft.Win32.Registry]::ClassesRoot.OpenSubKey("Excel.Application\CurVer")
    if ($null -eq $k) { return $null }
    $v = [string]$k.GetValue("")
    if ($v -match "Excel\.Application\.(\d+)") { return ($Matches[1] + ".0") }
    return $null
}

# Кодовая страница для программ без Юникода (нужна 1251 — русский)
function Test-RussianCodePage {
    $cp = [System.Text.Encoding]::Default.CodePage
    if ($cp -ne 1251) {
        Bad ("Кодовая страница Windows для программ без Юникода = " + $cp + " (нужна 1251).")
        Say "  Тексты макросов (русские сообщения и названия листов) будут испорчены." Yellow
        Say "  Исправьте: Панель управления -> Региональные стандарты -> Дополнительно ->" Yellow
        Say "  «Язык программ, не поддерживающих Юникод» -> Русский (Россия) -> перезагрузка." Yellow
        return $false
    }
    Ok "Кодовая страница 1251 (русская) — подходит"
    return $true
}

# Доступ к объектной модели VBA (временно, на время установки)
function Set-VBOM([string]$ver, [int]$value) {
    $path = "HKCU:\Software\Microsoft\Office\$ver\Excel\Security"
    if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
    $old = $null
    try { $old = (Get-ItemProperty -Path $path -Name AccessVBOM -ErrorAction Stop).AccessVBOM } catch { $old = $null }
    New-ItemProperty -Path $path -Name AccessVBOM -Value $value -PropertyType DWord -Force | Out-Null
    return $old
}
function Restore-VBOM([string]$ver, $old) {
    $path = "HKCU:\Software\Microsoft\Office\$ver\Excel\Security"
    if ($null -eq $old) { Remove-ItemProperty -Path $path -Name AccessVBOM -ErrorAction SilentlyContinue }
    else { New-ItemProperty -Path $path -Name AccessVBOM -Value ([int]$old) -PropertyType DWord -Force | Out-Null }
}

# Надёжное расположение (макросы из этой папки включаются без вопросов)
function Add-TrustedLocation([string]$ver, [string]$folder, [string]$name) {
    $base = "HKCU:\Software\Microsoft\Office\$ver\Excel\Security\Trusted Locations"
    if (-not (Test-Path $base)) { New-Item -Path $base -Force | Out-Null }
    $key = Join-Path $base $name
    if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null }
    $p = $folder; if (-not $p.EndsWith("\")) { $p = $p + "\" }
    New-ItemProperty -Path $key -Name Path -Value $p -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $key -Name AllowSubfolders -Value 1 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $key -Name Description -Value "ELKA v8" -PropertyType String -Force | Out-Null
    if ($p.StartsWith("\\")) {
        New-ItemProperty -Path $base -Name AllowNetworkLocations -Value 1 -PropertyType DWord -Force | Out-Null
    }
    Ok ("Надёжное расположение: " + $p)
}

function Ensure-Folder([string]$p) {
    if ([string]::IsNullOrWhiteSpace($p)) { return }
    if (-not (Test-Path -LiteralPath $p)) { New-Item -ItemType Directory -Path $p -Force | Out-Null }
}

function Start-Excel {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.ScreenUpdating = $false
    try { $xl.EnableEvents = $false } catch {}   # не запускать Workbook_Open при сборке
    try { $xl.AutomationSecurity = 1 } catch {}   # msoAutomationSecurityLow — макросы установщика выполняются
    return $xl
}
function Stop-Excel($xl) {
    if ($null -eq $xl) { return }
    try { $xl.DisplayAlerts = $true } catch {}
    try { foreach ($w in @($xl.Workbooks)) { try { $w.Close($false) } catch {} } } catch {}
    try { $xl.Quit() } catch {}
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

# Удалить стандартные модули ELKA (для обновления)
function Remove-ElkaModules($wb) {
    $names = @()
    foreach ($c in $wb.VBProject.VBComponents) { if ($c.Type -eq 1 -and $c.Name -like "mod*") { $names += $c.Name } }
    foreach ($n in $names) { $wb.VBProject.VBComponents.Remove($wb.VBProject.VBComponents.Item($n)) }
    if ($names.Count -gt 0) { Ok ("Удалены старые модули: " + ($names -join ", ")) }
}

# Импорт .bas (файлы в UTF-8) через AddFromString
function Import-Modules($wb, [string]$folder) {
    $files = Get-ChildItem -LiteralPath $folder -Filter *.bas | Sort-Object Name
    foreach ($f in $files) {
        $text = [System.IO.File]::ReadAllText($f.FullName, [System.Text.Encoding]::UTF8)
        $name = $null
        if ($text -match 'Attribute VB_Name = "([^"]+)"') { $name = $Matches[1] }
        if ($null -eq $name) { throw ("В файле нет Attribute VB_Name: " + $f.Name) }
        $code = ($text -replace '(?m)^Attribute VB_Name = "[^"]+"\r?\n', '')
        $code = $code -replace "`r?`n", "`r`n"
        $comp = $wb.VBProject.VBComponents.Add(1)
        $comp.Name = $name
        $cm = $comp.CodeModule
        if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }   # убрать авто-«Option Explicit»
        $cm.AddFromString($code)
        Ok ("Модуль " + $name + " (" + $comp.CodeModule.CountOfLines + " строк)")
    }
}

# Код «ЭтаКнига» / ThisWorkbook
function Set-WorkbookCode($wb, [string]$file) {
    $text = [System.IO.File]::ReadAllText($file, [System.Text.Encoding]::UTF8)
    $text = $text -replace "`r?`n", "`r`n"
    $comp = $null
    try { $cn = $wb.CodeName; if ($cn) { $comp = $wb.VBProject.VBComponents.Item($cn) } } catch { $comp = $null }
    if ($null -eq $comp) {
        foreach ($c in $wb.VBProject.VBComponents) {
            if ($c.Type -eq 100) {
                try { if ($c.Properties.Item("Name").Value -eq $wb.Name) { $comp = $c; break } } catch {}
            }
        }
    }
    if ($null -eq $comp) {
        foreach ($c in $wb.VBProject.VBComponents) {
            if ($c.Type -eq 100 -and ($c.Name -eq "ЭтаКнига" -or $c.Name -eq "ThisWorkbook")) { $comp = $c; break }
        }
    }
    if ($null -eq $comp) { throw "Не найден модуль книги (ЭтаКнига/ThisWorkbook)" }
    $cm = $comp.CodeModule
    if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
    $cm.AddFromString($text)
    Ok ("Код книги записан в модуль «" + $comp.Name + "»")
}

function New-Shortcut([string]$target, [string]$name) {
    try {
        $desk = [Environment]::GetFolderPath("Desktop")
        $sh = New-Object -ComObject WScript.Shell
        $lnk = $sh.CreateShortcut((Join-Path $desk ($name + ".lnk")))
        $lnk.TargetPath = $target
        $lnk.WorkingDirectory = Split-Path -Parent $target
        $lnk.Save()
        Ok ("Ярлык на рабочем столе: " + $name)
    } catch { Warn ("Ярлык не создан: " + $_.Exception.Message) }
}
