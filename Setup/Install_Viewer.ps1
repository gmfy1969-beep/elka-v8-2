# ============================================================
# ELKA v8.2 — УСТАНОВКА КНИГИ ПОЛЬЗОВАТЕЛЯ (ПРОСМОТРЩИКА)
# Запуск: двойной щелчок по 3_INSTALL_VIEWER.cmd
# Ставится на КАЖДЫЙ компьютер пользователя. Оператор не нужен.
# ============================================================
param([string]$Mode = "")
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $here "Common.ps1")
$root = Split-Path -Parent $here
$host.UI.RawUI.WindowTitle = "ELKA v8.2 — установка просмотрщика"

Say "=============================================================" Cyan
Say "  ELKA v8.2 — установка КНИГИ ПОЛЬЗОВАТЕЛЯ (просмотр)" Cyan
Say "  Enter — принять значение в [скобках]." Cyan
Say "=============================================================" Cyan

Step "1. Проверка компьютера"
$ver = Get-ExcelVersion
if ($null -eq $ver) { Bad "Microsoft Excel не найден."; exit 1 }
Ok ("Excel найден, версия " + $ver)
if (-not (Test-RussianCodePage)) {
    if (-not (AskYes "Продолжить всё равно (НЕ рекомендуется)?" $false)) { exit 1 }
}
if (Get-Process -Name EXCEL -ErrorAction SilentlyContinue) {
    Warn "Excel сейчас открыт. Сохраните свои файлы и закройте Excel."
    Read-Host "Закройте Excel и нажмите Enter"
}

Step "2. Параметры"
if ($Mode -eq "") {
    Say "  1 — обычная установка (данные будут загружены из папки публикации оператора)"
    Say "  2 — ДЕМО (уже содержит демо-данные; папка публикации — демо оператора)"
    $Mode = Ask "Ваш выбор" "1"
}
$installDir = Ask "Папка установки" "C:\ELKA"
if ($Mode -eq "2") { $installDir = Join-Path $installDir "DEMO" }
Ensure-Folder $installDir
Say "  Папку публикации сообщает оператор, например \\SERVER\ELKA\Publish"
$publish = Ask "Папка публикации данных" (Join-Path $installDir "Publish")
if (-not (Test-Path -LiteralPath $publish)) { Warn "Папка публикации сейчас недоступна — проверьте путь или сеть. Её можно изменить позже кнопкой «Настроить папки»." }
else {
    $n = @(Get-ChildItem -LiteralPath $publish -Filter "ELKA_data_v*.xlsx" -ErrorAction SilentlyContinue).Count
    if ($n -gt 0) { Ok ("В папке публикации найдено файлов данных: " + $n) } else { Warn "В папке пока нет данных (оператор ещё не публиковал)." }
}
$archive = Ask "Папка архива документов (Enter — взять из публикации оператора)" ""

$fileName = "ELKA_Viewer.xlsm"; if ($Mode -eq "2") { $fileName = "ELKA_Viewer_DEMO.xlsm" }
$target = Join-Path $installDir $fileName
if (Test-Path -LiteralPath $target) {
    Warn ("Файл уже существует: " + $target + " — будет заменён (в просмотрщике нет своих данных).")
    Remove-Item -LiteralPath $target -Force
}

Step "3. Настройка Excel"
Add-TrustedLocation $ver $installDir "ELKA_v8_Viewer"
$oldVBOM = Set-VBOM $ver 1
Ok "Временно включён доступ к проекту VBA"

$xl = $null; $wb = $null; $result = 1
try {
    Step "4. Сборка книги"
    $xl = Start-Excel
    $tpl = Join-Path $here "Workbooks\ELKA_Viewer_Template.xlsx"
    if ($Mode -eq "2") { $tpl = Join-Path $here "Workbooks\ELKA_Viewer_Demo.xlsx" }
    $wb = $xl.Workbooks.Open($tpl)
    Import-Modules $wb (Join-Path $here "VBA_Viewer")
    Set-WorkbookCode $wb (Join-Path $here "Object_Code\Viewer_ThisWorkbook.txt")
    $wb.SaveAs($target, 52)
    Ok ("Сохранено: " + $target)
    $cfg = $xl.Run("'" + $wb.Name + "'!Install_ConfigureViewer", $publish, $archive)
    if ($cfg -ne "OK") { throw ("Настройка: " + $cfg) }
    Ok "Папки записаны, кнопки созданы"
    $wb.Save()
    $wb.Close($false); $wb = $null
    $result = 0
}
catch {
    Bad ("ОШИБКА: " + $_.Exception.Message)
    Say "  Подробности — Docs\01_Install_RU.pdf, раздел «Если установка не прошла»." Yellow
    if ($wb) { try { $wb.Close($false) } catch {} }
}
finally {
    Stop-Excel $xl
    Restore-VBOM $ver $oldVBOM
    Ok "Доступ к проекту VBA возвращён в исходное состояние"
}
if ($result -ne 0) { exit 1 }

Step "5. Готово"
$lnk = "ELKA - просмотр"; if ($Mode -eq "2") { $lnk = "ELKA - ДЕМО просмотр" }
New-Shortcut $target $lnk
$ud = Join-Path $root "Docs\03_User_RU.pdf"
if (Test-Path -LiteralPath $ud) { Ensure-Folder (Join-Path $installDir "Docs"); Copy-Item -LiteralPath $ud -Destination (Join-Path $installDir "Docs") -Force -ErrorAction SilentlyContinue; Ok ("Инструкция: " + (Join-Path $installDir "Docs\03_User_RU.pdf")) }
Say ""
Say ("  Книга пользователя: " + $target) Green
Say "  Откройте её ярлыком на рабочем столе. При открытии данные обновятся сами;" Green
Say "  кнопка «ОБНОВИТЬ ДАННЫЕ» на листе «Панель» — обновить вручную." Green
exit 0
