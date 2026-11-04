# ============================================================
# ELKA v8.2 — УСТАНОВКА КНИГИ ОПЕРАТОРА
# Запуск: двойной щелчок по 1_INSTALL_OPERATOR.cmd (в корне комплекта)
# ============================================================
param([string]$Mode = "")
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $here "Common.ps1")
$root = Split-Path -Parent $here
$host.UI.RawUI.WindowTitle = "ELKA v8.2 — установка книги оператора"

Say "=============================================================" Cyan
Say "  ELKA v8.2 — установка КНИГИ ОПЕРАТОРА" Cyan
Say "  Отвечайте на вопросы. Enter — принять значение в [скобках]." Cyan
Say "=============================================================" Cyan

Step "1. Проверка компьютера"
$ver = Get-ExcelVersion
if ($null -eq $ver) { Bad "Microsoft Excel не найден. Установите Excel 2010 или новее и повторите."; exit 1 }
Ok ("Excel найден, версия " + $ver)
if (-not (Test-RussianCodePage)) {
    if (-not (AskYes "Продолжить всё равно (НЕ рекомендуется)?" $false)) { exit 1 }
}
$running = Get-Process -Name EXCEL -ErrorAction SilentlyContinue
if ($running) {
    Warn "Excel сейчас открыт. Сохраните свои файлы и закройте Excel."
    Read-Host "Закройте Excel и нажмите Enter"
}

Step "2. Что устанавливаем"
if ($Mode -eq "") {
    Say "  1 — РАБОЧАЯ книга (пустой журнал, справочник мест/операций/адресов заполнен примером)"
    Say "  2 — ДЕМО-книга (заполнена учебным маршрутом ремонта — чтобы потренироваться)"
    Say "  3 — ОБНОВИТЬ макросы в уже существующей книге (данные не меняются)"
    $Mode = Ask "Ваш выбор" "1"
}

$installDir = Ask "Папка установки (книга, архив, копии)" "C:\ELKA"
if ($Mode -eq "2") { $installDir = Join-Path $installDir "DEMO" }
Ensure-Folder $installDir

if ($Mode -eq "3") {
    $target = Ask "Полный путь к книге оператора" (Join-Path $installDir "ELKA_Operator.xlsm")
    if (-not (Test-Path -LiteralPath $target)) { Bad ("Файл не найден: " + $target); exit 1 }
    $bak = $target -replace '\.xlsm$', ("_до_обновления_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".xlsm")
    Copy-Item -LiteralPath $target -Destination $bak
    Ok ("Резервная копия: " + $bak)
    $archive = ""; $publish = ""; $backup = ""; $who = ""
} else {
    $fileName = "ELKA_Operator.xlsm"; if ($Mode -eq "2") { $fileName = "ELKA_Operator_DEMO.xlsm" }
    $target = Join-Path $installDir $fileName
    if (Test-Path -LiteralPath $target) {
        Warn ("Книга уже существует: " + $target)
        if (-not (AskYes "Переименовать старую книгу в *_old_<дата>.xlsm и установить новую?" $false)) {
            Say "  Установка отменена. Чтобы обновить макросы без потери данных — запустите снова и выберите 3." Yellow; exit 1
        }
        Rename-Item -LiteralPath $target -NewName ([IO.Path]::GetFileNameWithoutExtension($target) + "_old_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".xlsm")
    }
    Step "3. Папки"
    Say "  Архив документов — куда копируются прикреплённые файлы (накладные, акты, фото)."
    Say "  Чтобы пользователи открывали документы со своих компьютеров, укажите СЕТЕВУЮ папку,"
    Say "  например \\SERVER\ELKA\Архив. Для одного компьютера подойдёт локальная."
    $archive = Ask "Папка архива документов" (Join-Path $installDir "Архив")
    Say "  Публикация — куда кнопка «ОПУБЛИКОВАТЬ» кладёт данные для пользователей (тоже лучше сетевая)."
    $publish = Ask "Папка публикации" (Join-Path $installDir "Publish")
    $backup  = Ask "Папка резервных копий" (Join-Path $installDir "Backups")
    $who     = Ask "ФИО оператора (для столбца «Оператор»; Enter — имя пользователя Windows)" ""
    foreach ($p in @($archive, $publish, $backup)) {
        try { Ensure-Folder $p; Ok ("Папка: " + $p) } catch { Warn ("Нет доступа к папке " + $p + " — проверьте позже") }
    }
    if ($Mode -eq "2") {
        $demoArch = Join-Path $root "Demo\Архив"
        if (Test-Path -LiteralPath $demoArch) { Copy-Item -Path (Join-Path $demoArch "*") -Destination $archive -Recurse -Force; Ok "Демо-документы скопированы в архив" }
        $demoPub = Join-Path $root "Demo\Publish"
        if (Test-Path -LiteralPath $demoPub) { Copy-Item -Path (Join-Path $demoPub "*") -Destination $publish -Recurse -Force; Ok "Демо-публикация скопирована" }
    }
}

Step "4. Настройка Excel"
Add-TrustedLocation $ver $installDir "ELKA_v8"
$oldVBOM = Set-VBOM $ver 1
Ok "Временно включён доступ к проекту VBA (будет возвращён в конце)"

$xl = $null; $wb = $null; $result = 1
try {
    Step "5. Сборка книги"
    $xl = Start-Excel
    if ($Mode -eq "3") {
        $wb = $xl.Workbooks.Open($target)
        foreach ($required in @("Ввод", "Движение", "Места хранения", "Адреса", "Операции", "Некомплектные_ПФ", "Заказы")) {
            try { $null = $wb.Worksheets.Item($required).Name } catch { throw ("Книга не соответствует структуре v8.1/v8.2: нет листа " + $required + ". Установите новую рабочую книгу; не обновляйте v6.1 этим сценарием.") }
        }
        Remove-ElkaModules $wb
        # сохранить и переоткрыть: Excel окончательно удаляет модули только после сохранения
        $wb.Save(); $wb.Close($false); $wb = $null
        $wb = $xl.Workbooks.Open($target)
        foreach ($c in $wb.VBProject.VBComponents) { if ($c.Type -eq 1 -and $c.Name -like "mod*") { throw ("Старый модуль не удалён: " + $c.Name + ". Закройте Excel и повторите.") } }
    } else {
        $tpl = Join-Path $here "Workbooks\ELKA_Operator_Template.xlsx"
        if ($Mode -eq "2") { $tpl = Join-Path $here "Workbooks\ELKA_Operator_Demo.xlsx" }
        if (-not (Test-Path -LiteralPath $tpl)) { throw ("Нет файла шаблона: " + $tpl) }
        $wb = $xl.Workbooks.Open($tpl)
    }
    Import-Modules $wb (Join-Path $here "VBA_Operator")
    Set-WorkbookCode $wb (Join-Path $here "Object_Code\Operator_ThisWorkbook.txt")
    if ($Mode -ne "3") {
        $wb.SaveAs($target, 52)   # 52 = xlOpenXMLWorkbookMacroEnabled (.xlsm)
        Ok ("Сохранено: " + $target)
    }
    $name = $wb.Name
    $chk = $xl.Run("'" + $name + "'!Install_Check")
    Ok ("Проверка макросов: " + $chk)
    $cfg = $xl.Run("'" + $name + "'!Install_Configure", $archive, $publish, $backup, $who)
    if ($cfg -ne "OK") { throw ("Настройка книги: " + $cfg) }
    Ok "Папки записаны в «Настройки», кнопки и меню созданы"
    $wb.Save()
    $wb.Close($false); $wb = $null
    $result = 0
}
catch {
    Bad ("ОШИБКА: " + $_.Exception.Message)
    Say "  Частые причины: Excel был открыт; антивирус блокирует доступ к VBA; нет прав на папку." Yellow
    Say "  Подробности — Docs\01_Install_RU.pdf, раздел «Если установка не прошла»." Yellow
    if ($wb) { try { $wb.Close($false) } catch {} }
}
finally {
    Stop-Excel $xl
    Restore-VBOM $ver $oldVBOM
    Ok "Доступ к проекту VBA возвращён в исходное состояние"
}
if ($result -ne 0) { exit 1 }

Step "6. Готово"
$lnkName = "ELKA - книга оператора"; if ($Mode -eq "2") { $lnkName = "ELKA - ДЕМО оператора" }
if ($Mode -ne "3") { New-Shortcut $target $lnkName }
Copy-Item -Path (Join-Path $root "Docs\*") -Destination (Join-Path $installDir "Docs") -Recurse -Force -ErrorAction SilentlyContinue
Say ""
Say ("  Книга оператора: " + $target) Green
if ($Mode -ne "3") {
    Say ("  Архив документов: " + $archive) Green
    Say ("  Папка публикации: " + $publish) Green
    Say "  Сообщите пользователям путь к папке публикации — он нужен при установке просмотрщика." Green
}
Say "  Откройте книгу (ярлык на рабочем столе), прочитайте лист «Инструкция_оператора»." Green
exit 0
