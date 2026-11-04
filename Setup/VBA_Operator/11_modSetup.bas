Attribute VB_Name = "modSetup"
Option Explicit
' ============================================================
' ELKA v8.2 — КНОПКИ, НАВИГАЦИЯ, ПЕРВИЧНАЯ НАСТРОЙКА
' Setup_All запускается установщиком один раз. Его можно запускать
' повторно (Alt+F8 -> Setup_All): кнопки пересоздаются.
' ============================================================

Public Sub Setup_All()
    On Error Resume Next
    AddButtons
    Menu_Install
    RefreshStockRows
    RefreshInputFiles
    If Len(Txt(WS(SH_INPUT).Range(IN_DATE).Value)) = 0 Then WS(SH_INPUT).Range(IN_DATE).Value = Date
    EnsureFolder ArchiveRoot()
    EnsureFolder BackupFolder()
    WS(SH_PANEL).Activate
    WS(SH_PANEL).Range("A1").Select
End Sub

Private Sub Btn(ByVal sh As Worksheet, ByVal cellAddr As String, ByVal caption As String, ByVal macro As String, _
                Optional ByVal w As Double = 0, Optional ByVal h As Double = 26)
    Dim b As Object, rg As Range
    Set rg = sh.Range(cellAddr)
    If w = 0 Then w = rg.MergeArea.Width - 4
    Set b = sh.Buttons.Add(rg.Left + 2, rg.Top + 2, w, h)
    b.OnAction = "'" & ThisWorkbook.Name & "'!" & macro
    b.Characters.Text = caption
    b.Font.Bold = True
    b.Font.Size = 10
    b.Name = "ELKA_" & macro & "_" & sh.Buttons.Count
    b.Placement = 3
End Sub

Private Sub ClearButtons(ByVal sh As Worksheet)
    Dim i As Long
    On Error Resume Next
    For i = sh.Buttons.Count To 1 Step -1
        If Left$(sh.Buttons(i).Name, 5) = "ELKA_" Then sh.Buttons(i).Delete
    Next i
End Sub

Public Sub AddButtons()
    Dim sh As Worksheet
    ' ---- Ввод: кнопки справа от шапки (столбцы K..L) ----
    Set sh = WS(SH_INPUT): ClearButtons sh
    Btn sh, "K4", "ПРОВЕСТИ", "Post_FromInput", 190, 30
    Btn sh, "K6", "Прикрепить файл(ы)", "Files_AttachToInput", 190
    Btn sh, "K8", "Заполнить по спецификации", "Input_FillBySpec", 190
    Btn sh, "K10", "Заполнить по объединению", "Input_FillByBundle", 190
    Btn sh, "K12", "Заполнить остатками «Откуда»", "Input_FillFromStock", 190
    Btn sh, "M4", "СОБРАТЬ изделие (08+09/10)", "Input_Assemble", 200, 30
    Btn sh, "M6", "Скомплектовать объединение", "Input_BundleBuild", 200
    Btn sh, "M8", "Разобрать объединение", "Input_BundleSplit", 200
    Btn sh, "M12", "Очистить лист", "Input_Clear", 200
    ' ---- Движение ----
    Set sh = WS(SH_MOVE): ClearButtons sh
    Btn sh, "U1", "Открыть файл строки", "Files_OpenSelected", 170
    Btn sh, "W1", "Прикрепить файл к строке", "Files_AttachToSelectedRow", 180
    ' ---- Справочник ----
    Set sh = WS(SH_NOM): ClearButtons sh
    Btn sh, "N1", "Фото позиции", "Files_SetNomPhoto", 140
    Btn sh, "P1", "Удалить в корзину", "Nom_DeleteSelected", 150
    ' ---- Файлы ----
    Set sh = WS(SH_FILES): ClearButtons sh
    Btn sh, "K1", "Открыть файл", "Files_OpenSelected", 140
    Btn sh, "M1", "Переименовать", "Files_Rename", 140
    ' ---- Корзины ----
    Set sh = WS(SH_TRASH): ClearButtons sh
    Btn sh, "O1", "Восстановить выделенные", "Rows_RestoreSelected", 190
    Set sh = WS(SH_DELMOVE): ClearButtons sh
    Btn sh, "W1", "Восстановить выделенные", "Rows_RestoreSelected", 190
    ' ---- Контроль ----
    Set sh = WS(SH_CTRL): ClearButtons sh
    Btn sh, "H1", "ПРОВЕРИТЬ КНИГУ", "Control_Run", 180, 30
    ' ---- 1С ----
    Set sh = WS(SH_1C): ClearButtons sh
    Btn sh, "I1", "Загрузить из 1С", "Import1C_Dialog", 170, 30
    Set sh = WS(SH_CMP): ClearButtons sh
    Btn sh, "K1", "Загрузить из 1С", "Import1C_Dialog", 170, 30
    Btn sh, "M1", "Пересчитать сверку", "Compare1C_Run", 170, 30
    ' ---- отчёты ----
    Set sh = WS(SH_ADDRSTOCK): ClearButtons sh
    Btn sh, "H1", "Обновить отчёты", "Reports_RefreshAll", 160
    Set sh = WS(SH_VARSUM): ClearButtons sh
    Btn sh, "P1", "Обновить отчёты", "Reports_RefreshAll", 160
    Set sh = WS(SH_KITS): ClearButtons sh
    Btn sh, "N1", "Обновить отчёты", "Reports_RefreshAll", 160
    Set sh = WS(SH_STOCK): ClearButtons sh
    Btn sh, "U1", "Протянуть формулы", "Reports_RefreshAll", 160
    Set sh = WS(SH_PREVIEW): ClearButtons sh
    Btn sh, "F3", "Открыть в программе", "Preview_OpenInProgram", 170
    ' ---- Панель ----
    Set sh = WS(SH_PANEL): ClearButtons sh
    Btn sh, "E5", "Ввод операции", "Nav_Input", 200, 30
    Btn sh, "E7", "Проверить книгу", "Control_Run", 200, 30
    Btn sh, "E9", "Обновить отчёты", "Reports_RefreshAll", 200, 30
    Btn sh, "E11", "Загрузить из 1С и сверить", "Import1C_Dialog", 200, 30
    Btn sh, "E13", "ОПУБЛИКОВАТЬ для пользователей", "Publish_Data", 200, 30
    Btn sh, "E15", "Резервная копия сейчас", "Backup_Now", 200, 30
End Sub

Public Sub Nav_Input()
    GoSheet SH_INPUT
    WS(SH_INPUT).Range(IN_OP).Select
End Sub

Public Sub Backup_Now()
    Dim f As String
    gSilent = False
    f = MakeBackup("manual")
    If Len(f) > 0 Then
        Info "Резервная копия сохранена:" & vbLf & f
    Else
        Fail "Не удалось сохранить копию. Сначала сохраните книгу (Ctrl+S)."
    End If
End Sub

' ============================================================
' Вызывается УСТАНОВЩИКОМ (Install_Operator.ps1) после импорта модулей:
' записывает папки в «Настройки», создаёт папки, кнопки и меню.
' Можно вызвать вручную: Alt+F8 не показывает процедуры с параметрами,
' поэтому для ручной настройки просто измените лист «Настройки».
' ============================================================
Public Function Install_Configure(ByVal archiveFolder As String, ByVal publishFolder As String, _
                                  ByVal backupFolder As String, ByVal operatorName As String) As String
    Dim oldSilent As Boolean
    oldSilent = gSilent
    gSilent = True
    On Error GoTo Bad
    If Len(Trim$(archiveFolder)) > 0 Then SaveSetting2 "Папка архива документов", Trim$(archiveFolder)
    If Len(Trim$(publishFolder)) > 0 Then SaveSetting2 "Папка публикации", Trim$(publishFolder)
    If Len(Trim$(backupFolder)) > 0 Then SaveSetting2 "Папка резервных копий", Trim$(backupFolder)
    If Len(Trim$(operatorName)) > 0 Then SaveSetting2 "Оператор по умолчанию", Trim$(operatorName)
    EnsureFolder ArchiveRoot()
    EnsureFolder PublishFolder()
    EnsureFolder BackupFolder()
    Dim repairResult As String
    repairResult = Install_RepairDepartment()
    If repairResult <> "OK" Then Err.Raise 5, , repairResult
    Setup_All
    gSilent = oldSilent
    Install_Configure = "OK"
    Exit Function
Bad:
    gSilent = oldSilent
    Install_Configure = "ERROR: " & Err.Description
End Function

' Проверка, что все модули на месте (вызывается установщиком)
Public Function Install_Check() As String
    Install_Check = "ELKA " & ELKA_VERSION & " OK; листов " & ThisWorkbook.Worksheets.Count
End Function
