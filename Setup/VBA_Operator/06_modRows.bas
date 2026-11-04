Attribute VB_Name = "modRows"
Option Explicit
' ============================================================
' ELKA v8.2 — ВСТАВКА / УДАЛЕНИЕ / ВОССТАНОВЛЕНИЕ СТРОК
' Контекстное меню (правая кнопка мыши по ячейке или номеру строки):
'   «ELKA: вставить строки...»          — заданное количество (1..1000)
'   «ELKA: удалить выделенные строки»   — безопасно, с корзиной
'   «ELKA: восстановить из корзины»
'   «ELKA: открыть файл / фото»
'   «ELKA: прикрепить файл к операции»
' ============================================================
Private Const MENU_TAG As String = "ELKA_MENU"

' Листы, где разрешена вставка, и первая строка данных
Public Function RowsAllowed(ByVal sheetName As String, ByRef firstRow As Long) As Boolean
    firstRow = FIRST_ROW
    Select Case sheetName
    Case SH_INPUT
        firstRow = IN_FIRST: RowsAllowed = True
    Case SH_MOVE, SH_NOM, SH_SPEC, SH_BUNDLE, SH_LOC, SH_ADDR, SH_OPS, SH_FILES, SH_1C, SH_TRASH, SH_DELMOVE
        RowsAllowed = True
    End Select
End Function

' ---------------- контекстное меню ----------------
Public Sub Menu_Install()
    Dim barName As Variant, bar As Object, c As Object, pre As String
    On Error Resume Next
    Menu_Remove
    pre = "'" & ThisWorkbook.Name & "'!"
    For Each barName In Array("Cell", "Row")
        Set bar = Application.CommandBars(CStr(barName))
        If Not bar Is Nothing Then
            AddMenuItem bar, "ELKA: вставить строки...", pre & "Rows_InsertAsk", True, 296
            AddMenuItem bar, "ELKA: удалить выделенные строки", pre & "Rows_DeleteSelected", False, 293
            AddMenuItem bar, "ELKA: восстановить из корзины", pre & "Rows_RestoreSelected", False, 37
            AddMenuItem bar, "ELKA: открыть файл / фото", pre & "Files_OpenSelected", True, 23
            AddMenuItem bar, "ELKA: прикрепить файл к операции", pre & "Files_AttachToSelectedRow", False, 1079
        End If
    Next barName
End Sub

Private Sub AddMenuItem(ByVal bar As Object, ByVal caption As String, ByVal action As String, _
                        ByVal beginGroup As Boolean, ByVal faceId As Long)
    Dim c As Object
    On Error Resume Next
    Set c = bar.Controls.Add(Type:=1, Temporary:=True)
    c.Caption = caption
    c.OnAction = action
    c.Tag = MENU_TAG
    c.BeginGroup = beginGroup
    c.FaceId = faceId
End Sub

Public Sub Menu_Remove()
    Dim barName As Variant, c As Object
    On Error Resume Next
    For Each barName In Array("Cell", "Row")
        Do
            Set c = Nothing
            Set c = Application.CommandBars(CStr(barName)).FindControl(Tag:=MENU_TAG)
            If c Is Nothing Then Exit Do
            c.Delete
        Loop
    Next barName
End Sub

' ---------------- вставка ----------------
Public Sub Rows_InsertAsk()
    Dim n As Variant, sh As Worksheet, fr As Long
    Set sh = ActiveSheet
    If Not RowsAllowed(sh.Name, fr) Then Fail "На листе «" & sh.Name & "» вставка строк не нужна (лист формируется автоматически).": Exit Sub
    n = InputBox("Сколько строк вставить НАД выделенной строкой?" & vbLf & "(от 1 до 1000)", "ELKA: вставка строк", "10")
    If Len(n) = 0 Then Exit Sub
    If Not IsNumeric(n) Then Fail "Введите число от 1 до 1000.": Exit Sub
    If CLng(n) < 1 Or CLng(n) > 1000 Then Fail "Введите число от 1 до 1000.": Exit Sub
    InsertRowsCore sh, ActiveCell.Row, CLng(n)
End Sub

' Вставка n строк над строкой atRow с переносом форматов, проверок и формул
Public Sub InsertRowsCore(ByVal sh As Worksheet, ByVal atRow As Long, ByVal n As Long)
    Dim fr As Long, lastCol As Long, c As Long, src As Long
    If Not RowsAllowed(sh.Name, fr) Then Exit Sub
    If atRow < fr Then atRow = fr
    AppOff
    On Error GoTo Bad
    lastCol = sh.Cells(HDR_ROW, sh.Columns.Count).End(xlToLeft).Column
    If sh.Name = SH_INPUT Then lastCol = L_CHECK
    sh.Range(sh.Cells(atRow, 1), sh.Cells(atRow + n - 1, 1)).EntireRow.Insert
    ' образец — строка, сдвинутая вниз (бывшая atRow)
    src = atRow + n
    CopyRowFormat sh, src, atRow, atRow + n - 1, lastCol
    sh.Range(sh.Cells(atRow, 1), sh.Cells(atRow + n - 1, lastCol)).ClearContents
    For c = 1 To lastCol
        If sh.Cells(src, c).HasFormula Then
            sh.Range(sh.Cells(atRow, c), sh.Cells(atRow + n - 1, c)).FormulaR1C1 = sh.Cells(src, c).FormulaR1C1
        End If
    Next c
    If sh.Name = SH_MOVE Then EnsureJournalFormulasIfFilled atRow, atRow + n - 1
    If sh.Name = SH_INPUT Then RenumberInput
    AppOn
    If Not gSilent Then sh.Cells(atRow, IIf(sh.Name = SH_INPUT, L_CODE, 1)).Select
    Exit Sub
Bad:
    AppOn
    Fail "Не удалось вставить строки: " & Err.Description
End Sub

Private Sub EnsureJournalFormulasIfFilled(ByVal r1 As Long, ByVal r2 As Long)
    ' Пустые вставленные строки журнала формул не требуют — формулы
    ' появятся автоматически при вводе операции (событие листа).
End Sub

Public Sub RenumberInput()
    Dim sh As Worksheet, r As Long, lastR As Long
    Set sh = WS(SH_INPUT)
    lastR = sh.Cells(sh.Rows.Count, L_CHECK).End(xlUp).Row
    If lastR < IN_FIRST Then Exit Sub
    sh.Range(sh.Cells(IN_FIRST, L_NO), sh.Cells(lastR, L_NO)).Formula = "=ROW()-" & (IN_FIRST - 1)
End Sub

' ---------------- удаление ----------------
Public Sub Rows_DeleteSelected()
    Dim sh As Worksheet, fr As Long, r1 As Long, r2 As Long
    Set sh = ActiveSheet
    If Not RowsAllowed(sh.Name, fr) Then Fail "На листе «" & sh.Name & "» удаление строк не предусмотрено.": Exit Sub
    r1 = Selection.Row
    r2 = r1 + Selection.Rows.Count - 1
    If r1 < fr Then r1 = fr
    If r2 < r1 Then Exit Sub
    If r2 - r1 > 5000 Then Fail "За один раз можно удалить не более 5000 строк.": Exit Sub
    DeleteRowsCore sh, r1, r2, True
End Sub

Public Function DeleteRowsCore(ByVal sh As Worksheet, ByVal r1 As Long, ByVal r2 As Long, _
                               Optional ByVal confirm As Boolean = True) As Boolean
    Dim r As Long, cnt As Long, code As String, blocked As String
    Select Case sh.Name
    Case SH_MOVE
        If confirm Then
            If Not Ask("Удалить строки журнала " & r1 & "–" & r2 & "?" & vbLf & vbLf & _
                       "Они будут перенесены на лист «" & SH_DELMOVE & "» и их можно восстановить." & vbLf & _
                       "Остатки пересчитаются.") Then Exit Function
        End If
        MakeBackup "delete_movement"
        AppOff
        For r = r2 To r1 Step -1
            If Len(Txt(sh.Cells(r, J_OP).Value)) > 0 Then
                CopyRowTo sh, r, WS(SH_DELMOVE), J_LAST
                cnt = cnt + 1
            End If
            sh.Rows(r).Delete
        Next r
        AppOn
        Info "Удалено строк журнала: " & cnt & ". Восстановление — лист «" & SH_DELMOVE & "»."
    Case SH_NOM
        For r = r2 To r1 Step -1
            code = Txt(sh.Cells(r, N_CODE).Value)
            If Len(code) > 0 Then
                If Not NomDeleteCore(code, confirm) Then blocked = blocked & code & " "
            Else
                sh.Rows(r).Delete
            End If
        Next r
        If Len(blocked) > 0 Then Info "Не удалены (есть движения или отказ): " & blocked
    Case SH_INPUT
        AppOff
        sh.Range(sh.Cells(r1, 1), sh.Cells(r2, 1)).EntireRow.Delete
        EnsureInputRows IN_FIRST + 49
        RenumberInput
        AppOn
    Case SH_TRASH, SH_DELMOVE
        If confirm Then
            If Not Ask("Удалить строки корзины НАВСЕГДА?") Then Exit Function
        End If
        sh.Range(sh.Cells(r1, 1), sh.Cells(r2, 1)).EntireRow.Delete
    Case Else
        If confirm Then
            If Not Ask("Удалить строки " & r1 & "–" & r2 & " листа «" & sh.Name & "»?" & vbLf & _
                       "Перед удалением будет сделана резервная копия книги.") Then Exit Function
        End If
        MakeBackup "delete_" & sh.Name
        sh.Range(sh.Cells(r1, 1), sh.Cells(r2, 1)).EntireRow.Delete
    End Select
    DeleteRowsCore = True
End Function

' Копировать строку в конец листа-корзины с отметкой времени и пользователя
Private Sub CopyRowTo(ByVal src As Worksheet, ByVal r As Long, ByVal dst As Worksheet, ByVal nCols As Long)
    Dim t As Long, v As Variant
    t = LastRow(dst, 1) + 1
    If t < FIRST_ROW Then t = FIRST_ROW
    v = src.Range(src.Cells(r, 1), src.Cells(r, nCols)).Value
    dst.Range(dst.Cells(t, 1), dst.Cells(t, nCols)).Value = v
    dst.Cells(t, nCols + 1).Value = Now
    dst.Cells(t, nCols + 1).NumberFormat = "dd.mm.yyyy hh:mm"
    dst.Cells(t, nCols + 2).Value = UserName()
End Sub

' ---------------- номенклатура: удалить в корзину / восстановить ----------------
Public Function NomHasMovements(ByVal code As String) As Boolean
    Dim sh As Worksheet
    Set sh = WS(SH_MOVE)
    If FindRow(sh, J_CODE, code) > 0 Then NomHasMovements = True: Exit Function
    If FindRow(sh, J_ITEM, code) > 0 Then NomHasMovements = True: Exit Function
    If FindRow(WS(SH_SPEC), 1, code) > 0 Or FindRow(WS(SH_SPEC), 2, code) > 0 Then NomHasMovements = True: Exit Function
    If FindRow(WS(SH_BUNDLE), 1, code) > 0 Or FindRow(WS(SH_BUNDLE), 2, code) > 0 Then NomHasMovements = True
End Function

Public Function NomDeleteCore(ByVal code As String, Optional ByVal confirm As Boolean = True) As Boolean
    Dim sh As Worksheet, r As Long
    Set sh = WS(SH_NOM)
    r = NomRow(code)
    If r = 0 Then Fail "Код «" & code & "» не найден.": Exit Function
    If NomHasMovements(code) Then
        If Ask("По «" & code & "» есть движения или он входит в состав/спецификацию." & vbLf & _
               "Удалять такую позицию нельзя — история остатков сломается." & vbLf & vbLf & _
               "Сделать позицию НЕАКТИВНОЙ (Активно = Нет)? Она останется в истории, но не будет доступна для новых операций.") Then
            sh.Cells(r, N_ACTIVE).Value = "Нет"
        End If
        Exit Function
    End If
    If confirm Then
        If Not Ask("Перенести «" & code & " — " & Txt(sh.Cells(r, N_NAME).Value) & "» в корзину?") Then Exit Function
    End If
    MakeBackup "delete_nom"
    CopyRowTo sh, r, WS(SH_TRASH), N_LAST
    sh.Rows(r).Delete
    NomDeleteCore = True
End Function

Public Function NomRestoreCore(ByVal trashRow As Long) As Boolean
    Dim shT As Worksheet, shN As Worksheet, code As String, t As Long, v As Variant
    Set shT = WS(SH_TRASH): Set shN = WS(SH_NOM)
    code = Txt(shT.Cells(trashRow, N_CODE).Value)
    If Len(code) = 0 Then Exit Function
    If NomRow(code) > 0 Then Fail "Код «" & code & "» уже есть в справочнике — восстановление отменено.": Exit Function
    t = LastRow(shN, N_CODE) + 1
    v = shT.Range(shT.Cells(trashRow, 1), shT.Cells(trashRow, N_LAST)).Value
    shN.Range(shN.Cells(t, 1), shN.Cells(t, N_LAST)).Value = v
    If t > FIRST_ROW Then
        CopyRowFormat shN, t - 1, t, t, N_LAST
    End If
    shT.Rows(trashRow).Delete
    NomRestoreCore = True
End Function

Public Function MoveRestoreCore(ByVal delRow As Long) As Boolean
    Dim shD As Worksheet, shM As Worksheet, t As Long, v As Variant
    Set shD = WS(SH_DELMOVE): Set shM = WS(SH_MOVE)
    If Len(Txt(shD.Cells(delRow, J_OP).Value)) = 0 Then Exit Function
    t = LastRow(shM, J_OP) + 1
    v = shD.Range(shD.Cells(delRow, 1), shD.Cells(delRow, J_LAST)).Value
    shM.Range(shM.Cells(t, 1), shM.Cells(t, J_LAST)).Value = v
    EnsureJournalFormulas t, t
    FormatJournalRows t, t
    shD.Rows(delRow).Delete
    MoveRestoreCore = True
End Function

Public Sub Rows_RestoreSelected()
    Dim sh As Worksheet, r As Long, r1 As Long, r2 As Long, cnt As Long
    Set sh = ActiveSheet
    If sh.Name <> SH_TRASH And sh.Name <> SH_DELMOVE Then
        Fail "Перейдите на лист «" & SH_TRASH & "» (номенклатура) или «" & SH_DELMOVE & "» (строки журнала) и выделите строки.": Exit Sub
    End If
    r1 = Selection.Row: r2 = r1 + Selection.Rows.Count - 1
    If r1 < FIRST_ROW Then r1 = FIRST_ROW
    AppOff
    For r = r2 To r1 Step -1
        If sh.Name = SH_TRASH Then
            If NomRestoreCore(r) Then cnt = cnt + 1
        Else
            If MoveRestoreCore(r) Then cnt = cnt + 1
        End If
    Next r
    AppOn
    Info "Восстановлено строк: " & cnt
End Sub

Public Sub Nom_DeleteSelected()
    If ActiveSheet.Name <> SH_NOM Then Fail "Перейдите на лист «" & SH_NOM & "» и выделите позицию.": Exit Sub
    Rows_DeleteSelected
End Sub
