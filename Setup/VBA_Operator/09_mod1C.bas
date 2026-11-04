Attribute VB_Name = "mod1C"
Option Explicit
' ============================================================
' ELKA v8.2 — ИМПОРТ ОСТАТКОВ ИЗ 1С И СВЕРКА
' 1) В 1С сформируйте отчёт «Остатки товаров» (или «Ведомость
'    по товарам на складах») и сохраните его в Excel (.xlsx/.xls).
' 2) Кнопка «Загрузить из 1С» -> выберите файл.
'    Макрос сам найдёт строку заголовков и столбцы:
'      код:        «Артикул», «Код», «Код номенклатуры»
'      количество: «Остаток», «Конечный остаток», «Количество»
'      склад:      «Склад» (необязательно)
'      название:   «Номенклатура», «Наименование» (необязательно)
' 3) Склады 1С сопоставляются с местами книги по столбцу
'    «Склад в 1С» листа «Места хранения» (несколько — через «;»).
' 4) Результат — лист «Сравнение_1С»: разница = книга - 1С.
' ============================================================

Public Sub Import1C_Dialog()
    Dim f As Variant
    f = Application.GetOpenFilename("Файлы Excel (*.xlsx;*.xls;*.xlsm),*.xlsx;*.xls;*.xlsm,CSV (*.csv),*.csv", 1, "Выберите выгрузку остатков из 1С")
    If VarType(f) = vbBoolean Then Exit Sub
    If Import1C_Load(CStr(f)) Then
        Compare1C
        If Not gSilent Then WS(SH_CMP).Activate
        Info "Загрузка и сверка выполнены. Результат — лист «" & SH_CMP & "»." & vbLf & _
             "Расхождений: " & CountMismatch() & ".", "Сверка с 1С"
    End If
End Sub

Private Function HeaderKind(ByVal t As String) As String
    t = LCase$(Trim$(Replace(t, vbLf, " ")))
    If Len(t) = 0 Then Exit Function
    If t = "артикул" Or t = "код" Or t = "код номенклатуры" Or t = "номенклатура.код" Or t = "номенклатура.артикул" Or InStr(t, "артикул") > 0 Then HeaderKind = "code": Exit Function
    If InStr(t, "остат") > 0 Or InStr(t, "колич") > 0 Or t = "кол-во" Then HeaderKind = "qty": Exit Function
    If InStr(t, "склад") > 0 Then HeaderKind = "wh": Exit Function
    If InStr(t, "наимен") > 0 Or t = "номенклатура" Or t = "товар" Then HeaderKind = "name": Exit Function
    If t = "ед." Or t = "ед" Or InStr(t, "единиц") > 0 Then HeaderKind = "unit": Exit Function
End Function

Public Function Import1C_Load(ByVal path As String) As Boolean
    Dim outArr() As Variant
    Dim wb As Workbook, src As Worksheet, v As Variant, r As Long, c As Long, hr As Long
    Dim cCode As Long, cQty As Long, cWh As Long, cName As Long, cUnit As Long
    Dim sh As Worksheet, outR As Long, n As Long, k As String, lastWh As String
    Dim arr() As Variant, stamp As Date

    On Error GoTo Bad
    AppOff
    Set wb = Workbooks.Open(Filename:=path, ReadOnly:=True, UpdateLinks:=0)
    Set src = wb.Worksheets(1)
    v = src.UsedRange.Value
    ' поиск строки заголовков в первых 40 строках
    For r = 1 To Application.Min(40, UBound(v, 1))
        cCode = 0: cQty = 0: cWh = 0: cName = 0: cUnit = 0
        For c = 1 To UBound(v, 2)
            Select Case HeaderKind(Txt(v(r, c)))
            Case "code": If cCode = 0 Then cCode = c
            Case "qty": cQty = c         ' берём ПОСЛЕДНИЙ столбец количества (конечный остаток)
            Case "wh": If cWh = 0 Then cWh = c
            Case "name": If cName = 0 Then cName = c
            Case "unit": If cUnit = 0 Then cUnit = c
            End Select
        Next c
        If cCode > 0 And cQty > 0 Then hr = r: Exit For
    Next r
    If hr = 0 Then
        wb.Close False
        AppOn
        Fail "В файле не найдены заголовки «Артикул/Код» и «Остаток/Количество» в первых 40 строках." & vbLf & _
             "Переименуйте заголовки в выгрузке или сохраните отчёт 1С с шапкой таблицы."
        Exit Function
    End If
    stamp = Now
    ReDim arr(1 To UBound(v, 1), 1 To 7)
    For r = hr + 1 To UBound(v, 1)
        k = Txt(v(r, cCode))
        If cWh > 0 Then
            If Len(Txt(v(r, cWh))) > 0 And Len(k) = 0 Then lastWh = Txt(v(r, cWh))   ' группировка по складам в 1С
        End If
        If Len(k) > 0 And IsNumeric(v(r, cQty)) And LCase$(k) <> "итого" Then
            n = n + 1
            arr(n, 1) = k
            If cName > 0 Then arr(n, 2) = Txt(v(r, cName))
            If cWh > 0 Then
                arr(n, 3) = Txt(v(r, cWh))
                If Len(arr(n, 3)) = 0 Then arr(n, 3) = lastWh
            End If
            arr(n, 4) = CDbl(v(r, cQty))
            If cUnit > 0 Then arr(n, 5) = Txt(v(r, cUnit))
            arr(n, 6) = FileNameOnly(path)
            arr(n, 7) = stamp
        End If
    Next r
    wb.Close False
    Set wb = Nothing
    ThisWorkbook.Activate

    Set sh = WS(SH_1C)
    outR = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    If outR < FIRST_ROW Then outR = FIRST_ROW
    sh.Range(sh.Cells(FIRST_ROW, 1), sh.Cells(outR + 1, 7)).Clear
    If n > 0 Then
        ReDim outArr(1 To n, 1 To 7)
        For r = 1 To n
            For c = 1 To 7: outArr(r, c) = arr(r, c): Next c
        Next r
        sh.Range(sh.Cells(FIRST_ROW, 1), sh.Cells(FIRST_ROW + n - 1, 7)).Value = outArr
        sh.Range(sh.Cells(FIRST_ROW, 7), sh.Cells(FIRST_ROW + n - 1, 7)).NumberFormat = "dd.mm.yyyy hh:mm"
        sh.Range(sh.Cells(FIRST_ROW, 4), sh.Cells(FIRST_ROW + n - 1, 4)).NumberFormat = "# ##0.###"
        StyleReport sh, FIRST_ROW, FIRST_ROW + n - 1, 7
    End If
    AppOn
    Import1C_Load = (n > 0)
    If n = 0 Then Fail "В файле нет строк с кодом и числовым остатком."
    Exit Function
Bad:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    AppOn
    Fail "Не удалось прочитать файл 1С: " & Err.Description
End Function

' Место книги по названию склада 1С ("" — не сопоставлен, "*" — все учитываемые)
Public Function MapWarehouse(ByVal wh As String) As String
    Dim sh As Worksheet, r As Long, lastR As Long, parts() As String, i As Long
    If Len(Trim$(wh)) = 0 Then MapWarehouse = "*": Exit Function
    Set sh = WS(SH_LOC)
    lastR = LastRow(sh, 1)
    For r = FIRST_ROW To lastR
        parts = Split(Txt(sh.Cells(r, 5).Value), ";")
        For i = 0 To UBound(parts)
            If StrComp(Trim$(parts(i)), Trim$(wh), vbTextCompare) = 0 Then MapWarehouse = Txt(sh.Cells(r, 1).Value): Exit Function
        Next i
        If StrComp(Txt(sh.Cells(r, 1).Value), Trim$(wh), vbTextCompare) = 0 Then MapWarehouse = Txt(sh.Cells(r, 1).Value): Exit Function
    Next r
End Function

Public Sub Compare1C()
    Dim outArr() As Variant
    Dim sh1 As Worksheet, shC As Worksheet, lastR As Long, r As Long, i As Long, n As Long
    Dim keys As New Collection, kCode() As String, kWh() As String, kLoc() As String, kQ() As Double, kName() As String
    Dim code As String, wh As String, loc As String, key As String, bookQ As Double, diff As Double, tol As Double
    Dim arr() As Variant, outN As Long, hasWh As Boolean, shN As Worksheet, locs As Variant, j As Long
    Set sh1 = WS(SH_1C): Set shC = WS(SH_CMP): Set shN = WS(SH_NOM)
    tol = Nz(Setting("Допуск сравнения 1С", "0"))
    BuildBalances
    lastR = LastRow(sh1, 1)
    ReDim kCode(1 To lastR + 1): ReDim kWh(1 To lastR + 1): ReDim kLoc(1 To lastR + 1): ReDim kQ(1 To lastR + 1): ReDim kName(1 To lastR + 1)
    For r = FIRST_ROW To lastR
        code = Txt(sh1.Cells(r, 1).Value)
        If Len(code) > 0 Then
            wh = Txt(sh1.Cells(r, 3).Value)
            If Len(wh) > 0 Then hasWh = True
            loc = MapWarehouse(wh)
            key = code & "|" & wh
            i = KeyIndex(keys, key)
            If i = 0 Then
                n = n + 1: i = n: KeyAdd keys, key, n
                kCode(n) = code: kWh(n) = wh: kLoc(n) = loc: kName(n) = Txt(sh1.Cells(r, 2).Value)
            End If
            kQ(i) = kQ(i) + Nz(sh1.Cells(r, 4).Value)
        End If
    Next r
    ReDim arr(1 To n + LastRow(shN, N_CODE) * 10 + 10, 1 To 9)
    For i = 1 To n
        outN = outN + 1
        arr(outN, 1) = kCode(i)
        arr(outN, 2) = IIf(NomRow(kCode(i)) > 0, NomName(kCode(i)), kName(i))
        arr(outN, 3) = kWh(i)
        arr(outN, 5) = kQ(i)
        If NomRow(kCode(i)) = 0 Then
            arr(outN, 4) = "—": arr(outN, 6) = 0: arr(outN, 7) = -kQ(i)
            arr(outN, 8) = "НЕТ В КНИГЕ": arr(outN, 9) = "Заведите позицию в «Справочник» или исправьте артикул в 1С."
        ElseIf kLoc(i) = "" Then
            arr(outN, 4) = "?": arr(outN, 8) = "СКЛАД НЕ СОПОСТАВЛЕН"
            arr(outN, 9) = "Впишите «" & kWh(i) & "» в столбец «Склад в 1С» нужного места на листе «Места хранения» и повторите сверку."
        Else
            If kLoc(i) = "*" Then
                bookQ = BalGet("T|" & UCase$(kCode(i))): arr(outN, 4) = "все учитываемые"
            Else
                bookQ = BalGet("L|" & UCase$(kCode(i)) & "|" & UCase$(kLoc(i))): arr(outN, 4) = kLoc(i)
            End If
            diff = bookQ - kQ(i)
            arr(outN, 6) = bookQ: arr(outN, 7) = diff
            If Abs(diff) <= tol + 0.0000001 Then
                arr(outN, 8) = "СОВПАДАЕТ"
            Else
                arr(outN, 8) = IIf(diff > 0, "В КНИГЕ БОЛЬШЕ", "В КНИГЕ МЕНЬШЕ")
                arr(outN, 9) = "Проверьте операции после даты выгрузки 1С, склад и единицы измерения."
            End If
        End If
    Next i
    ' позиции с остатком в книге, которых нет в 1С
    locs = TrackedLocs()
    For r = FIRST_ROW To LastRow(shN, N_CODE)
        code = Txt(shN.Cells(r, N_CODE).Value)
        If Len(code) > 0 Then
            If hasWh Then
                For j = 0 To UBound(locs)
                    bookQ = BalGet("L|" & UCase$(code) & "|" & UCase$(locs(j)))
                    If Abs(bookQ) > tol + 0.0000001 And Not CodeLocIn1C(code, CStr(locs(j)), kCode, kLoc, n) Then
                        If Len(Txt(WS(SH_LOC).Cells(FindRow(WS(SH_LOC), 1, CStr(locs(j))), 5).Value)) > 0 Then
                            outN = outN + 1
                            arr(outN, 1) = code: arr(outN, 2) = NomName(code): arr(outN, 3) = "": arr(outN, 4) = locs(j)
                            arr(outN, 5) = 0: arr(outN, 6) = bookQ: arr(outN, 7) = bookQ
                            arr(outN, 8) = "НЕТ В 1С": arr(outN, 9) = "Позиция есть в книге, но отсутствует в выгрузке 1С по этому складу."
                        End If
                    End If
                Next j
            Else
                bookQ = BalGet("T|" & UCase$(code))
                If Abs(bookQ) > tol + 0.0000001 And KeyIndex(keys, code & "|") = 0 Then
                    outN = outN + 1
                    arr(outN, 1) = code: arr(outN, 2) = NomName(code): arr(outN, 4) = "все учитываемые"
                    arr(outN, 5) = 0: arr(outN, 6) = bookQ: arr(outN, 7) = bookQ
                    arr(outN, 8) = "НЕТ В 1С": arr(outN, 9) = "Позиция есть в книге, но отсутствует в выгрузке 1С."
                End If
            End If
        End If
    Next r
    AppOff
    lastR = shC.Cells(shC.Rows.Count, 1).End(xlUp).Row
    If lastR < FIRST_ROW Then lastR = FIRST_ROW
    shC.Range(shC.Cells(FIRST_ROW, 1), shC.Cells(lastR + 1, 9)).Clear
    If outN > 0 Then
        ReDim outArr(1 To outN, 1 To 9)
        For r = 1 To outN
            For j = 1 To 9: outArr(r, j) = arr(r, j): Next j
        Next r
        shC.Range(shC.Cells(FIRST_ROW, 1), shC.Cells(FIRST_ROW + outN - 1, 9)).Value = outArr
        StyleReport shC, FIRST_ROW, FIRST_ROW + outN - 1, 9
        shC.Range(shC.Cells(FIRST_ROW, 5), shC.Cells(FIRST_ROW + outN - 1, 7)).NumberFormat = "# ##0.###;[Red]-# ##0.###"
        For r = FIRST_ROW To FIRST_ROW + outN - 1
            If Txt(shC.Cells(r, 8).Value) = "СОВПАДАЕТ" Then
                shC.Cells(r, 8).Interior.Color = RGB(226, 240, 217)
            Else
                shC.Cells(r, 8).Interior.Color = RGB(244, 204, 204)
            End If
        Next r
    End If
    shC.Range("B2").Value = "Сверка выполнена " & Format$(Now, "dd.mm.yyyy hh:mm") & ". Разница = книга - 1С. Допуск: " & tol & "."
    AppOn
End Sub

Private Function CodeLocIn1C(ByVal code As String, ByVal loc As String, kCode() As String, kLoc() As String, ByVal n As Long) As Boolean
    Dim i As Long
    For i = 1 To n
        If StrComp(kCode(i), code, vbTextCompare) = 0 And StrComp(kLoc(i), loc, vbTextCompare) = 0 Then CodeLocIn1C = True: Exit Function
    Next i
End Function

Public Function CountMismatch() As Long
    Dim sh As Worksheet, r As Long
    Set sh = WS(SH_CMP)
    For r = FIRST_ROW To LastRow(sh, 1)
        If Len(Txt(sh.Cells(r, 8).Value)) > 0 And Txt(sh.Cells(r, 8).Value) <> "СОВПАДАЕТ" Then CountMismatch = CountMismatch + 1
    Next r
End Function

Public Sub Compare1C_Run()
    Compare1C
    Info "Сверка пересчитана. Расхождений: " & CountMismatch() & ".", "Сверка с 1С"
End Sub
