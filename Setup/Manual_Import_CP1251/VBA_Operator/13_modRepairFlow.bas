Attribute VB_Name = "modRepairFlow"
Option Explicit
' ELKA v8.2. Идемпотентное добавление ремонтного отдела в книги структуры v8.1/8.2.
' Не меняет проведённые движения, номенклатуру и существующие маршруты.
Public Function Install_RepairDepartment() As String
    Dim sh As Worksheet, r As Long, c As Long, i As Long, a As Variant, n As Long
    On Error GoTo Bad
    Set sh = WS(SH_LOC)
    r = FindRow(sh, 1, "РЕМ")
    If r > 0 Then
        If Not IsYes(sh.Cells(r, 4).Value) Then Err.Raise 5, , "Место РЕМ уже есть, но не учитывается в остатках. Проверьте справочник."
    Else
        n = 0
        For i = FIRST_ROW To LastRow(sh, 1)
            If IsYes(sh.Cells(i, 4).Value) Then n = n + 1
        Next i
        If n >= STOCK_LOC_COLS Then Err.Raise 5, , "Нет свободного столбца в Остатках для РЕМ."
    End If
    CheckRepairOp "45", "ПР", "РЕМ"
    CheckRepairOp "46", "РЕМ", "ПР"
    CheckRepairOp "47", "ИНВЕНТ", "РЕМ"
    For Each a In Array("РЕМ-01", "РЕМ-02")
        i = FindRow(WS(SH_ADDR), 1, CStr(a))
        If i > 0 Then
            If Txt(WS(SH_ADDR).Cells(i, 2).Value) <> "РЕМ" Then Err.Raise 5, , "Адрес " & a & " уже занят другим местом."
        End If
    Next a
    If r = 0 Then
        r = LastRow(sh, 1) + 1
        a = Array("РЕМ", "Ремонтный отдел", "Производство", "Да", "Ремонтный отдел", "Мастер ремонтного отдела", "Из производства в ремонт и обратно")
        For c = 0 To 6: sh.Cells(r, c + 1).Value = a(c): Next c
    End If
    For i = FIRST_ROW To LastRow(sh, 1)
        sh.Cells(i, 8).Formula = "=IF(D" & i & "=""Да"",COUNTIF($D$4:D" & i & ",""Да""),"""")"
    Next i
    For Each a In Array("РЕМ-01", "РЕМ-02")
        If FindRow(WS(SH_ADDR), 1, CStr(a)) = 0 Then
            i = LastRow(WS(SH_ADDR), 1) + 1
            WS(SH_ADDR).Cells(i, 1).Value = a
            WS(SH_ADDR).Cells(i, 2).Value = "РЕМ"
            WS(SH_ADDR).Cells(i, 3).Value = "Ремонт"
            WS(SH_ADDR).Cells(i, 8).Value = "Да"
        End If
    Next a
    AddRepairOp "45", "Передача в ремонт (ПР -> РЕМ)", "ПР", "РЕМ", "Да", "Передача из производства в ремонтный отдел"
    AddRepairOp "46", "Возврат из ремонта (РЕМ -> ПР)", "РЕМ", "ПР", "Да", "Возврат того же артикула после ремонта"
    AddRepairOp "47", "Входящий остаток РЕМ (свёртка периода)", "ИНВЕНТ", "РЕМ", "Нет", "Служебная операция входящего остатка"
    For c = 5 To 4 + STOCK_LOC_COLS
        WS(SH_STOCK).Cells(HDR_ROW, c).Formula = "=IFERROR(INDEX('Места хранения'!$A:$A,MATCH(" & (c - 4) & ",'Места хранения'!$H:$H,0)),"""")"
    Next c
    RefreshStockRows
    Application.Calculate
    Install_RepairDepartment = "OK"
    Exit Function
Bad:
    Install_RepairDepartment = "ERROR: " & Err.Description
End Function
Private Sub CheckRepairOp(ByVal code As String, ByVal fromLoc As String, ByVal toLoc As String)
    Dim r As Long
    r = FindRow(WS(SH_OPS), 1, code)
    If r = 0 Then Exit Sub
    If Txt(WS(SH_OPS).Cells(r, 3).Value) <> fromLoc Or Txt(WS(SH_OPS).Cells(r, 4).Value) <> toLoc Then
        Err.Raise 5, , "Код операции " & code & " уже имеет другой маршрут. Обновление остановлено."
    End If
End Sub
Private Sub AddRepairOp(ByVal code As String, ByVal title As String, ByVal fromLoc As String, ByVal toLoc As String, ByVal needOrder As String, ByVal note As String)
    Dim r As Long, c As Long, a As Variant
    If FindRow(WS(SH_OPS), 1, code) > 0 Then Exit Sub
    r = LastRow(WS(SH_OPS), 1) + 1
    a = Array(code, code & " " & title, fromLoc, toLoc, needOrder, "Нет", "Нет", "Ремонт", "Да", note)
    For c = 0 To 9: WS(SH_OPS).Cells(r, c + 1).Value = a(c): Next c
End Sub
