Attribute VB_Name = "modReports"
Option Explicit
' ============================================================
' ELKA v8.2 — ОТЧЁТЫ: остатки, адреса, варианты, комплектность
' Лист «Остатки» считается формулами (живой). Остальные отчёты
' строятся кнопкой «Обновить отчёты» (и перед публикацией).
' ============================================================
Public Const STOCK_LOC_COLS As Long = 12     ' столбцы мест хранения E..P на листе «Остатки»

Public Sub Reports_RefreshAll()
    AppOff
    On Error GoTo Bad
    RefreshStockRows
    BuildAddrStock
    BuildVariantSummary
    BuildKits
    AppOn
    If Not gSilent Then Application.Calculate
    Info "Отчёты обновлены: «" & SH_STOCK & "», «" & SH_ADDRSTOCK & "», «" & SH_VARSUM & "», «" & SH_KITS & "»."
    Exit Sub
Bad:
    AppOn
    Fail "Ошибка обновления отчётов: " & Err.Description
End Sub

' Формулы строки листа «Остатки»
Public Function StockFormula(ByVal col As Long, ByVal r As Long) As String
    Dim f As String, cl As String
    Select Case col
    Case 1: f = "=IF(INDEX(Справочник!$A:$A,ROW())="""","""",INDEX(Справочник!$A:$A,ROW()))"
    Case 2: f = "=IF($A{r}="""","""",INDEX(Справочник!$B:$B,ROW()))"
    Case 3: f = "=IF($A{r}="""","""",INDEX(Справочник!$C:$C,ROW()))"
    Case 4: f = "=IF($A{r}="""","""",INDEX(Справочник!$D:$D,ROW()))"
    Case 5 To 4 + STOCK_LOC_COLS
        cl = ColLetter(col)
        f = "=IF(OR($A{r}=""""," & cl & "$3=""""),"""",SUMIFS(Движение!$F:$F,Движение!$D:$D,$A{r},Движение!$I:$I," & cl & "$3)" & _
            "-SUMIFS(Движение!$F:$F,Движение!$D:$D,$A{r},Движение!$H:$H," & cl & "$3))"
    Case 5 + STOCK_LOC_COLS: f = "=IF($A{r}="""","""",SUM(E{r}:" & ColLetter(4 + STOCK_LOC_COLS) & "{r}))"
    Case 6 + STOCK_LOC_COLS: f = "=IF($A{r}="""","""",N(INDEX(Справочник!$H:$H,ROW())))"
    Case 7 + STOCK_LOC_COLS
        f = "=IF($A{r}="""","""",IF($C{r}=""Вариантный родитель"",""—"",IF(Q{r}<0,""МИНУС"",IF(AND(R{r}>0,Q{r}<R{r}),""НИЖЕ МИНИМУМА"",""OK""))))"
    End Select
    StockFormula = Replace(f, "{r}", CStr(r))
End Function

Public Function ColLetter(ByVal c As Long) As String
    Dim t As String
    Do While c > 0
        t = Chr$(65 + ((c - 1) Mod 26)) & t
        c = (c - 1) \ 26
    Loop
    ColLetter = t
End Function

' Протянуть формулы «Остатков» до конца справочника (+50 строк запаса)
Public Sub RefreshStockRows()
    Dim sh As Worksheet, need As Long, have As Long, r As Long, c As Long, arr() As Variant
    Set sh = WS(SH_STOCK)
    need = LastRow(WS(SH_NOM), N_CODE) + 50
    have = sh.Cells(sh.Rows.Count, 2).End(xlUp).Row
    If sh.Cells(have, 1).HasFormula = False And have >= FIRST_ROW Then have = FIRST_ROW - 1
    If have < FIRST_ROW Then have = FIRST_ROW - 1
    ' ищем последнюю строку с формулой в A
    r = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    Do While r >= FIRST_ROW
        If sh.Cells(r, 1).HasFormula Then Exit Do
        r = r - 1
    Loop
    have = r
    If have >= need Then Exit Sub
    ReDim arr(1 To need - have, 1 To 7 + STOCK_LOC_COLS)
    For r = have + 1 To need
        For c = 1 To 7 + STOCK_LOC_COLS
            arr(r - have, c) = StockFormula(c, r)
        Next c
    Next r
    sh.Range(sh.Cells(have + 1, 1), sh.Cells(need, 7 + STOCK_LOC_COLS)).Formula = arr
    If have >= FIRST_ROW Then
        CopyRowFormat sh, have, have + 1, need, 7 + STOCK_LOC_COLS
    End If
End Sub

' Учитываемые места (в порядке листа «Места хранения»)
Public Function TrackedLocs() As Variant
    Dim sh As Worksheet, r As Long, lastR As Long, t As String
    Set sh = WS(SH_LOC)
    lastR = LastRow(sh, 1)
    For r = FIRST_ROW To lastR
        If IsYes(sh.Cells(r, 4).Value) Then
            If Len(t) > 0 Then t = t & "|"
            t = t & Txt(sh.Cells(r, 1).Value)
        End If
    Next r
    TrackedLocs = Split(t, "|")
End Function

Private Sub ClearReport(ByVal sh As Worksheet, ByVal nCols As Long)
    Dim lastR As Long
    lastR = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    If lastR < FIRST_ROW Then lastR = FIRST_ROW
    sh.Range(sh.Cells(FIRST_ROW, 1), sh.Cells(lastR + 1, nCols)).Clear
End Sub

' ---------------- остатки по адресам ----------------
Public Sub BuildAddrStock()
    Dim ra As Long
    Dim outArr() As Variant
    Dim sh As Worksheet, shN As Worksheet, shA As Worksheet, locs As Variant
    Dim r As Long, lastN As Long, lastA As Long, i As Long, code As String, b As Double
    Dim outR As Long, arr() As Variant, n As Long, loc As String, a As String
    Set sh = WS(SH_ADDRSTOCK): Set shN = WS(SH_NOM): Set shA = WS(SH_ADDR)
    BuildBalances
    locs = TrackedLocs()
    lastN = LastRow(shN, N_CODE): lastA = LastRow(shA, 1)
    ReDim arr(1 To 6, 1 To 1)
    For r = FIRST_ROW To lastN
        code = Txt(shN.Cells(r, N_CODE).Value)
        If Len(code) > 0 Then
            For i = 0 To UBound(locs)
                loc = locs(i)
                ' адреса этого места
                For ra = FIRST_ROW To lastA
                    If StrComp(Txt(shA.Cells(ra, 2).Value), loc, vbTextCompare) = 0 Then
                        a = Txt(shA.Cells(ra, 1).Value)
                        b = BalGet("A|" & UCase$(code) & "|" & UCase$(loc) & "|" & UCase$(a))
                        If Abs(b) > 0.0000001 Then
                            n = n + 1: ReDim Preserve arr(1 To 6, 1 To n)
                            arr(1, n) = code: arr(2, n) = Txt(shN.Cells(r, N_NAME).Value)
                            arr(3, n) = loc: arr(4, n) = a: arr(5, n) = b: arr(6, n) = Txt(shN.Cells(r, N_UNIT).Value)
                        End If
                    End If
                Next ra
                b = BalGet("A|" & UCase$(code) & "|" & UCase$(loc) & "|")
                If Abs(b) > 0.0000001 Then
                    n = n + 1: ReDim Preserve arr(1 To 6, 1 To n)
                    arr(1, n) = code: arr(2, n) = Txt(shN.Cells(r, N_NAME).Value)
                    arr(3, n) = loc: arr(4, n) = "(без адреса)": arr(5, n) = b: arr(6, n) = Txt(shN.Cells(r, N_UNIT).Value)
                End If
            Next i
        End If
    Next r
    ClearReport sh, 6
    If n > 0 Then
        ReDim outArr(1 To n, 1 To 6)
        For r = 1 To n
            For i = 1 To 6: outArr(r, i) = arr(i, r): Next i
        Next r
        sh.Range(sh.Cells(FIRST_ROW, 1), sh.Cells(FIRST_ROW + n - 1, 6)).Value = outArr
        StyleReport sh, FIRST_ROW, FIRST_ROW + n - 1, 6
        sh.Range(sh.Cells(FIRST_ROW, 5), sh.Cells(FIRST_ROW + n - 1, 5)).NumberFormat = "# ##0.###"
    End If
End Sub

Public Sub StyleReport(ByVal sh As Worksheet, ByVal r1 As Long, ByVal r2 As Long, ByVal nCols As Long)
    On Error Resume Next
    With sh.Range(sh.Cells(r1, 1), sh.Cells(r2, nCols))
        .Font.Name = "Arial"
        .Font.Size = 10
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(191, 191, 191)
        .RowHeight = 20
    End With
End Sub

' ---------------- свод вариантов ----------------
Public Sub BuildVariantSummary()
    Dim sh As Worksheet, shN As Worksheet, locs As Variant, lastN As Long, r As Long, r2 As Long
    Dim outR As Long, i As Long, famCode As String, code As String, tot As Double, famTot As Double, nc As Long, famStart As Long
    Set sh = WS(SH_VARSUM): Set shN = WS(SH_NOM)
    BuildBalances
    locs = TrackedLocs()
    nc = 6 + UBound(locs) + 1
    ClearReport sh, 30
    ' заголовки мест
    sh.Range(sh.Cells(HDR_ROW, 7), sh.Cells(HDR_ROW, 30)).ClearContents
    For i = 0 To UBound(locs)
        sh.Cells(HDR_ROW, 7 + i).Value = locs(i)
    Next i
    lastN = LastRow(shN, N_CODE)
    outR = FIRST_ROW
    For r = FIRST_ROW To lastN
        If Txt(shN.Cells(r, N_TYPE).Value) = T_VARPARENT Then
            famCode = Txt(shN.Cells(r, N_CODE).Value)
            famStart = outR
            sh.Cells(outR, 1).Value = famCode
            sh.Cells(outR, 2).Value = Txt(shN.Cells(r, N_NAME).Value)
            sh.Cells(outR, 3).Value = "ИТОГО СЕМЕЙСТВО"
            sh.Range(sh.Cells(outR, 1), sh.Cells(outR, nc)).Font.Bold = True
            sh.Range(sh.Cells(outR, 1), sh.Cells(outR, nc)).Interior.Color = RGB(221, 235, 247)
            outR = outR + 1
            For r2 = FIRST_ROW To lastN
                If StrComp(Txt(shN.Cells(r2, N_BASE).Value), famCode, vbTextCompare) = 0 And r2 <> r Then
                    code = Txt(shN.Cells(r2, N_CODE).Value)
                    sh.Cells(outR, 1).Value = famCode
                    sh.Cells(outR, 2).Value = Txt(shN.Cells(r, N_NAME).Value)
                    sh.Cells(outR, 3).Value = code
                    sh.Cells(outR, 4).Value = Txt(shN.Cells(r2, N_NAME).Value)
                    sh.Cells(outR, 5).Value = Txt(shN.Cells(r2, N_VARIANT).Value)
                    tot = BalGet("T|" & UCase$(code))
                    sh.Cells(outR, 6).Value = tot
                    For i = 0 To UBound(locs)
                        sh.Cells(outR, 7 + i).Value = BalGet("L|" & UCase$(code) & "|" & UCase$(locs(i)))
                        sh.Cells(famStart, 7 + i).Value = Nz(sh.Cells(famStart, 7 + i).Value) + Nz(sh.Cells(outR, 7 + i).Value)
                    Next i
                    sh.Cells(famStart, 6).Value = Nz(sh.Cells(famStart, 6).Value) + tot
                    outR = outR + 1
                End If
            Next r2
        End If
    Next r
    If outR > FIRST_ROW Then
        StyleReport sh, FIRST_ROW, outR - 1, nc
        sh.Range(sh.Cells(FIRST_ROW, 6), sh.Cells(outR - 1, nc)).NumberFormat = "# ##0.###"
    End If
End Sub

' ---------------- комплектность по всем изделиям ----------------
' Для каждого изделия со спецификацией: наличие компонентов в местах
' из параметра «Места для комплектности» и сколько изделий можно собрать.
Public Sub BuildKits()
    Dim li As Long
    Dim sh As Worksheet, shS As Worksheet, locs As Variant, lastS As Long, r As Long, outR As Long
    Dim parents As New Collection, plist() As String, np As Long, p As String, i As Long, j As Long
    Dim comps() As String, norms() As Double, nComp As Long, have As Double, minK As Double, k As Double
    Dim hdrRow As Long, nc As Long, tot As Double
    Set sh = WS(SH_KITS): Set shS = WS(SH_SPEC)
    BuildBalances
    locs = Split(Setting("Места для комплектности", "ЦС|ПС|ПР"), "|")
    nc = 5 + UBound(locs) + 1 + 3
    ClearReport sh, 30
    sh.Range(sh.Cells(HDR_ROW, 6), sh.Cells(HDR_ROW, 30)).ClearContents
    For j = 0 To UBound(locs)
        sh.Cells(HDR_ROW, 6 + j).Value = "Есть: " & locs(j)
    Next j
    sh.Cells(HDR_ROW, 6 + UBound(locs) + 1).Value = "Всего доступно"
    sh.Cells(HDR_ROW, 7 + UBound(locs) + 1).Value = "Хватает на, шт"
    sh.Cells(HDR_ROW, 8 + UBound(locs) + 1).Value = "Статус"
    lastS = LastRow(shS, 1)
    ReDim plist(1 To lastS + 1)
    For r = FIRST_ROW To lastS
        p = Txt(shS.Cells(r, 1).Value)
        If Len(p) > 0 Then
            If KeyIndex(parents, p) = 0 Then np = np + 1: plist(np) = p: KeyAdd parents, p, np
        End If
    Next r
    For r = FIRST_ROW To LastRow(WS(SH_BUNDLE), 1)
        p = Txt(WS(SH_BUNDLE).Cells(r, 1).Value)
        If Len(p) > 0 Then
            If KeyIndex(parents, p) = 0 Then
                np = np + 1
                If np > UBound(plist) Then
                    ReDim Preserve plist(1 To np + 50)
                End If
                plist(np) = p: KeyAdd parents, p, np
            End If
        End If
    Next r
    outR = FIRST_ROW
    For i = 1 To np
        p = plist(i)
        If NomType(p) = T_BUNDLE Then
            nComp = GetComposition(SH_BUNDLE, p, comps, norms)
        Else
            nComp = GetComposition(SH_SPEC, p, comps, norms)
        End If
        hdrRow = outR
        sh.Cells(outR, 1).Value = p
        sh.Cells(outR, 2).Value = NomName(p)
        sh.Cells(outR, 3).Value = IIf(NomType(p) = T_BUNDLE, "ОБЪЕДИНЕНИЕ", "ИЗДЕЛИЕ")
        sh.Range(sh.Cells(outR, 1), sh.Cells(outR, nc)).Font.Bold = True
        sh.Range(sh.Cells(outR, 1), sh.Cells(outR, nc)).Interior.Color = RGB(221, 235, 247)
        outR = outR + 1
        minK = -1
        For j = 1 To nComp
            sh.Cells(outR, 1).Value = p
            sh.Cells(outR, 2).Value = comps(j)
            sh.Cells(outR, 3).Value = NomName(comps(j))
            sh.Cells(outR, 4).Value = norms(j)
            sh.Cells(outR, 5).Value = Txt(WS(SH_NOM).Cells(NomRow(comps(j)), N_UNIT).Value)
            tot = 0
            For li = 0 To UBound(locs)
                have = BalGet("L|" & UCase$(comps(j)) & "|" & UCase$(locs(li)))
                sh.Cells(outR, 6 + li).Value = have
                tot = tot + have
            Next li
            sh.Cells(outR, 6 + UBound(locs) + 1).Value = tot
            If norms(j) > 0 Then k = Int(tot / norms(j) + 0.0000001) Else k = 0
            If k < 0 Then k = 0
            sh.Cells(outR, 7 + UBound(locs) + 1).Value = k
            If minK < 0 Or k < minK Then minK = k
            sh.Cells(outR, 8 + UBound(locs) + 1).Value = IIf(k >= 1, "есть", "ДЕФИЦИТ")
            outR = outR + 1
        Next j
        If minK < 0 Then minK = 0
        sh.Cells(hdrRow, 7 + UBound(locs) + 1).Value = minK
        sh.Cells(hdrRow, 8 + UBound(locs) + 1).Value = IIf(minK >= 1, "можно собрать " & minK, "НЕЛЬЗЯ СОБРАТЬ")
    Next i
    If outR > FIRST_ROW Then
        StyleReport sh, FIRST_ROW, outR - 1, nc
        sh.Range(sh.Cells(FIRST_ROW, 4), sh.Cells(outR - 1, 4)).NumberFormat = "0.###"
        sh.Range(sh.Cells(FIRST_ROW, 6), sh.Cells(outR - 1, nc - 1)).NumberFormat = "# ##0.###"
    End If
End Sub
