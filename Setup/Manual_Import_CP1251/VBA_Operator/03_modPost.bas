Attribute VB_Name = "modPost"
Option Explicit
' ============================================================
' ELKA v8.2 — ПРОВЕДЕНИЕ ОПЕРАЦИЙ ДВИЖЕНИЯ
' Документ = одна операция + любое количество строк (позиций).
' Каждая строка документа записывается в «Движение» отдельной строкой.
' ============================================================

' ---- балансы: ключ -> сумма (Collection + массив) ----
Private mKeys As Collection
Private mVals() As Double
Private mCnt As Long

Public Sub BalReset()
    Set mKeys = New Collection
    mCnt = 0
    ReDim mVals(1 To 256)
End Sub

Public Sub BalAdd(ByVal key As String, ByVal delta As Double)
    Dim i As Long
    i = KeyIndex(mKeys, key)
    If i = 0 Then
        mCnt = mCnt + 1
        If mCnt > UBound(mVals) Then
            ReDim Preserve mVals(1 To UBound(mVals) * 2)
        End If
        KeyAdd mKeys, key, mCnt
        i = mCnt
    End If
    mVals(i) = mVals(i) + delta
End Sub

Public Function BalGet(ByVal key As String) As Double
    Dim i As Long
    If mKeys Is Nothing Then Exit Function
    i = KeyIndex(mKeys, key)
    If i > 0 Then BalGet = mVals(i)
End Function

' Откуда/куда по названию операции (кэш на время одного расчёта)
Public Sub OpRoute(ByVal opName As String, ByRef fromLoc As String, ByRef toLoc As String)
    Dim r As Long
    r = FindRow(WS(SH_OPS), 2, opName)
    If r = 0 Then fromLoc = "": toLoc = "": Exit Sub
    fromLoc = Txt(WS(SH_OPS).Cells(r, 3).Value)
    toLoc = Txt(WS(SH_OPS).Cells(r, 4).Value)
End Sub

' Пересчёт всех балансов по журналу:
'   L|код|место          — остаток позиции в месте хранения
'   A|код|место|адрес    — остаток позиции по адресу ("" = без адреса)
'   T|код                — итог по учитываемым местам
Public Sub BuildBalances()
    Dim sh As Worksheet, lastR As Long, v As Variant, i As Long
    Dim opNames As New Collection, opFrom() As String, opTo() As String, nOps As Long
    Dim shO As Worksheet, r As Long, k As Long, code As String, q As Double
    Dim fL As String, tL As String, trk As New Collection, lastO As Long

    BalReset
    Set shO = WS(SH_OPS)
    lastO = LastRow(shO, 2)
    ReDim opFrom(1 To lastO + 1): ReDim opTo(1 To lastO + 1)
    For r = FIRST_ROW To lastO
        nOps = nOps + 1
        KeyAdd opNames, Txt(shO.Cells(r, 2).Value), nOps
        opFrom(nOps) = Txt(shO.Cells(r, 3).Value)
        opTo(nOps) = Txt(shO.Cells(r, 4).Value)
    Next r
    Dim shL As Worksheet, lastL As Long
    Set shL = WS(SH_LOC)
    lastL = LastRow(shL, 1)
    For r = FIRST_ROW To lastL
        If IsYes(shL.Cells(r, 4).Value) Then KeyAdd trk, Txt(shL.Cells(r, 1).Value), 1
    Next r

    Set sh = WS(SH_MOVE)
    lastR = LastRow(sh, J_OP)
    If lastR < FIRST_ROW Then Exit Sub
    v = sh.Range(sh.Cells(FIRST_ROW, J_OP), sh.Cells(lastR, J_ATO)).Value
    For i = 1 To UBound(v, 1)
        k = KeyIndex(opNames, Txt(v(i, 1)))
        code = UCase$(Txt(v(i, 2)))
        q = Nz(v(i, J_QTY - J_OP + 1))
        If k > 0 And Len(code) > 0 And q <> 0 Then
            fL = UCase$(opFrom(k)): tL = UCase$(opTo(k))
            BalAdd "L|" & code & "|" & fL, -q
            BalAdd "L|" & code & "|" & tL, q
            BalAdd "A|" & code & "|" & fL & "|" & UCase$(Txt(v(i, J_AFROM - J_OP + 1))), -q
            BalAdd "A|" & code & "|" & tL & "|" & UCase$(Txt(v(i, J_ATO - J_OP + 1))), q
            If KeyIndex(trk, fL) > 0 Then BalAdd "T|" & code, -q
            If KeyIndex(trk, tL) > 0 Then BalAdd "T|" & code, q
        End If
    Next i
End Sub

Public Function StockAt(ByVal code As String, ByVal locCode As String) As Double
    BuildBalances
    StockAt = BalGet("L|" & UCase$(code) & "|" & UCase$(locCode))
End Function

' Все ключи балансов с префиксом (для отчётов)
Public Function BalKeysCount() As Long
    BalKeysCount = mCnt
End Function

' ============================================================
' ЯДРО ПРОВЕДЕНИЯ
' codes/qtys/afroms/atos — массивы 1..n
' Возвращает номер документа (D000001) или "" при ошибке (текст в gLastError)
' ============================================================
Public Function PostDoc(ByVal opText As String, ByVal dt As Date, ByVal order As String, _
        ByVal item As String, ByVal batch As String, ByVal comment As String, _
        codes As Variant, qtys As Variant, afroms As Variant, atos As Variant) As String

    Dim shO As Worksheet, shM As Worksheet, opName As String, opR As Long
    Dim fromLoc As String, toLoc As String, n As Long, i As Long, j As Long
    Dim errs As String, code As String, q As Double, nr As Long, t As String
    Dim fromTrk As Boolean, toTrk As Boolean, need As Double, have As Double
    Dim af() As String, at() As String

    gLastError = ""
    PostDoc = ""
    Set shO = WS(SH_OPS): Set shM = WS(SH_MOVE)

    ' ---- операция ----
    opName = OpFullName(opText)
    If Len(opName) = 0 Then Fail "Операция «" & opText & "» не найдена на листе «" & SH_OPS & "».": Exit Function
    opR = FindRow(shO, 2, opName)
    If Txt(shO.Cells(opR, 9).Value) <> "" And Not IsYes(shO.Cells(opR, 9).Value) Then
        Fail "Операция «" & opName & "» отключена (столбец «Активна» = Нет).": Exit Function
    End If
    fromLoc = Txt(shO.Cells(opR, 3).Value): toLoc = Txt(shO.Cells(opR, 4).Value)
    fromTrk = LocTracked(fromLoc): toTrk = LocTracked(toLoc)
    If IsYes(shO.Cells(opR, 5).Value) And Len(Trim$(order)) = 0 Then errs = errs & "• Для операции нужен № заказа." & vbLf
    If IsYes(shO.Cells(opR, 6).Value) And Len(Trim$(item)) = 0 Then errs = errs & "• Для операции нужен код изделия / ПФ." & vbLf
    If Len(Trim$(item)) > 0 Then
        If NomRow(item) = 0 Then errs = errs & "• Изделие «" & item & "» нет в справочнике." & vbLf
    End If
    If CDbl(dt) < 1 Then errs = errs & "• Не указана дата." & vbLf

    ' ---- строки ----
    n = UBound(codes) - LBound(codes) + 1
    If n < 1 Then Fail "В документе нет строк.": Exit Function
    ReDim af(1 To n): ReDim at(1 To n)
    BuildBalances
    For i = 1 To n
        code = Trim$(CStr(codes(LBound(codes) + i - 1)))
        q = Nz(qtys(LBound(qtys) + i - 1))
        af(i) = Trim$(CStr(afroms(LBound(afroms) + i - 1)))
        at(i) = Trim$(CStr(atos(LBound(atos) + i - 1)))
        nr = NomRow(code)
        If Len(code) = 0 Then
            errs = errs & "• Строка " & i & ": не указан код." & vbLf
        ElseIf nr = 0 Then
            errs = errs & "• Строка " & i & ": код «" & code & "» нет в справочнике." & vbLf
        Else
            t = Txt(WS(SH_NOM).Cells(nr, N_TYPE).Value)
            If t = T_VARPARENT Then errs = errs & "• Строка " & i & ": «" & code & "» — вариантный родитель, двигать можно только конкретный вариант." & vbLf
            If Txt(WS(SH_NOM).Cells(nr, N_ACTIVE).Value) <> "" And Not IsYes(WS(SH_NOM).Cells(nr, N_ACTIVE).Value) Then _
                errs = errs & "• Строка " & i & ": «" & code & "» неактивна (Активно = Нет)." & vbLf
        End If
        If q <= 0 Then errs = errs & "• Строка " & i & ": количество должно быть больше 0." & vbLf
        ' адрес «куда»
        If Len(at(i)) = 0 And toTrk And nr > 0 Then at(i) = DefaultAddr(code, toLoc)
        If Len(at(i)) > 0 Then
            If Not AddrOk(at(i), toLoc) Then errs = errs & "• Строка " & i & ": адрес «" & at(i) & "» не относится к месту «" & toLoc & "» или неактивен." & vbLf
        End If
        ' адрес «откуда»
        If Len(af(i)) > 0 Then
            If Not AddrOk(af(i), fromLoc) Then errs = errs & "• Строка " & i & ": адрес «" & af(i) & "» не относится к месту «" & fromLoc & "» или неактивен." & vbLf
        ElseIf fromTrk And nr > 0 And q > 0 Then
            af(i) = PickFromAddr(code, fromLoc, q)
        End If
    Next i

    ' ---- контроль остатков (с учётом нескольких строк одного кода) ----
    If fromTrk And Len(errs) = 0 Then
        For i = 1 To n
            code = Trim$(CStr(codes(LBound(codes) + i - 1)))
            need = 0
            For j = 1 To n
                If StrComp(Trim$(CStr(codes(LBound(codes) + j - 1))), code, vbTextCompare) = 0 Then need = need + Nz(qtys(LBound(qtys) + j - 1))
            Next j
            have = BalGet("L|" & UCase$(code) & "|" & UCase$(fromLoc))
            If need > have + 0.0000001 And SettingYes("Запрет отрицательных остатков", True) Then
                errs = errs & "• «" & code & "»: в месте «" & fromLoc & "» есть " & Format$(have, "0.###") & ", требуется " & Format$(need, "0.###") & "." & vbLf
            ElseIf SettingYes("Контроль остатков по адресам", True) Then
                need = 0
                For j = 1 To n
                    If StrComp(Trim$(CStr(codes(LBound(codes) + j - 1))), code, vbTextCompare) = 0 And StrComp(af(j), af(i), vbTextCompare) = 0 Then need = need + Nz(qtys(LBound(qtys) + j - 1))
                Next j
                have = BalGet("A|" & UCase$(code) & "|" & UCase$(fromLoc) & "|" & UCase$(af(i)))
                If need > have + 0.0000001 Then
                    errs = errs & "• «" & code & "»: по адресу «" & IIf(Len(af(i)) = 0, "(без адреса)", af(i)) & "» есть " & Format$(have, "0.###") & ", требуется " & Format$(need, "0.###") & ". Укажите адрес «откуда»." & vbLf
                End If
            End If
        Next i
    End If

    If Len(errs) > 0 Then
        Fail "Документ НЕ проведён. Исправьте:" & vbLf & vbLf & DedupLines(errs), "Проведение операции"
        Exit Function
    End If

    ' ---- нужен документ? ----
    If IsYes(shO.Cells(opR, 7).Value) Then
        If PendingFilesCount() = 0 Then
            If Not Ask("Для операции «" & opName & "» обычно прикладывают документ (накладную, акт)." & vbLf & _
                       "Файл не прикреплён. Провести без файла?", "Документ") Then Exit Function
        End If
    End If

    ' ---- запись ----
    PostDoc = WriteRows(opName, dt, order, item, batch, comment, codes, qtys, af, at, n)
End Function

Private Function DedupLines(ByVal t As String) As String
    Dim a() As String, i As Long, outS As String
    a = Split(t, vbLf)
    For i = 0 To UBound(a)
        If Len(a(i)) > 0 Then
            If InStr(1, vbLf & outS, vbLf & a(i) & vbLf, vbBinaryCompare) = 0 Then outS = outS & a(i) & vbLf
        End If
    Next i
    DedupLines = outS
End Function

' Подбор адреса «откуда», если оператор его не указал
Private Function PickFromAddr(ByVal code As String, ByVal loc As String, ByVal q As Double) As String
    Dim sh As Worksheet, r As Long, lastR As Long, a As String, b As Double
    ' 1) хватает «без адреса» — оставляем пусто
    If BalGet("A|" & UCase$(code) & "|" & UCase$(loc) & "|") >= q - 0.0000001 Then Exit Function
    ' 2) адрес по умолчанию
    a = DefaultAddr(code, loc)
    If Len(a) > 0 Then
        If BalGet("A|" & UCase$(code) & "|" & UCase$(loc) & "|" & UCase$(a)) >= q - 0.0000001 Then PickFromAddr = a: Exit Function
    End If
    ' 3) первый адрес места, где хватает
    Set sh = WS(SH_ADDR)
    lastR = LastRow(sh, 1)
    For r = FIRST_ROW To lastR
        If StrComp(Txt(sh.Cells(r, 2).Value), loc, vbTextCompare) = 0 Then
            a = Txt(sh.Cells(r, 1).Value)
            b = BalGet("A|" & UCase$(code) & "|" & UCase$(loc) & "|" & UCase$(a))
            If b >= q - 0.0000001 Then PickFromAddr = a: Exit Function
        End If
    Next r
End Function

Private Function WriteRows(ByVal opName As String, ByVal dt As Date, ByVal order As String, _
        ByVal item As String, ByVal batch As String, ByVal comment As String, _
        codes As Variant, qtys As Variant, af() As String, at() As String, ByVal n As Long) As String
    Dim shM As Worksheet, r0 As Long, i As Long, docId As String, no0 As Long
    Dim arr() As Variant, files As String, who As String, stamp As Date

    Set shM = WS(SH_MOVE)
    AppOff
    On Error GoTo Bad
    docId = NextDocId()
    no0 = NextRowNo()
    r0 = LastRow(shM, J_OP) + 1
    If r0 < FIRST_ROW Then r0 = FIRST_ROW
    LinkPendingFiles docId
    files = FilesTitles(docId)
    who = UserName()
    stamp = Now
    ReDim arr(1 To n, 1 To J_LAST)
    For i = 1 To n
        arr(i, J_NO) = no0 + i - 1
        arr(i, J_DATE) = CDbl(dt)
        arr(i, J_OP) = opName
        arr(i, J_CODE) = Trim$(CStr(codes(LBound(codes) + i - 1)))
        arr(i, J_NAME) = JFormula(J_NAME, r0 + i - 1)
        arr(i, J_QTY) = Nz(qtys(LBound(qtys) + i - 1))
        arr(i, J_UNIT) = JFormula(J_UNIT, r0 + i - 1)
        arr(i, J_FROM) = JFormula(J_FROM, r0 + i - 1)
        arr(i, J_TO) = JFormula(J_TO, r0 + i - 1)
        arr(i, J_AFROM) = af(i)
        arr(i, J_ATO) = at(i)
        arr(i, J_ORDER) = order
        arr(i, J_ITEM) = item
        arr(i, J_BATCH) = batch
        arr(i, J_DOC) = docId
        arr(i, J_FILES) = files
        arr(i, J_USER) = who
        arr(i, J_COMMENT) = comment
        arr(i, J_CHECK) = JFormula(J_CHECK, r0 + i - 1)
        arr(i, J_CREATED) = CDbl(stamp)
    Next i
    shM.Range(shM.Cells(r0, 1), shM.Cells(r0 + n - 1, J_LAST)).Formula = arr
    FormatJournalRows r0, r0 + n - 1
    AppOn
    WriteRows = docId
    Exit Function
Bad:
    AppOn
    Fail "Ошибка записи в журнал: " & Err.Description
    WriteRows = ""
End Function

Public Sub FormatJournalRows(ByVal r1 As Long, ByVal r2 As Long)
    Dim shM As Worksheet
    On Error Resume Next
    Set shM = WS(SH_MOVE)
    shM.Range(shM.Cells(r1, J_DATE), shM.Cells(r2, J_DATE)).NumberFormat = "dd.mm.yyyy"
    shM.Range(shM.Cells(r1, J_QTY), shM.Cells(r2, J_QTY)).NumberFormat = "# ##0.###"
    shM.Range(shM.Cells(r1, J_CREATED), shM.Cells(r2, J_CREATED)).NumberFormat = "dd.mm.yyyy hh:mm"
End Sub

' Формулы журнала для строки r
Public Function JFormula(ByVal col As Long, ByVal r As Long) As String
    Dim f As String
    Select Case col
    Case J_NAME
        f = "=IF($D{r}="""","""",IFERROR(INDEX(Справочник!$B:$B,MATCH($D{r},Справочник!$A:$A,0)),""НЕТ В СПРАВОЧНИКЕ""))"
    Case J_UNIT
        f = "=IF($D{r}="""","""",IFERROR(INDEX(Справочник!$D:$D,MATCH($D{r},Справочник!$A:$A,0)),""""))"
    Case J_FROM
        f = "=IF($C{r}="""","""",IFERROR(INDEX(Операции!$C:$C,MATCH($C{r},Операции!$B:$B,0)),""?""))"
    Case J_TO
        f = "=IF($C{r}="""","""",IFERROR(INDEX(Операции!$D:$D,MATCH($C{r},Операции!$B:$B,0)),""?""))"
    Case J_CHECK
        f = "=IF($C{r}="""","""",IF(COUNTIF(Операции!$B:$B,$C{r})=0,""НЕТ ОПЕРАЦИИ"",IF($D{r}="""",""НЕТ КОДА""," & _
            "IF(COUNTIF(Справочник!$A:$A,$D{r})=0,""НЕТ В СПРАВОЧНИКЕ"",IF(N($F{r})<=0,""КОЛИЧЕСТВО <= 0""," & _
            "IF($B{r}="""",""НЕТ ДАТЫ"",IF(AND($J{r}<>"""",COUNTIFS(Адреса!$A:$A,$J{r},Адреса!$B:$B,$H{r})=0),""АДРЕС ОТКУДА НЕ ИЗ ""&$H{r}," & _
            "IF(AND($K{r}<>"""",COUNTIFS(Адреса!$A:$A,$K{r},Адреса!$B:$B,$I{r})=0),""АДРЕС КУДА НЕ ИЗ ""&$I{r},""OK""))))))))"
    End Select
    JFormula = Replace(f, "{r}", CStr(r))
End Function

' Дописать формулы в строки журнала, где их нет (после ручного ввода/вставки)
Public Sub EnsureJournalFormulas(ByVal r1 As Long, ByVal r2 As Long)
    Dim shM As Worksheet, r As Long, c As Variant
    Set shM = WS(SH_MOVE)
    For r = r1 To r2
        For Each c In Array(J_NAME, J_UNIT, J_FROM, J_TO, J_CHECK)
            If Not shM.Cells(r, c).HasFormula Then shM.Cells(r, c).Formula = JFormula(CLng(c), r)
        Next c
    Next r
End Sub

Public Function NextDocId() As String
    Dim shM As Worksheet, lastR As Long, v As Variant, i As Long, mx As Long, t As String
    Set shM = WS(SH_MOVE)
    lastR = LastRow(shM, J_OP)
    If lastR >= FIRST_ROW Then
        v = shM.Range(shM.Cells(FIRST_ROW, J_DOC), shM.Cells(lastR + 1, J_DOC)).Value
        For i = 1 To UBound(v, 1)
            t = Txt(v(i, 1))
            If Left$(t, 1) = "D" And IsNumeric(Mid$(t, 2)) Then
                If CLng(Mid$(t, 2)) > mx Then mx = CLng(Mid$(t, 2))
            End If
        Next i
    End If
    t = Setting("Последний номер документа", "0")
    If IsNumeric(t) And Len(t) > 0 Then
        If CLng(t) > mx Then mx = CLng(t)
    End If
    NextDocId = "D" & Format$(mx + 1, "000000")
    SaveSetting2 "Последний номер документа", CStr(mx + 1)
End Function

Public Function NextRowNo() As Long
    Dim shM As Worksheet, lastR As Long, v As Variant, i As Long, mx As Long
    Set shM = WS(SH_MOVE)
    lastR = LastRow(shM, J_OP)
    If lastR >= FIRST_ROW Then
        v = shM.Range(shM.Cells(FIRST_ROW, J_NO), shM.Cells(lastR + 1, J_NO)).Value
        For i = 1 To UBound(v, 1)
            If Nz(v(i, 1)) > mx Then mx = CLng(Nz(v(i, 1)))
        Next i
    End If
    NextRowNo = mx + 1
End Function

Public Sub SaveSetting2(ByVal paramName As String, ByVal value As String)
    Dim sh As Worksheet, r As Long
    On Error Resume Next
    Set sh = WS(SH_SET)
    r = FindRow(sh, 1, paramName)
    If r = 0 Then
        r = LastRow(sh, 1) + 1
        sh.Cells(r, 1).Value = paramName
        sh.Cells(r, 3).Value = "Служебный параметр — заполняется макросами"
    End If
    sh.Cells(r, 2).Value = value
End Sub

' ============================================================
' КНОПКА «ПРОВЕСТИ» на листе «Ввод»
' ============================================================
Public Sub Post_FromInput()
    Dim sh As Worksheet, lastR As Long, r As Long, n As Long
    Dim codes() As String, qtys() As Double, af() As String, at() As String
    Dim docId As String, dt As Date, errs As String

    Set sh = WS(SH_INPUT)
    gLastError = ""
    If Len(Txt(sh.Range(IN_OP).Value)) = 0 Then Fail "Выберите операцию в ячейке " & IN_OP & ".": Exit Sub
    If IsDate(sh.Range(IN_DATE).Value) Then dt = CDate(sh.Range(IN_DATE).Value) Else dt = Date
    lastR = InputLastRow()
    If lastR < IN_FIRST Then Fail "Заполните хотя бы одну строку: код (столбец B) и количество (столбец E).": Exit Sub
    ReDim codes(1 To lastR - IN_FIRST + 1): ReDim qtys(1 To lastR - IN_FIRST + 1)
    ReDim af(1 To lastR - IN_FIRST + 1): ReDim at(1 To lastR - IN_FIRST + 1)
    For r = IN_FIRST To lastR
        If Len(Txt(sh.Cells(r, L_CODE).Value)) > 0 Or Len(Txt(sh.Cells(r, L_QTY).Value)) > 0 Then
            n = n + 1
            codes(n) = Txt(sh.Cells(r, L_CODE).Value)
            qtys(n) = Nz(sh.Cells(r, L_QTY).Value)
            af(n) = Txt(sh.Cells(r, L_AFROM).Value)
            at(n) = Txt(sh.Cells(r, L_ATO).Value)
            If Len(codes(n)) = 0 Then errs = errs & "• Строка ввода " & r & ": есть количество, но нет кода." & vbLf
        End If
    Next r
    If n = 0 Then Fail "Нет строк для проведения.": Exit Sub
    If Len(errs) > 0 Then Fail errs: Exit Sub
    ReDim Preserve codes(1 To n): ReDim Preserve qtys(1 To n)
    ReDim Preserve af(1 To n): ReDim Preserve at(1 To n)

    docId = PostDoc(Txt(sh.Range(IN_OP).Value), dt, Txt(sh.Range(IN_ORDER).Value), Txt(sh.Range(IN_ITEM).Value), _
                    Txt(sh.Range(IN_BATCH).Value), Txt(sh.Range(IN_COMMENT).Value), codes, qtys, af, at)
    If Len(docId) = 0 Then Exit Sub
    ClearInput False
    Info "Документ " & docId & " проведён." & vbLf & "Строк записано в «" & SH_MOVE & "»: " & n & "." & vbLf & vbLf & _
         "Остатки пересчитаны автоматически (лист «" & SH_STOCK & "»).", "Готово"
End Sub

Public Function InputLastRow() As Long
    Dim sh As Worksheet, r As Long, r2 As Long
    Set sh = WS(SH_INPUT)
    r = sh.Cells(sh.Rows.Count, L_CODE).End(xlUp).Row
    r2 = sh.Cells(sh.Rows.Count, L_QTY).End(xlUp).Row
    If r2 > r Then r = r2
    If r < IN_FIRST Then r = IN_FIRST - 1
    InputLastRow = r
End Function

' Очистить лист ввода. full=True — также дата, операция и файлы.
Public Sub ClearInput(Optional ByVal full As Boolean = False)
    Dim sh As Worksheet, lastR As Long
    Set sh = WS(SH_INPUT)
    AppOff
    lastR = InputLastRow()
    If lastR >= IN_FIRST Then
        sh.Range(sh.Cells(IN_FIRST, L_CODE), sh.Cells(lastR, L_CODE)).ClearContents
        sh.Range(sh.Cells(IN_FIRST, L_QTY), sh.Cells(lastR, L_ATO)).ClearContents
    End If
    sh.Range(IN_ORDER).ClearContents
    sh.Range(IN_ITEM).ClearContents
    sh.Range(IN_BATCH).ClearContents
    sh.Range(IN_COMMENT).ClearContents
    sh.Range(IN_KQTY).ClearContents
    If full Then
        sh.Range(IN_OP).ClearContents
        sh.Range(IN_DATE).Value = Date
        DiscardPendingFiles
    End If
    RefreshInputFiles
    AppOn
End Sub

Public Sub Input_Clear()
    If Ask("Очистить лист «Ввод» полностью (операция, строки, неприкреплённые файлы)?") Then ClearInput True
End Sub
