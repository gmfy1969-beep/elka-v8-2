Attribute VB_Name = "modControl"
Option Explicit
' ============================================================
' ELKA v8.2 — ЛИСТ «КОНТРОЛЬ»: проверка всей книги
' Сверху — сводка по проверкам, ниже — список каждой найденной
' проблемы со ссылкой на ячейку и подсказкой, как исправить.
' ============================================================
Private mOut As Long
Private mErr As Long
Private mWarn As Long
Private Const SUM_FIRST As Long = 5      ' первая строка сводки
Private Const SUM_ROWS As Long = 15      ' число проверок
Private Const DET_HDR As Long = 22       ' заголовок списка проблем

Private Function Checks() As Variant
    Checks = Array( _
      Array("C01", "ОШИБКА", "Дубли кодов в справочнике", "Код должен быть уникальным. Переименуйте один из дублей (например DET-001 -> DET-001A) или удалите лишнюю строку через «ELKA: удалить выделенные строки»."), _
      Array("C02", "ОШИБКА", "Справочник: нет наименования или неверный тип", "Заполните «Наименование». Тип выберите из выпадающего списка (Деталь, Материал, Полуфабрикат, ...)."), _
      Array("C03", "ОШИБКА", "Вариант без вариантного родителя", "В «Базовый код» укажите код позиции с типом «Вариантный родитель» или очистите «Базовый код» и «Вариант»."), _
      Array("C04", "ОШИБКА", "Спецификация: неизвестный код или норма <= 0", "Проверьте, что изделие и компонент есть в справочнике, норма больше 0 и компонент не равен самому изделию."), _
      Array("C05", "ОШИБКА", "Объединения: неизвестный код, норма <= 0 или неверный тип", "Код объединения должен иметь тип «Объединение»; компонент — существовать в справочнике; количество > 0."), _
      Array("C06", "ОШИБКА", "Адреса: дубль или неизвестное место хранения", "Адрес должен быть уникален, а «Место хранения» — из листа «Места хранения»."), _
      Array("C07", "ОШИБКА", "Операции: неизвестное место или дубль кода", "«Откуда»/«Куда» должны существовать на листе «Места хранения»; код операции уникален."), _
      Array("C08", "ОШИБКА", "Движение: ошибка в строке", "Откройте ссылку, прочитайте текст в столбце «Проверка» и исправьте операцию, код, количество, дату или адрес. Ошибочную строку можно удалить (она попадёт в «Движение_удалённые»)."), _
      Array("C09", "ОШИБКА", "Отрицательный остаток в месте хранения", "Списали больше, чем было. Найдите по коду строки в «Движение» (фильтр по коду), проверьте количество и операцию; при пропущенном приходе внесите его, лишнюю строку удалите."), _
      Array("C10", "ОШИБКА", "Отрицательный остаток по адресу", "Товар взяли не из той ячейки. Проведите операцию «Перемещение между адресами» или исправьте «Адрес откуда» в строке журнала."), _
      Array("C11", "ВНИМАНИЕ", "Остаток ниже минимального", "Закажите позицию у поставщика или уменьшите «Мин. остаток» в справочнике."), _
      Array("C12", "ОШИБКА", "Файл документа не найден", "Файл удалён или недоступна сетевая папка. Проверьте «Папка архива документов» на листе «Настройки» или прикрепите файл заново."), _
      Array("C13", "ВНИМАНИЕ", "Файлы прикреплены, но документ не проведён", "На листе «Ввод» нажмите «Провести» или «Очистить», чтобы удалить неиспользованные файлы."), _
      Array("C14", "ВНИМАНИЕ", "Расхождение с 1С", "Откройте «Сравнение_1С», проверьте дату выгрузки, склад и единицы. Ошибку книги исправьте операцией инвентаризации (15/16)."), _
      Array("C15", "ВНИМАНИЕ", "Позиция неактивна, но имеет остаток", "Переместите или спишите остаток либо снова сделайте позицию активной.") _
    )
End Function

Public Function ControlErrors() As Long
    RunControlCore
    ControlErrors = mErr
End Function

Public Sub Control_Run()
    RunControlCore
    If Not gSilent Then WS(SH_CTRL).Activate
    If mErr = 0 And mWarn = 0 Then
        Info "Проверка завершена: ошибок нет.", "Контроль"
    Else
        Info "Проверка завершена." & vbLf & "Ошибок: " & mErr & vbLf & "Предупреждений: " & mWarn & vbLf & vbLf & _
             "Список — на листе «" & SH_CTRL & "», начиная со строки " & DET_HDR + 1 & ". Щёлкните по ссылке в столбце «Где», чтобы перейти к ячейке.", "Контроль"
    End If
End Sub

Private Sub AddIssue(ByVal chk As String, ByVal sheetName As String, ByVal r As Long, ByVal c As Long, ByVal detail As String)
    Dim sh As Worksheet, ck As Variant, i As Long, lvl As String, hint As String
    Set sh = WS(SH_CTRL)
    ck = Checks()
    For i = 0 To UBound(ck)
        If ck(i)(0) = chk Then lvl = ck(i)(1): hint = ck(i)(3): Exit For
    Next i
    If lvl = "ОШИБКА" Then mErr = mErr + 1 Else mWarn = mWarn + 1
    sh.Cells(mOut, 1).Value = mOut - DET_HDR
    sh.Cells(mOut, 2).Value = lvl
    sh.Cells(mOut, 3).Value = chk
    sh.Cells(mOut, 4).Value = sheetName & IIf(r > 0, " !" & ColLetter(IIf(c < 1, 1, c)) & r, "")
    sh.Cells(mOut, 5).Value = detail
    sh.Cells(mOut, 6).Value = hint
    If r > 0 Then
        On Error Resume Next
        sh.Hyperlinks.Add Anchor:=sh.Cells(mOut, 4), Address:="", _
            SubAddress:="'" & sheetName & "'!" & ColLetter(IIf(c < 1, 1, c)) & r, TextToDisplay:=CStr(sh.Cells(mOut, 4).Value)
        On Error GoTo 0
    End If
    If lvl = "ОШИБКА" Then
        sh.Cells(mOut, 2).Interior.Color = RGB(244, 204, 204)
    Else
        sh.Cells(mOut, 2).Interior.Color = RGB(255, 242, 204)
    End If
    mOut = mOut + 1
End Sub

Public Sub RunControlCore()
    Dim ra As Long
    Dim sh As Worksheet, shN As Worksheet, lastR As Long, r As Long, r2 As Long, t As String, code As String
    Dim keys As New Collection, ck As Variant, i As Long, cnt() As Long, locs As Variant, b As Double
    Dim validTypes As String, shA As Worksheet, shL As Worksheet, shO As Worksheet, shM As Worksheet
    Dim v As Variant, opName As String, fL As String, tL As String

    AppOff
    On Error GoTo Bad
    Set sh = WS(SH_CTRL)
    mErr = 0: mWarn = 0
    lastR = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    If lastR < DET_HDR + 1 Then lastR = DET_HDR + 1
    sh.Range(sh.Cells(DET_HDR + 1, 1), sh.Cells(lastR + 5, 6)).Clear
    On Error Resume Next
    sh.Hyperlinks.Delete
    On Error GoTo Bad
    mOut = DET_HDR + 1
    validTypes = "|" & T_PART & "|" & T_MAT & "|" & T_PF & "|" & T_PFD & "|" & T_GP & "|" & T_BUNDLE & "|" & T_VARPARENT & "|"

    ' ---- справочник ----
    Set shN = WS(SH_NOM)
    lastR = LastRow(shN, N_CODE)
    For r = FIRST_ROW To lastR
        code = Txt(shN.Cells(r, N_CODE).Value)
        If Len(code) > 0 Then
            If KeyIndex(keys, code) > 0 Then
                AddIssue "C01", SH_NOM, r, N_CODE, "Код «" & code & "» уже есть в строке " & KeyIndex(keys, code)
            Else
                KeyAdd keys, code, r
            End If
            If Len(Txt(shN.Cells(r, N_NAME).Value)) = 0 Then AddIssue "C02", SH_NOM, r, N_NAME, "Нет наименования у «" & code & "»"
            If InStr(validTypes, "|" & Txt(shN.Cells(r, N_TYPE).Value) & "|") = 0 Then AddIssue "C02", SH_NOM, r, N_TYPE, "Неверный тип «" & Txt(shN.Cells(r, N_TYPE).Value) & "» у «" & code & "»"
            t = Txt(shN.Cells(r, N_BASE).Value)
            If Len(t) > 0 And StrComp(t, code, vbTextCompare) <> 0 Then
                If NomType(t) <> T_VARPARENT Then AddIssue "C03", SH_NOM, r, N_BASE, "«" & code & "»: базовый код «" & t & "» не является вариантным родителем"
            End If
        End If
    Next r

    ' ---- спецификация и объединения ----
    Dim shS As Worksheet, nm As Variant
    For Each nm In Array(SH_SPEC, SH_BUNDLE)
        Set shS = WS(CStr(nm))
        lastR = LastRow(shS, 1)
        For r = FIRST_ROW To lastR
            If Len(Txt(shS.Cells(r, 1).Value)) > 0 Or Len(Txt(shS.Cells(r, 2).Value)) > 0 Then
                t = ""
                If NomRow(Txt(shS.Cells(r, 1).Value)) = 0 Then t = "неизвестный код «" & Txt(shS.Cells(r, 1).Value) & "»"
                If NomRow(Txt(shS.Cells(r, 2).Value)) = 0 Then t = t & " неизвестный компонент «" & Txt(shS.Cells(r, 2).Value) & "»"
                If Nz(shS.Cells(r, 4).Value) <= 0 Then t = t & " норма <= 0"
                If StrComp(Txt(shS.Cells(r, 1).Value), Txt(shS.Cells(r, 2).Value), vbTextCompare) = 0 Then t = t & " компонент равен изделию"
                If nm = SH_BUNDLE And NomRow(Txt(shS.Cells(r, 1).Value)) > 0 Then
                    If NomType(Txt(shS.Cells(r, 1).Value)) <> T_BUNDLE Then t = t & " тип не «Объединение»"
                End If
                If Len(t) > 0 Then AddIssue IIf(nm = SH_SPEC, "C04", "C05"), CStr(nm), r, 1, Trim$(t)
            End If
        Next r
    Next nm

    ' ---- адреса ----
    Set shA = WS(SH_ADDR): Set shL = WS(SH_LOC)
    Dim akeys As New Collection
    lastR = LastRow(shA, 1)
    For r = FIRST_ROW To lastR
        t = Txt(shA.Cells(r, 1).Value)
        If Len(t) > 0 Then
            If KeyIndex(akeys, t) > 0 Then
                AddIssue "C06", SH_ADDR, r, 1, "Дубль адреса «" & t & "»"
            Else
                KeyAdd akeys, t, r
            End If
            If FindRow(shL, 1, Txt(shA.Cells(r, 2).Value)) = 0 Then AddIssue "C06", SH_ADDR, r, 2, "Место «" & Txt(shA.Cells(r, 2).Value) & "» не найдено"
        End If
    Next r

    ' ---- операции ----
    Set shO = WS(SH_OPS)
    Dim okeys As New Collection
    lastR = LastRow(shO, 2)
    For r = FIRST_ROW To lastR
        t = Txt(shO.Cells(r, 1).Value)
        If Len(t) > 0 Then
            If KeyIndex(okeys, t) > 0 Then
                AddIssue "C07", SH_OPS, r, 1, "Дубль кода операции «" & t & "»"
            Else
                KeyAdd okeys, t, r
            End If
            If FindRow(shL, 1, Txt(shO.Cells(r, 3).Value)) = 0 Then AddIssue "C07", SH_OPS, r, 3, "Место «" & Txt(shO.Cells(r, 3).Value) & "» не найдено"
            If FindRow(shL, 1, Txt(shO.Cells(r, 4).Value)) = 0 Then AddIssue "C07", SH_OPS, r, 4, "Место «" & Txt(shO.Cells(r, 4).Value) & "» не найдено"
        End If
    Next r

    ' ---- журнал движения (независимая проверка VBA) ----
    Set shM = WS(SH_MOVE)
    lastR = LastRow(shM, J_OP)
    If lastR >= FIRST_ROW Then
        v = shM.Range(shM.Cells(FIRST_ROW, 1), shM.Cells(lastR, J_LAST)).Value
        For i = 1 To UBound(v, 1)
            r = FIRST_ROW + i - 1
            opName = Txt(v(i, J_OP))
            If Len(opName) > 0 Or Len(Txt(v(i, J_CODE))) > 0 Then
                t = ""
                If FindRow(shO, 2, opName) = 0 Then
                    t = "неизвестная операция «" & opName & "»"
                Else
                    OpRoute opName, fL, tL
                    If Len(Txt(v(i, J_AFROM))) > 0 Then
                        If Not AddrOk(Txt(v(i, J_AFROM)), fL) Then t = t & " адрес откуда «" & Txt(v(i, J_AFROM)) & "» не из «" & fL & "»"
                    End If
                    If Len(Txt(v(i, J_ATO))) > 0 Then
                        If Not AddrOk(Txt(v(i, J_ATO)), tL) Then t = t & " адрес куда «" & Txt(v(i, J_ATO)) & "» не из «" & tL & "»"
                    End If
                End If
                If KeyIndex(keys, Txt(v(i, J_CODE))) = 0 Then t = t & " код «" & Txt(v(i, J_CODE)) & "» нет в справочнике"
                If Nz(v(i, J_QTY)) <= 0 Then t = t & " количество <= 0"
                If Not IsDate(v(i, J_DATE)) And Not IsNumeric(v(i, J_DATE)) Then t = t & " нет даты"
                If Len(t) > 0 Then AddIssue "C08", SH_MOVE, r, J_CHECK, "№ " & Txt(v(i, J_NO)) & ": " & Trim$(t)
            End If
        Next i
    End If

    ' ---- остатки ----
    BuildBalances
    locs = TrackedLocs()
    lastR = LastRow(shN, N_CODE)
    For r = FIRST_ROW To lastR
        code = Txt(shN.Cells(r, N_CODE).Value)
        If Len(code) > 0 Then
            For i = 0 To UBound(locs)
                b = BalGet("L|" & UCase$(code) & "|" & UCase$(locs(i)))
                If b < -0.0000001 Then AddIssue "C09", SH_NOM, r, N_CODE, "«" & code & "» в «" & locs(i) & "»: " & Format$(b, "0.###")
            Next i
            b = BalGet("T|" & UCase$(code))
            If Nz(shN.Cells(r, N_MIN).Value) > 0 And b < Nz(shN.Cells(r, N_MIN).Value) And Txt(shN.Cells(r, N_TYPE).Value) <> T_VARPARENT Then
                AddIssue "C11", SH_NOM, r, N_MIN, "«" & code & "»: остаток " & Format$(b, "0.###") & " < минимум " & Txt(shN.Cells(r, N_MIN).Value)
            End If
            If Txt(shN.Cells(r, N_ACTIVE).Value) <> "" And Not IsYes(shN.Cells(r, N_ACTIVE).Value) And Abs(b) > 0.0000001 Then
                AddIssue "C15", SH_NOM, r, N_ACTIVE, "«" & code & "» неактивна, остаток " & Format$(b, "0.###")
            End If
            ' адреса
            For ra = FIRST_ROW To LastRow(shA, 1)
                b = BalGet("A|" & UCase$(code) & "|" & UCase$(Txt(shA.Cells(ra, 2).Value)) & "|" & UCase$(Txt(shA.Cells(ra, 1).Value)))
                If b < -0.0000001 Then AddIssue "C10", SH_ADDR, ra, 1, "«" & code & "» по адресу «" & Txt(shA.Cells(ra, 1).Value) & "»: " & Format$(b, "0.###")
            Next ra
        End If
    Next r

    ' ---- файлы ----
    Dim shF As Worksheet
    Set shF = WS(SH_FILES)
    lastR = LastRow(shF, F_ID)
    For r = FIRST_ROW To lastR
        If Len(Txt(shF.Cells(r, F_PATH).Value)) > 0 Then
            If Not FileExists(FullFilePath(Txt(shF.Cells(r, F_PATH).Value))) Then AddIssue "C12", SH_FILES, r, F_PATH, "«" & Txt(shF.Cells(r, F_TITLE).Value) & "»: " & FullFilePath(Txt(shF.Cells(r, F_PATH).Value))
        End If
        If Txt(shF.Cells(r, F_DOC).Value) = PENDING_DOC Then AddIssue "C13", SH_FILES, r, F_DOC, "«" & Txt(shF.Cells(r, F_TITLE).Value) & "»"
    Next r

    ' ---- 1С ----
    Dim shC As Worksheet
    Set shC = WS(SH_CMP)
    lastR = LastRow(shC, 1)
    For r = FIRST_ROW To lastR
        t = Txt(shC.Cells(r, 8).Value)
        If Len(t) > 0 And t <> "СОВПАДАЕТ" Then AddIssue "C14", SH_CMP, r, 8, Txt(shC.Cells(r, 1).Value) & " / " & Txt(shC.Cells(r, 3).Value) & ": " & t & ", разница " & Txt(shC.Cells(r, 7).Value)
    Next r

    WriteSummary
    sh.Range("B3").Value = "Проверено: " & Format$(Now, "dd.mm.yyyy hh:mm") & "   •   ошибок: " & mErr & "   •   предупреждений: " & mWarn
    If mOut > DET_HDR + 1 Then
        StyleReport sh, DET_HDR + 1, mOut - 1, 6
        sh.Range(sh.Cells(DET_HDR + 1, 5), sh.Cells(mOut - 1, 6)).WrapText = True
        sh.Range(sh.Cells(DET_HDR + 1, 1), sh.Cells(mOut - 1, 6)).RowHeight = 30
    Else
        sh.Cells(DET_HDR + 1, 2).Value = "OK"
        sh.Cells(DET_HDR + 1, 5).Value = "Проблем не найдено"
    End If
    AppOn
    Exit Sub
Bad:
    AppOn
    Fail "Ошибка проверки: " & Err.Description
End Sub

Private Sub WriteSummary()
    Dim sh As Worksheet, ck As Variant, i As Long, r As Long, n As Long, lastR As Long
    Set sh = WS(SH_CTRL)
    ck = Checks()
    For i = 0 To UBound(ck)
        r = SUM_FIRST + i
        n = 0
        For lastR = DET_HDR + 1 To mOut - 1
            If Txt(sh.Cells(lastR, 3).Value) = ck(i)(0) Then n = n + 1
        Next lastR
        sh.Cells(r, 1).Value = ck(i)(0)
        sh.Cells(r, 2).Value = ck(i)(2)
        sh.Cells(r, 3).Value = ck(i)(1)
        sh.Cells(r, 4).Value = n
        sh.Cells(r, 5).Value = IIf(n = 0, "OK", IIf(ck(i)(1) = "ОШИБКА", "ИСПРАВИТЬ", "ПРОВЕРИТЬ"))
        sh.Cells(r, 6).Value = ck(i)(3)
        If n = 0 Then
            sh.Cells(r, 5).Interior.Color = RGB(226, 240, 217)
        ElseIf ck(i)(1) = "ОШИБКА" Then
            sh.Cells(r, 5).Interior.Color = RGB(244, 204, 204)
        Else
            sh.Cells(r, 5).Interior.Color = RGB(255, 242, 204)
        End If
    Next i
End Sub
