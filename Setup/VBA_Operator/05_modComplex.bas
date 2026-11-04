Attribute VB_Name = "modComplex"
Option Explicit
' ============================================================
' ELKA v8.2 — КОМПЛЕКТАЦИЯ, СБОРКА, ОБЪЕДИНЕНИЯ, НЕКОМПЛЕКТНЫЕ ПФ
' Спецификация: Код изделия | Код компонента | Наименование | Норма на 1 | Ед. | Примечание
' Объединения:  Код объединения | Код компонента | Наименование | Кол-во на 1 | Ед. | Примечание
' Некомплектные_ПФ: Код ПФД | Что собираем | Наим. | Код недостающего | Наим. | Кол-во на 1 | Ед. | Описание
' Вариантный компонент спецификации (тип «Вариантный родитель») при сборке
' списывается с имеющихся исполнений (порядок — как в «Справочнике»).
' ============================================================

' Состав изделия/объединения: возвращает число компонентов, заполняет массивы
Public Function GetComposition(ByVal sheetName As String, ByVal parentCode As String, _
                               ByRef comps() As String, ByRef norms() As Double) As Long
    Dim sh As Worksheet, r As Long, lastR As Long, n As Long
    Set sh = WS(sheetName)
    lastR = LastRow(sh, 1)
    ReDim comps(1 To lastR + 1): ReDim norms(1 To lastR + 1)
    For r = FIRST_ROW To lastR
        If StrComp(Txt(sh.Cells(r, 1).Value), parentCode, vbTextCompare) = 0 Then
            If Len(Txt(sh.Cells(r, 2).Value)) > 0 And Nz(sh.Cells(r, 4).Value) > 0 Then
                n = n + 1
                comps(n) = Txt(sh.Cells(r, 2).Value)
                norms(n) = Nz(sh.Cells(r, 4).Value)
            End If
        End If
    Next r
    GetComposition = n
End Function

' ---------------- некомплектные ПФ ----------------
' Что собирается из ПФД (код изделия/узла); "" — нет на листе
Public Function PFDTarget(ByVal pfd As String) As String
    Dim sh As Worksheet, r As Long
    Set sh = WS(SH_PFD)
    r = FindRow(sh, P_CODE, pfd)
    If r > 0 Then PFDTarget = Txt(sh.Cells(r, P_TARGET).Value)
End Function

' Недостающие компоненты ПФД (на 1 шт)
Public Function PFDMissing(ByVal pfd As String, ByRef comps() As String, ByRef norms() As Double) As Long
    Dim sh As Worksheet, r As Long, lastR As Long, n As Long
    Set sh = WS(SH_PFD)
    lastR = LastRow(sh, P_CODE)
    ReDim comps(1 To lastR + 1): ReDim norms(1 To lastR + 1)
    For r = FIRST_ROW To lastR
        If StrComp(Txt(sh.Cells(r, P_CODE).Value), pfd, vbTextCompare) = 0 And Len(Txt(sh.Cells(r, P_MISS).Value)) > 0 Then
            n = n + 1
            comps(n) = Txt(sh.Cells(r, P_MISS).Value)
            norms(n) = Nz(sh.Cells(r, P_QTY).Value)
        End If
    Next r
    PFDMissing = n
End Function

Private Function MissQty(ByVal pfd As String, ByVal comp As String) As Double
    Dim c() As String, q() As Double, n As Long, i As Long
    n = PFDMissing(pfd, c, q)
    For i = 1 To n
        If StrComp(c(i), comp, vbTextCompare) = 0 Then MissQty = MissQty + q(i)
    Next i
End Function

' Состав ПФД = спецификация собираемого минус недостающие компоненты
Public Function PFDComposition(ByVal pfd As String, ByRef comps() As String, ByRef norms() As Double) As Long
    Dim tgt As String, c() As String, q() As Double, n As Long, i As Long, k As Long, m As Double
    tgt = PFDTarget(pfd)
    If Len(tgt) = 0 Then Exit Function
    n = GetComposition(SH_SPEC, tgt, c, q)
    ReDim comps(1 To n + 1): ReDim norms(1 To n + 1)
    For i = 1 To n
        m = q(i) - MissQty(pfd, c(i))
        If m > 0.0000001 Then
            k = k + 1: comps(k) = c(i): norms(k) = m
        End If
    Next i
    PFDComposition = k
End Function

' ---------------- вариантные компоненты ----------------
Public Function IsVarParent(ByVal code As String) As Boolean
    IsVarParent = (NomType(code) = T_VARPARENT)
End Function

' Остаток с учётом семейства: для вариантного родителя — сумма исполнений
' (балансы должны быть рассчитаны BuildBalances)
Public Function FamilyBal(ByVal code As String, ByVal loc As String) As Double
    Dim shN As Worksheet, r As Long, lastR As Long, v As Variant
    If Not IsVarParent(code) Then FamilyBal = BalGet("L|" & UCase$(code) & "|" & UCase$(loc)): Exit Function
    Set shN = WS(SH_NOM)
    lastR = LastRow(shN, N_CODE)
    v = shN.Range(shN.Cells(FIRST_ROW, N_CODE), shN.Cells(lastR, N_BASE)).Value
    For r = 1 To UBound(v, 1)
        If StrComp(Txt(v(r, N_BASE)), code, vbTextCompare) = 0 Then
            FamilyBal = FamilyBal + BalGet("L|" & UCase$(Txt(v(r, N_CODE))) & "|" & UCase$(loc))
        End If
    Next r
End Function

' Заменить вариантных родителей на конкретные исполнения по наличию в loc.
' На входе comps/q (1..n), на выходе oc/oq (1..m). Ошибка — текст в errs.
Public Function ResolveVariants(comps() As String, q() As Double, ByVal n As Long, ByVal loc As String, _
                                ByRef oc() As String, ByRef oq() As Double, ByRef errs As String) As Long
    Dim shN As Worksheet, lastR As Long, v As Variant, i As Long, r As Long, m As Long
    Dim need As Double, have As Double, take As Double, used As New Collection, ku As Long
    Dim usedQ() As Double, nu As Long, code As String, avail As String
    BuildBalances
    Set shN = WS(SH_NOM)
    lastR = LastRow(shN, N_CODE)
    v = shN.Range(shN.Cells(FIRST_ROW, N_CODE), shN.Cells(lastR, N_ACTIVE)).Value
    ReDim oc(1 To n * 4 + 4): ReDim oq(1 To n * 4 + 4): ReDim usedQ(1 To UBound(v, 1) + 1)
    For i = 1 To n
        If IsVarParent(comps(i)) Then
            need = q(i): avail = ""
            For r = 1 To UBound(v, 1)
                If need <= 0.0000001 Then Exit For
                If StrComp(Txt(v(r, N_BASE)), comps(i), vbTextCompare) = 0 And Txt(v(r, N_ACTIVE)) <> "Нет" Then
                    code = Txt(v(r, N_CODE))
                    ku = KeyIndex(used, UCase$(code))
                    If ku = 0 Then nu = nu + 1: KeyAdd used, UCase$(code), nu: ku = nu
                    have = BalGet("L|" & UCase$(code) & "|" & UCase$(loc)) - usedQ(ku)
                    If have > 0.0000001 Then
                        take = IIf(have < need, have, need)
                        m = m + 1
                        If m > UBound(oc) Then
                ReDim Preserve oc(1 To m * 2)
                ReDim Preserve oq(1 To m * 2)
            End If
                        oc(m) = code: oq(m) = Round(take, 6)
                        usedQ(ku) = usedQ(ku) + take
                        need = need - take
                        avail = avail & code & "=" & Format$(take, "0.###") & " "
                    End If
                End If
            Next r
            If need > 0.0000001 Then
                errs = errs & "• «" & comps(i) & "» (любое исполнение): в «" & loc & "» не хватает " & Format$(need, "0.###") & _
                       " из " & Format$(q(i), "0.###") & ". Выдайте исполнения в «" & loc & "»." & vbLf
            End If
        Else
            m = m + 1
            If m > UBound(oc) Then
                ReDim Preserve oc(1 To m * 2)
                ReDim Preserve oq(1 To m * 2)
            End If
            oc(m) = comps(i): oq(m) = q(i)
        End If
    Next i
    ResolveVariants = m
End Function

' Провести документ, предварительно раскрыв вариантных родителей
Private Function PostResolved(ByVal op As String, ByVal dt As Date, ByVal order As String, ByVal item As String, _
                              ByVal comment As String, comps() As String, q() As Double, ByVal n As Long) As String
    Dim oc() As String, oq() As Double, m As Long, errs As String, fL As String, tL As String
    Dim af() As String, at() As String, cc() As String, qq() As Double, i As Long
    If Len(OpFullName(op)) = 0 Then Fail "Операция «" & op & "» не найдена на листе «" & SH_OPS & "» (см. «Настройки»).": Exit Function
    OpRoute OpFullName(op), fL, tL
    m = ResolveVariants(comps, q, n, fL, oc, oq, errs)
    If Len(errs) > 0 Then Fail "Документ НЕ проведён:" & vbLf & vbLf & errs: Exit Function
    ReDim cc(1 To m): ReDim qq(1 To m): ReDim af(1 To m): ReDim at(1 To m)
    For i = 1 To m: cc(i) = oc(i): qq(i) = oq(i): Next i
    PostResolved = PostDoc(op, dt, order, item, "", comment, cc, qq, af, at)
End Function

' Заполнить строки листа «Ввод» по спецификации (C9 — изделие, C13 — количество изделий)
Public Sub Input_FillBySpec()
    FillInputByComposition SH_SPEC
End Sub

Public Sub Input_FillByBundle()
    FillInputByComposition SH_BUNDLE
End Sub

Private Sub FillInputByComposition(ByVal sheetName As String)
    Dim sh As Worksheet, item As String, k As Double, comps() As String, norms() As Double
    Dim n As Long, i As Long, r As Long, q() As Double, oc() As String, oq() As Double, m As Long, errs As String
    Dim loc As String, nParent As Long
    Set sh = WS(SH_INPUT)
    item = Txt(sh.Range(IN_ITEM).Value)
    k = Nz(sh.Range(IN_KQTY).Value)
    If Len(item) = 0 Then Fail "Укажите код изделия / объединения в ячейке " & IN_ITEM & ".": Exit Sub
    If k <= 0 Then Fail "Укажите количество изделий в ячейке " & IN_KQTY & ".": Exit Sub
    If sheetName = SH_SPEC And NomType(item) = T_PFD Then
        n = PFDComposition(item, comps, norms)
    Else
        n = GetComposition(sheetName, item, comps, norms)
    End If
    If n = 0 Then Fail "Для «" & item & "» нет строк на листе «" & IIf(NomType(item) = T_PFD, SH_PFD, sheetName) & "».": Exit Sub
    If InputLastRow() >= IN_FIRST Then
        If Not Ask("На листе уже есть строки. Заменить их составом «" & item & "»?") Then Exit Sub
    End If
    ReDim q(1 To n)
    For i = 1 To n: q(i) = Round(norms(i) * k, 6): Next i
    ' вариантные родители -> конкретные исполнения по наличию в месте «Откуда»
    loc = Txt(sh.Range(IN_FROM).Value)
    If Len(loc) = 0 Then loc = "ПР"
    m = ResolveVariants(comps, q, n, loc, oc, oq, errs)
    If Len(errs) > 0 Then
        ' не хватает исполнений — оставляем родителя, оператор выберет исполнение сам
        m = n: ReDim oc(1 To n): ReDim oq(1 To n)
        For i = 1 To n
            oc(i) = comps(i): oq(i) = q(i)
            If IsVarParent(comps(i)) Then nParent = nParent + 1
        Next i
    End If
    AppOff
    r = InputLastRow()
    If r >= IN_FIRST Then
        sh.Range(sh.Cells(IN_FIRST, L_CODE), sh.Cells(r, L_CODE)).ClearContents
        sh.Range(sh.Cells(IN_FIRST, L_QTY), sh.Cells(r, L_ATO)).ClearContents
    End If
    EnsureInputRows IN_FIRST + m - 1
    For i = 1 To m
        sh.Cells(IN_FIRST + i - 1, L_CODE).Value = oc(i)
        sh.Cells(IN_FIRST + i - 1, L_QTY).Value = oq(i)
    Next i
    AppOn
    If nParent > 0 Then
        Info "Добавлено строк: " & m & ". ВНИМАНИЕ: " & nParent & " строк(и) с кодом вариантного родителя — " & _
             "в «" & loc & "» не хватает исполнений. Замените код родителя на код исполнения (например …_v1), иначе документ не проведётся." & vbLf & errs
    Else
        Info "Добавлено строк: " & m & " (состав «" & item & "» x " & k & "; вариантные компоненты заменены исполнениями из «" & loc & "»)." & vbLf & "Проверьте операцию и нажмите «Провести»."
    End If
End Sub

' Гарантировать, что на листе «Ввод» есть оформленные строки до lastNeeded
Public Sub EnsureInputRows(ByVal lastNeeded As Long)
    Dim sh As Worksheet, lastFormula As Long, r As Long
    Set sh = WS(SH_INPUT)
    lastFormula = sh.Cells(sh.Rows.Count, L_NO).End(xlUp).Row
    If lastFormula >= lastNeeded Then Exit Sub
    sh.Range(sh.Cells(lastFormula, 1), sh.Cells(lastFormula, L_CHECK)).Copy _
        sh.Range(sh.Cells(lastFormula + 1, 1), sh.Cells(lastNeeded, L_CHECK))
    sh.Range(sh.Cells(lastFormula + 1, L_CODE), sh.Cells(lastNeeded, L_CODE)).ClearContents
    sh.Range(sh.Cells(lastFormula + 1, L_QTY), sh.Cells(lastNeeded, L_ATO)).ClearContents
End Sub

' ============================================================
' СБОРКА
'   ПФ (узел):        07 списание компонентов ПР -> СБОРКА + 08 выпуск СБОРКА -> ПР
'   Готовое изделие:  07 + 09 выпуск СБОРКА -> ГП
'   ПФД (некомплект): 40 списание компонентов, КРОМЕ недостающих + 41 выпуск на ПФД
' ============================================================
Public Function AssembleCore(ByVal item As String, ByVal qty As Double, ByVal order As String, _
                             ByVal dt As Date, Optional ByVal comment As String = "") As String
    Dim comps() As String, norms() As Double, n As Long, i As Long, t As String
    Dim q() As Double, d1 As String, d2 As String, opIn As String, opOut As String
    t = NomType(item)
    opIn = OPC("списание в сборку", OP_TO_ASM)
    If t = T_PF Then
        opOut = OPC("выпуск полуфабриката", OP_PF_OUT)
        n = GetComposition(SH_SPEC, item, comps, norms)
    ElseIf t = T_GP Then
        opOut = OPC("выпуск готового изделия", OP_GP_OUT)
        n = GetComposition(SH_SPEC, item, comps, norms)
    ElseIf t = T_PFD Then
        opIn = OPC("ПФД — списание компонентов", OP_PFD_ASM)
        opOut = OPC("ПФД — выпуск", OP_PFD_OUT)
        If Len(PFDTarget(item)) = 0 Then Fail "«" & item & "» нет на листе «" & SH_PFD & "» (что собираем и чего не хватает).": Exit Function
        n = PFDComposition(item, comps, norms)
    Else
        Fail "«" & item & "» имеет тип «" & t & "». Собирать можно Полуфабрикат, Готовое изделие или Некомплектный ПФ.": Exit Function
    End If
    If n = 0 Then Fail "Для «" & item & "» нет спецификации.": Exit Function
    If qty <= 0 Then Fail "Количество должно быть больше 0.": Exit Function
    If Len(Trim$(order)) = 0 Then Fail "Укажите № заказа (лист «Ввод», ячейка " & IN_ORDER & ").": Exit Function
    ReDim q(1 To n)
    For i = 1 To n
        q(i) = Round(norms(i) * qty, 6)
    Next i
    d1 = PostResolved(opIn, dt, order, item, "Сборка " & item & " x " & qty & ". " & comment, comps, q, n)
    If Len(d1) = 0 Then Exit Function
    d2 = PostDoc(opOut, dt, order, item, "", "Выпуск по документу " & d1 & ". " & comment, Array(item), Array(qty), Array(""), Array(""))
    If Len(d2) = 0 Then
        DeleteDocRows d1
        Fail "Выпуск не проведён, списание " & d1 & " отменено." & vbLf & gLastError
        Exit Function
    End If
    AssembleCore = d1 & "+" & d2
End Function

' ============================================================
' ДОУКОМПЛЕКТАЦИЯ / ПЕРЕКЛАССИФИКАЦИЯ ПФД
'   newCode = ""            : 42 ПФД -> СБОРКА + 43 недостающие ПР -> СБОРКА + 08/09 выпуск собираемого
'   newCode = другой ПФД    : 42 + 43 (только то, что пришло) + 41 выпуск нового кода ПФД
' ============================================================
Public Function PFDCompleteCore(ByVal pfd As String, ByVal qty As Double, ByVal order As String, _
                                ByVal dt As Date, Optional ByVal newCode As String = "") As String
    Dim tgt As String, mc() As String, mq() As Double, n As Long, i As Long, opOut As String, outCode As String
    Dim q() As Double, cc() As String, k As Long, d1 As String, d2 As String, d3 As String, diff As Double
    If NomType(pfd) <> T_PFD Then Fail "«" & pfd & "» не является некомплектным ПФ (тип «" & T_PFD & "»).": Exit Function
    tgt = PFDTarget(pfd)
    If Len(tgt) = 0 Then Fail "«" & pfd & "» нет на листе «" & SH_PFD & "».": Exit Function
    If qty <= 0 Then Fail "Количество должно быть больше 0.": Exit Function
    If Len(Trim$(order)) = 0 Then Fail "Укажите № заказа (лист «Ввод», ячейка " & IN_ORDER & ").": Exit Function
    n = PFDMissing(pfd, mc, mq)
    ReDim cc(1 To n + 1): ReDim q(1 To n + 1)
    If Len(newCode) = 0 Then
        outCode = tgt
        If NomType(tgt) = T_GP Then opOut = OPC("выпуск готового изделия", OP_GP_OUT) Else opOut = OPC("выпуск полуфабриката", OP_PF_OUT)
        For i = 1 To n
            k = k + 1: cc(k) = mc(i): q(k) = Round(mq(i) * qty, 6)
        Next i
    Else
        If NomType(newCode) <> T_PFD Or StrComp(PFDTarget(newCode), tgt, vbTextCompare) <> 0 Then
            Fail "Новый код «" & newCode & "» должен быть ПФД того же изделия «" & tgt & "».": Exit Function
        End If
        If StrComp(newCode, pfd, vbTextCompare) = 0 Then Fail "Новый код совпадает с текущим.": Exit Function
        outCode = newCode
        opOut = OPC("ПФД — выпуск", OP_PFD_OUT)
        For i = 1 To n
            diff = mq(i) - MissQty(newCode, mc(i))
            If diff > 0.0000001 Then k = k + 1: cc(k) = mc(i): q(k) = Round(diff * qty, 6)
        Next i
    Dim nc() As String, nq() As Double, n2 As Long
        n2 = PFDMissing(newCode, nc, nq)
        For i = 1 To n2
            If nq(i) > MissQty(pfd, nc(i)) + 0.0000001 Then
                Fail "Переклассификация невозможна: у «" & newCode & "» не хватает «" & nc(i) & "», который в «" & pfd & "» уже установлен.": Exit Function
            End If
        Next i
    End If
    d1 = PostDoc(OPC("доукомплектация — списание ПФД", OP_PFD_BACK), dt, order, outCode, "", "Доукомплектация " & pfd & " x " & qty, Array(pfd), Array(qty), Array(""), Array(""))
    If Len(d1) = 0 Then Exit Function
    If k > 0 Then
        d2 = PostResolved(OPC("доукомплектация — списание детали", OP_PFD_ADD), dt, order, outCode, "Пришедшие детали для " & pfd & " (" & d1 & ")", cc, q, k)
        If Len(d2) = 0 Then DeleteDocRows d1: Exit Function
    End If
    d3 = PostDoc(opOut, dt, order, outCode, "", "Выпуск " & outCode & " из " & pfd & " (" & d1 & ")", Array(outCode), Array(qty), Array(""), Array(""))
    If Len(d3) = 0 Then
        If Len(d2) > 0 Then DeleteDocRows d2
        DeleteDocRows d1
        Fail "Выпуск не проведён, доукомплектация отменена." & vbLf & gLastError
        Exit Function
    End If
    PFDCompleteCore = d1 & IIf(Len(d2) > 0, "+" & d2, "") & "+" & d3
End Function

' ============================================================
' ОБЪЕДИНЕНИЯ: комплектование (32 + 33) и разбор (30 + 31)
' ============================================================
Public Function BundleBuildCore(ByVal bundle As String, ByVal qty As Double, ByVal dt As Date) As String
    Dim comps() As String, norms() As Double, n As Long, i As Long, q() As Double
    Dim d1 As String, d2 As String
    If NomType(bundle) <> T_BUNDLE Then Fail "«" & bundle & "» не является объединением.": Exit Function
    n = GetComposition(SH_BUNDLE, bundle, comps, norms)
    If n = 0 Then Fail "Для «" & bundle & "» нет состава на листе «" & SH_BUNDLE & "».": Exit Function
    If qty <= 0 Then Fail "Количество должно быть больше 0.": Exit Function
    ReDim q(1 To n)
    For i = 1 To n: q(i) = Round(norms(i) * qty, 6): Next i
    d1 = PostResolved(OPC("комплектование — списание", OP_KIT_OUT), dt, "", bundle, "Комплектование " & bundle & " x " & qty, comps, q, n)
    If Len(d1) = 0 Then Exit Function
    d2 = PostDoc(OPC("комплектование — приход", OP_KIT_IN), dt, "", bundle, "", "Объединение по документу " & d1, Array(bundle), Array(qty), Array(""), Array(""))
    If Len(d2) = 0 Then DeleteDocRows d1: Exit Function
    BundleBuildCore = d1 & "+" & d2
End Function

Public Function BundleSplitCore(ByVal bundle As String, ByVal qty As Double, ByVal dt As Date) As String
    Dim comps() As String, norms() As Double, n As Long, i As Long, q() As Double, af() As String, at() As String
    Dim d1 As String, d2 As String
    If NomType(bundle) <> T_BUNDLE Then Fail "«" & bundle & "» не является объединением.": Exit Function
    n = GetComposition(SH_BUNDLE, bundle, comps, norms)
    If n = 0 Then Fail "Для «" & bundle & "» нет состава на листе «" & SH_BUNDLE & "».": Exit Function
    For i = 1 To n
        If IsVarParent(comps(i)) Then Fail "В составе «" & bundle & "» указан вариантный родитель «" & comps(i) & "». Для разбора укажите конкретное исполнение.": Exit Function
    Next i
    d1 = PostDoc(OPC("разбор — списание", OP_SPLIT_OUT), dt, "", bundle, "", "Разбор " & bundle & " x " & qty, Array(bundle), Array(qty), Array(""), Array(""))
    If Len(d1) = 0 Then Exit Function
    ReDim q(1 To n): ReDim af(1 To n): ReDim at(1 To n): ReDim Preserve comps(1 To n)
    For i = 1 To n: q(i) = Round(norms(i) * qty, 6): Next i
    d2 = PostDoc(OPC("разбор — приход", OP_SPLIT_IN), dt, "", bundle, "", "Составляющие из разбора " & d1, comps, q, af, at)
    If Len(d2) = 0 Then DeleteDocRows d1: Exit Function
    BundleSplitCore = d1 & "+" & d2
End Function

' Физически удалить строки документа (откат незавершённой комплексной операции)
Public Sub DeleteDocRows(ByVal docId As String)
    Dim sh As Worksheet, r As Long
    Set sh = WS(SH_MOVE)
    For r = LastRow(sh, J_OP) To FIRST_ROW Step -1
        If StrComp(Txt(sh.Cells(r, J_DOC).Value), docId, vbTextCompare) = 0 Then sh.Rows(r).Delete
    Next r
End Sub

' ---- кнопки листа «Ввод» (используют C9 и C13) ----
Private Function InputItemQty(ByRef item As String, ByRef qty As Double) As Boolean
    item = Txt(WS(SH_INPUT).Range(IN_ITEM).Value)
    qty = Nz(WS(SH_INPUT).Range(IN_KQTY).Value)
    If Len(item) = 0 Then Fail "Укажите код в ячейке " & IN_ITEM & " («Изделие / ПФ»).": Exit Function
    If qty <= 0 Then Fail "Укажите количество в ячейке " & IN_KQTY & ".": Exit Function
    InputItemQty = True
End Function

Private Function InputDate() As Date
    If IsDate(WS(SH_INPUT).Range(IN_DATE).Value) Then InputDate = CDate(WS(SH_INPUT).Range(IN_DATE).Value) Else InputDate = Date
End Function

Public Sub Input_Assemble()
    Dim item As String, qty As Double, r As String
    If Not InputItemQty(item, qty) Then Exit Sub
    If Not Ask("Собрать «" & item & "» в количестве " & qty & "?" & vbLf & vbLf & _
               "1) компоненты по спецификации спишутся из «ПР» в сборку (операция " & OPC("списание в сборку", OP_TO_ASM) & "; для ПФД — " & OPC("ПФД — списание компонентов", OP_PFD_ASM) & ", без недостающих);" & vbLf & _
               "2) узел будет выпущен в «ПР» (" & OPC("выпуск полуфабриката", OP_PF_OUT) & "), готовое изделие — на «ГП» (" & OPC("выпуск готового изделия", OP_GP_OUT) & "), ПФД — на «ПФД» (" & OPC("ПФД — выпуск", OP_PFD_OUT) & ")." & vbLf & _
               "Вариантные компоненты (родители) спишутся с имеющихся в «ПР» исполнений.") Then Exit Sub
    r = AssembleCore(item, qty, Txt(WS(SH_INPUT).Range(IN_ORDER).Value), InputDate(), Txt(WS(SH_INPUT).Range(IN_COMMENT).Value))
    If Len(r) > 0 Then ClearInput False: Info "Сборка проведена. Документы: " & r
End Sub

' Доукомплектовать ПФД (C9 — код ПФД, C13 — количество, C8 — № заказа)
Public Sub Input_PFDComplete()
    Dim item As String, qty As Double, r As String, tgt As String
    If Not InputItemQty(item, qty) Then Exit Sub
    tgt = PFDTarget(item)
    If Len(tgt) = 0 Then Fail "В ячейке " & IN_ITEM & " укажите код некомплектного ПФ (ПФД-…), который есть на листе «" & SH_PFD & "».": Exit Sub
    If Not Ask("Доукомплектовать «" & item & "» x " & qty & " и выпустить «" & tgt & "»?" & vbLf & vbLf & _
               "1) ПФД спишется со склада ПФД в сборку (" & OPC("доукомплектация — списание ПФД", OP_PFD_BACK) & ");" & vbLf & _
               "2) недостающие детали спишутся из «ПР» (" & OPC("доукомплектация — списание детали", OP_PFD_ADD) & ");" & vbLf & _
               "3) «" & tgt & "» будет выпущен (" & IIf(NomType(tgt) = T_GP, OPC("выпуск готового изделия", OP_GP_OUT) & " на ГП", OPC("выпуск полуфабриката", OP_PF_OUT) & " в ПР") & ").") Then Exit Sub
    r = PFDCompleteCore(item, qty, Txt(WS(SH_INPUT).Range(IN_ORDER).Value), InputDate())
    If Len(r) > 0 Then ClearInput False: Info "Доукомплектация проведена. Документы: " & r
End Sub

' Переклассифицировать ПФД: пришла часть недостающих деталей -> новый код ПФД
Public Sub Input_PFDReclass()
    Dim item As String, qty As Double, r As String, newCode As String
    If Not InputItemQty(item, qty) Then Exit Sub
    If Len(PFDTarget(item)) = 0 Then Fail "В ячейке " & IN_ITEM & " укажите код некомплектного ПФ (ПФД-…).": Exit Sub
    newCode = Trim$(InputBox("Пришла часть недостающих деталей для «" & item & "»." & vbLf & _
              "Введите НОВЫЙ код ПФД (того же изделия «" & PFDTarget(item) & "»), в который переходит ПФ:", "Переклассификация ПФД"))
    If Len(newCode) = 0 Then Exit Sub
    r = PFDCompleteCore(item, qty, Txt(WS(SH_INPUT).Range(IN_ORDER).Value), InputDate(), newCode)
    If Len(r) > 0 Then ClearInput False: Info "Переклассификация проведена. Документы: " & r
End Sub

Public Sub Input_BundleBuild()
    Dim item As String, qty As Double, r As String
    If Not InputItemQty(item, qty) Then Exit Sub
    r = BundleBuildCore(item, qty, InputDate())
    If Len(r) > 0 Then ClearInput False: Info "Объединение скомплектовано. Документы: " & r
End Sub

Public Sub Input_BundleSplit()
    Dim item As String, qty As Double, r As String
    If Not InputItemQty(item, qty) Then Exit Sub
    r = BundleSplitCore(item, qty, InputDate())
    If Len(r) > 0 Then ClearInput False: Info "Объединение разобрано. Документы: " & r
End Sub

' Заполнить строки всеми позициями, которые есть в месте «Откуда» (удобно для инвентаризации/перемещения всего)
Public Sub Input_FillFromStock()
    Dim sh As Worksheet, loc As String, shN As Worksheet, r As Long, lastR As Long, n As Long, b As Double, code As String
    Set sh = WS(SH_INPUT): Set shN = WS(SH_NOM)
    loc = Txt(sh.Range(IN_FROM).Value)
    If Len(loc) = 0 Then Fail "Сначала выберите операцию — место «Откуда» определится автоматически.": Exit Sub
    BuildBalances
    AppOff
    lastR = LastRow(shN, N_CODE)
    For r = FIRST_ROW To lastR
        code = Txt(shN.Cells(r, N_CODE).Value)
        b = BalGet("L|" & UCase$(code) & "|" & UCase$(loc))
        If b > 0 Then
            n = n + 1
            EnsureInputRows IN_FIRST + n - 1
            sh.Cells(IN_FIRST + n - 1, L_CODE).Value = code
            sh.Cells(IN_FIRST + n - 1, L_QTY).Value = b
        End If
    Next r
    AppOn
    Info "Добавлено позиций с остатком в «" & loc & "»: " & n & ". Исправьте количества при необходимости."
End Sub
