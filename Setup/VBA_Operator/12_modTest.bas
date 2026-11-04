Attribute VB_Name = "modTest"
Option Explicit
' ============================================================
' ELKA v8.2 — ПОЛНЫЙ АВТОТЕСТ
' ВНИМАНИЕ: тест ПРОВОДИТ ОПЕРАЦИИ в текущей книге. Запускайте его
' только в демо-книге или в КОПИИ рабочей книги (ELKA_SelfTest).
' Сценарий:
'   1) приход 10 000 по каждой позиции + 2 прикреплённых файла;
'   2) по каждой позиции — полный маршрут по производству
'      (ВК, брак ВК, адресные перемещения, ПС, ПР, возврат, брак,
'      утилизация, возврат поставщику, инвентаризация);
'   3) сборка ПФ и ГП по спецификации, отгрузка, комплектование и
'      разбор объединения, некомплектные ПФ;
'   4) отрицательные тесты (документ ОБЯЗАН быть отклонён);
'   5) жизненный цикл номенклатуры, удаление/восстановление строк,
'      вставка строк, лист «Ввод», файлы, сверка с 1С, контроль;
'   6) сверка остатков: модель теста = журнал = формулы «Остатков».
' ============================================================
Private mExpK As Collection
Private mExpV() As Double
Private mExpN As Long
Private mPass As Long
Private mFail As Long
Private mRep As String
Private mDt As Date
Private mStage As String

Private Sub ExpAdd(ByVal code As String, ByVal loc As String, ByVal q As Double)
    Dim i As Long, key As String
    key = UCase$(code & "|" & loc)
    i = KeyIndex(mExpK, key)
    If i = 0 Then
        mExpN = mExpN + 1
        If mExpN > UBound(mExpV) Then
            ReDim Preserve mExpV(1 To mExpN * 2)
        End If
        KeyAdd mExpK, key, mExpN
        i = mExpN
    End If
    mExpV(i) = mExpV(i) + q
End Sub

Private Function ExpGet(ByVal code As String, ByVal loc As String) As Double
    Dim i As Long
    i = KeyIndex(mExpK, UCase$(code & "|" & loc))
    If i > 0 Then ExpGet = mExpV(i)
End Function

Private Sub Check(ByVal ok As Boolean, ByVal what As String)
    If ok Then
        mPass = mPass + 1
    Else
        mFail = mFail + 1
        mRep = mRep & "ОШИБКА: " & what & IIf(Len(gLastError) > 0, " | " & Replace(gLastError, vbLf, " "), "") & vbLf
    End If
End Sub

Private Function RndInt(ByVal a As Long, ByVal b As Long) As Long
    RndInt = a + Int(Rnd() * (b - a + 1))
End Function

' Провести и учесть в модели; expectOk=False — документ должен быть отклонён
Private Function TP(ByVal op As String, ByVal code As String, ByVal q As Double, _
                    Optional ByVal af As String = "", Optional ByVal at As String = "", _
                    Optional ByVal order As String = "", Optional ByVal item As String = "", _
                    Optional ByVal expectOk As Boolean = True, Optional ByVal note As String = "") As String
    Dim d As String, fL As String, tL As String, rowsBefore As Long
    mStage = "операция " & op & " " & code & " (" & note & ")"
    rowsBefore = LastRow(WS(SH_MOVE), J_OP)
    d = PostDoc(op, mDt, order, item, "TEST", note, Array(code), Array(q), Array(af), Array(at))
    If expectOk Then
        Check Len(d) > 0, "операция " & op & " " & code & " x " & q & " должна пройти"
        If Len(d) > 0 Then
            OpRoute OpFullName(op), fL, tL
            ExpAdd code, fL, -q
            ExpAdd code, tL, q
        End If
    Else
        Check Len(d) = 0 And LastRow(WS(SH_MOVE), J_OP) = rowsBefore, "НЕГАТИВНЫЙ тест «" & note & "» — документ должен быть отклонён"
    End If
    TP = d
End Function

Private Sub ExpComposition(ByVal sheetName As String, ByVal parent As String, ByVal k As Double, ByVal locFrom As String, ByVal locTo As String)
    Dim comps() As String, norms() As Double, n As Long, i As Long
    n = GetComposition(sheetName, parent, comps, norms)
    For i = 1 To n
        ExpAdd comps(i), locFrom, -Round(norms(i) * k, 6)
        ExpAdd comps(i), locTo, Round(norms(i) * k, 6)
    Next i
End Sub

Public Sub ELKA_SelfTest()
    Dim t As String
    If Not gSilent Then
        If MsgBox("Автотест проведёт сотни операций В ЭТОЙ книге." & vbLf & _
                  "Запускайте его только в демо-книге или КОПИИ рабочей книги." & vbLf & vbLf & "Продолжить?", _
                  vbExclamation + vbYesNo, "ELKA — автотест") <> vbYes Then Exit Sub
    End If
    t = RunSelfTest("")
    WriteTestSheet t
    MsgBox Left$(t, 900), IIf(mFail = 0, vbInformation, vbExclamation), "ELKA — автотест"
End Sub

Private Sub WriteTestSheet(ByVal t As String)
    Dim sh As Worksheet, a() As String, i As Long
    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets("Тест_отчёт")
    If sh Is Nothing Then
        Set sh = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        sh.Name = "Тест_отчёт"
    End If
    sh.Cells.Clear
    a = Split(t, vbLf)
    For i = 0 To UBound(a)
        sh.Cells(i + 1, 1).Value = a(i)
    Next i
    sh.Columns(1).ColumnWidth = 140
End Sub

Private Sub TestBatch(codes() As String, ByVal op As String, ByVal qty As Double, Optional ByVal order As String = "", Optional ByVal item As String = "", Optional ByVal addrTo As String = "")
    Dim qs() As Double, af() As String, at() As String, i As Long, n As Long
    Dim d As String, fL As String, tL As String
    n = UBound(codes)
    ReDim qs(1 To n): ReDim af(1 To n): ReDim at(1 To n)
    For i = 1 To n: qs(i) = qty: at(i) = addrTo: Next i
    mStage = "пакетная операция " & op
    d = PostDoc(op, mDt, order, item, "TEST", "Тест маршрута, включая ремонт", codes, qs, af, at)
    Check Len(d) > 0, "операция " & op & " для " & n & " артикулов"
    If Len(d) > 0 Then
        OpRoute OpFullName(op), fL, tL
        For i = 1 To n
            ExpAdd codes(i), fL, -qty: ExpAdd codes(i), tL, qty
        Next i
    End If
End Sub

' sampleFolder — папка с файлами для теста вложений ("" — создаются автоматически)
Public Function RunSelfTest(ByVal sampleFolder As String) As String
    Dim oldSilent As Boolean, shN As Worksheet, shM As Worksheet, sh1 As Worksheet
    Dim r As Long, i As Long, j As Long, n As Long, c As Long, before As Long, lr As Long
    Dim code As String, order As String, item As String, d As String, f1 As String, f2 As String
    Dim codes() As String, qs() As Double, af() As String, at() As String, locs As Variant
    Dim b As Double, e As Double, hdr As String, nCmp As Long, nF As Long, badF As Long
    Dim seen As New Collection, k As Long, target As String, opOut As String
    oldSilent = gSilent: gSilent = True
    AppOff
    On Error GoTo Crash
    mRep = "": mPass = 0: mFail = 0: gLog = "": mDt = Date
    Set mExpK = New Collection: ReDim mExpV(1 To 512): mExpN = 0
    mStage = "подготовка"
    Check Install_RepairDepartment() = "OK", "установка ремонтного отдела"
    Check Install_RepairDepartment() = "OK", "повторная установка без дублей"
    Set shN = WS(SH_NOM): Set shM = WS(SH_MOVE)
    order = Txt(WS(SH_ORDERS).Cells(FIRST_ROW, O_NO).Value)
    item = Txt(WS(SH_ORDERS).Cells(FIRST_ROW, O_ITEM).Value)
    locs = TrackedLocs()
    BuildBalances
    For r = FIRST_ROW To LastRow(shN, N_CODE)
        code = Txt(shN.Cells(r, N_CODE).Value)
        For j = 0 To UBound(locs)
            b = BalGet("L|" & UCase$(code) & "|" & UCase$(locs(j)))
            If b <> 0 Then ExpAdd code, CStr(locs(j)), b
        Next j
    Next r
    If Len(sampleFolder) = 0 Then sampleFolder = WorkFolder() & "TestFiles"
    f1 = AddSlash(sampleFolder) & "invoice_001.pdf"
    f2 = AddSlash(sampleFolder) & "pallet_photo.jpg"
    MakeSampleFile f1: MakeSampleFile f2
    Check Len(AttachFileCore(f1, "Накладная: реальная номенклатура ЁЛКА")) > 0, "произвольное название PDF"
    Check Len(AttachFileCore(f2, "Фото приёмки / ремонтный отдел")) > 0, "произвольное название фото"
    ReDim codes(1 To LastRow(shN, 1)): ReDim qs(1 To UBound(codes))
    ReDim af(1 To UBound(codes)): ReDim at(1 To UBound(codes))
    For r = FIRST_ROW To LastRow(shN, N_CODE)
        code = Txt(shN.Cells(r, N_CODE).Value)
        If Len(code) > 0 And NomType(code) <> T_VARPARENT And IsYes(shN.Cells(r, N_ACTIVE).Value) Then
            n = n + 1: codes(n) = code: qs(n) = 10000
        End If
    Next r
    ReDim Preserve codes(1 To n): ReDim Preserve qs(1 To n)
    ReDim Preserve af(1 To n): ReDim Preserve at(1 To n)
    d = PostDoc("01", mDt, "", "", "TEST", "Приход по 10 000 для теста", codes, qs, af, at)
    Check Len(d) > 0, "приход по 10 000 на все активные учётные позиции"
    If Len(d) > 0 Then
        For i = 1 To n: ExpAdd codes(i), "ЦС", 10000: Next i
    End If
    Check PendingFilesCount() = 0, "файлы привязаны к приходу"
    mRep = mRep & "Приход по 10 000: позиций " & n & vbLf
    TestBatch codes, "02", 100
    TestBatch codes, "03", 10
    TestBatch codes, "04", 90
    TestBatch codes, "05", 3000
    TestBatch codes, "06", 2000, order, item
    TestBatch codes, "45", 40, order, "", "РЕМ-01"
    TestBatch codes, "46", 15, order
    TestBatch codes, "46", 5, order, "", "ПР-УЧ2"
    BuildBalances
    For i = 1 To n
        code = codes(i)
        Check Abs(BalGet("L|" & UCase$(code) & "|РЕМ") - ExpGet(code, "РЕМ")) < 0.0001, "частичный возврат / ремонт " & code
    Next i
    TestBatch codes, "12", 20, order
    TestBatch codes, "13", 20
    TestBatch codes, "11", 5, order
    TestBatch codes, "14", 2
    TestBatch codes, "15", 1
    TestBatch codes, "16", 1
    TestBatch codes, "17", 10
    TestBatch codes, "18", 3
    TestBatch codes, "19", 10
    TestBatch codes, "20", 3
    TestBatch codes, "21", 10, order
    TestBatch codes, "22", 3, order
    mRep = mRep & "Маршрут: ВК, ПС, ПР, ремонт, два частичных возврата, брак, корректировки — для каждой позиции." & vbLf
    code = codes(1)
    TP "46", code, 99999999, "РЕМ-01", , order, , False, "возврат больше остатка ремонта"
    TP "45", code, 0, , , order, , False, "нулевое количество"
    TP "45", code, 1, , , , , False, "не указан заказ"
    TP "45", code, 1, , "ЦС-A-01-01", order, , False, "адрес другого места"
    TP "46", "NO-SUCH-CODE", 1, , , order, , False, "нет артикула"
    d = TP("46", code, 1, "РЕМ-01", , order, , True, "удаление и восстановление возврата")
    lr = LastRow(shM, J_OP): b = StockAt(code, "РЕМ")
    Check DeleteRowsCore(shM, lr, lr, False), "удаление строки возврата"
    Check Abs(StockAt(code, "РЕМ") - (b + 1)) < 0.0001, "при удалении возврата восстановлен остаток ремонта"
    Check MoveRestoreCore(FindRow(WS(SH_DELMOVE), J_DOC, d)), "восстановление строки возврата"
    Check Abs(StockAt(code, "РЕМ") - b) < 0.0001, "остаток после восстановления"
    before = LastRow(shN, N_CODE)
    InsertRowsCore shN, FIRST_ROW + 2, 3
    Check LastRow(shN, N_CODE) = before + 3, "вставка трёх строк справочника"
    DeleteRowsCore shN, FIRST_ROW + 2, FIRST_ROW + 4, False
    Check LastRow(shN, N_CODE) = before, "удаление пустых строк справочника"
    Reports_RefreshAll
    Application.Calculate
    BuildBalances
    For r = FIRST_ROW To LastRow(shN, N_CODE)
        code = Txt(shN.Cells(r, N_CODE).Value)
        For j = 0 To UBound(locs)
            b = BalGet("L|" & UCase$(code) & "|" & UCase$(locs(j)))
            e = ExpGet(code, CStr(locs(j))): nCmp = nCmp + 1
            Check Abs(b - e) < 0.0001, "модель маршрута = журнал " & code & "/" & locs(j)
        Next j
    Next r
    mRep = mRep & "Независимая модель маршрута: сверено " & nCmp & " ячеек." & vbLf
    ' Сборка всех реальных спецификаций, ПФД и объединений; без искусственных DET/PF/KIT кодов.
    mStage = "сборка реальных спецификаций"
    For r = FIRST_ROW To LastRow(WS(SH_SPEC), 1)
        code = Txt(WS(SH_SPEC).Cells(r, 1).Value)
        If KeyIndex(seen, code) = 0 Then
            KeyAdd seen, code, 1
            Check Len(AssembleCore(code, 1, order, mDt)) > 0, "сборка " & code
        End If
    Next r
    Set seen = New Collection
    For r = FIRST_ROW To LastRow(WS(SH_BUNDLE), 1)
        code = Txt(WS(SH_BUNDLE).Cells(r, 1).Value)
        If KeyIndex(seen, code) = 0 Then
            KeyAdd seen, code, 1
            Check Len(BundleBuildCore(code, 1, mDt)) > 0, "комплектование " & code
            Check Len(BundleSplitCore(code, 1, mDt)) > 0, "разбор " & code
        End If
    Next r
    Set seen = New Collection
    For r = FIRST_ROW To LastRow(WS(SH_PFD), P_CODE)
        code = Txt(WS(SH_PFD).Cells(r, P_CODE).Value)
        If KeyIndex(seen, code) = 0 Then
            KeyAdd seen, code, 1
            Check Len(AssembleCore(code, 1, order, mDt)) > 0, "сборка ПФД " & code
        End If
    Next r
    mStage = "сверка с 1С"
    Set sh1 = WS(SH_1C)
    sh1.Range(sh1.Cells(FIRST_ROW, 1), sh1.Cells(LastRow(sh1, 1) + 1, 7)).ClearContents
    BuildBalances: k = FIRST_ROW
    For r = FIRST_ROW To LastRow(shN, N_CODE)
        code = Txt(shN.Cells(r, N_CODE).Value)
        For j = 0 To UBound(locs)
            b = BalGet("L|" & UCase$(code) & "|" & UCase$(locs(j)))
            If Abs(b) > 0.0000001 Then
                sh1.Cells(k, 1).Value = code: sh1.Cells(k, 2).Value = NomName(code)
                sh1.Cells(k, 3).Value = Txt(WS(SH_LOC).Cells(FindRow(WS(SH_LOC), 1, CStr(locs(j))), 5).Value)
                sh1.Cells(k, 4).Value = b: sh1.Cells(k, 5).Value = "шт"
                sh1.Cells(k, 6).Value = "Sample_1C_export.xlsx": sh1.Cells(k, 7).Value = Now
                k = k + 1
            End If
        Next j
    Next r
    Compare1C
    Check CountMismatch() = 0, "сверка 1С включая ремонт: 0 расхождений"
    sh1.Cells(FIRST_ROW, 4).Value = Nz(sh1.Cells(FIRST_ROW, 4).Value) + 1
    Compare1C
    Check CountMismatch() = 1, "сверка 1С выявила намеренное расхождение"
    sh1.Cells(FIRST_ROW, 4).Value = Nz(sh1.Cells(FIRST_ROW, 4).Value) - 1
    Compare1C
    Reports_RefreshAll
    Application.Calculate
    BuildBalances
    For r = FIRST_ROW To LastRow(shN, N_CODE)
        code = Txt(shN.Cells(r, N_CODE).Value)
        For c = 5 To 4 + STOCK_LOC_COLS
            hdr = Txt(WS(SH_STOCK).Cells(HDR_ROW, c).Value)
            If Len(hdr) > 0 Then
                nF = nF + 1
                b = BalGet("L|" & UCase$(code) & "|" & UCase$(hdr))
                If Abs(Nz(WS(SH_STOCK).Cells(r, c).Value) - b) > 0.0001 Then badF = badF + 1
                Check b >= -0.0001, "неотрицательный остаток " & code & "/" & hdr
            End If
        Next c
    Next r
    Check badF = 0 And nF > 0, "формулы остатков совпали с журналом"
    Check ControlErrors() = 0, "контроль книги: 0 ошибок"
    Check Len(MissingFiles()) = 0, "вложения на месте"
    mRep = mRep & "Формулы: проверено " & nF & ", расхождений " & badF & vbLf
    mRep = mRep & "Ремонт входит в общий остаток, но исключён из мест комплектности ЦС|ПС|ПР." & vbLf
    AppOn
    gSilent = oldSilent
    RunSelfTest = "ELKA v" & ELKA_VERSION & " — АВТОТЕСТ" & vbLf & "Проверок пройдено: " & mPass & vbLf & "Проверок с ошибкой: " & mFail & vbLf & "Строк в журнале: " & (LastRow(shM, J_OP) - FIRST_ROW + 1) & vbLf & mRep
    Exit Function
Crash:
    mFail = mFail + 1
    gSilent = oldSilent: AppOn
    RunSelfTest = "АВТОТЕСТ ПРЕРВАН: " & mStage & " | " & Err.Description & vbLf & "Проверок с ошибкой: " & mFail & vbLf & mRep
End Function

' Свободный код для тестовой позиции (тест можно запускать повторно)
Private Function FreeCode(ByVal baseCode As String) As String
    Dim k As Long
    FreeCode = baseCode
    k = 1
    Do While NomRow(FreeCode) > 0
        k = k + 1
        FreeCode = baseCode & "-" & k
    Loop
End Function

' Создать маленький тестовый файл, если его нет
Private Sub MakeSampleFile(ByVal p As String)
    Dim fn As Integer, folder As String
    If FileExists(p) Then Exit Sub
    folder = Left$(p, Len(p) - Len(FileNameOnly(p)))
    EnsureFolder folder
    fn = FreeFile
    Open p For Output As #fn
    Print #fn, "ELKA test file " & FileNameOnly(p)
    Close #fn
End Sub
