Attribute VB_Name = "modUtil"
Option Explicit
' ============================================================
' ELKA v8.2 — СЛУЖЕБНЫЕ ФУНКЦИИ
' ============================================================
Private mCalc As Long
Private mDepth As Long

Public Function WS(ByVal sheetName As String) As Worksheet
    Set WS = ThisWorkbook.Worksheets(sheetName)
End Function

Public Function SheetExists(ByVal sheetName As String) As Boolean
    Dim s As Worksheet
    On Error Resume Next
    Set s = ThisWorkbook.Worksheets(sheetName)
    SheetExists = Not (s Is Nothing)
    On Error GoTo 0
End Function

' Последняя заполненная строка по столбцу (не меньше строки заголовка)
Public Function LastRow(ByVal sh As Worksheet, ByVal col As Long) As Long
    Dim r As Long
    r = sh.Cells(sh.Rows.Count, col).End(xlUp).Row
    Do While r >= FIRST_ROW
        If Len(Trim$(CStr(sh.Cells(r, col).Value))) > 0 Then Exit Do
        r = r - 1
    Loop
    If r < HDR_ROW Then r = HDR_ROW
    LastRow = r
End Function

' Строка, где в столбце col находится значение key (без учёта регистра); 0 — нет
Public Function FindRow(ByVal sh As Worksheet, ByVal col As Long, ByVal key As String, _
                        Optional ByVal firstRow As Long = 4) As Long
    Dim r As Long, lastR As Long, v As Variant, i As Long
    key = Trim$(key)
    If Len(key) = 0 Then Exit Function
    lastR = LastRow(sh, col)
    If lastR < firstRow Then Exit Function
    v = sh.Range(sh.Cells(firstRow, col), sh.Cells(lastR, col)).Value
    If Not IsArray(v) Then
        If StrComp(Trim$(CStr(v)), key, vbTextCompare) = 0 Then FindRow = firstRow
        Exit Function
    End If
    For i = 1 To UBound(v, 1)
        If StrComp(Trim$(CStr(v(i, 1))), key, vbTextCompare) = 0 Then
            FindRow = firstRow + i - 1
            Exit Function
        End If
    Next i
End Function

Public Function Nz(ByVal v As Variant, Optional ByVal dflt As Double = 0) As Double
    On Error GoTo Bad
    If IsError(v) Then GoTo Bad
    If IsEmpty(v) Then GoTo Bad
    If Len(Trim$(CStr(v))) = 0 Then GoTo Bad
    Nz = CDbl(v)
    Exit Function
Bad:
    Nz = dflt
End Function

Public Function Txt(ByVal v As Variant) As String
    On Error Resume Next
    If IsError(v) Then Txt = "" Else Txt = Trim$(CStr(v))
End Function

Public Function IsYes(ByVal v As Variant) As Boolean
    Dim t As String
    t = LCase$(Txt(v))
    IsYes = (t = "да" Or t = "yes" Or t = "1" Or t = "true" Or t = "истина")
End Function

Public Function UserName() As String
    Dim u As String
    u = Setting("Оператор по умолчанию", "")
    If Len(u) = 0 Then u = Environ$("USERNAME")
    If Len(u) = 0 Then
        On Error Resume Next
        u = Application.UserName
        On Error GoTo 0
    End If
    If Len(u) = 0 Then u = "оператор"
    UserName = u
End Function

' ---- ускорение массовых операций (вложенные вызовы допустимы) ----
Public Sub AppOff()
    mDepth = mDepth + 1
    If mDepth > 1 Then Exit Sub
    On Error Resume Next
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    mCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
End Sub

Public Sub AppOn()
    If mDepth > 0 Then mDepth = mDepth - 1
    If mDepth > 0 Then Exit Sub
    On Error Resume Next
    If mCalc = 0 Then mCalc = xlCalculationAutomatic
    Application.Calculation = mCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

Public Sub AppReset()
    mDepth = 1
    AppOn
End Sub

' ---- сообщения (в режиме автотеста не показываются) ----
Public Sub Info(ByVal msg As String, Optional ByVal title As String = "ELKA")
    If gSilent Then
        gLog = gLog & "[i] " & msg & vbLf
    Else
        MsgBox msg, vbInformation, title
    End If
End Sub

Public Sub Fail(ByVal msg As String, Optional ByVal title As String = "ELKA — ошибка")
    gLastError = msg
    If gSilent Then
        gLog = gLog & "[x] " & msg & vbLf
    Else
        MsgBox msg, vbExclamation, title
    End If
End Sub

Public Function Ask(ByVal msg As String, Optional ByVal title As String = "ELKA") As Boolean
    If gSilent Then
        Ask = True
    Else
        Ask = (MsgBox(msg, vbQuestion + vbYesNo, title) = vbYes)
    End If
End Function

' ---- папки и файлы ----
Public Function FolderExists(ByVal p As String) As Boolean
    Dim a As Long
    On Error GoTo NoFolder
    p = Trim$(p)
    If Len(p) = 0 Then Exit Function
    Do While Len(p) > 3 And (Right$(p, 1) = "\" Or Right$(p, 1) = "/")
        p = Left$(p, Len(p) - 1)
    Loop
    a = GetAttr(p)
    FolderExists = ((a And vbDirectory) = vbDirectory)
    Exit Function
NoFolder:
    FolderExists = False
End Function

Public Function FileExists(ByVal p As String) As Boolean
    Dim a As Long
    On Error GoTo NoFile
    If Len(Trim$(p)) = 0 Then Exit Function
    a = GetAttr(p)
    FileExists = ((a And vbDirectory) = 0)
    Exit Function
NoFile:
    FileExists = False
End Function

Public Function EnsureFolder(ByVal p As String) As Boolean
    Dim parts() As String, i As Long, cur As String, startI As Long, sep As String
    On Error Resume Next
    p = NormPath(AddSlash(p))
    If FolderExists(p) Then EnsureFolder = True: Exit Function
    sep = IIf(Left$(p, 1) = "/", "/", "\")
    parts = Split(p, sep)
    If sep = "/" Then
        cur = "/"
        startI = 1
    ElseIf Left$(p, 2) = "\\" Then
        cur = "\\" & parts(2) & "\" & parts(3) & "\"
        startI = 4
    Else
        cur = parts(0) & "\"
        startI = 1
    End If
    For i = startI To UBound(parts)
        If Len(parts(i)) > 0 Then
            cur = cur & parts(i) & sep
            If Not FolderExists(cur) Then MkDir cur
        End If
    Next i
    EnsureFolder = FolderExists(p)
End Function

' Путь в формате ОС: если корень архива — Unix-путь (тесты в LibreOffice), меняем \ на /
Public Function NormPath(ByVal p As String) As String
    If Left$(p, 1) = "/" Then NormPath = Replace(p, "\", "/") Else NormPath = p
End Function

Public Function FileExt(ByVal p As String) As String
    Dim k As Long
    k = InStrRev(p, ".")
    If k > 0 And k > InStrRev(p, "\") Then FileExt = LCase$(Mid$(p, k + 1))
End Function

Public Function FileNameOnly(ByVal p As String) As String
    Dim k As Long
    k = InStrRev(Replace(p, "/", "\"), "\")
    FileNameOnly = Mid$(p, k + 1)
End Function

' Имя файла без запрещённых символов Windows
Public Function SafeName(ByVal t As String) As String
    Dim bad As Variant, i As Long
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|", vbTab, vbCr, vbLf)
    For i = 0 To UBound(bad)
        t = Replace(t, bad(i), "_")
    Next i
    t = Trim$(t)
    If Len(t) > 80 Then t = Left$(t, 80)
    If Len(t) = 0 Then t = "file"
    SafeName = t
End Function

' ---- резервная копия книги ----
Public Function MakeBackup(ByVal reason As String) As String
    Dim folder As String, fn As String
    On Error GoTo Bad
    If gSilent Then Exit Function
    If Len(ThisWorkbook.Path) = 0 Then Exit Function
    folder = BackupFolder()
    If Not EnsureFolder(folder) Then Exit Function
    fn = folder & "ELKA_" & Format$(Now, "yyyymmdd_hhmmss") & "_" & SafeName(reason) & ".xlsm"
    ThisWorkbook.SaveCopyAs fn
    MakeBackup = fn
    CleanupBackups folder
    Exit Function
Bad:
    MakeBackup = ""
End Function

Private Sub CleanupBackups(ByVal folder As String)
    Dim f As String, names() As String, n As Long, i As Long, j As Long, t As String, keep As Long
    On Error Resume Next
    keep = CLng(Nz(Setting("Хранить резервных копий", "30"), 30))
    If keep < 3 Then keep = 3
    ReDim names(1 To 1000)
    f = Dir$(folder & "ELKA_*.xlsm")
    Do While Len(f) > 0 And n < 1000
        n = n + 1: names(n) = f
        f = Dir$()
    Loop
    If n <= keep Then Exit Sub
    For i = 1 To n - 1
        For j = i + 1 To n
            If names(i) > names(j) Then t = names(i): names(i) = names(j): names(j) = t
        Next j
    Next i
    For i = 1 To n - keep
        Kill folder & names(i)
    Next i
End Sub

' ---- простой словарь на Collection (работает и в Excel, и в LibreOffice) ----
Public Function KeyIndex(ByVal col As Collection, ByVal key As String) As Long
    On Error GoTo NotFound
    KeyIndex = col(UCase$(key))
    Exit Function
NotFound:
    KeyIndex = 0
End Function

Public Sub KeyAdd(ByVal col As Collection, ByVal key As String, ByVal idx As Long)
    On Error Resume Next
    col.Add idx, UCase$(key)
End Sub

' Код операции «05» из полного названия «05 Передача...»
Public Function OpCode(ByVal opName As String) As String
    Dim t As String, k As Long
    t = Trim$(opName)
    k = InStr(t, " ")
    If k > 0 Then OpCode = Left$(t, k - 1) Else OpCode = t
End Function

' Полное название операции по коду или названию; "" — нет такой
Public Function OpFullName(ByVal opText As String) As String
    Dim sh As Worksheet, r As Long, lastR As Long, c As String
    opText = Trim$(opText)
    If Len(opText) = 0 Then Exit Function
    Set sh = WS(SH_OPS)
    lastR = LastRow(sh, 2)
    If IsNumeric(opText) And Len(opText) = 1 Then opText = "0" & opText
    For r = FIRST_ROW To lastR
        c = Txt(sh.Cells(r, 1).Value)
        If Len(c) = 1 And IsNumeric(c) Then c = "0" & c
        If StrComp(Txt(sh.Cells(r, 2).Value), opText, vbTextCompare) = 0 Or StrComp(c, opText, vbTextCompare) = 0 Then
            OpFullName = Txt(sh.Cells(r, 2).Value)
            Exit Function
        End If
    Next r
End Function

Public Function OpRow(ByVal opText As String) As Long
    Dim fullName As String
    fullName = OpFullName(opText)
    If Len(fullName) > 0 Then OpRow = FindRow(WS(SH_OPS), 2, fullName)
End Function

' Место хранения учитывается в остатках?
Public Function LocTracked(ByVal locCode As String) As Boolean
    Dim r As Long
    r = FindRow(WS(SH_LOC), 1, locCode)
    If r > 0 Then LocTracked = IsYes(WS(SH_LOC).Cells(r, 4).Value)
End Function

Public Function NomRow(ByVal code As String) As Long
    NomRow = FindRow(WS(SH_NOM), N_CODE, code)
End Function

Public Function NomName(ByVal code As String) As String
    Dim r As Long
    r = NomRow(code)
    If r > 0 Then NomName = Txt(WS(SH_NOM).Cells(r, N_NAME).Value)
End Function

Public Function NomType(ByVal code As String) As String
    Dim r As Long
    r = NomRow(code)
    If r > 0 Then NomType = Txt(WS(SH_NOM).Cells(r, N_TYPE).Value)
End Function

' Адрес принадлежит месту хранения и активен?
Public Function AddrOk(ByVal addr As String, ByVal locCode As String) As Boolean
    Dim r As Long, sh As Worksheet
    Set sh = WS(SH_ADDR)
    r = FindRow(sh, 1, addr)
    If r = 0 Then Exit Function
    If StrComp(Txt(sh.Cells(r, 2).Value), locCode, vbTextCompare) <> 0 Then Exit Function
    If Txt(sh.Cells(r, 8).Value) <> "" And Not IsYes(sh.Cells(r, 8).Value) Then Exit Function
    AddrOk = True
End Function

' Адрес по умолчанию для позиции, если он относится к месту хранения
Public Function DefaultAddr(ByVal code As String, ByVal locCode As String) As String
    Dim r As Long, a As String
    r = NomRow(code)
    If r = 0 Then Exit Function
    a = Txt(WS(SH_NOM).Cells(r, N_DEFADDR).Value)
    If Len(a) > 0 Then
        If AddrOk(a, locCode) Then DefaultAddr = a: Exit Function
    End If
    ' иначе — первый активный адрес этого места, если он единственный
    Dim sh As Worksheet, lastR As Long, found As String, cnt As Long
    Set sh = WS(SH_ADDR)
    lastR = LastRow(sh, 1)
    For r = FIRST_ROW To lastR
        If StrComp(Txt(sh.Cells(r, 2).Value), locCode, vbTextCompare) = 0 Then
            If Txt(sh.Cells(r, 8).Value) = "" Or IsYes(sh.Cells(r, 8).Value) Then
                cnt = cnt + 1
                If cnt = 1 Then found = Txt(sh.Cells(r, 1).Value)
            End If
        End If
    Next r
    If cnt = 1 Then DefaultAddr = found
End Function

Public Sub GoSheet(ByVal sheetName As String)
    On Error Resume Next
    WS(sheetName).Activate
    WS(sheetName).Range("A4").Select
End Sub

' ------------------------------------------------------------
' Перенести оформление строки srcRow на строки r1..r2 (столбцы 1..lastCol)
' без буфера обмена: формат чисел, шрифт, заливка, границы, проверка данных.
' ------------------------------------------------------------
Public Sub CopyRowFormat(ByVal sh As Worksheet, ByVal srcRow As Long, ByVal r1 As Long, ByVal r2 As Long, ByVal lastCol As Long)
    Dim c As Long, s As Range, d As Range, e As Variant, vt As Long, f1 As String
    On Error Resume Next
    For c = 1 To lastCol
        Set s = sh.Cells(srcRow, c)
        Set d = sh.Range(sh.Cells(r1, c), sh.Cells(r2, c))
        d.NumberFormat = s.NumberFormat
        d.HorizontalAlignment = s.HorizontalAlignment
        d.WrapText = s.WrapText
        d.Font.Bold = s.Font.Bold
        d.Font.Italic = s.Font.Italic
        d.Font.Color = s.Font.Color
        If s.Interior.ColorIndex = xlNone Then
            d.Interior.ColorIndex = xlNone
        Else
            d.Interior.Color = s.Interior.Color
        End If
        For Each e In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideHorizontal)
            If e = xlInsideHorizontal Then
                d.Borders(e).LineStyle = s.Borders(xlEdgeBottom).LineStyle
                d.Borders(e).Color = s.Borders(xlEdgeBottom).Color
            Else
                d.Borders(e).LineStyle = s.Borders(e).LineStyle
                d.Borders(e).Color = s.Borders(e).Color
            End If
        Next e
        Err.Clear
        vt = -1
        vt = s.Validation.Type
        If Err.Number = 0 And vt >= 0 Then
            f1 = s.Validation.Formula1
            d.Validation.Delete
            If vt = xlValidateList Then
                d.Validation.Add Type:=xlValidateList, AlertStyle:=s.Validation.AlertStyle, Formula1:=f1
                d.Validation.IgnoreBlank = True
                d.Validation.InCellDropdown = True
            ElseIf vt > 0 Then
                d.Validation.Add Type:=vt, AlertStyle:=s.Validation.AlertStyle, Operator:=s.Validation.Operator, Formula1:=f1
            End If
            d.Validation.ErrorMessage = s.Validation.ErrorMessage
            d.Validation.ShowError = s.Validation.ShowError
        End If
        Err.Clear
    Next c
End Sub
