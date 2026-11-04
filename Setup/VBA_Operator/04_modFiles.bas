Attribute VB_Name = "modFiles"
Option Explicit
' ============================================================
' ELKA v8.2 — ДОКУМЕНТЫ И ФОТО
' • к операции можно приложить ЛЮБОЕ количество файлов;
' • каждому файлу оператор даёт ПРОИЗВОЛЬНОЕ название;
' • файл копируется в архив: <Папка архива>\ГГГГ\ММ\F000001_имя.ext;
' • в книге хранится путь ОТНОСИТЕЛЬНО папки архива, поэтому ссылки
'   работают и у оператора, и у пользователей просмотрщика;
' • фото открываются на листе «Просмотр», остальные файлы —
'   программой Windows по умолчанию (PDF, Word, Excel, ...).
' ============================================================

Public Function IsImageExt(ByVal ext As String) As Boolean
    IsImageExt = (InStr("|jpg|jpeg|png|bmp|gif|tif|tiff|", "|" & LCase$(ext) & "|") > 0)
End Function

Public Function NextFileId() As String
    Dim sh As Worksheet, lastR As Long, r As Long, mx As Long, t As String
    Set sh = WS(SH_FILES)
    lastR = LastRow(sh, F_ID)
    For r = FIRST_ROW To lastR
        t = Txt(sh.Cells(r, F_ID).Value)
        If Left$(t, 1) = "F" And IsNumeric(Mid$(t, 2)) Then
            If CLng(Mid$(t, 2)) > mx Then mx = CLng(Mid$(t, 2))
        End If
    Next r
    NextFileId = "F" & Format$(mx + 1, "000000")
End Function

' Полный путь из сохранённого (относительного или абсолютного)
Public Function FullFilePath(ByVal stored As String) As String
    stored = Trim$(stored)
    If Len(stored) = 0 Then Exit Function
    If Left$(stored, 2) = "\\" Or Mid$(stored, 2, 1) = ":" Or Left$(stored, 1) = "/" Then
        FullFilePath = stored
    Else
        FullFilePath = NormPath(ArchiveRoot() & stored)
    End If
End Function

' ============================================================
' ЯДРО: зарегистрировать файл. docId = "" -> ожидает проведения
' Возвращает ID файла или "".
' ============================================================
Public Function AttachFileCore(ByVal srcPath As String, ByVal title As String, _
                               Optional ByVal docId As String = "", Optional ByVal note As String = "") As String
    Dim sh As Worksheet, r As Long, fid As String, ext As String, rel As String
    Dim dstFolder As String, dstName As String, root As String

    gLastError = ""
    If Not FileExists(srcPath) Then Fail "Файл не найден:" & vbLf & srcPath: Exit Function
    title = Trim$(title)
    If Len(title) = 0 Then title = FileNameOnly(srcPath)
    ext = FileExt(srcPath)
    fid = NextFileId()

    If SettingYes("Копировать файлы в архив", True) Then
        root = ArchiveRoot()
        rel = Format$(Date, "yyyy") & "\" & Format$(Date, "mm") & "\"
        dstFolder = root & rel
        If Not EnsureFolder(dstFolder) Then Fail "Не удалось создать папку архива:" & vbLf & dstFolder: Exit Function
        dstName = fid & "_" & SafeName(title)
        If Len(ext) > 0 Then dstName = dstName & "." & ext
        On Error GoTo CopyFail
        FileCopy srcPath, NormPath(dstFolder & dstName)
        On Error GoTo 0
        rel = rel & dstName
    Else
        rel = srcPath
    End If

    Set sh = WS(SH_FILES)
    r = LastRow(sh, F_ID) + 1
    sh.Cells(r, F_ID).Value = fid
    sh.Cells(r, F_DOC).Value = IIf(Len(docId) = 0, PENDING_DOC, docId)
    sh.Cells(r, F_TITLE).Value = title
    sh.Cells(r, F_EXT).Value = UCase$(ext)
    sh.Cells(r, F_PATH).Value = rel
    sh.Cells(r, F_ORIG).Value = FileNameOnly(srcPath)
    sh.Cells(r, F_DATE).Value = Now
    sh.Cells(r, F_DATE).NumberFormat = "dd.mm.yyyy hh:mm"
    sh.Cells(r, F_USER).Value = UserName()
    sh.Cells(r, F_NOTE).Value = note
    AttachFileCore = fid
    If Len(docId) > 0 Then UpdateJournalFiles docId
    Exit Function
CopyFail:
    Fail "Не удалось скопировать файл в архив:" & vbLf & Err.Description
End Function

Public Function PendingFilesCount() As Long
    Dim sh As Worksheet, r As Long, lastR As Long
    Set sh = WS(SH_FILES)
    lastR = LastRow(sh, F_ID)
    For r = FIRST_ROW To lastR
        If Txt(sh.Cells(r, F_DOC).Value) = PENDING_DOC Then PendingFilesCount = PendingFilesCount + 1
    Next r
End Function

Public Sub LinkPendingFiles(ByVal docId As String)
    Dim sh As Worksheet, r As Long, lastR As Long
    Set sh = WS(SH_FILES)
    lastR = LastRow(sh, F_ID)
    For r = FIRST_ROW To lastR
        If Txt(sh.Cells(r, F_DOC).Value) = PENDING_DOC Then sh.Cells(r, F_DOC).Value = docId
    Next r
End Sub

Public Sub DiscardPendingFiles()
    Dim sh As Worksheet, r As Long
    Set sh = WS(SH_FILES)
    For r = LastRow(sh, F_ID) To FIRST_ROW Step -1
        If Txt(sh.Cells(r, F_DOC).Value) = PENDING_DOC Then sh.Rows(r).Delete
    Next r
End Sub

' Названия всех файлов документа через «; »
Public Function FilesTitles(ByVal docId As String) As String
    Dim sh As Worksheet, r As Long, lastR As Long, t As String
    Set sh = WS(SH_FILES)
    lastR = LastRow(sh, F_ID)
    For r = FIRST_ROW To lastR
        If StrComp(Txt(sh.Cells(r, F_DOC).Value), docId, vbTextCompare) = 0 Then
            If Len(t) > 0 Then t = t & "; "
            t = t & Txt(sh.Cells(r, F_TITLE).Value)
        End If
    Next r
    FilesTitles = t
End Function

' Обновить столбец «Файлы» у всех строк журнала документа
Public Sub UpdateJournalFiles(ByVal docId As String)
    Dim sh As Worksheet, r As Long, lastR As Long, t As String
    Set sh = WS(SH_MOVE)
    t = FilesTitles(docId)
    lastR = LastRow(sh, J_OP)
    For r = FIRST_ROW To lastR
        If StrComp(Txt(sh.Cells(r, J_DOC).Value), docId, vbTextCompare) = 0 Then sh.Cells(r, J_FILES).Value = t
    Next r
End Sub

Public Sub RefreshInputFiles()
    Dim sh As Worksheet, fs As Worksheet, r As Long, lastR As Long, t As String
    On Error Resume Next
    Set sh = WS(SH_INPUT): Set fs = WS(SH_FILES)
    lastR = LastRow(fs, F_ID)
    For r = FIRST_ROW To lastR
        If Txt(fs.Cells(r, F_DOC).Value) = PENDING_DOC Then
            If Len(t) > 0 Then t = t & "; "
            t = t & Txt(fs.Cells(r, F_TITLE).Value)
        End If
    Next r
    If Len(t) = 0 Then t = "(файлы не прикреплены)"
    sh.Range(IN_FILES).Value = t
End Sub

' ============================================================
' КНОПКИ
' ============================================================
' Прикрепить файлы к документу, который сейчас заполняется на листе «Ввод»
Public Sub Files_AttachToInput()
    AttachDialog ""
    RefreshInputFiles
End Sub

' Прикрепить файл к уже проведённому документу (выделите строку на листе «Движение»)
Public Sub Files_AttachToSelectedRow()
    Dim docId As String
    If ActiveSheet.Name <> SH_MOVE Then Fail "Перейдите на лист «" & SH_MOVE & "» и выделите строку операции.": Exit Sub
    If ActiveCell.Row < FIRST_ROW Then Fail "Выделите строку с операцией.": Exit Sub
    docId = Txt(WS(SH_MOVE).Cells(ActiveCell.Row, J_DOC).Value)
    If Len(docId) = 0 Then
        ' строка введена вручную — присвоим номер документа
        docId = NextDocId()
        WS(SH_MOVE).Cells(ActiveCell.Row, J_DOC).Value = docId
    End If
    AttachDialog docId
End Sub

Private Sub AttachDialog(ByVal docId As String)
    Dim f As Variant, i As Long, title As String, cnt As Long
    f = Application.GetOpenFilename("Все файлы (*.*),*.*,PDF (*.pdf),*.pdf,Изображения (*.jpg;*.jpeg;*.png;*.bmp),*.jpg;*.jpeg;*.png;*.bmp", _
                                    1, "Выберите один или несколько файлов", , True)
    If Not IsArray(f) Then Exit Sub
    For i = LBound(f) To UBound(f)
        title = InputBox("Введите ПРОИЗВОЛЬНОЕ название файла." & vbLf & vbLf & _
                         "Например: «Накладная №125 от поставщика», «Фото брака партии 7»." & vbLf & _
                         "Файл: " & FileNameOnly(CStr(f(i))), "Название файла", FileNameOnly(CStr(f(i))))
        If Len(Trim$(title)) > 0 Then
            If Len(AttachFileCore(CStr(f(i)), title, docId)) > 0 Then cnt = cnt + 1
        End If
    Next i
    If cnt > 0 Then Info "Прикреплено файлов: " & cnt & IIf(Len(docId) = 0, vbLf & "Они будут привязаны к документу при нажатии «Провести».", vbLf & "Документ: " & docId)
End Sub

' Переименовать (задать новое произвольное название) файл — выделите строку на листе «Файлы»
Public Sub Files_Rename()
    Dim sh As Worksheet, r As Long, t As String, docId As String
    Set sh = WS(SH_FILES)
    If ActiveSheet.Name <> SH_FILES Or ActiveCell.Row < FIRST_ROW Then Fail "Выделите строку файла на листе «" & SH_FILES & "».": Exit Sub
    r = ActiveCell.Row
    If Len(Txt(sh.Cells(r, F_ID).Value)) = 0 Then Exit Sub
    t = InputBox("Новое название файла:", "Переименовать", Txt(sh.Cells(r, F_TITLE).Value))
    If Len(Trim$(t)) = 0 Then Exit Sub
    sh.Cells(r, F_TITLE).Value = Trim$(t)
    docId = Txt(sh.Cells(r, F_DOC).Value)
    If docId <> PENDING_DOC Then
        UpdateJournalFiles docId
    Else
        RefreshInputFiles
    End If
End Sub

' Открыть файл(ы) строки: работает на листах Движение, Файлы, Справочник (фото)
Public Sub Files_OpenSelected()
    OpenForRow ActiveSheet, ActiveCell.Row
End Sub

Public Function OpenForRow(ByVal sh As Worksheet, ByVal r As Long) As Boolean
    Dim docId As String, p As String
    If r < FIRST_ROW Then Exit Function
    Select Case sh.Name
    Case SH_FILES
        p = Txt(sh.Cells(r, F_PATH).Value)
        If Len(p) = 0 Then Exit Function
        OpenStored p, Txt(sh.Cells(r, F_TITLE).Value)
        OpenForRow = True
    Case SH_MOVE
        docId = Txt(sh.Cells(r, J_DOC).Value)
        If Len(docId) = 0 Then Exit Function
        OpenFilesOfDoc docId
        OpenForRow = True
    Case SH_NOM
        p = Txt(sh.Cells(r, N_PHOTO).Value)
        If Len(p) = 0 Then Fail "У позиции нет фото. Путь к фото вносится в столбец «Фото» (кнопка «Фото позиции»).": Exit Function
        OpenStored p, Txt(sh.Cells(r, N_NAME).Value)
        OpenForRow = True
    End Select
End Function

Public Sub OpenFilesOfDoc(ByVal docId As String)
    Dim sh As Worksheet, r As Long, lastR As Long, idx() As Long, n As Long, lst As String, k As Variant
    Set sh = WS(SH_FILES)
    lastR = LastRow(sh, F_ID)
    ReDim idx(1 To lastR + 1)
    For r = FIRST_ROW To lastR
        If StrComp(Txt(sh.Cells(r, F_DOC).Value), docId, vbTextCompare) = 0 Then
            n = n + 1: idx(n) = r
            lst = lst & n & " — " & Txt(sh.Cells(r, F_TITLE).Value) & " (" & Txt(sh.Cells(r, F_EXT).Value) & ")" & vbLf
        End If
    Next r
    If n = 0 Then Fail "К документу " & docId & " файлы не прикреплены.": Exit Sub
    If n = 1 Then
        k = 1
    Else
        k = InputBox("К документу " & docId & " прикреплено файлов: " & n & vbLf & vbLf & lst & vbLf & "Введите номер файла:", "Открыть файл", "1")
        If Not IsNumeric(k) Then Exit Sub
        If CLng(k) < 1 Or CLng(k) > n Then Exit Sub
    End If
    OpenStored Txt(sh.Cells(idx(CLng(k)), F_PATH).Value), Txt(sh.Cells(idx(CLng(k)), F_TITLE).Value)
End Sub

' Открыть файл: изображения — на листе «Просмотр», остальное — программой Windows
Public Sub OpenStored(ByVal stored As String, ByVal title As String)
    Dim p As String
    p = FullFilePath(stored)
    If Not FileExists(p) Then
        Fail "Файл не найден:" & vbLf & p & vbLf & vbLf & _
             "Проверьте параметр «Папка архива документов» на листе «" & SH_SET & "» и доступ к сетевой папке."
        Exit Sub
    End If
    If IsImageExt(FileExt(p)) And SheetExists(SH_PREVIEW) Then
        ShowImage p, title
    Else
        OpenExternal p
    End If
End Sub

Public Sub OpenExternal(ByVal p As String)
    On Error GoTo Bad
    ThisWorkbook.FollowHyperlink Address:=p, NewWindow:=True
    Exit Sub
Bad:
    On Error Resume Next
    Shell "explorer.exe """ & p & """", vbNormalFocus
End Sub

' Встроенный просмотр изображения
Public Sub ShowImage(ByVal p As String, ByVal title As String)
    Dim sh As Worksheet, shp As Object, maxW As Double, maxH As Double, k As Double, i As Long
    Set sh = WS(SH_PREVIEW)
    On Error Resume Next
    For i = sh.Shapes.Count To 1 Step -1
        If Left$(sh.Shapes(i).Name, 9) = "ELKA_IMG_" Then sh.Shapes(i).Delete
    Next i
    On Error GoTo Bad
    sh.Range("B3").Value = title
    sh.Range("B4").Value = p
    Set shp = sh.Shapes.AddPicture(p, False, True, sh.Range("A7").Left + 4, sh.Range("A7").Top + 4, -1, -1)
    shp.Name = "ELKA_IMG_1"
    shp.LockAspectRatio = True
    maxW = 900: maxH = 560
    k = 1
    If shp.Width > maxW Then k = maxW / shp.Width
    If shp.Height * k > maxH Then k = maxH / shp.Height
    shp.Width = shp.Width * k
    sh.Activate
    sh.Range("A1").Select
    Exit Sub
Bad:
    OpenExternal p
End Sub

Public Sub Preview_OpenInProgram()
    Dim p As String
    p = Txt(WS(SH_PREVIEW).Range("B4").Value)
    If Len(p) > 0 Then OpenExternal p
End Sub

' Фото для позиции справочника (выделите строку на листе «Справочник»)
Public Sub Files_SetNomPhoto()
    Dim f As Variant, r As Long, fid As String, sh As Worksheet
    If ActiveSheet.Name <> SH_NOM Or ActiveCell.Row < FIRST_ROW Then Fail "Выделите строку позиции на листе «" & SH_NOM & "».": Exit Sub
    r = ActiveCell.Row
    Set sh = WS(SH_NOM)
    If Len(Txt(sh.Cells(r, N_CODE).Value)) = 0 Then Exit Sub
    f = Application.GetOpenFilename("Изображения (*.jpg;*.jpeg;*.png;*.bmp),*.jpg;*.jpeg;*.png;*.bmp,Все файлы (*.*),*.*", 1, "Фото позиции")
    If VarType(f) = vbBoolean Then Exit Sub
    fid = AttachFileCore(CStr(f), "Фото " & Txt(sh.Cells(r, N_CODE).Value) & " — " & Txt(sh.Cells(r, N_NAME).Value), "ФОТО-" & Txt(sh.Cells(r, N_CODE).Value))
    If Len(fid) = 0 Then Exit Sub
    sh.Cells(r, N_PHOTO).Value = WS(SH_FILES).Cells(LastRow(WS(SH_FILES), F_ID), F_PATH).Value
    Info "Фото привязано. Двойной щелчок по ячейке «Фото» открывает его."
End Sub

' Проверка: все ли файлы на месте
Public Function MissingFiles() As String
    Dim sh As Worksheet, r As Long, lastR As Long, t As String
    Set sh = WS(SH_FILES)
    lastR = LastRow(sh, F_ID)
    For r = FIRST_ROW To lastR
        If Len(Txt(sh.Cells(r, F_PATH).Value)) > 0 Then
            If Not FileExists(FullFilePath(Txt(sh.Cells(r, F_PATH).Value))) Then t = t & r & ";"
        End If
    Next r
    MissingFiles = t
End Function
