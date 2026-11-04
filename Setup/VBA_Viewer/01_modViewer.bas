Attribute VB_Name = "modViewer"
Option Explicit
' ============================================================
' ELKA v8.2 — КНИГА-ПРОСМОТРЩИК ДЛЯ ПОЛЬЗОВАТЕЛЕЙ (ELKA_Viewer.xlsm)
' • «ОБНОВИТЬ ДАННЫЕ» — берёт последнюю версию ELKA_data_vNNNNNN.xlsx
'   из папки публикации (копирует к себе во временную папку, поэтому
'   оператору не мешает) и заменяет листы данных;
' • двойной щелчок по документу/файлу — открывает файл из архива:
'   фото — на листе «Просмотр», остальное — программой Windows;
' • данные только для чтения, правки вносит оператор.
' ============================================================
Public Const VIEWER_VERSION As String = "8.2"
Public Const V_DATA_PREFIX As String = "ELKA_data_v"
Public Const V_SHEETS As String = "Остатки|Остатки_адреса|Движение|Комплектность|Файлы"
Public Const V_HIDDEN_SHEETS As String = "|Файлы|"   ' служебные: нужны для открытия документов, пользователю не показываются
Public vSilent As Boolean          ' True — без окон (автообновление при открытии, автотест)
Public vLastMsg As String
Public Const V_PANEL As String = "Панель"
Public Const V_PREVIEW As String = "Просмотр"
Public Const V_CFG As String = "_cfg"
' столбцы опубликованных листов
Private Const VJ_DOC As Long = 15, VJ_FILES As Long = 16
Private Const VF_DOC As Long = 2, VF_TITLE As Long = 3, VF_EXT As Long = 4, VF_PATH As Long = 5
Private Const V_FIRST As Long = 4

' Ссылка на саму книгу-просмотрщик (запоминается до открытия файла данных)
Private mMe As Workbook
Public Function VB() As Workbook
    If mMe Is Nothing Then Set mMe = ThisWorkbook
    Set VB = mMe
End Function

' ---------------- сообщения ----------------
Private Sub VMsg(ByVal msg As String, Optional ByVal icon As Long = 64)
    vLastMsg = msg
    If vSilent Then
        Application.StatusBar = "ELKA: " & Replace(msg, vbLf, " ")
    Else
        MsgBox msg, icon, "ELKA"
    End If
End Sub

Public Function V_FileExists(ByVal p As String) As Boolean
    On Error GoTo No
    If Len(Trim$(p)) = 0 Then Exit Function
    V_FileExists = ((GetAttr(p) And vbDirectory) = 0)
    Exit Function
No:
    V_FileExists = False
End Function

' ---------------- настройки (скрытый лист _cfg) ----------------
Private Function Cfg() As Worksheet
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = VB().Worksheets(V_CFG)
    If sh Is Nothing Then
        Set sh = VB().Worksheets.Add(After:=VB().Worksheets(VB().Worksheets.Count))
        sh.Name = V_CFG
        sh.Range("A1").Value = "Папка публикации"
        sh.Range("A2").Value = "Папка архива (из публикации)"
        sh.Range("A3").Value = "Папка архива (вручную)"
        sh.Range("A4").Value = "Загруженная версия"
        sh.Range("A5").Value = "Дата публикации"
        sh.Range("A6").Value = "Автообновление при открытии (Да/Нет)"
        sh.Range("B6").Value = "Да"
        sh.Visible = 2   ' xlSheetVeryHidden
    End If
    Set Cfg = sh
End Function

Private Function Slash(ByVal p As String) As String
    p = Trim$(p)
    If Len(p) > 0 And Right$(p, 1) <> "\" And Right$(p, 1) <> "/" Then p = p & IIf(Left$(p, 1) = "/", "/", "\")
    Slash = p
End Function

Public Function V_PublishFolder() As String
    V_PublishFolder = Slash(CStr(Cfg().Range("B1").Value))
End Function

Public Function V_ArchiveRoot() As String
    Dim p As String
    p = Trim$(CStr(Cfg().Range("B3").Value))
    If Len(p) = 0 Then p = Trim$(CStr(Cfg().Range("B2").Value))
    V_ArchiveRoot = Slash(p)
End Function

' Кнопка «Настроить папки»
Public Sub Viewer_Setup()
    Dim p As String
    p = InputBox("Папка публикации данных (её сообщает оператор)." & vbLf & _
                 "Пример: \\SERVER\ELKA\Publish", "ELKA — папка данных", CStr(Cfg().Range("B1").Value))
    If Len(Trim$(p)) = 0 Then Exit Sub
    Cfg().Range("B1").Value = Trim$(p)
    p = InputBox("Папка архива документов (оставьте пустым — взять из публикации):", "ELKA — архив документов", CStr(Cfg().Range("B3").Value))
    Cfg().Range("B3").Value = Trim$(p)
    WritePanelInfo
    MsgBox "Настройки сохранены. Нажмите «ОБНОВИТЬ ДАННЫЕ».", vbInformation, "ELKA"
End Sub

' ---------------- обновление данных ----------------
Public Function LatestDataFile() As String
    Dim f As String, t As String, mx As Long, best As String, folder As String
    folder = V_PublishFolder()
    If Len(folder) = 0 Then Exit Function
    On Error Resume Next
    f = Dir$(folder & V_DATA_PREFIX & "*.xlsx")
    Do While Len(f) > 0
        t = Mid$(f, Len(V_DATA_PREFIX) + 1, 6)
        If t Like "######" Then
            If CLng(t) > mx Then mx = CLng(t): best = f
        End If
        f = Dir$()
    Loop
    LatestDataFile = best
End Function

Public Sub Viewer_Refresh()
    Dim src As String, tmp As String, wb As Workbook, names() As String, i As Long, cnt As Long, t0 As Single
    t0 = Timer
    Set mMe = ThisWorkbook
    If Len(V_PublishFolder()) = 0 And Not vSilent Then Viewer_Setup
    If Len(V_PublishFolder()) = 0 Then Exit Sub
    src = LatestDataFile()
    If Len(src) = 0 Then
        VMsg "В папке " & V_PublishFolder() & " нет файлов данных ELKA_data_v*.xlsx." & vbLf & vbLf & _
               "Причины: нет доступа к папке или оператор ещё не публиковал данные." & vbLf & _
               "Ранее загруженные данные остались без изменений.", vbExclamation
        Exit Sub
    End If
    On Error GoTo Bad
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    tmp = Environ$("TEMP")
    If Len(tmp) = 0 Then tmp = V_PublishFolder() Else tmp = Slash(tmp)
    tmp = tmp & "ELKA_" & Format$(Now, "yyyymmdd_hhmmss") & "_" & src
    FileCopy V_PublishFolder() & src, tmp
    Set wb = Workbooks.Open(tmp, 0, True)   ' файл, без обновления связей, только чтение
    names = Split(V_SHEETS, "|")
    For i = 0 To UBound(names)
        If ReplaceSheet(wb, names(i)) Then cnt = cnt + 1
    Next i
    On Error Resume Next
    Cfg().Range("B2").Value = CStr(wb.Worksheets("О_файле").Range("B6").Value)
    Cfg().Range("B5").Value = CStr(wb.Worksheets("О_файле").Range("B3").Value)
    On Error GoTo Bad
    wb.Close SaveChanges:=False
    Set wb = Nothing
    Kill tmp
    Cfg().Range("B4").Value = src
    WritePanelInfo
    VB().Worksheets(V_PANEL).Activate
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    VMsg "Данные обновлены." & vbLf & "Версия: " & src & vbLf & "Опубликовано: " & Cfg().Range("B5").Value & vbLf & "Листов: " & cnt & vbLf & "Время: " & Format$(Timer - t0, "0.0") & " сек."
    Exit Sub
Bad:
    Dim d As String
    d = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    If Len(tmp) > 0 Then Kill tmp
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    VMsg "Не удалось обновить данные: " & d & vbLf & "Ранее загруженные данные остались без изменений.", vbExclamation
End Sub

Private Function ReplaceSheet(ByVal wb As Workbook, ByVal nm As String) As Boolean
    Dim src As Worksheet, old As Worksheet, pos As Long, nw As Worksheet
    On Error Resume Next
    Set src = wb.Worksheets(nm)
    Set old = VB().Worksheets(nm)
    On Error GoTo 0
    If src Is Nothing Then Exit Function
    If Not old Is Nothing Then
        old.Visible = -1   ' xlSheetVisible
        pos = old.Index
        old.Unprotect
        old.Name = nm & "_old"
        src.Copy Before:=VB().Worksheets(pos)
        Set nw = VB().Worksheets(pos)
        old.Delete
    Else
        pos = VB().Worksheets(V_PREVIEW).Index
        src.Copy Before:=VB().Worksheets(pos)
        Set nw = VB().Worksheets(pos)
    End If
    nw.Name = nm
    On Error Resume Next
    nw.Protect DrawingObjects:=True, Contents:=True, AllowFiltering:=True, AllowSorting:=True, AllowFormattingColumns:=True
    If Err.Number <> 0 Then Err.Clear: nw.Protect
    On Error GoTo 0
    If InStr(V_HIDDEN_SHEETS, "|" & nm & "|") > 0 Then nw.Visible = 2   ' xlSheetVeryHidden
    ReplaceSheet = True
End Function

Public Sub WritePanelInfo()
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = VB().Worksheets(V_PANEL)
    sh.Unprotect
    sh.Range("C5").Value = IIf(Len(CStr(Cfg().Range("B5").Value)) > 0, Cfg().Range("B5").Value, "данные ещё не загружались")
    sh.Range("C6").Value = Cfg().Range("B4").Value
    sh.Range("C7").Value = V_PublishFolder()
    sh.Range("C8").Value = V_ArchiveRoot()
    sh.Protect DrawingObjects:=False, Contents:=True
End Sub

' ---------------- поиск ----------------
Public Sub Viewer_Find()
    Dim t As String, sh As Worksheet
    t = InputBox("Введите код или часть наименования (например DET-001 или «болт»):", "ELKA — поиск")
    If Len(Trim$(t)) = 0 Then Exit Sub
    Set sh = VB().Worksheets("Остатки")
    sh.Activate
    On Error Resume Next
    If sh.AutoFilterMode Then sh.AutoFilter.ShowAllData
    If Not sh.AutoFilterMode Then sh.Range("A3").CurrentRegion.AutoFilter
    If InStr(t, "-") > 0 Then
        sh.Range("A3").AutoFilter Field:=1, Criteria1:="*" & Trim$(t) & "*"
    Else
        sh.Range("A3").AutoFilter Field:=2, Criteria1:="*" & Trim$(t) & "*"
    End If
End Sub

Public Sub Viewer_ClearFilters()
    Dim sh As Worksheet
    On Error Resume Next
    For Each sh In VB().Worksheets
        If sh.AutoFilterMode Then sh.AutoFilter.ShowAllData
    Next sh
End Sub

' ---------------- файлы ----------------
Private Function ResolveV(ByVal stored As String) As String
    stored = Trim$(stored)
    If Left$(stored, 2) = "\\" Or Mid$(stored, 2, 1) = ":" Or Left$(stored, 1) = "/" Then
        ResolveV = stored
    Else
        ResolveV = V_ArchiveRoot() & stored
        If Left$(ResolveV, 1) = "/" Then ResolveV = Replace(ResolveV, "\", "/")
    End If
End Function

Public Sub Viewer_OpenSelected()
    V_OpenForRow ActiveSheet, ActiveCell.Row
End Sub

Public Function V_OpenForRow(ByVal sh As Worksheet, ByVal r As Long) As Boolean
    If r < V_FIRST Then Exit Function
    If sh.Name = "Файлы" Then
        If Len(CStr(sh.Cells(r, VF_PATH).Value)) > 0 Then V_OpenFile CStr(sh.Cells(r, VF_PATH).Value), CStr(sh.Cells(r, VF_TITLE).Value): V_OpenForRow = True
    ElseIf sh.Name = "Движение" Then
        If Len(CStr(sh.Cells(r, VJ_DOC).Value)) > 0 Then V_OpenDoc CStr(sh.Cells(r, VJ_DOC).Value): V_OpenForRow = True
    End If
End Function

Public Sub V_OpenDoc(ByVal docId As String)
    Dim sh As Worksheet, r As Long, lastR As Long, idx() As Long, n As Long, lst As String, k As Variant
    Set sh = VB().Worksheets("Файлы")
    lastR = sh.Cells(sh.Rows.Count, 1).End(xlUp).Row
    If lastR < V_FIRST Then
        VMsg "Файлов нет."
        Exit Sub
    End If
    ReDim idx(1 To lastR)
    For r = V_FIRST To lastR
        If StrComp(Trim$(CStr(sh.Cells(r, VF_DOC).Value)), docId, vbTextCompare) = 0 Then
            n = n + 1: idx(n) = r
            lst = lst & n & " — " & sh.Cells(r, VF_TITLE).Value & " (" & sh.Cells(r, VF_EXT).Value & ")" & vbLf
        End If
    Next r
    If n = 0 Then
        VMsg "К документу " & docId & " файлы не прикреплены."
        Exit Sub
    End If
    k = 1
    If n > 1 And Not vSilent Then
        k = InputBox("Файлов: " & n & vbLf & vbLf & lst & vbLf & "Введите номер:", "Открыть файл", "1")
        If Not IsNumeric(k) Then Exit Sub
        If CLng(k) < 1 Or CLng(k) > n Then Exit Sub
    End If
    V_OpenFile CStr(sh.Cells(idx(CLng(k)), VF_PATH).Value), CStr(sh.Cells(idx(CLng(k)), VF_TITLE).Value)
End Sub

Public Sub V_OpenFile(ByVal stored As String, ByVal title As String)
    Dim p As String, ext As String
    p = ResolveV(stored)
    If Not V_FileExists(p) Then
        VMsg "Файл не найден:" & vbLf & p & vbLf & vbLf & "Нет доступа к сетевой папке архива. Обратитесь к оператору или" & vbLf & _
               "укажите папку архива кнопкой «Настроить папки».", vbExclamation
        Exit Sub
    End If
    ext = LCase$(Mid$(p, InStrRev(p, ".") + 1))
    If InStr("|jpg|jpeg|png|bmp|gif|tif|tiff|", "|" & ext & "|") > 0 Then
        V_ShowImage p, title
    Else
        On Error Resume Next
        vLastMsg = "OPEN " & p
        If Not vSilent Then
            VB().FollowHyperlink Address:=p, NewWindow:=True
            If Err.Number <> 0 Then Shell "explorer.exe """ & p & """", vbNormalFocus
        End If
    End If
End Sub

Public Sub V_ShowImage(ByVal p As String, ByVal title As String)
    Dim sh As Worksheet, shp As Object, k As Double, i As Long
    Set sh = VB().Worksheets(V_PREVIEW)
    On Error Resume Next
    sh.Unprotect
    For i = sh.Shapes.Count To 1 Step -1
        If Left$(sh.Shapes(i).Name, 9) = "ELKA_IMG_" Then sh.Shapes(i).Delete
    Next i
    sh.Range("B3").Value = title
    sh.Range("B4").Value = p
    Err.Clear
    On Error GoTo NoPic
    Set shp = sh.Shapes.AddPicture(p, False, True, sh.Range("A7").Left + 4, sh.Range("A7").Top + 4, -1, -1)
    shp.Name = "ELKA_IMG_1"
    shp.LockAspectRatio = True
    k = 1
    If shp.Width > 900 Then k = 900 / shp.Width
    If shp.Height * k > 560 Then k = 560 / shp.Height
    shp.Width = shp.Width * k
    vLastMsg = "IMAGE " & p
    On Error Resume Next
    sh.Protect DrawingObjects:=False, Contents:=True
    sh.Activate
    Exit Sub
NoPic:
    ' не удалось вставить картинку — открыть стандартной программой Windows
    On Error Resume Next
    sh.Protect DrawingObjects:=False, Contents:=True
    vLastMsg = "OPEN " & p
    If Not vSilent Then VB().FollowHyperlink Address:=p, NewWindow:=True
End Sub

Public Sub Viewer_PreviewExternal()
    Dim p As String
    p = CStr(VB().Worksheets(V_PREVIEW).Range("B4").Value)
    If Len(p) > 0 Then VB().FollowHyperlink Address:=p, NewWindow:=True
End Sub

' ---------------- кнопки (создаются установщиком) ----------------
Private Sub VBtn(ByVal sh As Worksheet, ByVal cellAddr As String, ByVal caption As String, ByVal macro As String, ByVal w As Double, Optional ByVal h As Double = 28)
    Dim b As Object
    Set b = sh.Buttons.Add(sh.Range(cellAddr).Left + 2, sh.Range(cellAddr).Top + 2, w, h)
    b.OnAction = "'" & VB().Name & "'!" & macro
    b.Characters.Text = caption
    b.Font.Bold = True
    b.Font.Size = 10
    b.Name = "ELKA_" & macro
    b.Placement = 3
End Sub

Public Sub Viewer_SetupButtons()
    Dim sh As Worksheet, i As Long
    Set sh = VB().Worksheets(V_PANEL)
    sh.Unprotect
    On Error Resume Next
    For i = sh.Buttons.Count To 1 Step -1: sh.Buttons(i).Delete: Next i
    On Error GoTo 0
    VBtn sh, "E4", "ОБНОВИТЬ ДАННЫЕ", "Viewer_Refresh", 210, 34
    VBtn sh, "E7", "Найти позицию", "Viewer_Find", 210
    VBtn sh, "E9", "Сбросить фильтры", "Viewer_ClearFilters", 210
    VBtn sh, "E11", "Открыть файл выделенной строки", "Viewer_OpenSelected", 210
    VBtn sh, "E13", "Настроить папки", "Viewer_Setup", 210
    sh.Protect DrawingObjects:=False, Contents:=True
    Set sh = VB().Worksheets(V_PREVIEW)
    sh.Unprotect
    For i = sh.Buttons.Count To 1 Step -1: sh.Buttons(i).Delete: Next i
    VBtn sh, "F3", "Открыть в программе", "Viewer_PreviewExternal", 170, 26
    sh.Protect DrawingObjects:=False, Contents:=True
    WritePanelInfo
End Sub

Public Sub Viewer_InitConfig(ByVal publishFolder As String)
    Cfg().Range("B1").Value = publishFolder
    WritePanelInfo
End Sub

' Автообновление при открытии книги (если есть более новая публикация)
Public Sub Viewer_AutoRefresh()
    Dim latest As String
    On Error Resume Next
    WritePanelInfo
    If LCase$(Trim$(CStr(Cfg().Range("B6").Value))) = "нет" Then
        latest = LatestDataFile()
        If Len(latest) > 0 And latest <> CStr(Cfg().Range("B4").Value) Then
            Application.StatusBar = "ELKA: доступны новые данные — нажмите «ОБНОВИТЬ ДАННЫЕ» на листе «Панель»"
        End If
        Exit Sub
    End If
    latest = LatestDataFile()
    If Len(latest) = 0 Then Exit Sub
    If latest = CStr(Cfg().Range("B4").Value) Then
        Application.StatusBar = "ELKA: данные актуальны (" & latest & ")"
        Exit Sub
    End If
    vSilent = True
    Viewer_Refresh
    vSilent = False
End Sub

' Вызывается УСТАНОВЩИКОМ (Install_Viewer.ps1)
Public Function Install_ConfigureViewer(ByVal publishFolder As String, ByVal archiveFolder As String) As String
    On Error GoTo Bad
    Set mMe = ThisWorkbook
    Cfg().Range("B1").Value = Trim$(publishFolder)
    Cfg().Range("B3").Value = Trim$(archiveFolder)
    If Len(CStr(Cfg().Range("B6").Value)) = 0 Then Cfg().Range("B6").Value = "Да"
    Viewer_SetupButtons
    Install_ConfigureViewer = "OK"
    Exit Function
Bad:
    Install_ConfigureViewer = "ERROR: " & Err.Description
End Function
