Attribute VB_Name = "modPublish"
Option Explicit
' ============================================================
' ELKA v8.2 — ПУБЛИКАЦИЯ ДАННЫХ ДЛЯ ПОЛЬЗОВАТЕЛЕЙ
' Книгу оператора открывает ОДИН человек. Пользователи смотрят данные
' в книге-просмотрщике ELKA_Viewer.xlsm, которая забирает последнюю
' опубликованную копию ELKA_data_v000123.xlsx из «Папки публикации».
' В копии — только значения (без формул и макросов) и только листы
' для просмотра: остатки, остатки по адресам, движение, комплектность
' и (скрытый) список файлов для открытия документов.
' ============================================================

Public Sub Publish_Data()
    Dim folder As String, wb As Workbook, names() As String, i As Long, sh As Worksheet
    Dim verNo As Long, target As String, nErr As Long, t0 As Single, shp As Object, arch As String, tmp As String
    Dim nRows As Long, who As String

    t0 = Timer
    folder = PublishFolder()
    If Not EnsureFolder(folder) Then Fail "Папка публикации недоступна:" & vbLf & folder & vbLf & "Проверьте параметр «Папка публикации» на листе «" & SH_SET & "».": Exit Sub
    Reports_RefreshAllSilent
    nErr = ControlErrors()
    If nErr > 0 Then
        If Not Ask("На листе «" & SH_CTRL & "» найдено ошибок: " & nErr & "." & vbLf & "Пользователи увидят данные с этими ошибками. Всё равно опубликовать?", "Публикация") Then Exit Sub
    End If
    arch = ArchiveRoot()
    If Left$(arch, 2) <> "\\" And Not gSilent Then
        If Not Ask("Папка архива документов — локальная (" & arch & ")." & vbLf & _
                   "Пользователи на других компьютерах НЕ смогут открыть прикреплённые файлы." & vbLf & _
                   "Рекомендуется сетевая папка вида \\сервер\ELKA\Архив." & vbLf & vbLf & "Продолжить публикацию?", "Публикация") Then Exit Sub
    End If

    On Error GoTo Bad
    AppOff
    Application.DisplayAlerts = False
    ' 1) снимок книги во временный файл (книга оператора не трогается)
    nRows = LastRow(WS(SH_MOVE), J_OP) - FIRST_ROW + 1
    who = UserName()
    tmp = NormPath(folder & "~ELKA_publish_" & Format$(Now, "yyyymmdd_hhmmss") & ".xlsm")
    ThisWorkbook.SaveCopyAs tmp
    Application.EnableEvents = False
    Set wb = Workbooks.Open(tmp, 0, False)   ' файл, без обновления связей, для записи
    Application.EnableEvents = True
    names = Split(PUBLISH_SHEETS, "|")
    ' 2) нужные листы -> только значения (до удаления остальных, чтобы не было #ССЫЛКА!)
    For Each sh In wb.Worksheets
        If InList(sh.Name, names) Then
            sh.Unprotect
            sh.Visible = xlSheetVisible
            If sh.AutoFilterMode Then sh.AutoFilterMode = False
            sh.UsedRange.Value = sh.UsedRange.Value
            sh.Range("A2").Value = "Данные на " & Format$(Now, "dd.mm.yyyy hh:mm") & " (публикация оператора). Только просмотр: изменения вносит оператор. Фильтр — стрелки в заголовках."
            For Each shp In sh.Shapes
                shp.Delete
            Next shp
            On Error Resume Next
            sh.Cells.Validation.Delete
            sh.Hyperlinks.Delete
            On Error GoTo Bad
        End If
    Next sh
    ' 3) удалить все остальные листы (операторские) и имена
    For i = wb.Worksheets.Count To 1 Step -1
        If Not InList(wb.Worksheets(i).Name, names) Then
            wb.Worksheets(i).Visible = xlSheetVisible
            wb.Worksheets(i).Delete
        End If
    Next i
    On Error Resume Next
    For i = wb.Names.Count To 1 Step -1
        wb.Names(i).Delete
    Next i
    On Error GoTo Bad
    AddAboutSheet wb, arch, nRows, who
    verNo = LatestVersion(folder) + 1
    target = folder & DATA_PREFIX & Format$(verNo, "000000") & ".xlsx"
    Do While FileExists(target)
        verNo = verNo + 1
        target = folder & DATA_PREFIX & Format$(verNo, "000000") & ".xlsx"
    Loop
    wb.SaveAs Filename:=target, FileFormat:=51
    wb.Close SaveChanges:=False
    Set wb = Nothing
    On Error Resume Next
    Kill tmp
    On Error GoTo Bad
    gLastPublished = target
    CleanupVersions folder
    Application.DisplayAlerts = True
    AppOn
    ThisWorkbook.Activate
    Info "Данные опубликованы." & vbLf & vbLf & "Файл: " & target & vbLf & "Время: " & Format$(Timer - t0, "0.0") & " сек." & vbLf & vbLf & _
         "Пользователи получат их кнопкой «ОБНОВИТЬ ДАННЫЕ» в ELKA_Viewer.xlsm.", "Публикация"
    Exit Sub
Bad:
    Dim d As String
    d = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    If Len(tmp) > 0 Then Kill tmp
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    AppOn
    Fail "Публикация не выполнена: " & d
End Sub

Private Sub Reports_RefreshAllSilent()
    Dim old As Boolean
    old = gSilent
    gSilent = True
    Reports_RefreshAll
    gSilent = old
End Sub

Private Sub AddAboutSheet(ByVal wb As Workbook, ByVal arch As String, ByVal nRows As Long, ByVal who As String)
    Dim sh As Worksheet
    Set sh = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    sh.Name = "О_файле"
    sh.Range("A1").Value = "ELKA — опубликованные данные (только для чтения)"
    sh.Range("A1").Font.Bold = True: sh.Range("A1").Font.Size = 14
    sh.Range("A3").Value = "Опубликовано": sh.Range("B3").NumberFormat = "@": sh.Range("B3").Value = Format$(Now, "dd.mm.yyyy hh:mm:ss")
    sh.Range("A4").Value = "Оператор": sh.Range("B4").Value = who
    sh.Range("A5").Value = "Версия ELKA": sh.Range("B5").Value = "v" & ELKA_VERSION
    sh.Range("A6").Value = "Папка архива документов": sh.Range("B6").Value = arch
    sh.Range("A7").Value = "Строк журнала": sh.Range("B7").Value = nRows
    sh.Columns("A").ColumnWidth = 30
    sh.Columns("B").ColumnWidth = 70
End Sub

Public Function LatestVersion(ByVal folder As String) As Long
    Dim f As String, t As String, mx As Long
    f = Dir$(folder & DATA_PREFIX & "*.xlsx")
    Do While Len(f) > 0
        t = Mid$(f, Len(DATA_PREFIX) + 1, 6)
        If t Like "######" Then
            If CLng(t) > mx Then mx = CLng(t)
        End If
        f = Dir$()
    Loop
    LatestVersion = mx
End Function

Private Sub CleanupVersions(ByVal folder As String)
    Dim keep As Long, latest As Long, f As String, t As String, lst As String, a() As String, i As Long
    keep = CLng(Nz(Setting("Хранить версий публикации", "20")))
    If keep < 3 Then keep = 3
    latest = LatestVersion(folder)
    f = Dir$(folder & DATA_PREFIX & "*.xlsx")
    Do While Len(f) > 0
        t = Mid$(f, Len(DATA_PREFIX) + 1, 6)
        If t Like "######" Then
            If CLng(t) <= latest - keep Then lst = lst & f & "|"
        End If
        f = Dir$()
    Loop
    If Len(lst) = 0 Then Exit Sub
    a = Split(Left$(lst, Len(lst) - 1), "|")
    On Error Resume Next
    For i = 0 To UBound(a)
        Kill folder & a(i)
    Next i
End Sub

Private Function InList(ByVal nm As String, names() As String) As Boolean
    Dim i As Long
    For i = LBound(names) To UBound(names)
        If StrComp(nm, names(i), vbTextCompare) = 0 Then InList = True: Exit Function
    Next i
End Function
