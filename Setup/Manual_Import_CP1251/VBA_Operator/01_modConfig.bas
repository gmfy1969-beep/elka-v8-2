Attribute VB_Name = "modConfig"
Option Explicit
' ============================================================
' ELKA v8.2 — НАСТРОЙКИ И КОНСТАНТЫ (книга оператора)
' Единственное место, где описаны имена листов и столбцов.
' ============================================================
Public Const ELKA_VERSION As String = "8.2"

' ---- листы книги оператора ----
Public Const SH_PANEL As String = "Панель"
Public Const SH_INPUT As String = "Ввод"
Public Const SH_MOVE As String = "Движение"
Public Const SH_STOCK As String = "Остатки"
Public Const SH_ADDRSTOCK As String = "Остатки_адреса"
Public Const SH_NOM As String = "Справочник"
Public Const SH_SPEC As String = "Спецификация"
Public Const SH_BUNDLE As String = "Объединения"
Public Const SH_VARSUM As String = "Свод_вариантов"
Public Const SH_KITS As String = "Комплектность"
Public Const SH_LOC As String = "Места хранения"
Public Const SH_ADDR As String = "Адреса"
Public Const SH_OPS As String = "Операции"
Public Const SH_FILES As String = "Файлы"
Public Const SH_PREVIEW As String = "Просмотр"
Public Const SH_1C As String = "Импорт_1С"
Public Const SH_CMP As String = "Сравнение_1С"
Public Const SH_CTRL As String = "Контроль"
Public Const SH_SET As String = "Настройки"
Public Const SH_TRASH As String = "Корзина"
Public Const SH_DELMOVE As String = "Движение_удалённые"
Public Const SH_HELP As String = "Инструкция_оператора"
Public Const SH_PFD As String = "Некомплектные_ПФ"
Public Const SH_ORDERS As String = "Заказы"

' Листы, которые публикуются пользователям (книга-просмотрщик)
Public Const PUBLISH_SHEETS As String = _
    "Остатки|Остатки_адреса|Движение|Комплектность|Файлы"
Public Const DATA_PREFIX As String = "ELKA_data_v"

Public Const HDR_ROW As Long = 3          ' строка заголовков всех таблиц
Public Const FIRST_ROW As Long = 4        ' первая строка данных

' ---- столбцы листа «Движение» ----
Public Const J_NO As Long = 1        ' A  №
Public Const J_DATE As Long = 2      ' B  Дата
Public Const J_OP As Long = 3        ' C  Операция
Public Const J_CODE As Long = 4      ' D  Код
Public Const J_NAME As Long = 5      ' E  Наименование (формула)
Public Const J_QTY As Long = 6       ' F  Количество
Public Const J_UNIT As Long = 7      ' G  Ед. (формула)
Public Const J_FROM As Long = 8      ' H  Откуда (формула)
Public Const J_TO As Long = 9        ' I  Куда (формула)
Public Const J_AFROM As Long = 10    ' J  Адрес откуда
Public Const J_ATO As Long = 11      ' K  Адрес куда
Public Const J_ORDER As Long = 12    ' L  № заказа
Public Const J_ITEM As Long = 13     ' M  Изделие / ПФ
Public Const J_BATCH As Long = 14    ' N  Партия / поставка
Public Const J_DOC As Long = 15      ' O  Документ №
Public Const J_FILES As Long = 16    ' P  Файлы (произвольные названия)
Public Const J_USER As Long = 17     ' Q  Оператор
Public Const J_COMMENT As Long = 18  ' R  Комментарий
Public Const J_CHECK As Long = 19    ' S  Проверка (формула)
Public Const J_CREATED As Long = 20  ' T  Создано
Public Const J_LAST As Long = 20

' ---- столбцы «Справочник» ----
Public Const N_CODE As Long = 1, N_NAME As Long = 2, N_TYPE As Long = 3, N_UNIT As Long = 4
Public Const N_GROUP As Long = 5, N_BASE As Long = 6, N_VARIANT As Long = 7, N_MIN As Long = 8
Public Const N_DEFADDR As Long = 9, N_ACTIVE As Long = 10, N_PHOTO As Long = 11, N_NOTE As Long = 12
Public Const N_LAST As Long = 12

' ---- типы номенклатуры ----
Public Const T_PART As String = "Деталь"
Public Const T_MAT As String = "Материал"
Public Const T_PF As String = "Полуфабрикат"
Public Const T_PFD As String = "Некомплектный ПФ"
Public Const T_GP As String = "Готовое изделие"
Public Const T_BUNDLE As String = "Объединение"
Public Const T_VARPARENT As String = "Вариантный родитель"

' ---- операции, используемые автоматикой (коды по умолчанию) ----
' Реальные коды берутся с листа «Настройки» (функция OPC), здесь — значения
' по умолчанию, совпадающие со справочником операций ЁЛКА.
Public Const OP_RECEIPT As String = "01"
Public Const OP_TO_ASM As String = "07"
Public Const OP_PF_OUT As String = "08"
Public Const OP_GP_OUT As String = "09"
Public Const OP_SPLIT_OUT As String = "30"
Public Const OP_SPLIT_IN As String = "31"
Public Const OP_KIT_OUT As String = "32"
Public Const OP_KIT_IN As String = "33"
Public Const OP_PFD_ASM As String = "40"
Public Const OP_PFD_OUT As String = "41"
Public Const OP_PFD_BACK As String = "42"
Public Const OP_PFD_ADD As String = "43"

' ---- столбцы листа «Некомплектные_ПФ» ----
Public Const P_CODE As Long = 1, P_TARGET As Long = 2, P_MISS As Long = 4, P_QTY As Long = 6

' ---- столбцы листа «Заказы» ----
Public Const O_NO As Long = 1, O_DATE As Long = 2, O_ITEM As Long = 3, O_PLAN As Long = 5

' ---- лист «Ввод» ----
Public Const IN_DATE As String = "C4"
Public Const IN_OP As String = "C5"
Public Const IN_FROM As String = "C6"
Public Const IN_TO As String = "C7"
Public Const IN_ORDER As String = "C8"
Public Const IN_ITEM As String = "C9"
Public Const IN_BATCH As String = "C10"
Public Const IN_COMMENT As String = "C11"
Public Const IN_FILES As String = "C12"
Public Const IN_KQTY As String = "C13"
Public Const IN_HDR As Long = 15         ' заголовок строк документа
Public Const IN_FIRST As Long = 16       ' первая строка документа
' столбцы строк документа
Public Const L_NO As Long = 1, L_CODE As Long = 2, L_NAME As Long = 3, L_UNIT As Long = 4
Public Const L_QTY As Long = 5, L_AFROM As Long = 6, L_ATO As Long = 7, L_AVAIL As Long = 8, L_CHECK As Long = 9

' ---- лист «Файлы» ----
Public Const F_ID As Long = 1, F_DOC As Long = 2, F_TITLE As Long = 3, F_EXT As Long = 4
Public Const F_PATH As Long = 5, F_ORIG As Long = 6, F_DATE As Long = 7, F_USER As Long = 8, F_NOTE As Long = 9
Public Const PENDING_DOC As String = "ОЖИДАЕТ"

' ---- режим без диалогов (автотесты) ----
Public gSilent As Boolean
Public gLastError As String
Public gLog As String
Public gLastPublished As String

' ============================================================
' Чтение параметра с листа «Настройки» (столбец A — имя, B — значение)
' ============================================================
Public Function Setting(ByVal paramName As String, Optional ByVal defaultValue As String = "") As String
    Dim ws As Worksheet, r As Long, lastR As Long
    On Error GoTo Fallback
    Set ws = ThisWorkbook.Worksheets(SH_SET)
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = FIRST_ROW To lastR
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), paramName, vbTextCompare) = 0 Then
            Setting = Trim$(CStr(ws.Cells(r, 2).Value))
            If Len(Setting) = 0 Then Setting = defaultValue
            Exit Function
        End If
    Next r
Fallback:
    Setting = defaultValue
End Function

' Код операции для автоматики: значение с листа «Настройки» или код по умолчанию
Public Function OPC(ByVal role As String, ByVal defaultCode As String) As String
    OPC = Setting("Код операции: " & role, defaultCode)
End Function

Public Function SettingYes(ByVal paramName As String, Optional ByVal defaultYes As Boolean = True) As Boolean
    Dim v As String
    v = LCase$(Setting(paramName, IIf(defaultYes, "Да", "Нет")))
    SettingYes = (v = "да" Or v = "yes" Or v = "1" Or v = "true")
End Function

Public Function WorkFolder() As String
    Dim p As String
    p = ThisWorkbook.Path
    If Len(p) = 0 Then p = "C:\ELKA"
    WorkFolder = AddSlash(p)
End Function

Public Function ArchiveRoot() As String
    ArchiveRoot = AddSlash(Setting("Папка архива документов", WorkFolder() & "Архив"))
End Function

Public Function PublishFolder() As String
    PublishFolder = AddSlash(Setting("Папка публикации", WorkFolder() & "Publish"))
End Function

Public Function BackupFolder() As String
    BackupFolder = AddSlash(Setting("Папка резервных копий", WorkFolder() & "Backups"))
End Function

Public Function AddSlash(ByVal p As String) As String
    p = Trim$(p)
    If Len(p) = 0 Then AddSlash = "": Exit Function
    If Right$(p, 1) <> "\" And Right$(p, 1) <> "/" Then p = p & IIf(Left$(p, 1) = "/", "/", "\")
    AddSlash = p
End Function
