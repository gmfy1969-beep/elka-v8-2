import sys
sys.path.insert(0,"/data/v82/harness")
from lo_run import *
p,desk=start()
try:
    doc=desk.loadComponentFromURL(uno.systemPathToFileUrl("/data/v82/predemo.xlsx"),"_blank",0,(pv("Hidden",True),pv("MacroExecutionMode",4)))
    L=doc.BasicLibraries
    if not L.hasByName("Standard"): L.createLibrary("Standard")
    L=L.getByName("Standard")
    L.insertByName("P0",'''Option VBASupport 1
Function T(n As Integer) As String
On Error GoTo E
Select Case n
Case 1: T = ActiveWorkbook.Name
Case 2: T = ThisWorkbook.Name
Case 3: T = ThisWorkbook.Worksheets("Настройки").Cells(4,1).Value
Case 4: T = Application.ActiveWorkbook.Worksheets.Count
End Select
Exit Function
E: T = "ERR " & Err.Number & " " & Err.Description
End Function
''')
    for i in range(1,5): print(i,repr(call(doc,"P0","T",(i,))))
finally:
    try: desk.terminate()
    except Exception: pass
    time.sleep(1); p.kill()
