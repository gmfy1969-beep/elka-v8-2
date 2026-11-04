import sys, time
sys.path.insert(0,"/data/v82/harness")
from lo_run import *
src,f1c=sys.argv[1],sys.argv[2]
extra='''
Function Do1C(ByVal p As String) As String
    gSilent = True
    If Import1C_Load(p) Then
        On Error Resume Next
        Compare1C
        If Err.Number <> 0 Then Do1C = "ERR " & Err.Number & " " & Err.Description & " | "
        On Error GoTo 0
        Do1C = Do1C & "lr1c=" & LastRow(WS(SH_1C), 1) & " wb=" & WS(SH_1C).Parent.Name & " A5=" & WS(SH_1C).Range("A5").Value & " mv=" & LastRow(WS(SH_MOVE), 2) & " bal=" & BalKeysCount() & " out=" & LastRow(WS(SH_CMP), 1) & " B2=" & WS(SH_CMP).Range("B2").Value & " loaded rows=" & (LastRow(WS(SH_1C), 1) - 3) & " mismatch=" & CountMismatch()
    Else
        Do1C = "load failed " & gLog
    End If
End Function
Function Dbg() As String
    Dim sh As Object
    Set sh = WS(SH_MOVE)
    BuildBalances
    Dbg = "keys=" & BalKeysCount() & " lastOp=" & LastRow(sh, J_OP) & " r4=" & sh.Cells(4, J_OP).Value & "|" & sh.Cells(4, J_OP + 1).Value & "|" & sh.Cells(4, J_QTY).Value & " ops=" & LastRow(WS(SH_OPS), 2) & " op4=" & WS(SH_OPS).Cells(4, 2).Value & " stock=" & StockAt("DET-001", "ЦС")
End Function
Function Mis() As String
    Dim sh As Object, r As Long, s As String
    Set sh = WS(SH_CMP)
    For r = 1 To LastRow(sh, 1)
        If InStr(sh.Cells(r, 8).Value, "СОВПАДАЕТ") = 0 And r > 3 Then s = s & sh.Cells(r, 1).Value & "/" & sh.Cells(r, 3).Value & ":" & sh.Cells(r, 8).Value & "; "
    Next r
    Mis = s
End Function
'''
p,desk=start()
try:
    doc=desk.loadComponentFromURL(uno.systemPathToFileUrl(src),"_blank",0,(pv("Hidden",False),pv("MacroExecutionMode",4)))
    load_modules(doc,"/data/v82/src/VBA_Operator",extra)
    print(call(doc,"LOHarness","Do1C",(f1c,)))
    print(call(doc,"LOHarness","Mis"))
    doc.storeToURL(uno.systemPathToFileUrl("/tmp/c1.xlsx"),(pv("FilterName","Calc MS Excel 2007 XML"),))
    doc.close(True)
finally:
    try: desk.terminate()
    except Exception: pass
    time.sleep(1); p.kill()
