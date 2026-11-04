import sys, os, shutil, time
sys.path.insert(0,"/data/v82/harness")
from lo_run import *
src,out,pub=sys.argv[1],sys.argv[2],sys.argv[3]
extra='''
Function VInit(ByVal pub As String) As String
    Viewer_InitConfig pub
    VInit = V_PublishFolder() & " latest=" & LatestDataFile()
End Function
Function VRefresh() As String
    vSilent = True
    Viewer_Refresh
    VRefresh = vLastMsg
End Function
Function VList() As String
    Dim sh As Object, s As String
    For Each sh In ThisWorkbook.Worksheets
        s = s & sh.Name & "(" & sh.Visible & ") "
    Next sh
    VList = s & " | panel C5=" & ThisWorkbook.Worksheets("Панель").Range("C5").Value & " C6=" & ThisWorkbook.Worksheets("Панель").Range("C6").Value & " C8=" & ThisWorkbook.Worksheets("Панель").Range("C8").Value
End Function
Function VOpen(ByVal docId As String) As String
    vSilent = True
    vLastMsg = ""
    V_OpenDoc docId
    VOpen = vLastMsg & " | B3=" & ThisWorkbook.Worksheets("Просмотр").Range("B3").Value & " shapes=" & ThisWorkbook.Worksheets("Просмотр").Shapes.Count
End Function
Function VAuto() As String
    vLastMsg = ""
    Viewer_AutoRefresh
    VAuto = "msg=" & vLastMsg
End Function
'''
p,desk=start()
try:
    doc=desk.loadComponentFromURL(uno.systemPathToFileUrl(src),"_blank",0,(pv("Hidden",False),pv("MacroExecutionMode",4)))
    load_modules(doc,"/data/v82/src/VBA_Viewer",extra)
    print("init:",call(doc,"LOHarness","VInit",(pub,)))
    print("before:",call(doc,"LOHarness","VList"))
    t=time.time(); print("refresh:",call(doc,"LOHarness","VRefresh"),"%.1fs"%(time.time()-t))
    print("after:",call(doc,"LOHarness","VList"))
    for d in sys.argv[4:]:
        print("open",d,":",call(doc,"LOHarness","VOpen",(d,)))
    print("auto:",call(doc,"LOHarness","VAuto"))
    doc.storeToURL(uno.systemPathToFileUrl(out),(pv("FilterName","Calc MS Excel 2007 XML"),))
    doc.close(True)
finally:
    try: desk.terminate()
    except Exception: pass
    time.sleep(1); p.kill()
