import sys, time
sys.path.insert(0,"/data/v82/harness")
from lo_run import *
extra='''
Function T1() As String
    gSilent = True
    T1 = Install_Configure("/tmp/tpl/Архив", "/tmp/tpl/Publish", "/tmp/tpl/Backups", "Тест") & " | " & Install_Check()
    RunControlCore
    T1 = T1 & " | ctrlErr=" & ControlErrors()
    Reports_RefreshAll
    T1 = T1 & " | refresh ok"
End Function
'''
p,desk=start()
try:
    doc=desk.loadComponentFromURL(uno.systemPathToFileUrl(sys.argv[1]),"_blank",0,(pv("Hidden",False),pv("MacroExecutionMode",4)))
    load_modules(doc,"/data/v82/src/VBA_Operator",extra)
    print(call(doc,"LOHarness","T1"))
    doc.close(True)
finally:
    try: desk.terminate()
    except Exception: pass
    time.sleep(1); p.kill()
