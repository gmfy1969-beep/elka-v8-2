import sys
sys.path.insert(0,"/data/v82/harness")
from lo_run import *
code=open(sys.argv[1]).read()
p,desk=start()
try:
    doc=desk.loadComponentFromURL(uno.systemPathToFileUrl(sys.argv[2] if len(sys.argv)>2 else "/data/v82/predemo.xlsx"),"_blank",0,(pv("Hidden",False),pv("MacroExecutionMode",4)))
    load_modules(doc,"/tmp/srcdbg","Function D() As String\nOn Error GoTo E\nDim s As String\n"+code+"\nD = s\nExit Function\nE: D = s & \" | ERR \" & Err.Number & \" \" & Err.Description & \" line \" & Erl\nEnd Function\n")
    print(call(doc,"LOHarness","D"))
finally:
    try: desk.terminate()
    except Exception: pass
    time.sleep(1); p.kill()
