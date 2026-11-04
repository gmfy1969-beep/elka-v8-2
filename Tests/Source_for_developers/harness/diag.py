import sys
sys.path.insert(0,"/data/v82/harness")
from lo_run import *
import glob
p,desk=start()
try:
    doc=desk.loadComponentFromURL(uno.systemPathToFileUrl("/data/v82/predemo.xlsx"),"_blank",0,(pv("Hidden",True),pv("MacroExecutionMode",4)))
    lib=doc.BasicLibraries
    if not lib.hasByName("Standard"): lib.createLibrary("Standard")
    L=lib.getByName("Standard"); print("existing",L.getElementNames())
    L.insertByName("P0","Option VBASupport 1\nFunction Ping() As String\nPing = \"pong\" & Worksheets.Count\nEnd Function\n")
    print("ping:",repr(call(doc,"P0","Ping")))
    files=sorted(glob.glob("/data/v82/src/VBA_Operator/*.bas"))
    for f in files:
        txt=open(f,encoding="utf-8-sig").read()
        m=re.search(r'^Attribute VB_Name = "([^"]+)"',txt,re.M); name=m.group(1)
        txt=re.sub(r'^Attribute VB_Name = "[^"]+"\n','',txt,count=1,flags=re.M)
        L.insertByName(name,"Option VBASupport 1\n"+txt)
        try: print(name,"ping:",repr(call(doc,"P0","Ping")))
        except Exception as e: print(name,"ERR",str(e)[:500])
finally:
    try: desk.terminate()
    except Exception: pass
    time.sleep(1); p.kill()
