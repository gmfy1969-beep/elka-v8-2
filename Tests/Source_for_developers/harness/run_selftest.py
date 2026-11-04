import sys, os, shutil, time
sys.stdout.reconfigure(line_buffering=True)
sys.path.insert(0,"/data/v82/harness")
from lo_run import *
src,out,arch=sys.argv[1],sys.argv[2],sys.argv[3]
shutil.rmtree(arch,ignore_errors=True); os.makedirs(arch)
if len(sys.argv)>4: shutil.copytree(sys.argv[4], arch+'/Архив')
import subprocess as _sp
_sp.run(['/usr/local/bin/python3','/data/v82/harness/make_samples.py',arch+'/TestFiles'],env={'PATH':'/usr/local/bin:/usr/bin:/bin','HOME':os.environ.get('HOME','/tmp')})
extra='''
Function Prep(ByVal arch As String, ByVal pub As String) As String
    Prep = Install_Configure(arch, pub, Replace(arch, "Архив", "Backups"), "Иванова А.С.") & " | " & Install_Check() & " | " & Setting("Папка архива документов")
End Function
Function Ping() As String
    Ping = "pong " & ELKA_VERSION
End Function
Function DoPublish() As String
    gSilent = True
    gLastError = ""
    Publish_Data
    DoPublish = gLastPublished & " | " & gLastError
End Function
Function GetLog() As String
    GetLog = gLog
End Function
'''
p,desk=start()
try:
    doc=desk.loadComponentFromURL(uno.systemPathToFileUrl(src),"_blank",0,(pv("Hidden",False),pv("MacroExecutionMode",4)))
    names=load_modules(doc,"/data/v82/src/VBA_Operator",extra)
    print("modules:",names)
    print(call(doc,"LOHarness","Ping"))
    print("prep:",call(doc,"LOHarness","Prep",(arch+"/Архив",arch+"/Publish")))
    t=time.time()
    rep=call(doc,"modTest","RunSelfTest",(arch+"/TestFiles",))
    print("selftest secs %.1f"%(time.time()-t))
    print(rep)
    open(out+".report.txt","w").write(rep)
    open(out+".log.txt","w").write(call(doc,"LOHarness","GetLog"))
    t=time.time()
    try: print("publish:",call(doc,"LOHarness","DoPublish"),"%.1fs"%(time.time()-t))
    except Exception as e: print("publish EXC",e)
    doc.storeToURL(uno.systemPathToFileUrl(out),(pv("FilterName","Calc MS Excel 2007 XML"),))
    doc.close(True)
finally:
    try: desk.terminate()
    except Exception: pass
    time.sleep(1); p.kill()
