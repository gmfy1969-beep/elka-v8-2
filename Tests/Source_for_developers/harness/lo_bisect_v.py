import sys
sys.path.insert(0,"/data/v82/harness")
from lo_run import *
import glob
mods=sys.argv[1:]
def split(txt):
    lines=txt.split("\n"); chunks=[]; cur=[]; head=[]; inproc=False
    for ln in lines:
        s=ln.strip()
        if not inproc and re.match(r'^(Public |Private |Friend )?(Static )?(Sub|Function|Property) ',s,re.I):
            inproc=True; cur=[ln]; continue
        if inproc:
            cur.append(ln)
            if re.match(r'^End (Sub|Function|Property)\b',s,re.I): chunks.append("\n".join(cur)); inproc=False
        else: head.append(ln)
    return "\n".join(head),chunks
p,desk=start()
try:
    doc=desk.loadComponentFromURL(uno.systemPathToFileUrl("/data/v82/viewer_test.xlsx"),"_blank",0,(pv("Hidden",False),pv("MacroExecutionMode",4)))
    lib=doc.BasicLibraries
    if not lib.hasByName("Standard"): lib.createLibrary("Standard")
    L=lib.getByName("Standard")
    L.insertByName("P0","Option VBASupport 1\nFunction Ping() As String\nPing = \"pong\"\nEnd Function\n")
    def ok(): return call(doc,"P0","Ping")=="pong"
    for f in sorted(glob.glob("/data/v82/src/VBA_Viewer/*.bas")):
        txt=open(f,encoding="utf-8-sig").read()
        name=re.search(r'^Attribute VB_Name = "([^"]+)"',txt,re.M).group(1)
        txt=re.sub(r'^Attribute VB_Name = "[^"]+"\n','',txt,count=1,flags=re.M)
        full="Option VBASupport 1\n"+txt
        if name not in mods:
            L.insertByName(name,full)
            if not ok(): print("base module fails",name)
            continue
        head,chunks=split(full)
        L.insertByName(name,head)
        if not ok(): print(name,"HEADER FAILS"); 
        acc=head
        for c in chunks:
            L.replaceByName(name,acc+"\n"+c)
            if ok(): acc=acc+"\n"+c
            else: print(name,"FAIL in:\n",c[:1500],"\n-----")
        L.replaceByName(name,acc)
finally:
    try: desk.terminate()
    except Exception: pass
    time.sleep(1); p.kill()
