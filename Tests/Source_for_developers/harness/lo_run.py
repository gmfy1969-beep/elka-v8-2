import uno, subprocess, time, os, sys, glob, re, json
from com.sun.star.beans import PropertyValue
def pv(n,v):
    x=PropertyValue();x.Name=n;x.Value=v;return x
PORT=int(os.environ.get("ELKA_TEST_PORT", "2003"))
def start():
    p=subprocess.Popen(["/opt/libreoffice26.8/program/soffice","-env:UserInstallation=file:///tmp/elka82_"+str(PORT),"--headless","--invisible","--norestore","--nologo",f"--accept=socket,host=localhost,port={PORT};urp;"],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
    local=uno.getComponentContext()
    res=local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",local)
    for i in range(120):
        try:
            ctx=res.resolve(f"uno:socket,host=localhost,port={PORT};urp;StarOffice.ComponentContext");break
        except Exception: time.sleep(0.5)
    desk=ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop",ctx)
    return p,desk
def load_modules(doc,folder,extra=""):
    libs=doc.BasicLibraries
    if not libs.hasByName("Standard"): libs.createLibrary("Standard")
    lib=libs.getByName("Standard")
    names=[]
    for f in sorted(glob.glob(folder+"/*.bas")):
        txt=open(f,encoding="utf-8-sig").read().replace("\r\n","\n")
        m=re.search(r'^Attribute VB_Name = "([^"]+)"',txt,re.M)
        name=m.group(1)
        txt=re.sub(r'^Attribute VB_Name = "[^"]+"\n','',txt,count=1,flags=re.M)
        txt="Option VBASupport 1\n"+txt
        if lib.hasByName(name): lib.removeByName(name)
        lib.insertByName(name,txt); names.append(name)
    if extra:
        lib.insertByName("LOHarness","Option VBASupport 1\n"+extra); names.append("LOHarness")
    return names
def call(doc,mod,fn,args=()):
    sp=doc.getScriptProvider()
    scr=sp.getScript(f"vnd.sun.star.script:Standard.{mod}.{fn}?language=Basic&location=document")
    return scr.invoke(args,(),())[0]
