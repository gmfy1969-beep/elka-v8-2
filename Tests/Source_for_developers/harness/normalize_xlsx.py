import sys,time
from lo_run import *
p,desk=start()
try:
 for src,out in zip(sys.argv[1::2],sys.argv[2::2]):
  doc=desk.loadComponentFromURL(uno.systemPathToFileUrl(src),'_blank',0,(pv('Hidden',True),pv('MacroExecutionMode',4)))
  doc.storeToURL(uno.systemPathToFileUrl(out),(pv('FilterName','Calc MS Excel 2007 XML'),))
  doc.close(True);print(out,flush=True)
finally:
 try:desk.terminate()
 except:pass
 time.sleep(1);p.kill()
