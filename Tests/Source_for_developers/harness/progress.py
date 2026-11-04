import uno
l=uno.getComponentContext();r=l.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver',l)
c=r.resolve('uno:socket,host=localhost,port=2003;urp;StarOffice.ComponentContext');d=c.ServiceManager.createInstanceWithContext('com.sun.star.frame.Desktop',c)
doc=d.getCurrentComponent();s=doc.Sheets.getByName('Движение');cur=s.createCursor();cur.gotoEndOfUsedArea(False);last=cur.RangeAddress.EndRow;print(last,s.getCellByPosition(2,last).String,flush=True)
