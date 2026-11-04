import sys,os,time
sys.stdout.reconfigure(line_buffering=True)
from lo_run import *
extra='''
Function RepairSmoke(ByVal root As String) As String
 Dim code As String, ord As String, item As String, d As String, rep As String, n As Long
 gSilent = True
 code = WS(SH_NOM).Cells(4, 1).Value
 ord = WS(SH_ORDERS).Cells(4, 1).Value
 item = WS(SH_ORDERS).Cells(4, 3).Value
 rep = Install_Configure(root & "/Архив", root & "/Publish", root & "/Backups", "Тест ремонта")
 rep = rep & vbLf & "migration-repeat=" & Install_RepairDepartment()
 d = PostDoc("01", Date, "", "", "", "Учебный приход", Array(code), Array(100), Array(""), Array(""))
 rep = rep & vbLf & "01=" & d
 d = PostDoc("05", Date, "", "", "", "На ПС", Array(code), Array(100), Array(""), Array(""))
 rep = rep & vbLf & "05=" & d
 d = PostDoc("06", Date, ord, item, "", "В производство", Array(code), Array(100), Array(""), Array(""))
 rep = rep & vbLf & "06=" & d
 d = PostDoc("45", Date, ord, "", "", "Передача в ремонт", Array(code), Array(10), Array(""), Array("РЕМ-01"))
 rep = rep & vbLf & "45=" & d & "; ПР=" & StockAt(code, "ПР") & "; РЕМ=" & StockAt(code, "РЕМ")
 d = PostDoc("46", Date, ord, "", "", "Частичный возврат", Array(code), Array(6), Array("РЕМ-01"), Array("ПР-УЧ1"))
 rep = rep & vbLf & "46=" & d & "; ПР=" & StockAt(code, "ПР") & "; РЕМ=" & StockAt(code, "РЕМ")
 n = LastRow(WS(SH_MOVE), J_OP)
 d = PostDoc("46", Date, ord, "", "", "NEG overdraw", Array(code), Array(5), Array("РЕМ-01"), Array("ПР-УЧ1"))
 rep = rep & vbLf & "negative-overdraw=" & (Len(d) = 0 And n = LastRow(WS(SH_MOVE), J_OP))
 d = PostDoc("45", Date, "", "", "", "NEG no order", Array(code), Array(1), Array(""), Array("РЕМ-01"))
 rep = rep & vbLf & "negative-order=" & (Len(d) = 0 And n = LastRow(WS(SH_MOVE), J_OP))
 d = PostDoc("45", Date, ord, "", "", "NEG wrong address", Array(code), Array(1), Array(""), Array("ЦС-A-01-01"))
 rep = rep & vbLf & "negative-address=" & (Len(d) = 0 And n = LastRow(WS(SH_MOVE), J_OP))
 d = PostDoc("46", Date, ord, "", "", "Окончательный возврат", Array(code), Array(4), Array("РЕМ-01"), Array("ПР-УЧ2"))
 rep = rep & vbLf & "46-full=" & d & "; ПР=" & StockAt(code, "ПР") & "; РЕМ=" & StockAt(code, "РЕМ")
 d = PostDoc("45", Date, ord, "", "", "Остаток в ремонте для просмотра", Array(code), Array(3), Array(""), Array("РЕМ-01"))
 Reports_RefreshAll
 rep = rep & vbLf & "control-errors=" & ControlErrors()
 Publish_Data
 rep = rep & vbLf & "published=" & gLastPublished
 RepairSmoke = rep
End Function
'''
p,desk=start()
try:
 doc=desk.loadComponentFromURL(uno.systemPathToFileUrl('/data/v82/out/ELKA_Operator_Template.xlsx'),'_blank',0,(pv('Hidden',False),pv('MacroExecutionMode',4)))
 load_modules(doc,'/data/v82/src/VBA_Operator',extra)
 print('started',flush=True)
 print(call(doc,'LOHarness','RepairSmoke',('/data/v82/smoke',)))
 doc.storeToURL(uno.systemPathToFileUrl('/data/v82/smoke_demo.xlsx'),(pv('FilterName','Calc MS Excel 2007 XML'),))
 doc.close(True)
finally:
 try:desk.terminate()
 except:pass
 time.sleep(1);p.kill()
