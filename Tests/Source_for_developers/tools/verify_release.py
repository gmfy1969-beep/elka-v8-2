"""Independent ledger checks for the final demo and its published snapshot."""
from pathlib import Path
import openpyxl,json,collections
V=Path('/data/v82')
w=openpyxl.load_workbook(V/'lo_demo.xlsx',data_only=True)
ops={str(r[1]):(str(r[2]),str(r[3])) for r in w['Операции'].iter_rows(min_row=4,values_only=True) if r[0]}
tracked=[r[0] for r in w['Места хранения'].iter_rows(min_row=4,values_only=True) if r[0] and r[3]=='Да']
nom={r[0]:r for r in w['Справочник'].iter_rows(min_row=4,values_only=True) if r[0]}
bal=collections.defaultdict(float);addr=collections.defaultdict(float);counts=collections.Counter();errors=[];repair=[]
for row in w['Движение'].iter_rows(min_row=4,values_only=True):
 if not row[2]: continue
 op=str(row[2]); code=row[3]; q=float(row[5]); f,t=ops[op]; counts[op.split()[0]]+=1
 if q<=0:errors.append('nonpositive quantity')
 if f!=row[7] or t!=row[8]:errors.append(f'route cache {op}')
 bal[code,f]-=q;bal[code,t]+=q
 addr[code,f,row[9] or '']-=q;addr[code,t,row[10] or '']+=q
 if op.startswith(('45 ','46 ')):
  if {f,t}!={'ПР','РЕМ'}:errors.append('repair route')
  repair.append((op,code,q))
for (code,loc),v in bal.items():
 if loc in tracked and v<-.0001:errors.append(f'negative balance {code}/{loc} {v}')
for (code,loc,a),v in addr.items():
 if loc in tracked and v<-.0001:errors.append(f'negative address {code}/{loc}/{a} {v}')
s=w['Остатки']; matched=0
for row in s.iter_rows(min_row=4,values_only=True):
 if not row[0]:continue
 for c in range(4,16):
  loc=s.cell(3,c+1).value
  if not loc:continue
  expected=bal[row[0],loc]; actual=float(row[c] or 0);matched+=1
  if abs(expected-actual)>.0001:errors.append(f'formula {row[0]}/{loc}: {actual}!={expected}')
 if abs(float(row[16] or 0)-sum(bal[row[0],loc] for loc in tracked))>.0001:errors.append('total formula')
for sheet in w:
 for row in sheet:
  for cell in row:
   if cell.data_type=='e':errors.append(f'{sheet.title}!{cell.coordinate} {cell.value}')
report={'version':'8.2','nomenclature':len(nom),'tracked_locations':tracked,'movement_rows':sum(counts.values()),'operation_row_counts':dict(counts),'repair_rows':len(repair),'formula_cells_compared':matched,'repair_ending_quantity':sum(v for (c,l),v in bal.items() if l=='РЕМ'),'errors':errors,'environment':'LibreOffice VBA compatibility + independent Python ledger; not Windows Excel'}
assert set(['45','46'])<=set(counts)
assert len(nom)==255
(V/'tests_out/independent_ledger.json').write_text(json.dumps(report,ensure_ascii=False,indent=2))
print(json.dumps({k:v for k,v in report.items() if k!='operation_row_counts'},ensure_ascii=False,indent=2))
assert not errors
