"""Извлечь результаты автотеста из книги LibreOffice (lo_demo.xlsx) и публикации в data.json"""
import openpyxl, json, sys, datetime
src, report, pubfile, out = sys.argv[1:5]
wb = openpyxl.load_workbook(src, data_only=True)
def rows(ws, ncol, r0=4, key=1):
    res = []
    for r in range(r0, ws.max_row + 1):
        v = [ws.cell(r, c).value for c in range(1, ncol + 1)]
        if v[key - 1] in (None, ""): continue
        res.append(v)
    return res
def ser(x):
    if isinstance(x, (datetime.datetime, datetime.date)): return {"__dt__": x.isoformat()}
    return x
D = {}
D["nom"] = rows(wb["Справочник"], 12)
D["moves"] = rows(wb["Движение"], 20, key=3)
D["addrstock"] = rows(wb["Остатки_адреса"], 6)
ws = wb["Свод_вариантов"]; head = []
for c in range(7, 31):
    v = ws.cell(3, c).value
    if v in (None, ""): break
    head.append(v)
D["varsum"] = {"head": head, "rows": rows(ws, 6 + len(head))}
ws = wb["Комплектность"]; head = []
for c in range(6, 30):
    v = ws.cell(3, c).value
    if v in (None, ""): break
    head.append(v)
D["kits"] = {"head": head, "rows": rows(ws, 5 + len(head))}
D["files"] = rows(wb["Файлы"], 9)
D["imp1c"] = rows(wb["Импорт_1С"], 7)
D["cmp1c"] = rows(wb["Сравнение_1С"], 9)
ws = wb["Контроль"]
D["control"] = {"status": ws["B3"].value,
                "summary": [[ws.cell(r, c).value for c in range(1, 7)] for r in range(5, 20) if ws.cell(r, 1).value],
                "details": [[ws.cell(r, c).value for c in range(1, 7)] for r in range(23, ws.max_row + 1) if ws.cell(r, 1).value not in (None, "")]}
D["trash"] = rows(wb["Корзина"], 14)
D["delmoves"] = rows(wb["Движение_удалённые"], 22, key=3)
ws = wb["Настройки"]
for r in range(4, 30):
    if ws.cell(r, 1).value == "Последний номер документа": D["lastdoc"] = ws.cell(r, 2).value
D["spec"] = [[x[0], x[1], x[3], x[6]] for x in rows(wb["Спецификация"], 7)]
D["bundle"] = [[x[0], x[1], x[3], x[6]] for x in rows(wb["Объединения"], 7)]
D["testreport"] = open(report, encoding="utf-8").read()
# публикация
pw = openpyxl.load_workbook(pubfile, data_only=True)
pub = {}
ws = pw["Остатки"]
pub["Остатки_head"] = [ws.cell(3, c).value or "" for c in range(1, 20)]
pub["Остатки"] = rows(ws, 19)
pub["Остатки_адреса"] = rows(pw["Остатки_адреса"], 6)
pub["Движение"] = rows(pw["Движение"], 20, key=3)
ws = pw["Комплектность"]; head = []
for c in range(6, 30):
    v = ws.cell(3, c).value
    if v in (None, ""): break
    head.append(v)
pub["Комплектность_head"] = head
pub["Комплектность"] = rows(ws, 5 + len(head))
pub["Файлы"] = rows(pw["Файлы"], 9)
ab = pw["О_файле"]
D["published_sheets"] = pub
D["published"] = ab["B3"].value
D["version"] = "ELKA_data_v000001.xlsx"
D["pubfolder"] = "C:\\ELKA\\Publish"
D["archive"] = "C:\\ELKA\\Архив\\"
def walk(o):
    if isinstance(o, list): return [walk(x) for x in o]
    if isinstance(o, dict): return {k: walk(v) for k, v in o.items()}
    return ser(o)
json.dump(walk(D), open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=0)
print({k: (len(v) if isinstance(v, list) else "") for k, v in D.items()})
print("pub", {k: len(v) for k, v in pub.items()})
