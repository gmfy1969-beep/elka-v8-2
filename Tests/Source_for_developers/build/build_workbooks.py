# -*- coding: utf-8 -*-
"""ELKA v8.2 — сборщик книг (оператор: шаблон/демо, пользователь: просмотрщик).
python3 build_workbooks.py <mode> <out.xlsx> [data.json]
mode: template | predemo | demo | viewer | viewer_demo
"""
import sys, json, datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.properties import CalcProperties
import refdata as R

BLUE="1F4E78"; HBLUE="2F5597"; LIGHT="DDEBF7"; INPUT="FFF8DC"; CALC="E2F0D9"; BAD="F4CCCC"; WARN="FFF2CC"; GRAY="F2F2F2"
thin=Side(style="thin",color="BFBFBF"); BORDER=Border(left=thin,right=thin,top=thin,bottom=thin)
F_TITLE=Font(name="Arial",size=16,bold=True,color="FFFFFF")
F_SUB=Font(name="Arial",size=11,italic=True,color="1F1F1F")
F_HDR=Font(name="Arial",size=11,bold=True,color="FFFFFF")
F_DATA=Font(name="Arial",size=10)
F_BOLD=Font(name="Arial",size=10,bold=True)
FILL=lambda c: PatternFill("solid",fgColor=c)
VERSION="8.2"
FIRST=4

def sheet(wb,name,headers,widths,subtitle,input_cols=(),calc_cols=(),data_rows=40,tab=None):
    ws=wb.create_sheet(name)
    n=len(headers)
    ws.sheet_view.showGridLines=False
    ws.sheet_view.zoomScale=90
    ws.freeze_panes="A4"
    title(ws,f"ELKA v{VERSION} — {name}",subtitle,n)
    for i,h in enumerate(headers,1):
        c=ws.cell(3,i,h); c.font=F_HDR; c.fill=FILL(HBLUE); c.border=BORDER
        c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    ws.row_dimensions[3].height=48
    for i,w in enumerate(widths,1): ws.column_dimensions[L(i)].width=w
    if n: ws.auto_filter.ref=f"A3:{L(n)}{3+max(data_rows,1)}"
    for c in input_cols: ws.cell(3,c).fill=FILL("7F6000")
    for c in calc_cols: ws.cell(3,c).fill=FILL("385723")
    if tab: ws.sheet_properties.tabColor=tab
    return ws

def title(ws,t,sub,n):
    n=max(n,4)
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=n)
    ws.merge_cells(start_row=2,start_column=1,end_row=2,end_column=n)
    c=ws.cell(1,1,t); c.font=F_TITLE; c.alignment=Alignment(vertical="center",indent=1)
    c2=ws.cell(2,1,sub); c2.font=F_SUB; c2.alignment=Alignment(vertical="center",wrap_text=True,indent=1)
    for i in range(1,n+1):
        ws.cell(1,i).fill=FILL(BLUE); ws.cell(2,i).fill=FILL(LIGHT)
    ws.row_dimensions[1].height=32; ws.row_dimensions[2].height=38

def style_rows(ws,r1,r2,ncols,input_cols=(),calc_cols=(),numfmt=None,height=21):
    numfmt=numfmt or {}
    for r in range(r1,r2+1):
        ws.row_dimensions[r].height=height
        for c in range(1,ncols+1):
            cell=ws.cell(r,c); cell.font=F_DATA; cell.border=BORDER
            cell.alignment=Alignment(vertical="center",wrap_text=False)
            if c in input_cols: cell.fill=FILL(INPUT)
            elif c in calc_cols: cell.fill=FILL(CALC)
            if c in numfmt: cell.number_format=numfmt[c]

def dv_list(ws,rng,formula=None,values=None,prompt=None,err=True):
    if values: dv=DataValidation(type="list",formula1='"'+",".join(values)+'"',allow_blank=True)
    else: dv=DataValidation(type="list",formula1=formula,allow_blank=True)
    dv.showErrorMessage=err
    dv.errorStyle="stop" if err else "warning"
    dv.error="Выберите значение из списка"; dv.errorTitle="ELKA"
    if prompt: dv.prompt=prompt; dv.promptTitle="Подсказка"; dv.showInputMessage=True
    ws.add_data_validation(dv); dv.add(rng)

# ---------------- формулы (совпадают с VBA: modPost.JFormula, modReports.StockFormula) ----------------
def jf(col,r):
    f={5:'=IF($D{r}="","",IFERROR(INDEX(Справочник!$B:$B,MATCH($D{r},Справочник!$A:$A,0)),"НЕТ В СПРАВОЧНИКЕ"))',
       7:'=IF($D{r}="","",IFERROR(INDEX(Справочник!$D:$D,MATCH($D{r},Справочник!$A:$A,0)),""))',
       8:'=IF($C{r}="","",IFERROR(INDEX(Операции!$C:$C,MATCH($C{r},Операции!$B:$B,0)),"?"))',
       9:'=IF($C{r}="","",IFERROR(INDEX(Операции!$D:$D,MATCH($C{r},Операции!$B:$B,0)),"?"))',
       19:('=IF($C{r}="","",IF(COUNTIF(Операции!$B:$B,$C{r})=0,"НЕТ ОПЕРАЦИИ",IF($D{r}="","НЕТ КОДА",'
           'IF(COUNTIF(Справочник!$A:$A,$D{r})=0,"НЕТ В СПРАВОЧНИКЕ",IF(N($F{r})<=0,"КОЛИЧЕСТВО <= 0",'
           'IF($B{r}="","НЕТ ДАТЫ",IF(AND($J{r}<>"",COUNTIFS(Адреса!$A:$A,$J{r},Адреса!$B:$B,$H{r})=0),"АДРЕС ОТКУДА НЕ ИЗ "&$H{r},'
           'IF(AND($K{r}<>"",COUNTIFS(Адреса!$A:$A,$K{r},Адреса!$B:$B,$I{r})=0),"АДРЕС КУДА НЕ ИЗ "&$I{r},"OK"))))))))')}[col]
    return f.replace("{r}",str(r))

NLOC=12
def sf(col,r):
    if col==1: f='=IF(INDEX(Справочник!$A:$A,ROW())="","",INDEX(Справочник!$A:$A,ROW()))'
    elif col==2: f='=IF($A{r}="","",INDEX(Справочник!$B:$B,ROW()))'
    elif col==3: f='=IF($A{r}="","",INDEX(Справочник!$C:$C,ROW()))'
    elif col==4: f='=IF($A{r}="","",INDEX(Справочник!$D:$D,ROW()))'
    elif 5<=col<=4+NLOC:
        cl=L(col)
        f=(f'=IF(OR($A{{r}}="",{cl}$3=""),"",SUMIFS(Движение!$F:$F,Движение!$D:$D,$A{{r}},Движение!$I:$I,{cl}$3)'
           f'-SUMIFS(Движение!$F:$F,Движение!$D:$D,$A{{r}},Движение!$H:$H,{cl}$3))')
    elif col==5+NLOC: f='=IF($A{r}="","",SUM(E{r}:'+L(4+NLOC)+'{r}))'
    elif col==6+NLOC: f='=IF($A{r}="","",N(INDEX(Справочник!$H:$H,ROW())))'
    else: f='=IF($A{r}="","",IF($C{r}="Вариантный родитель","—",IF(Q{r}<0,"МИНУС",IF(AND(R{r}>0,Q{r}<R{r}),"НИЖЕ МИНИМУМА","OK"))))'
    return f.replace("{r}",str(r))

MOVE_HEAD=["№","Дата","Операция","Код","Наименование","Количество","Ед.","Откуда","Куда","Адрес откуда","Адрес куда","№ заказа","Изделие / ПФ","Партия / поставка","Документ №","Файлы (двойной щелчок — открыть)","Оператор","Комментарий","Проверка","Создано"]
MOVE_W=[7,12,34,13,32,12,7,10,13,14,14,15,13,16,12,40,14,30,24,16]
MOVE_IN=(1,2,3,4,6,10,11,12,13,14,18); MOVE_CALC=(5,7,8,9,19)
FILES_HEAD=["ID файла","Документ №","Название (произвольное)","Тип","Путь (относительно папки архива)","Исходное имя файла","Дата","Оператор","Примечание"]
ADDRSTOCK_HEAD=["Код","Наименование","Место","Адрес","Остаток","Ед."]
VARSUM_HEAD=["Базовый код","Семейство","Код варианта","Наименование","Вариант","Итого (учёт.)"]
KITS_HEAD=["Изделие / объединение","Компонент","Наименование","Норма на 1","Ед."]
CMP_HEAD=["Артикул","Наименование","Склад 1С","Место в книге","Остаток 1С","Остаток книги","Разница (книга − 1С)","Статус","Что проверить"]
IMP_HEAD=["Артикул","Наименование","Склад 1С","Остаток 1С","Ед.","Файл-источник","Дата импорта"]

def stock_sheet(wb,nrows,values=None):
    heads=["Код","Наименование","Тип","Ед."]+[""]*NLOC+["Итого (учитываемые места)","Мин. остаток","Статус"]
    ws=sheet(wb,"Остатки",heads,[13,32,19,7]+[11]*NLOC+[14,11,16],
        "Живые формулы: строка = строка «Справочника», столбцы = места с «В остатках = Да». Ничего не вводите вручную. Новые позиции появляются автоматически.",
        calc_cols=(),data_rows=nrows,tab="70AD47")
    for k in range(1,NLOC+1):
        ws.cell(3,4+k).value=f"=IFERROR(INDEX('Места хранения'!$A:$A,MATCH({k},'Места хранения'!$H:$H,0)),\"\")"
    for r in range(FIRST,FIRST+nrows):
        for c in range(1,8+NLOC):
            ws.cell(r,c).value=sf(c,r)
    style_rows(ws,FIRST,FIRST+nrows-1,7+NLOC,numfmt={**{c:'# ##0.###;[Red]-# ##0.###' for c in range(5,6+NLOC)},6+NLOC:'# ##0.###'})
    for r in range(FIRST,FIRST+nrows): ws.cell(r,5+NLOC).font=F_BOLD
    rng=f"E{FIRST}:{L(5+NLOC)}{FIRST+nrows-1}"
    ws.conditional_formatting.add(rng,CellIsRule(operator="lessThan",formula=["0"],fill=FILL(BAD)))
    ws.conditional_formatting.add(f"{L(7+NLOC)}{FIRST}:{L(7+NLOC)}{FIRST+nrows-1}",CellIsRule(operator="equal",formula=['"НИЖЕ МИНИМУМА"'],fill=FILL(WARN)))
    ws.conditional_formatting.add(f"{L(7+NLOC)}{FIRST}:{L(7+NLOC)}{FIRST+nrows-1}",CellIsRule(operator="equal",formula=['"МИНУС"'],fill=FILL(BAD)))
    return ws

def put_rows(ws,rows,r0=FIRST):
    for i,row in enumerate(rows):
        for j,v in enumerate(row):
            if v is not None and v!="": ws.cell(r0+i,j+1,v)

def build(mode,out,data=None):
    data=data or {}
    operator = mode in ("template","predemo","demo")
    if mode=="pubfile": operator=False
    demo_ref = mode in ("predemo","demo","viewer_demo")
    wb=Workbook(); wb.remove(wb.active)
    wb.calculation=CalcProperties(fullCalcOnLoad=True)
    nom = data.get("nom") or R.NOM
    moves = data.get("moves", [])
    nm=max(len(nom)+60,300)

    if operator:
        # ---------------- ПАНЕЛЬ ----------------
        ws=wb.create_sheet("Панель"); ws.sheet_view.showGridLines=False; ws.sheet_properties.tabColor="1F4E78"
        title(ws,f"ELKA v{VERSION} — УЧЁТ ДВИЖЕНИЯ, КОМПЛЕКТАЦИИ И АДРЕСНОГО ХРАНЕНИЯ",
              "Книга ОПЕРАТОРА. Щёлкните по названию листа, чтобы перейти. Кнопки справа выполняют основные действия. Подробно — лист «Инструкция_оператора».",6)
        ws.column_dimensions["A"].width=4; ws.column_dimensions["B"].width=30; ws.column_dimensions["C"].width=70; ws.column_dimensions["D"].width=3; ws.column_dimensions["E"].width=34; ws.column_dimensions["F"].width=4
        nav=[("Ввод","Ввод операций: приход, перемещение, выдача, брак, отгрузка, сборка"),("Движение","Журнал всех операций (каждая строка документа)"),
             ("Остатки","Остатки по местам хранения — считаются автоматически"),("Остатки_адреса","Остатки по ячейкам (адресное хранение)"),
             ("Справочник","Номенклатура: детали, материалы, ПФ, изделия, варианты, объединения"),("Спецификация","Комплектация изделий и ПФ (нормы на 1 шт)"),
             ("Объединения","Состав объединений (комплектов)"),("Свод_вариантов","Остатки по вариантам каждого семейства"),
             ("Комплектность","Хватает ли компонентов и сколько изделий можно собрать"),("Места хранения","Склады, цех, транзитные места"),
             ("Адреса","Ячейки хранения: место, зона, стеллаж, ярус, ячейка"),("Операции","Справочник операций: откуда -> куда"),
             ("Файлы","Все прикреплённые документы и фото"),("Импорт_1С","Загруженная выгрузка остатков из 1С"),("Сравнение_1С","Сверка книги с 1С"),
             ("Контроль","Проверка книги и подсказки по исправлению"),("Настройки","Папки, запреты, допуски"),("Инструкция_оператора","Пошаговая инструкция")]
        hdr=ws.cell(4,2,"Лист"); hdr2=ws.cell(4,3,"Для чего")
        for c in (hdr,hdr2): c.font=F_HDR; c.fill=FILL(HBLUE); c.border=BORDER; c.alignment=Alignment(vertical="center",horizontal="center")
        ws.cell(4,5,"Действия (кнопки)").font=F_HDR; ws.cell(4,5).fill=FILL(HBLUE); ws.cell(4,5).alignment=Alignment(horizontal="center")
        ws.row_dimensions[4].height=30
        for i,(s,d) in enumerate(nav):
            r=5+i
            c=ws.cell(r,2,f'=HYPERLINK("#\'{s}\'!A1","{s}")'); c.font=Font(name="Arial",size=11,bold=True,color="0563C1",underline="single"); c.border=BORDER
            c2=ws.cell(r,3,d); c2.font=F_DATA; c2.border=BORDER
            ws.row_dimensions[r].height=24
        ws.cell(5+len(nav)+1,2,"Режим").font=F_BOLD
        ws.cell(5+len(nav)+1,3,"ДЕМО: справочники и результаты полного автотеста" if mode!="template" else "РАБОЧИЙ ШАБЛОН: справочник номенклатуры и журнал пусты").font=F_DATA
        ws.cell(5+len(nav)+2,2,"Версия").font=F_BOLD; ws.cell(5+len(nav)+2,3,VERSION).font=F_DATA

        # ---------------- ВВОД ----------------
        ws=wb.create_sheet("Ввод"); ws.sheet_view.showGridLines=False; ws.sheet_view.zoomScale=90; ws.sheet_properties.tabColor="C00000"
        title(ws,f"ELKA v{VERSION} — ВВОД ОПЕРАЦИИ (ДОКУМЕНТА)",
              "1) Дата и операция  2) Заказ/изделие, если нужно  3) Строки: код + количество (адреса — по желанию)  4) Файлы  5) «ПРОВЕСТИ». Жёлтые ячейки заполняете вы, зелёные — считаются.",13)
        widths=[6,26,40,9,13,16,16,14,30,2,32,2,32]
        for i,w in enumerate(widths,1): ws.column_dimensions[L(i)].width=w
        labels=[("Дата","",'Дата документа (по умолчанию сегодня)'),("Операция","","Выберите из списка или введите код: 6 или 06"),
                ("Откуда (авто)",'=IFERROR(INDEX(Операции!$C:$C,MATCH($C$5,Операции!$B:$B,0)),"")',"Определяется операцией"),
                ("Куда (авто)",'=IFERROR(INDEX(Операции!$D:$D,MATCH($C$5,Операции!$B:$B,0)),"")',"Определяется операцией"),
                ("№ заказа","","Обязателен по справочнику операций; для ремонта 45/46 — обязательно"),("Изделие / ПФ (код)","","Для сборки, выпуска, комплектования и разбора"),
                ("Партия / поставка","","Например: номер накладной или партии"),("Комментарий","","Любой текст — попадёт в каждую строку журнала"),
                ("Файлы документа","(файлы не прикреплены)","Кнопка «Прикрепить файл(ы)» — любое количество, любые названия"),
                ("Кол-во изделий","","Для кнопок «Заполнить по спецификации», «Собрать», «Скомплектовать», «Разобрать»")]
        for i,(lab,val,hint) in enumerate(labels):
            r=4+i
            a=ws.cell(r,2,lab); a.font=F_BOLD; a.alignment=Alignment(vertical="center",horizontal="right"); a.border=BORDER; a.fill=FILL(GRAY)
            b=ws.cell(r,3,val if val else None); b.font=Font(name="Arial",size=11,bold=(i in (1,))); b.border=BORDER
            b.fill=FILL(CALC if i in (2,3,8) else INPUT); b.alignment=Alignment(vertical="center")
            ws.merge_cells(start_row=r,start_column=4,end_row=r,end_column=9)
            h=ws.cell(r,4,hint); h.font=Font(name="Arial",size=9,italic=True,color="595959"); h.alignment=Alignment(vertical="center",indent=1)
            ws.row_dimensions[r].height=24
        ws["C4"]=datetime.date.today(); ws["C4"].number_format="dd.mm.yyyy"; ws["C13"].number_format="0.###"
        ws.row_dimensions[14].height=10
        lh=["№","Код","Наименование","Ед.","Количество","Адрес откуда","Адрес куда","Есть в «Откуда»","Проверка строки"]
        for i,h in enumerate(lh,1):
            c=ws.cell(15,i,h); c.font=F_HDR; c.fill=FILL(HBLUE); c.border=BORDER; c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
        ws.row_dimensions[15].height=40
        NIN=200
        for r in range(16,16+NIN):
            ws.cell(r,1,f"=ROW()-15")
            ws.cell(r,3,f'=IF($B{r}="","",IFERROR(INDEX(Справочник!$B:$B,MATCH($B{r},Справочник!$A:$A,0)),"НЕТ В СПРАВОЧНИКЕ"))')
            ws.cell(r,4,f'=IF($B{r}="","",IFERROR(INDEX(Справочник!$D:$D,MATCH($B{r},Справочник!$A:$A,0)),""))')
            ws.cell(r,8,f'=IF(OR($B{r}="",$C$6=""),"",SUMIFS(Движение!$F:$F,Движение!$D:$D,$B{r},Движение!$I:$I,$C$6)-SUMIFS(Движение!$F:$F,Движение!$D:$D,$B{r},Движение!$H:$H,$C$6))')
            ws.cell(r,9,(f'=IF($B{r}="","",IF(COUNTIF(Справочник!$A:$A,$B{r})=0,"НЕТ В СПРАВОЧНИКЕ",IF(N($E{r})<=0,"УКАЖИТЕ КОЛИЧЕСТВО",'
                         f'IF(AND($F{r}<>"",COUNTIFS(Адреса!$A:$A,$F{r},Адреса!$B:$B,$C$6)=0),"АДРЕС ОТКУДА НЕ ИЗ "&$C$6,'
                         f'IF(AND($G{r}<>"",COUNTIFS(Адреса!$A:$A,$G{r},Адреса!$B:$B,$C$7)=0),"АДРЕС КУДА НЕ ИЗ "&$C$7,'
                         f'IF(AND(COUNTIFS(\'Места хранения\'!$A:$A,$C$6,\'Места хранения\'!$D:$D,"Да")>0,N($E{r})>N($H{r})),"НЕ ХВАТАЕТ: есть "&N($H{r}),"OK"))))))'))
        style_rows(ws,16,15+NIN,9,input_cols=(2,5,6,7),calc_cols=(1,3,4,8,9),numfmt={5:'# ##0.###',8:'# ##0.###'})
        ws.conditional_formatting.add(f"I16:I{15+NIN}",FormulaRule(formula=['AND($I16<>"",$I16<>"OK")'],fill=FILL(BAD)))
        dv_list(ws,"C5",formula="=Операции!$B$4:$B$300",err=False,prompt="Выберите операцию или введите её код (например 6)")
        dv_list(ws,"C9",formula="=Справочник!$A$4:$A$5000",err=False)
        dv_list(ws,f"B16:B2000",formula="=Справочник!$A$4:$A$5000",prompt="Код из справочника")
        dv_list(ws,f"F16:G2000",formula="=Адреса!$A$4:$A$3000",prompt="Адрес из листа «Адреса» (необязательно)")
        ws.freeze_panes="A16"
        for r,txt in ((3,"ПРОВЕДЕНИЕ И ФАЙЛЫ"),):
            ws.cell(r,11,txt).font=Font(name="Arial",size=10,bold=True,color="1F4E78"); ws.cell(r,13,"СБОРКА / ОБЪЕДИНЕНИЯ").font=Font(name="Arial",size=10,bold=True,color="1F4E78")

        # ---------------- ДВИЖЕНИЕ ----------------
        ws=sheet(wb,"Движение",MOVE_HEAD,MOVE_W,
            "Журнал. Обычно заполняется кнопкой «ПРОВЕСТИ» с листа «Ввод». Можно вводить и вручную: операция, код, количество — формулы и № появятся сами. Удаление — правой кнопкой «ELKA: удалить выделенные строки».",
            input_cols=MOVE_IN,calc_cols=MOVE_CALC,data_rows=max(len(moves),1),tab="C00000")
        for i,m in enumerate(moves):
            r=FIRST+i
            for c in range(1,21):
                if c in MOVE_CALC: ws.cell(r,c,jf(c,r))
                else:
                    v=m[c-1]
                    if v not in (None,""): ws.cell(r,c,v)
        nmv=max(len(moves),1)
        style_rows(ws,FIRST,FIRST+nmv-1,20,input_cols=MOVE_IN,calc_cols=MOVE_CALC,numfmt={2:"dd.mm.yyyy",6:'# ##0.###',20:"dd.mm.yyyy hh:mm"})
        for c in (2,): 
            for r in range(FIRST,FIRST+nmv): ws.cell(r,c).number_format="dd.mm.yyyy"
        ws.conditional_formatting.add("S4:S200000",FormulaRule(formula=['AND($S4<>"",$S4<>"OK")'],fill=FILL(BAD)))
        dv_list(ws,"C4:C200000",formula="=Операции!$B$4:$B$300",err=False,prompt="Операция из списка или код (6)")
        dv_list(ws,"D4:D200000",formula="=Справочник!$A$4:$A$5000")
        dv_list(ws,"J4:K200000",formula="=Адреса!$A$4:$A$3000")
        ws["P3"].comment=Comment("Названия файлов произвольные. Двойной щелчок по ячейке открывает файл (фото — на листе «Просмотр»).","ELKA")

        stock_sheet(wb,nm)
        ws=sheet(wb,"Остатки_адреса",ADDRSTOCK_HEAD,[13,34,10,16,13,8],"Отчёт по ячейкам. Обновляется кнопкой «Обновить отчёты» и перед публикацией. «(без адреса)» — остаток, принятый без указания ячейки.",data_rows=max(len(data.get("addrstock",[])),1),tab="70AD47")
        put_rows(ws,data.get("addrstock",[])); style_rows(ws,FIRST,FIRST+max(len(data.get("addrstock",[])),1)-1,6,numfmt={5:'# ##0.###;[Red]-# ##0.###'})

        # ---------------- СПРАВОЧНИК ----------------
        ws=sheet(wb,"Справочник",R.NOM_HEAD,[13,36,20,8,15,16,14,11,15,10,26,34],
            "Каждая позиция — одна строка. Код уникален. Вариант: заведите «Вариантный родитель», затем позиции с его кодом в «Базовый код». Удаление — кнопкой «Удалить в корзину».",
            input_cols=tuple(range(1,13)),data_rows=nm,tab="FFC000")
        put_rows(ws,nom); style_rows(ws,FIRST,FIRST+nm-1,12,input_cols=tuple(range(1,13)),numfmt={8:'# ##0.###'})
        dv_list(ws,f"C4:C10000",values=R.TYPES,prompt="Тип позиции")
        dv_list(ws,f"J4:J10000",values=["Да","Нет"])
        dv_list(ws,f"I4:I10000",formula="=Адреса!$A$4:$A$3000",prompt="Ячейка, куда по умолчанию принимается позиция")
        dv_list(ws,f"F4:F10000",formula="=Справочник!$A$4:$A$5000",err=False,prompt="Код вариантного родителя (только для вариантов)")

        for nmz,head,rows,sub in (("Спецификация",R.SPEC_HEAD,data.get("spec") or R.SPEC,"Комплектация: одна строка = один компонент изделия. Норма — на 1 изделие. Количество строк не ограничено. Наименования подставляются автоматически."),
                                  ("Объединения",R.BUNDLE_HEAD,data.get("bundle") or R.BUNDLE,"Объединение (комплект): одна строка = один компонент. Кнопки на «Вводе»: «Скомплектовать» и «Разобрать»." )):
            n=max(len(rows)+100,200)
            ws=sheet(wb,nmz,head,[16,16,36,13,8,32,30],sub,input_cols=(1,2,4,7),calc_cols=(3,5,6),data_rows=n,tab="FFC000")
            for i,row in enumerate(rows):
                r=FIRST+i; ws.cell(r,1,row[0]); ws.cell(r,2,row[1]); ws.cell(r,4,row[2]); 
                if row[3]: ws.cell(r,7,row[3])
            for r in range(FIRST,FIRST+n):
                ws.cell(r,3,f'=IF($B{r}="","",IFERROR(INDEX(Справочник!$B:$B,MATCH($B{r},Справочник!$A:$A,0)),"НЕТ В СПРАВОЧНИКЕ"))')
                ws.cell(r,5,f'=IF($B{r}="","",IFERROR(INDEX(Справочник!$D:$D,MATCH($B{r},Справочник!$A:$A,0)),""))')
                ws.cell(r,6,f'=IF($A{r}="","",IFERROR(INDEX(Справочник!$B:$B,MATCH($A{r},Справочник!$A:$A,0)),"НЕТ В СПРАВОЧНИКЕ"))')
            style_rows(ws,FIRST,FIRST+n-1,7,input_cols=(1,2,4,7),calc_cols=(3,5,6),numfmt={4:'0.######'})
            dv_list(ws,f"A4:B5000",formula="=Справочник!$A$4:$A$5000")

        ws=sheet(wb,"Некомплектные_ПФ",R.PFD_HEAD,[22,22,36,22,36,14,8,45],"Код ПФД, целевое изделие и недостающие компоненты на 1 ПФД. Для сборки используется спецификация целевого изделия за вычетом дефицита.",input_cols=(1,2,4,6,8),calc_cols=(3,5,7),data_rows=200,tab="FFC000")
        for i,row in enumerate(R.PFD):
            r=FIRST+i
            for c,v in [(1,row[0]),(2,row[1]),(4,row[2]),(6,row[3]),(8,row[4])]: ws.cell(r,c,v)
        for r in range(FIRST,FIRST+200):
            for c,key,ret in [(3,"B","B"),(5,"D","B"),(7,"D","D")]:
                ws.cell(r,c,f'=IF(${key}{r}="","",IFERROR(INDEX(Справочник!${ret}:${ret},MATCH(${key}{r},Справочник!$A:$A,0)),"НЕТ В СПРАВОЧНИКЕ"))')
        style_rows(ws,FIRST,FIRST+199,8,input_cols=(1,2,4,6,8),calc_cols=(3,5,7),numfmt={6:'0.######'})
        ws=sheet(wb,"Заказы",R.ORD_HEAD,[20,14,24,36,14,14,16,18,18,18,45],"Заказы: номер, дата, код изделия и план. Показатели выпуска и отгрузки рассчитываются по журналу. План можно менять.",input_cols=(1,2,3,5,11),calc_cols=(4,6,7,8,9,10),data_rows=200,tab="FFC000")
        for i,row in enumerate(R.ORDERS):
            r=FIRST+i
            for c,v in [(1,row[0]),(2,datetime.date.fromisoformat(row[1])),(3,row[2]),(5,row[3]),(11,row[4])]: ws.cell(r,c,v)
        for r in range(FIRST,FIRST+200):
            ws.cell(r,4,f'=IF($C{r}="","",IFERROR(INDEX(Справочник!$B:$B,MATCH($C{r},Справочник!$A:$A,0)),"НЕТ В СПРАВОЧНИКЕ"))')
            ws.cell(r,6,f'=IF($A{r}="","",SUMIFS(Движение!$F:$F,Движение!$L:$L,$A{r},Движение!$D:$D,$C{r},Движение!$I:$I,"ГП",Движение!$H:$H,"СБОРКА"))')
            ws.cell(r,7,f'=IF($A{r}="","",SUMIFS(Движение!$F:$F,Движение!$L:$L,$A{r},Движение!$M:$M,$C{r},Движение!$I:$I,"ПФД")-SUMIFS(Движение!$F:$F,Движение!$L:$L,$A{r},Движение!$M:$M,$C{r},Движение!$H:$H,"ПФД"))')
            ws.cell(r,8,f'=IF($A{r}="","",SUMIFS(Движение!$F:$F,Движение!$L:$L,$A{r},Движение!$D:$D,$C{r},Движение!$H:$H,"ГП",Движение!$I:$I,"ПОКУПАТЕЛЬ"))')
            ws.cell(r,9,f'=IF($A{r}="","",MAX(0,E{r}-F{r}))')
            ws.cell(r,10,f'=IF($A{r}="","",IF(I{r}=0,"Выпущен","В работе"))')
        style_rows(ws,FIRST,FIRST+199,11,input_cols=(1,2,3,5,11),calc_cols=(4,6,7,8,9,10),numfmt={2:'dd.mm.yyyy',5:'0.###',6:'0.###',7:'0.###',8:'0.###',9:'0.###'})

        vs=data.get("varsum",{"head":[],"rows":[]})
        ws=sheet(wb,"Свод_вариантов",VARSUM_HEAD+vs.get("head",[]),[13,28,13,30,12,13]+[10]*12,"Остатки по вариантам. Строка «ИТОГО СЕМЕЙСТВО» — сумма вариантов. Обновляется кнопкой «Обновить отчёты».",data_rows=max(len(vs["rows"]),1),tab="70AD47")
        put_rows(ws,vs["rows"]); ncol=6+len(vs.get("head",[]))
        style_rows(ws,FIRST,FIRST+max(len(vs["rows"]),1)-1,ncol,numfmt={c:'# ##0.###' for c in range(6,ncol+1)})
        for i,row in enumerate(vs["rows"]):
            if row[2]=="ИТОГО СЕМЕЙСТВО":
                for c in range(1,ncol+1): ws.cell(FIRST+i,c).font=F_BOLD; ws.cell(FIRST+i,c).fill=FILL(LIGHT)
        kt=data.get("kits",{"head":[],"rows":[]})
        ws=sheet(wb,"Комплектность",KITS_HEAD+kt.get("head",[]),[16,14,34,11,8]+[12]*10,"Для каждого изделия и объединения: наличие компонентов и сколько штук можно собрать. Обновляется кнопкой «Обновить отчёты».",data_rows=max(len(kt["rows"]),1),tab="70AD47")
        put_rows(ws,kt["rows"]); ncol=5+len(kt.get("head",[]))
        style_rows(ws,FIRST,FIRST+max(len(kt["rows"]),1)-1,ncol,numfmt={**{4:'0.###'},**{c:'# ##0.###' for c in range(6,ncol)}})
        for i,row in enumerate(kt["rows"]):
            if row[2] in ("ИЗДЕЛИЕ","ОБЪЕДИНЕНИЕ"):
                for c in range(1,ncol+1): ws.cell(FIRST+i,c).font=F_BOLD; ws.cell(FIRST+i,c).fill=FILL(LIGHT)
            if row and row[-1]=="ДЕФИЦИТ": ws.cell(FIRST+i,ncol).fill=FILL(BAD)

        # ---------------- МЕСТА, АДРЕСА, ОПЕРАЦИИ ----------------
        ws=sheet(wb,"Места хранения",R.LOC_HEAD,[14,30,14,12,26,18,30,12],"«В остатках = Да» — место получает столбец на листе «Остатки» (до 12 мест). «Склад в 1С» — название склада в выгрузке 1С для сверки.",input_cols=(1,2,3,4,5,6,7),calc_cols=(8,),data_rows=100,tab="A5A5A5")
        put_rows(ws,R.LOC)
        for r in range(FIRST,FIRST+100): ws.cell(r,8,f'=IF($D{r}="Да",COUNTIF($D$4:$D{r},"Да"),"")')
        style_rows(ws,FIRST,FIRST+99,8,input_cols=tuple(range(1,8)),calc_cols=(8,))
        for rr in range(FIRST,FIRST+len(R.LOC)):
            ws.row_dimensions[rr].height=42
            for cc in (2,5,6,7): ws.cell(rr,cc).alignment=Alignment(vertical="center",wrap_text=True)
        dv_list(ws,"C4:C200",values=["Склад","Контроль","Производство","Транзит","Внешний"]); dv_list(ws,"D4:D200",values=["Да","Нет"])
        ws=sheet(wb,"Адреса",R.ADDR_HEAD,[16,14,8,10,10,10,12,12,30],"Адресное хранение: одна строка = одна ячейка. Формат адреса: МЕСТО-ЗОНА-СТЕЛЛАЖ-ЯЧЕЙКА (например ЦС-A-01-02). Адрес обязан относиться к своему месту хранения.",input_cols=tuple(range(1,10)),data_rows=500,tab="A5A5A5")
        put_rows(ws,R.ADDR); style_rows(ws,FIRST,FIRST+499,9,input_cols=tuple(range(1,10)))
        dv_list(ws,"B4:B3000",formula="='Места хранения'!$A$4:$A$200"); dv_list(ws,"H4:H3000",values=["Да","Нет"])
        ws=sheet(wb,"Операции",R.OPS_HEAD,[7,42,13,15,11,12,12,14,9,52],"Справочник операций. Столбец B — то, что выбирается в списке. Новую операцию добавьте строкой: код, «код + название», откуда, куда. Не удаляйте операции, по которым есть движения.",input_cols=tuple(range(1,11)),data_rows=100,tab="A5A5A5")
        put_rows(ws,[[o[0],o[0]+" "+o[1]]+o[2:] for o in R.OPS]); style_rows(ws,FIRST,FIRST+99,10,input_cols=tuple(range(1,11)))
        for rr in range(FIRST,FIRST+len(R.OPS)):
            ws.row_dimensions[rr].height=48
            for cc in (2,10): ws.cell(rr,cc).alignment=Alignment(vertical="center",wrap_text=True)
        for c in "CD": dv_list(ws,f"{c}4:{c}300",formula="='Места хранения'!$A$4:$A$200")
        dv_list(ws,"E4:G300",values=["Да","Нет"]); dv_list(ws,"I4:I300",values=["Да","Нет"])

        # ---------------- ФАЙЛЫ, ПРОСМОТР ----------------
        fl=data.get("files",[])
        ws=sheet(wb,"Файлы",FILES_HEAD,[11,13,44,7,46,28,16,14,24],"Реестр прикреплённых файлов. Название — произвольное (можно переименовать кнопкой). Двойной щелчок по названию открывает файл. «ОЖИДАЕТ» — файл прикреплён на «Вводе», документ ещё не проведён.",data_rows=max(len(fl),1),tab="5B9BD5")
        put_rows(ws,fl); style_rows(ws,FIRST,FIRST+max(len(fl),1)-1,9,numfmt={7:"dd.mm.yyyy hh:mm"})
        preview_sheet(wb)

        # ---------------- 1С ----------------
        im=data.get("imp1c",[])
        ws=sheet(wb,"Импорт_1С",IMP_HEAD,[14,34,24,13,7,30,16],"Сюда кнопкой «Загрузить из 1С» загружается выгрузка остатков (Excel из 1С). Вручную не заполняйте — лист перезаписывается.",data_rows=max(len(im),1),tab="7030A0")
        put_rows(ws,im); style_rows(ws,FIRST,FIRST+max(len(im),1)-1,7,numfmt={4:'# ##0.###',7:"dd.mm.yyyy hh:mm"})
        cm=data.get("cmp1c",[])
        ws=sheet(wb,"Сравнение_1С",CMP_HEAD,[14,32,22,16,13,13,14,22,52],data.get("cmp_sub","Разница = книга − 1С. «СОВПАДАЕТ» — зелёный, расхождения — красный. Кнопка «Пересчитать сверку» после исправлений."),data_rows=max(len(cm),1),tab="7030A0")
        put_rows(ws,cm); style_rows(ws,FIRST,FIRST+max(len(cm),1)-1,9,numfmt={5:'# ##0.###;[Red]-# ##0.###',6:'# ##0.###;[Red]-# ##0.###',7:'# ##0.###;[Red]-# ##0.###'})
        for i,row in enumerate(cm):
            ws.cell(FIRST+i,8).fill=FILL(CALC if row[7]=="СОВПАДАЕТ" else BAD)

        control_sheet(wb,data.get("control"))

        ws=sheet(wb,"Настройки",["Параметр","Значение","Описание"],[34,40,90],"Параметры книги. Меняйте только столбец «Значение». Установщик сам записывает сюда папки.",input_cols=(2,),data_rows=len(R.SETTINGS),tab="A5A5A5")
        st=[list(x) for x in R.SETTINGS]
        if data.get("lastdoc"): st[-1][1]=str(data["lastdoc"])
        put_rows(ws,st); style_rows(ws,FIRST,FIRST+len(st)-1,3,input_cols=(2,))
        for r in range(FIRST,FIRST+len(st)): ws.cell(r,1).font=F_BOLD; ws.cell(r,2).number_format="@"
        tr=data.get("trash",[])
        ws=sheet(wb,"Корзина",R.NOM_HEAD+["Удалено","Кем"],[13,34,18,8,14,14,12,10,14,9,20,26,16,14],"Удалённая номенклатура. Выделите строку и нажмите «Восстановить выделенные» (или правой кнопкой).",data_rows=max(len(tr),1),tab="808080")
        put_rows(ws,tr); style_rows(ws,FIRST,FIRST+max(len(tr),1)-1,14,numfmt={13:"dd.mm.yyyy hh:mm"})
        dm=data.get("delmoves",[])
        ws=sheet(wb,"Движение_удалённые",MOVE_HEAD+["Удалено","Кем"],MOVE_W+[16,14],"Удалённые строки журнала. Выделите и нажмите «Восстановить выделенные» — строка вернётся в конец «Движения».",data_rows=max(len(dm),1),tab="808080")
        put_rows(ws,dm); style_rows(ws,FIRST,FIRST+max(len(dm),1)-1,22,numfmt={2:"dd.mm.yyyy",20:"dd.mm.yyyy hh:mm",21:"dd.mm.yyyy hh:mm"})
        help_sheet(wb,"Инструкция_оператора",OPERATOR_HELP)
        if data.get("testreport"):
            ws=wb.create_sheet("Тест_отчёт"); ws.sheet_view.showGridLines=False
            title(ws,f"ELKA v{VERSION} — ОТЧЁТ АВТОТЕСТА","Результат макроса RunSelfTest, выполненного над этой книгой.",2)
            ws.column_dimensions["A"].width=150
            for i,line in enumerate(data["testreport"].split("\n")):
                c=ws.cell(3+i,1,line); c.font=Font(name="Consolas",size=10,bold=line.startswith("Проверок") or line.startswith("ОШИБКА"),color="C00000" if line.startswith("ОШИБКА") else "000000")
        wb.active=1
        wb.code_name="ThisWorkbook"
    elif mode=="pubfile":
        ws=wb.create_sheet("О_файле")
        ws["A1"]="ELKA — опубликованные данные (только для чтения)"; ws["A1"].font=Font(name="Arial",size=14,bold=True)
        for i,(a,b) in enumerate([("Опубликовано",data.get("published","")),("Оператор",data.get("operator","")),("Версия ELKA","v"+VERSION),("Папка архива документов",data.get("archive","")),("Строк журнала",len(data.get("published_sheets",{}).get("Движение",[])))]):
            ws.cell(3+i,1,a); ws.cell(3+i,2,b)
        ws.column_dimensions["A"].width=30; ws.column_dimensions["B"].width=70
        data_sheets(wb,mode,data,nom)
    else:
        build_viewer(wb,mode,data,nom,moves)
    wb.save(out)

def preview_sheet(wb,viewer=False):
    ws=wb.create_sheet("Просмотр"); ws.sheet_view.showGridLines=False; ws.sheet_properties.tabColor="5B9BD5"
    title(ws,f"ELKA v{VERSION} — ПРОСМОТР ФОТО",("Фото открываются здесь двойным щелчком по столбцу «Документ №» или «Файлы» на листе «Движение». PDF, Word, Excel открываются программой Windows." if viewer else "Фото открываются здесь двойным щелчком по файлу (листы «Файлы», «Движение»). PDF, Word, Excel открываются программой Windows."),8)
    ws.column_dimensions["A"].width=14; ws.column_dimensions["B"].width=80
    ws["A3"]="Название"; ws["A4"]="Файл"; ws["A3"].font=F_BOLD; ws["A4"].font=F_BOLD
    ws["B3"].font=Font(name="Arial",size=12,bold=True); ws["B4"].font=Font(name="Arial",size=9,color="595959")
    ws["A6"]="Изображение:"; ws["A6"].font=F_BOLD

def control_sheet(wb,ctl):
    ws=wb.create_sheet("Контроль"); ws.sheet_view.showGridLines=False; ws.sheet_view.zoomScale=90; ws.sheet_properties.tabColor="C00000"
    title(ws,f"ELKA v{VERSION} — КОНТРОЛЬ КНИГИ","Нажмите «ПРОВЕРИТЬ КНИГУ». Исправляйте сверху вниз: сначала ОШИБКИ, затем ВНИМАНИЕ. Ссылка в столбце «Где» ведёт прямо к ячейке. Подробно — «Инструкция_оператора», раздел 12.",6)
    for i,w in enumerate([8,46,12,12,48,80],1): ws.column_dimensions[L(i)].width=w
    ws["A3"]="Состояние:"; ws["A3"].font=F_BOLD
    ws.merge_cells("B3:F3"); ws["B3"]=(ctl or {}).get("status","Проверка ещё не запускалась"); ws["B3"].font=Font(name="Arial",size=11,bold=True,color="1F4E78")
    for i,h in enumerate(["Код","Проверка","Уровень","Найдено","Статус","Что делать, если не OK"],1):
        c=ws.cell(4,i,h); c.font=F_HDR; c.fill=FILL(HBLUE); c.border=BORDER; c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    ws.row_dimensions[4].height=34
    summ=(ctl or {}).get("summary",[])
    put_rows(ws,summ,5)
    style_rows(ws,5,19,6,height=34)
    for r in range(5,20):
        ws.cell(r,6).alignment=Alignment(wrap_text=True,vertical="center"); ws.cell(r,2).alignment=Alignment(wrap_text=True,vertical="center")
        v=ws.cell(r,5).value
        if v: ws.cell(r,5).fill=FILL(CALC if v=="OK" else (BAD if v=="ИСПРАВИТЬ" else WARN))
    ws["A21"]="СПИСОК НАЙДЕННЫХ ПРОБЛЕМ"; ws["A21"].font=Font(name="Arial",size=12,bold=True,color="1F4E78")
    for i,h in enumerate(["№","Уровень","Код","Где (ссылка)","Описание","Как исправить"],1):
        c=ws.cell(22,i,h); c.font=F_HDR; c.fill=FILL(HBLUE); c.border=BORDER; c.alignment=Alignment(horizontal="center",vertical="center")
    det=(ctl or {}).get("details",[])
    put_rows(ws,det,23)
    if det:
        style_rows(ws,23,22+len(det),6,height=30)
        for i,row in enumerate(det):
            ws.cell(23+i,2).fill=FILL(BAD if row[1]=="ОШИБКА" else WARN)
            for c in (5,6): ws.cell(23+i,c).alignment=Alignment(wrap_text=True,vertical="center")
    ws.freeze_panes="A5"

def help_sheet(wb,name,rows):
    ws=wb.create_sheet(name); ws.sheet_view.showGridLines=False; ws.sheet_properties.tabColor="00B050"
    title(ws,f"ELKA v{VERSION} — {name.replace('_',' ').upper()}","Краткая пошаговая инструкция внутри книги. Полное руководство с иллюстрациями — PDF в папке Docs комплекта.",2)
    ws.column_dimensions["A"].width=30; ws.column_dimensions["B"].width=120
    for i,h in enumerate(["Раздел","Пошагово"],1):
        c=ws.cell(3,i,h); c.font=F_HDR; c.fill=FILL(HBLUE); c.border=BORDER; c.alignment=Alignment(horizontal="center",vertical="center")
    ws.row_dimensions[3].height=30
    for i,(a,b) in enumerate(rows):
        r=4+i
        ca=ws.cell(r,1,a); cb=ws.cell(r,2,b)
        ca.font=F_BOLD; cb.font=Font(name="Arial",size=10)
        for c in (ca,cb): c.border=BORDER; c.alignment=Alignment(wrap_text=True,vertical="top")
        ca.fill=FILL(LIGHT)
        lines=sum(max(1,len(x)//135+1) for x in b.split("\n"))
        ws.row_dimensions[r].height=max(24,14*lines+8)
    ws.freeze_panes="A4"

def data_sheets(wb,mode,data,nom):
    # листы данных (в демо — опубликованные значения; в шаблоне — пустые заголовки)
    pub=data.get("published_sheets",{})
    nrow=max(len(nom),1)
    heads=pub.get("Остатки_head") or (["Код","Наименование","Тип","Ед."]+[""]*NLOC+["Итого (учитываемые места)","Мин. остаток","Статус"])
    ws=sheet(wb,"Остатки",heads,[13,32,19,7]+[11]*NLOC+[14,11,16],"Остатки по местам хранения на момент публикации. Фильтр — стрелки в заголовках. Поиск — кнопка «Найти позицию» на «Панели».",data_rows=max(len(pub.get("Остатки",[])),1),tab="70AD47")
    rows=pub.get("Остатки",[]); put_rows(ws,rows); style_rows(ws,FIRST,FIRST+max(len(rows),1)-1,7+NLOC,numfmt={c:'# ##0.###;[Red]-# ##0.###' for c in range(5,7+NLOC)})
    for r in range(FIRST,FIRST+len(rows)): ws.cell(r,5+NLOC).font=F_BOLD
    if mode!="pubfile" and (mode=="viewer_demo" or rows): ws.protection.sheet=True; ws.protection.autoFilter=False; ws.protection.sort=False; ws.protection.formatColumns=False
    spec_def=[("Остатки_адреса",ADDRSTOCK_HEAD,[13,34,10,16,13,8],"Где лежит позиция: остатки по ячейкам.",{5:'# ##0.###'}),
              ("Движение",MOVE_HEAD,MOVE_W,"Журнал операций. Двойной щелчок по столбцу «Файлы» или «Документ №» открывает прикреплённый документ.",{2:"dd.mm.yyyy",6:'# ##0.###',20:"dd.mm.yyyy hh:mm"}),
              ("Комплектность",KITS_HEAD+pub.get("Комплектность_head",[]),[16,14,34,11,8]+[12]*10,"Сколько изделий и объединений можно собрать из имеющихся компонентов.",{}),
              ("Файлы",FILES_HEAD,[11,13,44,7,46,28,16,14,24],"Документы и фото. Двойной щелчок по названию — открыть.",{7:"dd.mm.yyyy hh:mm"})]
    for nm_,head,w,sub,fmt in spec_def:
        rows=pub.get(nm_,[])
        ws=sheet(wb,nm_,head,w,sub,data_rows=max(len(rows),1),tab="5B9BD5" if nm_=="Файлы" else "70AD47")
        put_rows(ws,rows); style_rows(ws,FIRST,FIRST+max(len(rows),1)-1,len(head),numfmt=fmt)
        if nm_ in ("Свод_вариантов","Комплектность"):
            for i,row in enumerate(rows):
                if len(row)>2 and row[2] in ("ИТОГО СЕМЕЙСТВО","ИЗДЕЛИЕ","ОБЪЕДИНЕНИЕ"):
                    for c in range(1,len(head)+1): ws.cell(FIRST+i,c).font=F_BOLD; ws.cell(FIRST+i,c).fill=FILL(LIGHT)
        if mode!="pubfile" and (mode=="viewer_demo" or rows): ws.protection.sheet=True; ws.protection.autoFilter=False; ws.protection.sort=False; ws.protection.formatColumns=False
        if nm_=="Файлы" and mode!="pubfile": ws.sheet_state="veryHidden"

def build_viewer(wb,mode,data,nom,moves):
    ws=wb.create_sheet("Панель"); ws.sheet_view.showGridLines=False; ws.sheet_properties.tabColor="1F4E78"
    title(ws,"ELKA — ПРОСМОТР ОСТАТКОВ, ДВИЖЕНИЯ И КОМПЛЕКТАЦИИ","Книга ПОЛЬЗОВАТЕЛЯ: только просмотр. Нажмите «ОБНОВИТЬ ДАННЫЕ», чтобы получить последние данные оператора. Инструкция — лист «Инструкция».",6)
    for c,w in zip("ABCDEF",[4,28,52,3,36,4]): ws.column_dimensions[c].width=w
    info=[("Данные опубликованы",data.get("published","данные ещё не загружались")),("Загруженная версия",data.get("version","")),("Папка данных",""),("Папка архива документов","")]
    for i,(a,b) in enumerate(info):
        r=5+i; ws.cell(r,2,a).font=F_BOLD; ws.cell(r,3,b).font=Font(name="Arial",size=11,color="1F4E78",bold=(i==0))
        for c in (2,3): ws.cell(r,c).border=BORDER
        ws.row_dimensions[r].height=24
    ws.cell(4,2,"СОСТОЯНИЕ").font=Font(name="Arial",size=11,bold=True,color="1F4E78")
    ws.cell(11,2,"ЛИСТЫ").font=Font(name="Arial",size=11,bold=True,color="1F4E78")
    nav=[("Остатки","Остатки по местам хранения (ЦС, ПС, цех, ПФ, ГП, брак…)"),("Остатки_адреса","Где лежит: остатки по ячейкам склада"),("Движение","Все операции; двойной щелчок по «Документ №» / «Файлы» — открыть документ или фото"),
         ("Комплектность","Комплектация: сколько изделий и комплектов можно собрать, чего не хватает"),("Просмотр","Здесь открываются фото документов"),("Инструкция","Как пользоваться этой книгой")]
    for i,(s,d) in enumerate(nav):
        r=12+i
        c=ws.cell(r,2,f'=HYPERLINK("#\'{s}\'!A1","{s}")'); c.font=Font(name="Arial",size=11,bold=True,color="0563C1",underline="single"); c.border=BORDER
        ws.cell(r,3,d).font=F_DATA; ws.cell(r,3).border=BORDER; ws.row_dimensions[r].height=22
    data_sheets(wb,mode,data,nom)
    preview_sheet(wb,True)
    help_sheet(wb,"Инструкция",USER_HELP)
    # «Файлы» — служебный (скрытый) лист: переносим в конец
    wb.move_sheet("Файлы",offset=len(wb.sheetnames)-1-wb.sheetnames.index("Файлы"))
    cfg=wb.create_sheet("_cfg"); cfg.sheet_state="veryHidden"
    for i,(a,b) in enumerate([("Папка публикации",data.get("pubfolder","")),("Папка архива (из публикации)",data.get("archive","")),("Папка архива (вручную)",""),("Загруженная версия",data.get("version","")),("Дата публикации",data.get("published","")),("Автообновление при открытии (Да/Нет)","Да")]):
        cfg.cell(1+i,1,a); cfg.cell(1+i,2,b)
    wb.active=0
    wb.code_name="ThisWorkbook"

from helptext import OPERATOR_HELP, USER_HELP

if __name__=="__main__":
    mode,out=sys.argv[1],sys.argv[2]
    def _hook(d):
        if "__dt__" in d and len(d)==1:
            return datetime.datetime.fromisoformat(d["__dt__"])
        return d
    data=json.load(open(sys.argv[3],encoding="utf-8"),object_hook=_hook) if len(sys.argv)>3 else None
    build(mode,out,data)
    print("built",mode,out)
