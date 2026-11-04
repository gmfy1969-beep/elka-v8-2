import openpyxl, subprocess, glob, os, shutil, datetime
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from PIL import Image, ImageChops
IMG="/data/v82/docs/img/"; T="/tmp/shots/"; shutil.rmtree(T,ignore_errors=True); os.makedirs(T)
BF=PatternFill("solid",fgColor="E1E1E1"); BD=Border(*(Side("medium",color="7F7F7F"),)*4)
def btn(ws,addr,text,merge=1):
    c=ws[addr]; c.value=text; c.fill=BF; c.border=BD; c.font=Font(bold=True,size=10,color="000000"); c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    if merge>1:
        col=c.column; ws.merge_cells(start_row=c.row,start_column=col,end_row=c.row,end_column=col+merge-1)
def render(path,view,name):
    od=T+name; subprocess.run(["python3","/skills/sheets/scripts/render_workbook.py",path,"--out-dir",od,"--view",view,"--dpi","110"],capture_output=True)
    png=sorted(glob.glob(od+"/*.png"))[0]; im=Image.open(png).convert("RGB")
    bg=Image.new("RGB",im.size,(255,255,255)); bb=ImageChops.difference(im,bg).getbbox()
    im=im.crop((max(bb[0]-10,0),max(bb[1]-10,0),min(bb[2]+10,im.width),min(bb[3]+10,im.height))); im.save(IMG+name+".png"); print(name,im.size)
# operator panel
wb=openpyxl.load_workbook("/data/v82/out/ELKA_Operator_Demo.xlsx"); ws=wb["Панель"]
for a,t in [("E5","Ввод операции"),("E7","Проверить книгу"),("E9","Обновить отчёты"),("E11","Загрузить из 1С и сверить"),("E13","ОПУБЛИКОВАТЬ для пользователей"),("E15","Резервная копия сейчас")]: btn(ws,a,t)
wi=wb["Ввод"]
wi["C4"]=datetime.datetime(2026,10,5); wi["C5"]="07 Выдача в производство"; wi["C8"]="З-2026-015"; wi["C9"]="PF-001"; wi["C13"]=10
wi["C12"]="Требование-накладная М-11 №45"
for i,(c,q,f) in enumerate([("DET-001",40,""),("MAT-002",20,""),("DET-003",500,"")]):
    wi.cell(16+i,2,c); wi.cell(16+i,5,q)
for a,t in [("K4","ПРОВЕСТИ"),("K6","Прикрепить файл(ы)"),("K8","Заполнить по спецификации"),("K10","Заполнить по объединению"),("K12","Заполнить остатками «Откуда»"),("M4","СОБРАТЬ изделие (08+09/10)"),("M6","Скомплектовать объединение"),("M8","Разобрать объединение"),("M12","Очистить лист")]: btn(wi,a,t)
wb.save(T+"op.xlsx")
render(T+"op.xlsx","Панель!A1:F26","op_panel")
render(T+"op.xlsx","Ввод!A1:M22","op_input")
render(T+"op.xlsx","Движение!A1:M16","op_moves")
render(T+"op.xlsx","Остатки!A1:N20","op_stock")
wc=openpyxl.load_workbook("/tmp/demo_retest.xlsx"); wc.save(T+"ctl.xlsx")
render(T+"ctl.xlsx","Контроль!A1:H30","op_control")
wv=openpyxl.load_workbook("/data/v82/out/ELKA_Viewer_Demo.xlsx"); wp=wv["Панель"]
wp["C7"]="C:\\ELKA\\DEMO\\Publish\\"; wp["C8"]="C:\\ELKA\\DEMO\\Архив\\"
for a,t in [("E4","ОБНОВИТЬ ДАННЫЕ"),("E7","Найти позицию"),("E9","Сбросить фильтры"),("E11","Открыть файл выделенной строки"),("E13","Настроить папки")]: btn(wp,a,t)
wv.save(T+"v.xlsx")
render(T+"v.xlsx","Панель!A1:F16","viewer_panel")
render(T+"v.xlsx","Остатки!A1:N20","viewer_stock")
render(T+"v.xlsx","Движение!A1:M16","viewer_moves")
render(T+"v.xlsx","Комплектность!A1:K14","viewer_kits")
