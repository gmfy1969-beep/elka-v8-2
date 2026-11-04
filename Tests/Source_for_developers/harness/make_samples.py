import sys,os
from PIL import Image, ImageDraw, ImageFont
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import docx
d=sys.argv[1]; os.makedirs(d,exist_ok=True)
pdfmetrics.registerFont(TTFont("Arial","/usr/share/fonts/msttcore/arial.ttf") if os.path.exists("/usr/share/fonts/msttcore/arial.ttf") else TTFont("Arial",[p for p in os.popen("fc-list : file | grep -i 'arial.ttf'").read().split(":") if p.strip().endswith(".ttf")][0].strip()))
c=canvas.Canvas(os.path.join(d,"invoice_001.pdf"),pagesize=A4); c.setFont("Arial",16)
c.drawString(60,780,"ТОВАРНАЯ НАКЛАДНАЯ № 125 (демо)"); c.setFont("Arial",11)
y=740
for line in ["Поставщик: ООО «Металлснаб»","Получатель: склад ЦС","DET-001 Болт М8x30 — 10 000 шт","DET-002 Гайка М8 — 10 000 шт","... (демонстрационный документ ELKA)"]:
    c.drawString(60,y,line); y-=22
c.save()
fp=[p for p in os.popen("fc-list : file | grep -i 'arial'").read().split("\n") if p.strip()]
font=ImageFont.truetype(fp[0].split(":")[0].strip(),40) if fp else None
im=Image.new("RGB",(1280,860),(222,233,245)); dr=ImageDraw.Draw(im)
for i in range(4):
    dr.rectangle([80+i*290,300,320+i*290,700],fill=(181,137,86),outline=(90,60,30),width=6)
dr.text((80,90),"Фото паллет при приёмке (демо ELKA)",fill=(20,40,80),font=font)
im.save(os.path.join(d,"pallet_photo.jpg"),"JPEG",quality=85)
doc=docx.Document(); doc.add_heading("Требование-накладная М-11 № 44 (демо)",1)
doc.add_paragraph("Отпуск материалов на производственный склад ПС.")
t=doc.add_table(rows=1,cols=3); t.rows[0].cells[0].text="Код"; t.rows[0].cells[1].text="Наименование"; t.rows[0].cells[2].text="Кол-во"
for r in (("DET-003","Шайба плоская 8 мм","40"),("MAT-001","Лист стальной 2 мм","25,5"),("PNL-RED","Панель боковая красная","12")):
    cells=t.add_row().cells
    for i,v in enumerate(r): cells[i].text=v
doc.save(os.path.join(d,"defect_act.docx"))
print("samples ok")
