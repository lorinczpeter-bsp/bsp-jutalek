"""Tömör, arculatos elszámolások: fekvő A4, egyetlen táblázatszélesség."""
import io
from pathlib import Path
from xml.sax.saxutils import escape
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether, PageBreak, Flowable
from core import export_columns, export_rows, compact_fields, NAVY, YELLOW, dec

ASSETS=Path(__file__).parent/"assets"

def ft(n,digits=0):
    return f"{float(n):,.{digits}f}".replace(","," ").replace(".",",")

def excel_safe(value):
    if isinstance(value,str) and value.startswith(("=","+","-","@")):
        return "'"+value
    return value

def field_weight(field):
    s=field["source"]
    if s=="Fuvarszám":return 43
    if s in ["Megbízó","Alvállalkozó"]:return 133
    if isinstance(s,list):return 80
    if s=="Teljesítés időpontja":return 70
    if field.get("currency"):return 70
    return 70

def display_rows(plan):
    rows=export_rows(plan)
    for row,line in zip(rows,plan["lines"]):
        for i,f in enumerate(compact_fields(plan)):
            if f.get("currency"):
                currency=str(line["original"].get(f["currency"]) or "").strip()
                if isinstance(row[i],(float,int)):
                    row[i]=ft(row[i],2)+" "+currency
                elif row[i]!="":
                    row[i]=str(row[i])+" "+currency
    return rows

def centered_from(plan):
    return next((i for i,f in enumerate(compact_fields(plan)) if f["source"]=="Alvállalkozói díj"),len(compact_fields(plan)))

def summary_items(plan):
    lines=plan["lines"]
    return [
        ("Elszámolt fuvarok",len(lines),"db"),
        ("Árrés összesen, forintban",sum(dec(x["margin_huf"]) for x in lines),"Ft"),
        ("Negatív árrésű fuvarok",sum(x["negative"] for x in lines),"db"),
        ("Negatív árrésű tételek jutaléka",sum(x["commission"] for x in lines if x["negative"]),"Ft"),
        ("Tárgyhavi jutalék",plan["current"],"Ft"),
    ] + ([] if plan.get("snapshot") else [("Korábbi hónapok korrekciója",plan["corrections"],"Ft")]) + [
        ("ÖSSZES JUTALÉK",plan["total"],"Ft")]

def summary_cards(plan):
    lines=plan["lines"]
    negative=sum(x["negative"] for x in lines)
    rate=plan["rates"].get(plan["month"])
    rate_value=dec(rate["rate"]) if rate else ("Több havi árfolyam" if plan["rates"] else "Csak HUF")
    return [
        ("MÉG FIZETENDŐ JUTALÉK" if plan.get("paid_keys") else "ÖSSZES JUTALÉK",plan["total"],"Ft","primary"),
        ("Árrés extrákkal, forintban",sum(dec(x["margin_huf"]) for x in lines),"Ft","neutral"),
        ("Elszámolt fuvarok",len(lines),"db","neutral"),
    ] + ([
        ("Számított teljes jutalék",plan["earned_total"],"Ft","neutral"),
        ("Kifizetettként kezelt rész",plan["paid_total"],"Ft","neutral"),
    ] if plan.get("paid_keys") else []) + ([
        ("Negatív árrésű fuvarok",negative,"db","negative"),
        ("Negatív tételek jutaléka",sum(x["commission"] for x in lines if x["negative"]),"Ft","negative"),
    ] if negative else []) + [
        ("MNB havi EUR/HUF átlag",rate_value,"Ft/EUR" if rate else "","neutral"),
    ]

class SummaryCards(Flowable):
    """Two rows of three cards; all metrics use the same data as the workbook."""
    def __init__(self,plan,width):
        Flowable.__init__(self)
        self.plan=plan;self.width=width
        columns=3 if len(summary_cards(plan))>4 else 2
        rows=(len(summary_cards(plan))+columns-1)//columns
        self.height=rows*104+(rows-1)*12
    def draw(self):
        c=self.canv;c.saveState()
        cards=summary_cards(self.plan);columns=3 if len(cards)>4 else 2
        gap=12;cw=(self.width-(columns-1)*gap)/columns;ch=104
        for i,(label,value,unit,tone) in enumerate(cards):
            x=(i%columns)*(cw+gap);y=self.height-(i//columns+1)*ch-(i//columns)*gap
            bg=NAVY if tone=="primary" else ("FFF0ED" if tone=="negative" else "F2F4F7")
            fg="FFFFFF" if tone=="primary" else ("9C2222" if tone=="negative" else NAVY)
            c.setFillColor(colors.HexColor("#"+bg));c.roundRect(x,y,cw,ch,9,fill=1,stroke=0)
            c.setFillColor(colors.HexColor("#"+fg));c.setFont("Fira",10)
            c.drawString(x+16,y+ch-25,label)
            text=value if isinstance(value,str) else ft(value,2 if unit=="Ft/EUR" else 0)
            text=str(text)+(" "+unit if unit else "")
            size=24 if tone=="primary" else (19 if unit=="Ft/EUR" else 22)
            while pdfmetrics.stringWidth(text,"FiraBold",size)>cw-32 and size>12:size-=1
            c.setFont("FiraBold",size)
            c.setFillColor(colors.HexColor("#"+YELLOW if tone=="primary" else "#"+fg))
            c.drawString(x+16,y+29,text)
            if tone=="primary":
                c.setFillColor(colors.HexColor("#"+YELLOW));c.roundRect(x+16,y+ch-8,35,3,1,fill=1,stroke=0)
        c.restoreState()

def make_excel(plan):
    wb=openpyxl.Workbook();ws=wb.active;ws.title="Jutalékelszámolás"
    cols,rows=export_columns(plan),export_rows(plan)
    fields=compact_fields(plan);n=len(cols);center_start=centered_from(plan)
    ws.sheet_view.showGridLines=False
    logo=XLImage(io.BytesIO((ASSETS/"logo.png").read_bytes()))
    logo.width,logo.height=200,200*logo.height/logo.width
    ws.add_image(logo,"A1")
    for r in range(1,4):
        ws.row_dimensions[r].height=26
        for j in range(1,n+1):ws.cell(r,j).fill=PatternFill("solid",fgColor="FFFFFF")
    for r,text,size in [(1,"ALVÁLLALKOZÓI JUTALÉKELSZÁMOLÁS",15),
                         (2,f'{plan["month"]} | {plan["contact"]}',12),
                         (3,f'{plan["status"]} | Jutalék: 10% | Készült: {plan["created"][:19]}',10)]:
        ws.merge_cells(start_row=r,start_column=4,end_row=r,end_column=n)
        ws.cell(r,4,excel_safe(text)).font=Font(name="Montserrat" if r==1 else "Fira Sans",size=size,color=NAVY,bold=r<3)
    for j in range(1,n+1):ws.cell(4,j).fill=PatternFill("solid",fgColor=NAVY)
    ws.row_dimensions[4].height=4
    rate_text="; ".join(f'{m}: {ft(p["rate"],2)} Ft/EUR' for m,p in sorted(plan["rates"].items()))
    ws.merge_cells(start_row=5,start_column=1,end_row=5,end_column=n)
    ws.cell(5,1,"MNB havi átlag: "+(rate_text or "csak HUF-tételek"))
    ws.cell(5,1).alignment=Alignment(wrap_text=True)
    ws.row_dimensions[5].height=max(26,14*((len(rate_text)//110)+1))
    ws.merge_cells(start_row=6,start_column=1,end_row=6,end_column=n)
    ws.cell(6,1,"A negatív árrésű tételek jutaléka 0 Ft. A jutalék az adott elszámolásban ténylegesen rendezendő összeg.")
    ws.cell(6,1).font=Font(name="Fira Sans",size=9,color=NAVY)
    header=8
    for j,label in enumerate(cols,1):
        c=ws.cell(header,j,label)
        c.fill=PatternFill("solid",fgColor=NAVY)
        c.font=Font(name="Fira Sans",bold=True,size=10,color="FFFFFF")
        c.alignment=Alignment(wrap_text=True,vertical="center",horizontal="center" if j-1>=center_start else "left")
        ws.column_dimensions[get_column_letter(j)].width=(field_weight(fields[j-1])/4 if j<=len(fields) else 20)
    ws.row_dimensions[8].height=34
    for i,(row,line) in enumerate(zip(rows,plan["lines"]),9):
        for j,v in enumerate(row,1):
            c=ws.cell(i,j,excel_safe(v))
            c.font=Font(name="Fira Sans",size=10,color="9C2222" if line["negative"] else NAVY)
            c.fill=PatternFill("solid",fgColor="FFF0ED" if line["negative"] else ("F2F4F7" if i%2 else "FFFFFF"))
            c.alignment=Alignment(wrap_text=True,vertical="top",horizontal="center" if j-1>=center_start else "left")
            if j==n:
                c.number_format='#,##0" Ft"' if not line.get("paid") else '#,##0" Ft (kifizetve)"'
                c.font=Font(name="Fira Sans",size=10,bold=True,color="9C2222" if line["negative"] or v<0 else NAVY)
            elif fields[j-1].get("currency") and isinstance(v,(int,float)):
                currency=str(line["original"].get(fields[j-1]["currency"]) or "").replace('"',"")
                c.number_format='#,##0.00" '+currency+'"'
        # A cégnevek és a több dátum is teljes egészében olvasható marad.
        height=1
        for j,v in enumerate(row,1):
            width=ws.column_dimensions[get_column_letter(j)].width
            height=max(height,sum(max(1,int(len(part)/max(8,width-2))+1) for part in str(v).split("\n")))
        ws.row_dimensions[i].height=max(28,14*height)
    bottom=8+len(rows)+2
    totals=[("Tárgyhavi jutalék (Ft)",plan["current"]),
            ("Korábbi hónapok korrekciója (Ft)",plan["corrections"]),
            ("ÖSSZES FIZETENDŐ JUTALÉK (Ft)",plan["total"])]
    for k,(label,value) in enumerate(totals):
        r=bottom+k
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=n-1)
        ws.cell(r,1,label);ws.cell(r,n,value).number_format='#,##0" Ft"'
        for c in ws[r]:
            c.fill=PatternFill("solid",fgColor=YELLOW if k==2 else "E9ECF1")
            c.font=Font(name="Fira Sans",bold=True,size=11,color=NAVY)
        ws.row_dimensions[r].height=25
    ws.freeze_panes="B9"
    ws.auto_filter.ref=f"A8:{get_column_letter(n)}{max(8,8+len(rows))}"
    ws.print_title_rows="1:8"
    ws.page_setup.orientation="landscape";ws.page_setup.paperSize=ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth=1;ws.page_setup.fitToHeight=0
    ws.sheet_properties.pageSetUpPr.fitToPage=True
    ws.page_margins.left=0.2;ws.page_margins.right=0.2
    ws.print_area=f"A1:{get_column_letter(n)}{bottom+2}"
    ws.oddFooter.center.text="Bábolna Sped Kft. | &P / &N"
    # Az árfolyamnapló megmarad ellenőrzésre, de nem növeli a nyomtatott lista oszlopszámát.
    audit=wb.create_sheet("Árfolyamnapló")
    audit.append(["Hónap","EUR/HUF átlag","Módszer","Forrás","Rögzítés","Dátum","Napi EUR/HUF"])
    ar=2
    for m,payload in sorted(plan["rates"].items()):
        audit.cell(ar,1,m);audit.cell(ar,2,float(dec(payload["rate"]))).number_format="0.00"
        audit.cell(ar,3,payload.get("method",""));audit.cell(ar,4,payload["source"]);audit.cell(ar,5,payload.get("retrieved",""))
        for d,v in sorted(payload.get("daily",{}).items()):
            audit.cell(ar,6,d);audit.cell(ar,7,float(dec(v))).number_format="0.00";ar+=1
        ar=max(ar,audit.max_row+1)
    audit.sheet_state="hidden"
    summary=wb.create_sheet("Összesítő",0)
    summary.sheet_view.showGridLines=False
    cards=summary_cards(plan);columns=3 if len(cards)>4 else 2
    for col in "ABCDEFGHIJK":summary.column_dimensions[col].width=15
    for col in ("DH" if columns==3 else "F"):summary.column_dimensions[col].width=3
    image=XLImage(io.BytesIO((ASSETS/"logo.png").read_bytes()))
    image.width,image.height=200,200*image.height/image.width
    summary.add_image(image,"A1")
    for r in range(1,5):summary.row_dimensions[r].height=24
    for r,text in [(1,"ALVÁLLALKOZÓI JUTALÉKELSZÁMOLÁS"),(2,plan["contact"]),(3,plan["month"]+" | "+plan["status"])]:
        summary.merge_cells(start_row=r,start_column=5,end_row=r,end_column=11)
        summary.cell(r,5,excel_safe(text)).font=Font(name="Fira Sans",size=14 if r==1 else 12,bold=True,color=NAVY)
    summary.merge_cells("A6:K6");summary["A6"]="ÖSSZESÍTŐ"
    summary["A6"].font=Font(name="Fira Sans",bold=True,size=17,color=NAVY)
    for i,(label,value,unit,tone) in enumerate(cards):
        top=8+(i//columns)*6;left=1+(i%columns)*(4 if columns==3 else 6);right=left+(2 if columns==3 else 4)
        bg=NAVY if tone=="primary" else ("FFF0ED" if tone=="negative" else "F2F4F7")
        fg="FFFFFF" if tone=="primary" else ("9C2222" if tone=="negative" else NAVY)
        for row in summary.iter_rows(min_row=top,max_row=top+3,min_col=left,max_col=right):
            for c in row:c.fill=PatternFill("solid",fgColor=bg)
        summary.merge_cells(start_row=top,start_column=left,end_row=top,end_column=right)
        summary.cell(top,left,label).font=Font(name="Fira Sans",size=11,color=fg)
        summary.cell(top,left).alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
        summary.merge_cells(start_row=top+1,start_column=left,end_row=top+2,end_column=right)
        cell=summary.cell(top+1,left,float(value) if not isinstance(value,str) else value)
        cell.number_format=('#,##0.00' if unit=="Ft/EUR" else '#,##0')+'" '+unit+'"'
        cell.font=Font(name="Fira Sans",bold=True,size=23 if unit!="Ft/EUR" else 19,color=YELLOW if tone=="primary" else fg)
        cell.alignment=Alignment(horizontal="center",vertical="center")
        for r in range(top,top+4):summary.row_dimensions[r].height=23
    r=8+((len(cards)+columns-1)//columns)*6
    notes=["MNB havi átlag: "+(rate_text or "csak HUF-tételek"),
           "Jutalék: 10%. A negatív árrésű tételek jutaléka 0 Ft. Tételenként egész forintra kerekítve.",
           ("A feltöltött korábbi kimutatásban szereplő kifizetéseket nem fizetjük ki újra." if plan.get("paid_keys") else "A kiválasztott hónap és fuvarszervező jutalékelszámolása.") if plan.get("snapshot") else "A korrekciók különbözeti összeget tartalmaznak."]
    if plan["blockers"]:notes += ["HIÁNYOS KIMUTATÁS: "+"; ".join(plan["blockers"])]
    notes+=plan["warnings"]
    for text in notes:
        summary.merge_cells(start_row=r,start_column=1,end_row=r,end_column=11)
        summary.cell(r,1,excel_safe(text)).alignment=Alignment(wrap_text=True,vertical="center")
        summary.cell(r,1).font=Font(name="Fira Sans",size=10,color=NAVY)
        summary.row_dimensions[r].height=max(28,14*(len(text)//140+1));r+=1
    summary.print_area=f"A1:K{r}"
    summary.page_setup.orientation="landscape";summary.page_setup.paperSize=summary.PAPERSIZE_A4
    summary.page_setup.fitToWidth=1;summary.page_setup.fitToHeight=1
    summary.sheet_properties.pageSetUpPr.fitToPage=True
    summary.oddFooter.center.text="Bábolna Sped Kft. | Összesítő"
    if plan.get("paid_keys"):
        payment_audit=wb.create_sheet("Kifizetési ellenőrzés")
        payment_audit.append(["Fuvarszám","Számított jutalék (Ft)","Korábbi tényleges kifizetés (Ft)","Kifizetés dátuma","Fuvarszervező","Forrás","Újra fizetendő (Ft)"])
        for line in plan["lines"]:
            if line.get("paid"):
                paid=line["payment"]
                payment_audit.append([line["key"],line["earned_commission"],float(dec(paid["amount"])),paid["paid_date"],excel_safe(paid["contact"]),excel_safe(paid["source"]),0])
        for column in payment_audit.columns:
            payment_audit.column_dimensions[column[0].column_letter].width=26
        payment_audit.freeze_panes="A2"
    calculations=wb.create_sheet("Számítási napló")
    calculations.append(["Fuvarszám","Elszámolási hónap","Módszer","Forrás árrés","Forrás pénznem","Számított árrés (Ft)","MNB átlag (pontos)","Számított jutalék (Ft)","Újra fizetendő (Ft)"])
    for line in plan["lines"]:
        source=line.get("source_original",line["original"])
        mixed=source.get("Díj pénzneme")!=source.get("Alvállalkozói díj pénzneme")
        calculations.append([line["key"],line["origin"],"Vegyes díjak forintosítva" if mixed else "Exportált árrés",source.get("Árrés extra számlasorokkal (állapot szerint)"),source.get("Árrés extra számlasorokkal (állapot szerint) pénzneme"),float(dec(line["margin_huf"])),float(dec(line["rate"])),line.get("earned_commission",line["commission"]),line["commission"]])
        calculations.cell(calculations.max_row,7).number_format="0.00"
    calculations.sheet_state="hidden"
    wb.active=0
    out=io.BytesIO();wb.save(out);return out.getvalue()

def make_pdf(plan):
    for name,file in [("Fira","FiraSans-Regular.ttf"),("FiraBold","FiraSans-Bold.ttf")]:
        if name not in pdfmetrics.getRegisteredFontNames():
            pdfmetrics.registerFont(TTFont(name,str(ASSETS/file)))
    navy=colors.HexColor("#"+NAVY);yellow=colors.HexColor("#"+YELLOW)
    width,height=landscape(A4);available=width-36
    body=ParagraphStyle("body",fontName="Fira",fontSize=7.5,leading=9.2,textColor=navy,splitLongWords=False)
    head=ParagraphStyle("head",parent=body,fontName="FiraBold",fontSize=7.5,leading=9.1,textColor=colors.white)
    center=ParagraphStyle("center",parent=body,alignment=1)
    right=ParagraphStyle("right",parent=body,alignment=2)
    bold=ParagraphStyle("bold",parent=right,fontName="FiraBold",fontSize=8,leading=10)
    note=ParagraphStyle("note",parent=body,fontSize=7,leading=9)
    def p(v,style=body):
        return Paragraph(escape(str(v if v is not None else "")).replace("\n","<br/>"),style)
    out=io.BytesIO()
    doc=SimpleDocTemplate(out,pagesize=(width,height),leftMargin=18,rightMargin=18,topMargin=87,bottomMargin=27,
                         title="Alvállalkozói jutalékelszámolás",author="Bábolna Sped Kft.")
    story=[p("ÖSSZESÍTŐ",ParagraphStyle("summarytitle",parent=body,fontName="FiraBold",fontSize=19,leading=24)),Spacer(1,18)]
    story += [SummaryCards(plan,available),Spacer(1,18)]
    if plan.get("snapshot"):
        story += [p("A feltöltött korábbi kimutatásban szereplő kifizetéseket nem fizetjük ki újra." if plan.get("paid_keys") else "A kiválasztott hónap és fuvarszervező jutalékelszámolása.",note),Spacer(1,5)]
    if plan["rates"]:
        rate_text="; ".join(f'{m}: {ft(r["rate"],2)} Ft/EUR' for m,r in sorted(plan["rates"].items()))
        story += [p("MNB havi átlag: "+rate_text,note),Spacer(1,5)]
    if plan["blockers"]:
        story += [p("HIÁNYOS KIMUTATÁS: nyitott ellenőrzések vannak. Az összeg nem tartalmazza a hibás vagy nem számítható tételeket.",note),Spacer(1,5)]
    for msg in plan["blockers"]:story += [p(msg,note),Spacer(1,3)]
    for w in plan["warnings"]:story += [p(w,note),Spacer(1,3)]
    story += [p("Jutalék: 10%. A negatív árrésű tételek jutaléka 0 Ft. Tételenként egész forintra kerekítve.",note),
              PageBreak(),p("RÉSZLETEZŐ",ParagraphStyle("detailtitle",parent=body,fontName="FiraBold",fontSize=13,leading=17)),Spacer(1,8)]
    cols,rows=export_columns(plan),display_rows(plan);fields=compact_fields(plan);center_start=centered_from(plan)
    weights=[field_weight(f) for f in fields]+[67]
    widths=[available*x/sum(weights) for x in weights]
    table_data=[[p(h,ParagraphStyle("hc",parent=head,alignment=1) if i>=center_start else head) for i,h in enumerate(cols)]]
    for row,line in zip(rows,plan["lines"]):
        cells=[]
        for i,v in enumerate(row):
            if i==len(cols)-1:
                cells.append(p(ft(v)+" Ft"+("\nkifizetve" if line.get("paid") else ""),ParagraphStyle("commission",parent=bold,alignment=1,textColor=colors.HexColor("#9C2222") if line["negative"] else navy)))
            else:
                if fields[i]["source"]=="Teljesítés időpontja":
                    v=str(v).replace(", ","\n").replace(",","\n")
                cells.append(p(v,ParagraphStyle("cell",parent=center if i>=center_start else body,textColor=colors.HexColor("#9C2222") if line["negative"] else navy)))
        table_data.append(cells)
    if not rows:
        table_data.append([p("Nincs tétel")]+[p("") for _ in cols[1:]])
    table=Table(table_data,colWidths=widths,repeatRows=1,hAlign="LEFT")
    commands=[("BACKGROUND",(0,0),(-1,0),navy),("VALIGN",(0,0),(-1,-1),"TOP"),
              ("TOPPADDING",(0,0),(-1,0),7),("BOTTOMPADDING",(0,0),(-1,0),7),
              ("TOPPADDING",(0,1),(-1,-1),4),("BOTTOMPADDING",(0,1),(-1,-1),4),
              ("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4),
              ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white,colors.HexColor("#F2F4F7")]),
              ("LINEBELOW",(0,0),(-1,0),1,yellow)]
    for i,line in enumerate(plan["lines"],1):
        if line["negative"]:
            commands += [("BACKGROUND",(0,i),(-1,i),colors.HexColor("#FFF0ED"))]
    table.setStyle(TableStyle(commands))
    story.append(table)
    sums=[]
    if plan["corrections"]:
        sums += [[p("Tárgyhavi jutalék"),p(ft(plan["current"])+" Ft",right)],
                 [p("Korábbi hónapok korrekciója"),p(ft(plan["corrections"])+" Ft",right)]]
    sums += [[p("ÖSSZES FIZETENDŐ JUTALÉK",body),p(ft(plan["total"])+" Ft",bold)]]
    totals=Table(sums,colWidths=[available-135,135])
    totals.setStyle(TableStyle([("BACKGROUND",(0,-1),(-1,-1),yellow),("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                               ("TOPPADDING",(0,0),(-1,-1),8),("BOTTOMPADDING",(0,0),(-1,-1),8)]))
    story.append(KeepTogether([Spacer(1,7),totals,Spacer(1,4),
        p("A negatív árrésű tételek jutaléka 0 Ft. A kimutatás a feltöltött forrás aktuális havi állapotát mutatja." if plan.get("snapshot") else "A negatív árrésű tételek jutaléka 0 Ft. Korrekciónál a különbözet szerepel.",note)]))
    def page(canvas,doc):
        # Az engedélyezett vízszintes logó fehér alapon, torzítás nélkül.
        iw,ih=ImageReader(str(ASSETS/"logo.png")).getSize()
        lw=167;lh=lw*ih/iw
        canvas.drawImage(str(ASSETS/"logo.png"),18,height-77,width=lw,height=lh,mask="auto")
        canvas.setFillColor(navy);canvas.setFont("FiraBold",13)
        canvas.drawString(199,height-26,"Alvállalkozói jutalékelszámolás")
        canvas.setFont("FiraBold",10)
        canvas.drawString(199,height-44,plan["contact"])
        canvas.setFont("Fira",9)
        canvas.drawString(199,height-60,"Elszámolási hónap: "+plan["month"]+"   |   Jutalék: 10%")
        canvas.setFont("FiraBold",8)
        canvas.drawRightString(width-18,height-25,plan["status"])
        canvas.setStrokeColor(navy);canvas.setLineWidth(1.5)
        canvas.line(18,height-76,width-18,height-76)
        canvas.setFillColor(navy);canvas.setFont("Fira",7)
        canvas.drawString(18,12,"Bábolna Sped Kft. | Készült: "+plan["created"][:19])
        canvas.drawRightString(width-18,12,str(doc.page)+". oldal")
    doc.build(story,onFirstPage=page,onLaterPages=page)
    return out.getvalue()
