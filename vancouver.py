import sys
import pandas as pd
import sqlite3
from PyQt5.QtWidgets import *
from PyQt5 import uic
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import BaseDocTemplate, Table, TableStyle, Paragraph, Frame, PageTemplate
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont



pdfmetrics.registerFont(TTFont('Arial', 'Arial.ttf'))
pdfmetrics.registerFont(TTFont('ArialBd', 'ArialBd.ttf'))





class VanInvoice(QWidget):
    def __init__(self):
        super(VanInvoice, self).__init__()
        uic.loadUi("UI/vaninv.ui",self)
        self.show()
        self.selectbtn.clicked.connect(self.getData)
        self.pdfbtn.clicked.connect(self.toPDF)
        #load excel file
        conn = sqlite3.connect(':memory:')
        cur = conn.cursor()
        df = pd.read_excel("Data/OHEAD.xlsx",sheet_name="OHEAD")
        df.to_sql(name='OHEAD', con=conn, if_exists='append', index=False)
        cur.execute("SELECT OINO FROM OHEAD")
        inv = cur.fetchall()
        #final = pd.DataFrame(inv)

        #list1 = [list(row) for row in inv]
        #print(list)
        #completer = QCompleter(list)
        #self.lineEdit.setCompleter(completer)





    def getData(self):
        try:
            global invdate, salesman,addedby, invnum, qoutenum, shipping,get,total
            conn = sqlite3.connect(':memory:')
            cur = conn.cursor()
            userinput = self.lineEdit.text()

            # get header info
            # load excel file
            df = pd.read_excel("Data/OHEAD.xlsx", sheet_name="OHEAD")
            df.to_sql(name='OHEAD', con=conn, if_exists='append')
            cur.execute("SELECT OIDATE,OSLSMAN,ADDEDBY,OINO,OQNO FROM OHEAD WHERE OINO =?", (userinput,))
            result = cur.fetchall()
            invdate = result[0][0]
            salesman = result[0][1]
            addedby = result[0][2]
            invnum = result[0][3]
            qoutenum = result[0][4]


            # get sold to & ship to
            cur.execute("SELECT ONAME || ' ' || OFNAME,OADDR1, OADDR2,OADDR3,OPCODE,OBPHONE,OPSTNO,OSTREET,OINAME,OIADDR1,OIADDR2,OIADDR3,OIPCODE FROM OHEAD WHERE OINO =?",(userinput,))
            shipping = cur.fetchall()

            #get inv details
            df1 = pd.read_excel("Data/ODETL.xlsx", sheet_name="ODETL")
            df1.to_sql(name='ODETL', con=conn, if_exists='append')
            cur.execute("SELECT OD_ITEM, "
                        "CASE WHEN round(OD_QTY,2) = 0 AND round(OD_PRICE,2) = 0 AND round(OD_AMOUNT,2) = 0 THEN ''"
                        "ELSE round(OD_QTY)"
                        "END, "
                        "COALESCE(OD_DESCR, ' ') as descr,"
                        "CASE WHEN round(OD_QTY,2) = 0 AND round(OD_PRICE,2) = 0 AND round(OD_AMOUNT,2) = 0 THEN '' "
                        "ELSE round(OD_PRICE,2)"
                        "END, "
                        "CASE WHEN round(OD_QTY,2) = 0 AND round(OD_PRICE,2) = 0 AND round(OD_AMOUNT,2) = 0 THEN '' "
                        "ELSE round(OD_AMOUNT,2) "
                        "END "
                       "FROM ODETL "
                        "INNER JOIN OHEAD ON ODETL.OD_UNO = OHEAD.ONUMBER "
                        "WHERE OHEAD.OINO = ?",(userinput,))
            table = cur.fetchall()

            self.tableWidget.setRowCount(0)
            for row_number, row_data in enumerate(table):
                self.tableWidget.insertRow(row_number)
                for column_number, data in enumerate(row_data):
                    self.tableWidget.setItem(row_number, column_number, QTableWidgetItem(str(data)))

            #get gst and total
            cur.execute("SELECT sum(round(OD_AMOUNT,2)), sum(round(OHEAD.OGST,2)) FROM ODETL INNER JOIN OHEAD ON ODETL.OD_UNO = OHEAD.ONUMBER WHERE OHEAD.OINO = ?",(invnum,))
            get = cur.fetchall()
            total = get[0][0] + get[0][1]



            QMessageBox.information(self, "Done!", "Data fetched.")
        except Exception as e:
                print(e)
                QMessageBox.information(self, "Invalid Input", "No invoice number matched.")

    def toPDF(self):

        filename, _ = QFileDialog.getSaveFileName(self, 'Save file', '', 'PDF File (*.pdf)')
        if filename != '':
            try:

                style2 = ParagraphStyle(
                    name='Normal',
                    fontName='ArialBd',
                    fontSize=22,
                    alignment=1,
                    spaceAfter=15,
                )
                style3 = ParagraphStyle(
                    name='Normal',
                    fontName='Arial',
                    fontSize=12,
                    alignment=1,
                    spaceBefore=20,
                )

                addstyle = ParagraphStyle(
                    name='Normal',
                    fontName='ArialBd',
                    fontSize=10,
                    alignment=1,
                )

                def header(canvas, pdf):

                    # Draw heading
                    heading = Paragraph("VANCOUVER GLASS (1990) LTD.", style2)
                    heading.wrap(pdf.width, inch * 0.3)
                    heading.drawOn(canvas, pdf.leftMargin, pdf.height + inch)

                    # Draw subheading.
                    subheading = Paragraph("INVOICE", style3)
                    subheading.wrap(pdf.width, inch * 0.2)

                    subheading.drawOn(canvas, pdf.leftMargin, pdf.height + inch * 0.5)
                    canvas.line(1.7 * inch, 2.5 *inch, 1.7 *inch, 7.6 *inch)
                    canvas.line(2.26 * inch, 2.5 * inch, 2.26 * inch, 7.6 * inch)
                    canvas.line(6.41 * inch, 2.5 * inch, 6.41 * inch, 7.6 * inch)
                    canvas.line(7.19 * inch, 2.5 * inch, 7.19 * inch, 7.6 * inch)

                    tablelist1 = [["Date: "+invdate, "Invoice#: "+str(invnum)],
                                  ["Salesman: "+str(salesman)],
                                  ["GST#: 121989834RT","Quote#: "+str(qoutenum)],
                                  [addedby,"Page: "+"%d " % doc.page]
                                  ]
                    tablelist2 = [["Sold To:", "Location: " + str(shipping[0][7])],
                                  [str(shipping[0][0]), "Ship To:"],
                                  [str(shipping[0][1]), str(shipping[0][8])],
                                  [str(shipping[0][2]), str(shipping[0][9])],
                                  [str(shipping[0][3]), str(shipping[0][10])],
                                  [str(shipping[0][4]), str(shipping[0][11])],
                                  ["Phone: " + str(shipping[0][5]), str(shipping[0][12])],
                                  ["PST Exempt# " + str(shipping[0][6])]]

                    footerlist = [["", "Sub-Total:", get[0][0]],
                                 ["Overdue accounts will be charged 2% per month.", "GST", round(get[0][1],2)],
                                 ["Please enclose a copy of the invoice with the cheque."],
                                 [""],
                                 [""],
                                 ["Charge to Acount", "Total Amount:", round(total,2)]]



                    tablestyle1 = TableStyle([

                        ('FONTSIZE', (0, 0), (-1, -1), 10),
                        ('FONTNAME', (0, 0), (-1, -1), 'ArialBd'),

                    ])

                    tablestyle2 = TableStyle([

                        ('FONTSIZE', (0, 0), (-1, -1), 10),
                        ('FONTNAME', (0, 0), (-1, -1), 'ArialBd'),
                        ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
                        ('LINEBEFORE', (1, 0), (1, -1), 0.5, colors.black)

                    ])

                    footerstyle = TableStyle([

                        ('FONTSIZE', (0, 0), (-1, -1), 8),
                        ('FONTSIZE', (1, 0), (2, -1), 9),
                        ('FONTNAME', (0, 0), (-1, -1), 'ArialBd'),
                        ('LINEBEFORE', (0, 0), (0, 4), 1, colors.black),
                        ('LINEBEFORE', (2, 0), (-1, -1), 1, colors.black),
                        ('LINEAFTER', (2, 0), (-1, -1), 1, colors.black),
                        ('LINEABOVE', (0, 5), (-1, 5), 1, colors.black),
                        ('LINEBELOW', (2, 5), (-1, 5), 1, colors.black),
                        ('ALIGN', (2, 0), (2, -1), "RIGHT"),
                        ])

                    table1 = Table(tablelist1, colWidths=[400, 200], rowHeights=[10, 10, 10, 10],
                                   hAlign='CENTER', spaceBefore=5, style=tablestyle1)
                    table1.wrap(pdf.width, inch)
                    table1.drawOn(canvas, pdf.leftMargin, pdf.height - inch * 0.4)

                    table2 = Table(tablelist2, colWidths=[268, 268], rowHeights=[20, 15, 10, 10, 10, 10, 10, 20],
                                   hAlign='CENTER', spaceBefore=5, style=tablestyle2)
                    table2.wrap(pdf.width, inch)
                    table2.drawOn(canvas, pdf.leftMargin, pdf.height - inch * 2)
                    footertable = Table(footerlist, colWidths=[410, 71, 56], rowHeights=[20, 10, 10, 10, 10, 20],
                                        hAlign='LEFT', spaceBefore=5, style=footerstyle)
                    footertable.wrap(pdf.width, inch)
                    footertable.drawOn(canvas, pdf.leftMargin, 1.38 * inch)


                    addressnote = Paragraph("1706 E. HASTINGS, VAN, B.C. V5L 1S9 Phone (604)253-7707 Fax (604)253-8448",addstyle)
                    addressnote.wrap(pdf.width, inch)
                    addressnote.drawOn(canvas, pdf.leftMargin, 0.7 * inch)

                doc = BaseDocTemplate(filename, leftMargin=0.5 * inch, rightMargin=0.5 * inch)

                frame = Frame(

                    0.5 * inch,  # x
                    2.5 * inch,  # y at bottom
                    7.45 * inch,  # width
                    5.1 * inch,  # height
                    showBoundary=1
                )

                template = PageTemplate(id='all_pages', frames=frame, onPage=header)
                doc.addPageTemplates([template])


                tableheader = [["Code", "Qty", "Description", "Unit Price", "Amount"]]

                tablelist3 = []
                for row in range(self.tableWidget.rowCount()):
                    df_list2 = []
                    for col in range(self.tableWidget.columnCount()):
                        table_item = self.tableWidget.item(row, col)
                        df_list2.append('' if table_item is None else str(table_item.text()))
                    tablelist3.append(df_list2)


                tablestyle3 = TableStyle([


                    ('FONTSIZE', (0, 0), (-1, -1), 8),
                    ('FONTNAME', (0, 1), (-1, -1), 'Arial'),

                    ('ALIGN',(1,0),(1,-1), "RIGHT"),
                    ('ALIGN', (3, 0), (4, -1), "RIGHT"),


                ])

                tableheaderstyle = TableStyle([

                    ('FONTSIZE', (0, 0), (-1, 0), 9),
                    ('FONTNAME', (0, 0), (-1, 0), 'ArialBd'),
                    ('LINEBELOW', (0, 0), (-1, 0), 0.5, colors.black),


                ])


                table3 = Table(tablelist3,
                               hAlign='LEFT', spaceBefore=5, repeatRows=1, style=tablestyle3,
                               colWidths=[80, 40, 300, 55, 50], rowHeights=10)
                tableheader1 = Table(tableheader,
                                     hAlign='LEFT', spaceBefore=5, repeatRows=1, style=tableheaderstyle,
                                     colWidths=[80, 40, 300, 55, 50])


                Elements = []
                Elements.append(tableheader1)
                Elements.append(table3)
                doc.build(Elements)

                QMessageBox.information(self, "Done!", "File Exported.")

            except Exception as e:
                print(e)
                QMessageBox.information(self, "Error", "Failed to run script.")


app = QApplication(sys.argv)
design = VanInvoice()

app.exec_()

