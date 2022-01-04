from pywebio import *
from pywebio.utils import pyinstaller_datas
from pywebio.input import *
from pywebio.output import *
from pywebio.session import *
from pywebio.pin import *
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment
import warnings
from pathlib import Path
import argparse
import io
from pywebio.exceptions import SessionClosedException
from pywebio import start_server
from flask import Flask, send_from_directory, send_file, make_response, request
import flask.ext.excel
from pywebio import STATIC_PATH
from pywebio.platform.flask import webio_view
from django.http import FileResponse


DownloadFolder = str(Path.home() / "Downloads")
app = Flask(__name_)
@app.route("/download", methods=['GET'])

def main():
    data = input_group("Thesis/Design Information",[
        input('Title of Thesis/Design', name='title'),
        checkbox('Advisor', options=['Caya, Meo Vincent ', 'Hortinela, Carlos IV', 'Linsangan, Noel', 'Manlises, Cyrel', 'Maramba, Rafael', 'Pellegrino, Rosemarie', 'Torres, Jumelyn', 'Villaverde, Jocelyn', 'Yumang, Analyn'], name='advisor'),
        checkbox('Panels', options=['Caya, Meo Vincent ', 'Hortinela, Carlos IV', 'Linsangan, Noel', 'Manlises, Cyrel', 'Maramba, Rafael', 'Pellegrino, Rosemarie', 'Torres, Jumelyn', 'Villaverde, Jocelyn', 'Yumang, Analyn'], name='panels'),
        input("Payor Name", name='payor'),
        input("Reference Number", name='refnum'),
    ])
    
    if not data['title']:
        popup('Error', [
            put_text('Missing Title')])
        main()
    elif not data['advisor']:
        popup('Error', [
            put_text('Missing Advisor')])
        main()
    elif not data['panels']:
        popup('Error', [
            put_text('Missing Panels')])
        main()
    elif not data['payor']:
        popup('Error', [
            put_text('Missing Payor')])
        main()
    elif not data['refnum']:
        popup('Error', [
            put_text('Missing Reference Number')])
        main()

    Title = data['title']
    Advisor = data['advisor']
    Panels = data['panels']
    Payor = data['payor']
    Ref = data['refnum']
    column = len(Advisor) + len(Panels) + 4
    Y = ["A", 'B', 'C', 'D', 'E', 'F']
    fname = Title + ".xlsx"
    
    
    wb =  Workbook()
    ws = wb.active
    ws.title = Title

    ws.merge_cells(start_row=1, start_column=1, end_row=column, end_column=1)
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.column_dimensions['A'].width = 4
    ws['A1'] = 1


    ws.merge_cells('B1:B2')
    ws['B1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.column_dimensions['B'].width = 20
    ws['B1']= "FACULTY NAME"
    x = 3
    for adv in Advisor:
        ws['B'+ str(x)]=adv
        ws['C'+ str(x)]="Advisor"
        ws['E'+ str(x)]= "PHP 3,000"
        ws['F'+ str(x)]= "ATM"
        x+=1
    y = len(Advisor) + 4
    for pan in Panels:
        ws['B'+ str(y)]=pan
        ws['C'+ str(y)]="Panel"
        ws['E'+ str(y)]="PHP 1,500"
        ws['F'+ str(y)]="ATM"
        y+=1

    ws.merge_cells('C1:C2')
    counterC=1
    while counterC!=column:
        ws['C'+str(counterC)].alignment = Alignment(vertical='center', horizontal='center', wrapText=True)
        counterC+=1
    ws.column_dimensions['C'].width = 10
    ws['C1']= "ADVISOR or PANEL"

    

    ws.merge_cells('D1:D2')
    counterD=1
    while counterD<=column:
        ws['D'+str(counterD)].alignment = Alignment(vertical='center', horizontal='center', wrapText=True)
        counterD+=1
    ws.merge_cells(start_row=3, start_column=4, end_row=column-2, end_column=4)
    ws.column_dimensions['D'].width = 30
    ws['D1'] = 'TITLE OF THESIS/DESIGN'
    ws['D3'] = Title
    ws['D' + str(column-1)] = Payor
    ws['D' + str(column)] = Ref

    ws.merge_cells('E1:E2')
    counterE=1
    while counterE<column:
        ws['E'+str(counterE)].alignment = Alignment(vertical='center', horizontal='center')
        counterE+=1
    ws.column_dimensions['E'].width = len("HONORARIUM")+5
    ws['E1']="HONORARIUM"

    ws.merge_cells('F1:F2')
    counterF=1
    while counterF<column:
        ws['F'+str(counterF)].alignment = Alignment(vertical='center', horizontal='center', wrapText=True)
        counterF+=1
    ws.column_dimensions['F'].width = 12
    ws['F1']="MODE OF PAYMENT"

    for C in Y:
        for R in range(column):
            ws[str(C)+str(R+1)].font=Font(name="Arial")
            ws[str(C)+str(R+1)].border=Border(left=Side(border_style='thin',color='00000000'),right=Side(border_style='thin',color='00000000'), top=Side(border_style='thin',color='00000000'), bottom=Side(border_style='thin',color='00000000'))
    
   
    #output = wb.save(file_name)
    #buffer = io.BytesIO()
    output = wb.save(fname.strip())
    
    return excel.make_response_from_array([[1, 2], [3, 4]], "csv")
    #buffer.seek(0)
    #return FileResponse(buffer, as_attachment=True, filename=fname)
    #return output
    #return(data['title'],data['advisor'],data['panels'],data['payor'],data['refnum'])

def process():
    main()
    termi = True
    while termi:
        put_text('Success')
        keep_main = radio('Continue?', ['Yes', 'No'])
        if keep_main == 'Yes':
            main()
        else:
            put_text('Thank you!')
            termi = False
            
if __name__ == '__main__':
    app.run()
    parser = argparse.ArgumentParser()
    parser.add_argument("-p", "--port", type=int, default=8080)
    args = parser.parse_args()

    start_server(process, port=args.port)
