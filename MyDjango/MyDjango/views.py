from django.shortcuts import render
from django.http import HttpResponse
from . import excel
from openpyxl import Workbook

def index(request):
    return render(request,'excel_index.html')

def upload(request):
    return render(request, 'upload.html')

def write_excel(data_list):
    wb = Workbook()
    ws = wb.active
    for row in data_list:
        ws.append(row)
    return wb

def deal_excel(request):
    #https://www.cnblogs.com/yoyo008/p/9232805.html
    files = request.FILES.getlist('myfiles')
    if not files:
        return render(request, 'upload.html')
    columns=['车型','VIN','测试者','测试日期','后外倾角左','后外倾角右','前外倾角左','前外倾角右']
    data_list=[]
    data_list.append(columns)

    for file in files:
        data=excel.operate_excel(file)
        data_list+=data

    wb=write_excel(data_list)
    response = HttpResponse(content_type='application/msexcel')
    response['Content-Disposition'] = 'attachment;filename=result.xlsx'
    wb.save(response)
    return response




