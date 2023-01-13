import os
import sys
import easygui as E
import xlrd
import openpyxl
import datetime
import calendar
import re
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, Protection
num_in_each_room = [0 for i in range(31)]
col_index = [0 for i in range(31)]
file_path = ''
wb = None
date = datetime.datetime.now().date()

def data_collect(book):
    type_list = []
    name_list = []
    sh=book.sheet_by_index(0)
    for i in range(sh.nrows):
        type_list.append(sh.cell_value(i,0))
    for i in range(len(type_list)):
        j=1
        tmp=[]
        for j in range(1,sh.ncols):
            if sh.cell_type(i,j)==0:
                continue
            tmp.append(sh.cell_value(i,j))
        name_list.append(tmp)
    return (type_list,name_list)

def del_lines(ws):
    #删除原表中多余信息
    # ws.page_setup.fitToHeight = 1
    ws.unmerge_cells('A1:Q1')
    ws.delete_cols(14)#巡回护士
    ws.delete_cols(13)#洗手护士
    ws.delete_cols(12)#手术助手
    ws.delete_cols(10)#备注
    for i in range(1, ws.max_row+1):
        ws.row_dimensions[i].height = 0
    for row in range(ws.max_row,2,-1):#最后一行开始删除局麻
        type = ws.cell(row = row, column = ws.max_column).value
        # print(type)
        if type=='局麻' or type == '局部麻醉':
            ws.delete_rows(row)
    colors=['E2EFDA','FFF2CC','E6E6FA']
    vis_name = set()
    for index,row in enumerate(ws.rows):
        if index <= 1:
            continue
        row[10].font=Font('仿宋',14,True)
        row[11].font=Font('仿宋',14,True)
        row[10].alignment=Alignment('center','center')
        row[11].alignment=Alignment('center','center')
        room_name=row[0].value
        vis_name.add(room_name);
        row[0].fill=PatternFill('solid',fgColor=colors[len(vis_name)%2])
        # flag=0
        # for i,each in enumerate(room_partition):
        #     if room_name in each:
        #         flag=1
        #         for cell in row:
        #             cell.fill=PatternFill('solid',fgColor=colors[i%2])
        #         break
        # if flag==0:
        #     for cell in row:
        #         cell.fill=PatternFill('solid',fgColor=colors[2])

    string = ws['A1'].value
    string += '  麻醉总数:' + str(ws.max_row-2)
    ws['A1'] = string
    ws.merge_cells('A1:M1')

    ws.column_dimensions['K'].width=20
    ws.column_dimensions['L'].width=20

    for i in range(1, ws.max_row+1):
        ws.row_dimensions[i].height = 28

# def relieve():
#     pass

def avai_nextday(ws):
    global date
    type_list,name_list = data_collect(xlrd.open_workbook('data.xls',formatting_info=True))
    doc_dict = dict(zip(type_list, name_list))
    # print(doc_dict)
    avai_doc=doc_dict['三线']+doc_dict['二线（白）']+doc_dict['一线']
    avai_ass=doc_dict['学生-1']
    print(avai_doc,avai_ass)
    #去除请假人员
    # date = datetime.datetime.now().date()
    xls_name = str(date.year)+'年'+str(date.month)+'月值班表.xls'
    if not os.path.isfile('./xls/'+xls_name):
        E.msgbox(msg='找不到当月排班表！')
        return
    sh0=xlrd.open_workbook('./xls/'+xls_name).sheet_by_index(1)
    date_list = calendar.monthcalendar(date.year,date.month)
    qj = []
    height = 10
    for i in range(len(date_list)):
        for j in range(7):
            if date_list[i][j]==date.day:#请假
                for k in range(1,10):
                    qj.append(sh0.cell_value(i*(height+1)+2+k,j))
    print(qj)
    avai_doc=list(set(avai_doc)-set(qj))
    avai_ass=list(set(avai_ass)-set(qj))
    #去除下夜班人员
    xyb,zb,bb = [],[],[]
    sh1=xlrd.open_workbook('./xls/'+xls_name).sheet_by_index(0)
    for i in range(len(date_list)):
        for j in range(7):
            if date_list[i][j]==date.day-1:#下夜班
                for k in range(len(type_list)):
                    xyb.append(sh1.cell_value(i%5*(len(type_list)+1)+k+1+1+1,j+1))
            if date_list[i][j]==date.day:#值班
                for k in range(len(type_list)):
                    zb.append(sh1.cell_value(i%5*(len(type_list)+1)+k+1+1+1,j+1))
            if date_list[i][j]==date.day+2:#备班
                for k in range(len(type_list)):
                    bb.append(sh1.cell_value(i%5*(len(type_list)+1)+k+1+1+1,j+1))
    print(xyb)
    avai_doc=list(set(avai_doc)-set(xyb))
    avai_ass=list(set(avai_ass)-set(xyb))
    print(avai_doc,avai_ass)
    ws['A1']=date.strftime('%Y/%m/%d')
    data = {'请假':qj,'下夜班':xyb,'值班':zb,'备班':bb,'可安排医师':avai_doc,'可安排助手':avai_ass}
    for i,item in enumerate(data.keys()):
        ws.cell(2,i+1).value=item
    for i,l in enumerate(data.values()):
        for j in range(len(l)):
            ws.cell(3+j,i+1).value=l[j]
            if i>3:
                if l[j] in zb:
                    ws.cell(3+j,i+1).font = Font(color='FF0000')
                if l[j] in bb:
                    ws.cell(3+j,i+1).font = Font(color='006400')

def open_file():
    global file_path,wb,date
    file_path = E.fileopenbox(title='打开文件',msg='打开要填写的表格',default='*.xlsx')
    if file_path==None:
        sys.exit()
    date = datetime.datetime.strptime(file_path[-16:-6],'%Y-%m-%d').date()
    print(date)
    wb = openpyxl.load_workbook(file_path)

def save_file():
    global file_path,wb
    new_str=file_path[:-5] + '(new).xlsx'
    flag=1
    while flag:
        try:
            wb.save(new_str)
            flag=0
        except:
            E.msgbox(msg='请先关闭用Excel打开的已生成的表格！')
    os.startfile(new_str)

def main():
    global wb
    # avai_nextday()
    open_file()
    del_lines(wb.active)
    avai_nextday(wb.create_sheet('排班'))
    save_file()
    # #init
    # data_path = E.fileopenbox(title='打开文件',msg='请选择数据文件',default='*.xls')
    # book = xlrd.open_workbook(data_path)
    # assistant = Assistant.import_all(book.sheet_by_index(2))
    # doctor    =    Doctor.import_all(book.sheet_by_index(1))

    # filling_blanks(assistant,doctor)

if(__name__=='__main__'):
    main()