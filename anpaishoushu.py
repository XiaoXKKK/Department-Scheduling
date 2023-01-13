import os
import sys
import easygui as E
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, Protection
date_str = ''
num_in_each_room = [0 for i in range(31)]
col_index = [0 for i in range(31)]

# room_partition=[['02间','03间','03A间'],
#                 ['05间','06间'],
#                 ['07间','08间','09间'],
#                 ['10间','11间','12间'],
#                 ['13间','13A间','15间','16间'],
#                 ['17间','18间','19间'],
#                 ['20间','21间'],
#                 ['22间','23间','23A'],
#                 ['25间','26间','27间'],
#                 ['28间','29间'],
#                 ['30间']]
# 2-3a 5-6
# 7-9,17-19
# 13-16,10-12
# 20 21,22,28,29
# 23,25,26,27,30

# class Assistant:

#     AssCount = 0

#     @classmethod
#     def import_all(sh) -> list:
#         all_assistant=[]
#         for i in range(1,sh.nrows):
#             if sh.cell_type(i,1)==0 :
#                 continue
#             tmp=[]
#             for j in range(1,5):
#                 tmp.append(sh.cell_value(i,j))
#                 ass=Assistant(tmp)
#                 print(ass)
#                 all_assistant.append(ass)
#         all_assistant.sort(reverse=True)
#         return all_assistant

#     def __init__(self,init_list) -> None:
#         self.name=init_list[0]
#         self.priority_num=init_list[1]
#         self.mentor=init_list[2]
#         self.section=init_list[3]
#         Assistant.AssCount += 1

#     def __eq__(self, __o: object) -> bool:
#         return self.priority_num == __o.priority_num

#     def __gt__(self, __o: object) -> bool:
#         return self.priority_num > __o.priority_num

#     def __lt__(self, __o: object) -> bool:
#         return self.priority_num < __o.priority_num

# class Doctor:

#     DocCount = 0

#     @classmethod
#     def import_all(sh) -> list:
#         all_doctor=[]
#         for i in range(1,sh.nrows):
#             if sh.cell_type(i,1)==0 :
#                 continue
#             tmp=[]
#             for j in range(1,4):
#                 tmp.append(sh.cell_value(i,j))
#                 doc=Doctor(tmp)
#                 print(doc)
#                 all_doctor.append(doc)
#         all_doctor.sort(reverse=True)
#         return all_doctor

#     def __init__(self,init_list) -> None:
#         self.name=init_list[0]
#         self.priority_num=init_list[1]
#         self.section=init_list[2]
#         Doctor.DocCount += 1

#     def __eq__(self, __o: object) -> bool:
#         return self.priority_num == __o.priority_num

#     def __gt__(self, __o: object) -> bool:
#         return self.priority_num > __o.priority_num

#     def __lt__(self, __o: object) -> bool:
#         return self.priority_num < __o.priority_num

# class Room:

#     last_now_num = 1

#     def __init__(self,name,num,ward) -> None:
#         name = name
#         row_num = num
#         ward = ward
#         Surgery_num = Room.last_now_num-row_num
#         Room.last_now_num = row_num

#     def __eq__(self, __o: object) -> bool:
#         return self.Surgery_num == __o.Surgery_num

#     def __gt__(self, __o: object) -> bool:
#         return self.Surgery_num > __o.Surgery_num

#     def __lt__(self, __o: object) -> bool:
#         return self.Surgery_num < __o.Surgery_num

# def change_prior():
#     pass

# def filling_blanks(assistant,doctor):
#     file_path = E.fileopenbox(title='打开文件',msg='打开要填写的表格',default='*.xlsx')
#     schedule = xlrd.open_workbook(file_path)
#     date_str = file_path[-15:-5]
#     y = int(date_str[:4])
#     m = int(date_str[5:7])
#     d = int(date_str[-2:])
#     sh = schedule.sheet_by_index(0)
#     wb = xlutils.copy.copy(schedule)
#     sheet1 = wb.get_sheet(0)
#     #read
#     #(name,row_id) -> list
#     last = None
#     rooms = []
#     for i in range (2,sh.nrows):
#         if sh.cell_type(i,0)==0 :
#             break
#         now = sh.cell_value(i,0)
#         if now!=last:
#             ward=sh.cell_value(i,3)
#             rooms.append(Room(now,i,ward))
#             last = now
#     #doctor
#     special  = [doc for doc in doctor if doc.section]
#     for person in special:
#         for room in rooms:
#             if person.section == room.section:
#                 sheet1.write(room.row_num, 14, doctor.name)
#                 doctor.remove(person)
#                 rooms.remove(room)
#     for each in doctor:
#         for room in rooms:
#             pass

#     #assistant


#     #write


#     wb.save(file_path)

def del_lines():
    #删除原表中多余信息
    file_path = E.fileopenbox(title='打开文件',msg='打开要填写的表格',default='*.xlsx')
    if file_path==None:
        sys.exit()
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active
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

    new_str=file_path[:-5] + '(new).xlsx'
    wb.save(new_str)
    os.startfile(new_str)

# def relieve():
#     pass

def main():
    del_lines()
    # #init
    # data_path = E.fileopenbox(title='打开文件',msg='请选择数据文件',default='*.xls')
    # book = xlrd.open_workbook(data_path)
    # assistant = Assistant.import_all(book.sheet_by_index(2))
    # doctor    =    Doctor.import_all(book.sheet_by_index(1))

    # filling_blanks(assistant,doctor)

if(__name__=='__main__'):
    main()