import xlrd
import xlwt
import xlutils
import calendar
import easygui as E
import os
import sys

week_name = ['星期一','星期二','星期三','星期四','星期五','星期六','星期日']
type_list = []
name_list = []
list_p = []
# 输入指定年月
yy = E.integerbox(msg="输入年份: (2000~9999)",title='读取信息',lowerbound=2000,upperbound=9999)
if not yy:
    exit()
mm = E.integerbox(msg="输入月份: (1~12)",title='读取信息',lowerbound=1,upperbound=12)
if not mm:
    exit()
date_list = calendar.monthcalendar(yy,mm)
WPSid=345470524
wps_path = os.environ['USERPROFILE']+'\\Documents\\WPSDrive\\'+str(WPSid)+'\\WPS云盘\\共享文~1\\群共享文件\\'

alignment = xlwt.Alignment()
alignment.horz = 0x02
alignment.vert = 0x01
font0 = xlwt.Font()
font0.name = u'仿宋'
font0.height = 20*12
borders = xlwt.Borders()
borders.bottom=1
borders.left=1
borders.right=1
borders.top=1
style0=xlwt.XFStyle()
style0.font=font0
style0.alignment=alignment
style0.borders=borders

def data_collect(book):
    sh=book.sheet_by_index(0)
    for i in range(sh.nrows):
        type_list.append(sh.cell_value(i,0))
    for i in range(len(type_list)):
        j=1
        tmp=[]
        found=0
        for j in range(1,sh.ncols):
            if sh.cell_type(i,j)==0:
                continue
            tmp.append(sh.cell_value(i,j))
            color=book.xf_list[sh.cell_xf_index(i,j)].background.pattern_colour_index
            if color!=64:
                if found:
                    E.msgbox(msg=str(i+1)+' 行有多个开始的人!请重新检查data.xls',title='错误')
                    exit()
                list_p.append(j-1)
                found=1
        if not found:
            print(type_list)
            E.msgbox(msg=str(i+1)+' 行没有设置从谁开始轮班!请重新检查data.xls',title='错误')
            exit()
        name_list.append(tmp)

def print_struct(sh):
    sh.write(1,0,'WEEK' , style0)
    for j in range(7):
        sh.write(1,j+1,week_name[j] , style0)
    for i in range(len(date_list)):
        if i<=4:
            sh.write(i*(len(type_list)+1)+2,0,'DAY' , style0)

        for j in range(7):
            if date_list[i][j]:
                if i<=4:
                    sh.write(i*(len(type_list)+1)+2,j+1,date_list[i][j] , style0)
                else:
                    sh.write(i%5*(len(type_list)+1)+2,j+1,date_list[i][j] , style0)
            elif(not i):
                if i<=4:
                    sh.write(i*(len(type_list)+1)+2,j+1,'',style0)
        for k in range(len(type_list)):
            if i<=4:
                sh.write(i*(len(type_list)+1)+k+1+2,0,type_list[k] , style0)


def suffle(l):
    tmp=name_list[l][0]
    name_list[l]=name_list[l][1:]
    name_list[l].append(tmp)

def print_names(sh):
#    print("输入上月最后一天值班情况:\n")
#    for i in range(len(type_list)):
#        list_p.append(int(input(type_list[i]+':')))
    day_save=[]
    day_save.append([])
    for i in range(len(date_list)):
        for j in range(7):
            if date_list[i][j]:
                tmp=[]
                for k in range(len(type_list)):
                    tmp.append(name_list[k][list_p[k]])
                    sh.write(i%5*(len(type_list)+1)+k+1+1+1,j+1,name_list[k][list_p[k]] , style0)
                    list_p[k]+=1
                    list_p[k]%=len(name_list[k])
                    if(j==6 and len(name_list[k])==7):
                        suffle(k)
                day_save.append(tmp)
            elif(not i):
                for k in range(len(type_list)):
                    sh.write(i%5*(len(type_list)+1)+k+1+1+1,j+1,'',style0)
    print(day_save)

def write_Title(sh):
    font=xlwt.Font()
    font.height = 20*18
    font.name=u'仿宋'
    font.bold=True
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    style=xlwt.XFStyle()
    style.font = font
    style.pattern = pattern
    style.alignment = alignment
    title='麻醉科'+str(yy)+'年'+str(mm)+'月排班表'
    sh.write_merge(0,0,0,7,title,style)

def print_blanks(sh,height):
    for j in range(7):
        sh.write(1,j,week_name[j])
    for i in range(len(date_list)):

        for j in range(7):
            if date_list[i][j]:
                sh.write(i*(height+1)+2,j,date_list[i][j])

def save_file(wb,file_path):
    flag=1
    dir_name=os.path.dirname(wps_path + file_path)
    file_name=os.path.basename(wps_path + file_path)
    if os.path.isfile(wps_path + file_path):
        if not E.ccbox(msg='当前路径：'+dir_name+'\n已有同名表格:'+file_name+'，是否覆盖？',choices=('是','否')):
            return
    while flag:
        try:
            wb.save(wps_path + file_path)
            flag=0
        except PermissionError as e:
            r=E.msgbox(msg=repr(e)+'\n\n文件可能正在被其他应用或其他人占用！')
            if r is None:
                return
        except FileNotFoundError as e:
            r=E.msgbox(msg=repr(e)+'\n\n路径错误！请检查路径选择在“群共享文件”一级')
            return
        except Exception as e:
            r=E.msgbox(msg=repr(e))
            return
    os.startfile(wps_path + file_path)


def main():
    global wps_path
    calendar.setfirstweekday(firstweekday=1)
    if not os.path.isfile('data.xls'):
        E.msgbox(msg='没有准备数据库文件！（data.xls）')
        return
    data_collect(xlrd.open_workbook('data.xls',formatting_info=True))
    wps_path = E.diropenbox(title='打开存储文件夹',msg='打开表格存储位置')
    print(wps_path)
    if wps_path==None:
        sys.exit()
    book=xlwt.Workbook()
    sheet=book.add_sheet('排程',cell_overwrite_ok=True)
    for i in range(8):
        sheet.col(i).width=256*12
    sheet.row(0).height_mismatch = True
    sheet.row(0).height = 20*28
    write_Title(sheet)
    print_struct(sheet)
    print_names(sheet)
    book2=xlwt.Workbook()
    sheet2=book2.add_sheet('请假',cell_overwrite_ok=True)
    style1=xlwt.XFStyle()
    style1.alignment=alignment
    sheet2.write_merge(0,0,0,6,'麻醉科'+str(yy)+'年'+str(mm)+'月请假表',style1)
    print_blanks(sheet2,10)
    save_file(book,'\\'+str(yy)+'年值班表\\'+str(yy)+'年'+str(mm)+'月值班表.xls')
    save_file(book2,'\\'+str(yy)+'年请假表\\'+str(yy)+'年'+str(mm)+'月请假表.xls')
if __name__=='__main__':
    main()

