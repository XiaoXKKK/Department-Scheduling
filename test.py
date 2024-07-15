from tkinter import *
 
master = Tk()
 
frame = Frame(master)
frame.pack(padx=10,pady=10)
 
v1 = StringVar()
v2 = StringVar()
v3 = StringVar()
 
'''
注意在这里我们不能使用entry控件的get()方法来获取输入的内容
因为当validate选项指定为key的时候，有任何的输入操作都会被拦截
到这个函数当中，也就是说先拦截，只有这个函数返回True，那么输入的内容
才会到变量里面去，也就是说我们使用get函数并无法get到数据，get函数在这个
函数之后才会有效，get函数得到的是变量的值。
所以只有使用%P来获得最新的输入框的内容
'''
def test(content):
     if len(content)==0 or content.isdigit():
          return True
     else:
          return False
#使用了特殊技能的函数需要使用register将其封装起来才可以
testCMD = master.register(test)
e1 = Entry(frame,width=10,validate='key',\
           validatecommand=(testCMD,'%P')).grid(row=0,column=0)
Label(frame,text='+').grid(row=0,column=1)
 
e2 = Entry(frame,width=10,textvariable=v2,validate='key',\
           validatecommand=(testCMD,'%P')).grid(row=0,column=2)
Label(frame,text='=').grid(row=0,column=3)
#e3输入框是显示结果文本框，所以不允许修改里面的值。将state属性设置为state='readonly'只读
e3 = Entry(frame,width=10,textvariable=v3,state='readonly').grid(row=0,column=4)
 
def calc():
     result=int(v1.get())+int(v2.get())
     v3.set(str(result))
Button(frame,text='计算结果',command=calc).grid(row=1,column=2,padx=5)
 
mainloop()