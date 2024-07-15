import tkinter as tk
from tkinter import ttk,filedialog

root = tk.Tk()
root.resizable(False, False)
root.geometry('400x300')
root.title("手术排程工具")
# 设置窗口图标，确保 icon.ico 文件在脚本运行的目录下
root.iconbitmap("icon.ico")

notebook = ttk.Notebook(root)
notebook.pack(expand=1, fill="x")

# 创建第一个选项卡
tab1 = ttk.Frame(notebook, padding="3 3 12 12")
notebook.add(tab1, text='排班助手')

# 验证函数
def num_validated(d):
    return d.isdigit() or len(d) == 0
# 注册验证函数
numCMD = root.register(num_validated)
# 添加内容
tk.Label(tab1, text='年份', bg='lightblue', fg='red').grid(row=0, column=0, padx=(0, 10), pady=(0, 5), sticky='e')
yearEntry = tk.Entry(tab1, validate='key', validatecommand=(numCMD, '%P'))
yearEntry.grid(row=0, column=1, padx=(0, 10), pady=(0, 5))

tk.Label(tab1, text='月份', bg='lightblue', fg='red').grid(row=1, column=0, padx=(0, 10), pady=(0, 5), sticky='e')
monthEntry = tk.Entry(tab1, validate='key', validatecommand=(numCMD, '%P'))
monthEntry.grid(row=1, column=1, padx=(0, 10), pady=(0, 5))

genButton = tk.Button(tab1, text="生成排班表")
genButton.grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

# 创建第二个选项卡，使用不同的变量名
tab2 = ttk.Frame(notebook, padding="3 3 12 12")
notebook.add(tab2, text='手术排程')

def FileOpen():
    r = filedialog.askopenfilename(title='打开要填写的表格',
                                           filetypes=[('Excel', '*.xls'), ('All files', '*')])
    print(r)

openButton = tk.Button(tab2, text="打开表格", command=FileOpen)
openButton.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")




root.mainloop()