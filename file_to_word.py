#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@文件        :file_to_word.py
@说明        :Word模板生成管理文件
@时间        :2023/06/28 10:58:46
@作者        :跳跃的🐸
@版本        :2.0
'''


import ttkbootstrap as tk
from tkinter import messagebox
from ttkbootstrap.tooltip import ToolTip
import webbrowser
import datetime
from seatable_api import Base, context
from mailmerge import MailMerge
from tkinter.filedialog import askopenfilename
import mistune
from bs4 import BeautifulSoup
import time

# 初始化seatable-api连接数据
server_url = context.server_url or 'https://seatable.10an.fun'
api_token = context.api_token or 'f5184823fc756daa0d84edbcac4f97ab51734a57'
base = Base(api_token, server_url)
base.auth()

# 初始化GUI
root = tk.Window(themename="minty",iconphoto=None)
root.iconbitmap('export_word.ico')
root.title("WGWL管理文件word模板生成            作者：跳跃的🐸")
root.resizable(False, False)


screenwidth = root.winfo_screenwidth()  # 屏幕宽度
screenheight = root.winfo_screenheight()  # 屏幕高度
width = 1000
height = 800
x = int((screenwidth - width) / 2)
y = int((screenheight - height) / 2)
root.geometry('{}x{}+{}+{}'.format(width, height, x, y))  # 大小以及位置

# 创建一个菜单栏
search_frame = tk.Frame(root)
search_frame.pack(side="left", padx=20, fill='y')
# 创建一个结果窗口
tabel_frame = tk.Frame(root)
tabel_frame.pack(padx=20, fill="x")
# 创建一个按钮栏
button_frame = tk.Frame(root)
button_frame.pack(padx=20, fill="x")
# 创建一个信息栏
info_frame = tk.Frame(root)
info_frame.pack(padx=20, fill="x")

# 初始化Word模板文件路径
path = None

# 判断时间格式是否正确
def isVaildDate(date):
        try:
            if date:
                time.strptime(date, "%Y-%m-%d")
            return "日期格式正确！"
        except:
            return "请检查日期是否输入正确！" 

# 创建多个StringVar对象，用于绑定entry框和变量
filenum_var = tk.StringVar()
filenum_var.trace("w", lambda name, index, mode: update_sql())
version_var = tk.StringVar()
version_var.trace("w", lambda name, index, mode: update_sql())
issuer_var = tk.StringVar()
issuer_var.trace("w", lambda name, index, mode: update_sql())
filename_var = tk.StringVar()
filename_var.trace("w", lambda name, index, mode: update_sql())
author_var = tk.StringVar()
author_var.trace("w", lambda name, index, mode: update_sql())
author_date_var = tk.StringVar()
author_date_var.trace("w", lambda name, index, mode: update_sql())
reviewer_var = tk.StringVar()
reviewer_var.trace("w", lambda name, index, mode: update_sql())
review_date_var = tk.StringVar()
review_date_var.trace("w", lambda name, index, mode: update_sql())
publish_date_var = tk.StringVar()
publish_date_var.trace("w", lambda name, index, mode: update_sql())
implement_date_var = tk.StringVar()
implement_date_var.trace("w", lambda name, index, mode: update_sql())
content_var = tk.StringVar()
content_var.trace("w", lambda name, index, mode: update_sql())
# 创建多个Entry组件，用于输入搜索条件
filenum_entry = tk.Entry(search_frame, textvariable=filenum_var, width=13)
filenum_entry.grid(row=0, column=1)
version_entry = tk.Entry(search_frame, textvariable=version_var, width=13)
version_entry.grid(row=1, column=1)
issuer_entry = tk.Entry(search_frame, textvariable=issuer_var, width=13)
issuer_entry.grid(row=2, column=1)
filename_entry = tk.Entry(search_frame, textvariable=filename_var, width=13)
filename_entry.grid(row=3, column=1)
ToolTip(filename_entry, text="支持文件名模糊查找!", bootstyle=("info","inverse"))
author_entry = tk.Entry(search_frame, textvariable=author_var, width=13)
author_entry.grid(row=4, column=1)
author_date_entry = tk.Entry(
    search_frame, textvariable=author_date_var, width=13)
author_date_entry.grid(row=5, column=1)
reviewer_entry = tk.Entry(search_frame, textvariable=reviewer_var, width=13)
reviewer_entry.grid(row=6, column=1)
review_date_entry = tk.Entry(
    search_frame, textvariable=review_date_var, width=13)
review_date_entry.grid(row=7, column=1)
publish_date_entry = tk.Entry(
    search_frame, textvariable=publish_date_var, width=13)
publish_date_entry.grid(row=8, column=1)
implement_date_entry = tk.Entry(
    search_frame, textvariable=implement_date_var, width=13)
implement_date_entry.grid(row=9, column=1)
content_entry = tk.Entry(search_frame, textvariable=content_var, width=13)
content_entry.grid(row=10, column=1)
ToolTip(content_entry, text="支持文件内容模糊查找!", bootstyle=("info","inverse"))


# 创建多个Label组件，用于显示搜索条件的名称
filenum_label = tk.Label(search_frame, text="文件编号")
filenum_label.grid(row=0, column=0, sticky="w")
version_label = tk.Label(search_frame, text="版本")
version_label.grid(row=1, column=0, sticky="w")
issuer_label = tk.Label(search_frame, text="签发人")
issuer_label.grid(row=2, column=0, sticky="w")
filename_label = tk.Label(search_frame, text="管理文件名")
filename_label.grid(row=3, column=0, sticky="w")
author_label = tk.Label(search_frame, text="制订人")
author_label.grid(row=4, column=0, sticky="w")
author_date_label = tk.Label(search_frame, text="制订日期")
author_date_label.grid(row=5, column=0, sticky="w")
reviewer_label = tk.Label(search_frame, text="审核人")
reviewer_label.grid(row=6, column=0, sticky="w")
reviewer_date_label = tk.Label(search_frame, text="审核日期")
reviewer_date_label.grid(row=7, column=0, sticky="w")
publish_date_label = tk.Label(search_frame, text="发布日期")
publish_date_label.grid(row=8, column=0, sticky="w")
implement_date_label = tk.Label(search_frame, text="实施日期")
implement_date_label.grid(row=9, column=0, sticky="w")
content_label = tk.Label(search_frame, text="内容包含")
content_label.grid(row=10, column=0, sticky="w")

# 创建一个Label组件放在LabelFrame框内，用于显示SQL语句
sql_info = tk.LabelFrame(search_frame, text='SQL语句', padding=5,bootstyle="primary")
sql_info.grid(row=12, columnspan=2)
sql_label = tk.Label(
    sql_info, text="", wraplength=150)
sql_label.pack()

# 定义一个函数，用于更新SQL语句


def update_sql():
    # 获取搜索条件
    filenum = filenum_var.get()
    version = version_var.get()
    issuer = issuer_var.get()
    filename = filename_var.get()
    author = author_var.get()
    author_date = author_date_var.get()
    reviewer = reviewer_var.get()
    review_date = review_date_var.get()
    publish_date = publish_date_var.get()
    implement_date = implement_date_var.get()
    content = content_var.get()
    # 初始化SQL语句
    sql = 'select filename, version,issuer,file_number,content,author,author_date,reviewer,review_date,publish_date,implement_date from file'
    # 初始化条件列表
    conditions = []
    # 判断条件列表添加条件
    if filenum:
        conditions.append(f"file_number = '{filenum}'")
    if version:
        conditions.append(f"version = '{version}'")
    if issuer:
        conditions.append(f"issuer = '{issuer}'")
    if filename:
        conditions.append(f"filename like '%{filename}%'")
    if author:
        conditions.append(f"author = '{author}'")
    if author_date:
        conditions.append(f"author_date = '{author_date}'")
        ToolTip(author_date_entry, text=isVaildDate(author_date), bootstyle=("danger","inverse"))
    if reviewer:
        conditions.append(f"reviewer = '{reviewer}'")
    if review_date:
        conditions.append(f"review_date = '{review_date}'")
        ToolTip(review_date_entry, text=isVaildDate(review_date), bootstyle=("danger","inverse"))
    if publish_date:
        conditions.append(f"publish_date = '{publish_date}'")
        ToolTip(publish_date_entry, text=isVaildDate(publish_date), bootstyle=("danger","inverse"))
    if implement_date:
        conditions.append(f"implement_date = '{implement_date}'")
        ToolTip(implement_date_entry, text=isVaildDate(implement_date), bootstyle=("danger","inverse"))
    if content:
        conditions.append(f"content like '%{content}%'")
    # 如果条件列表不为空，则在SQL语句中添加WHERE子句，并用AND连接所有的条件
    if conditions:
        sql += " WHERE " + " AND ".join(conditions)
    # 在Label组件中显示SQL语句
    sql_label.config(text=sql)

# 搜索函数


def search():
    global rowCount
    try:
        sql = sql_label.cget("text")
        rows = base.query(sql)
        tree.delete(*tree.get_children())
        for row in rows:
            # if(rowCount %2 ==1):
            tree.insert('', 'end', values=(row['filename'],
                                        row['version'],
                                        row['issuer'],
                                        row['file_number'],
                                        md_word(row['content']),
                                        row['author'],
                                        row['author_date'],
                                        row['reviewer'],
                                        row['review_date'],
                                        row['publish_date'],
                                        row['implement_date']))
    except:
        messagebox.showerror("错误", "获取数据失败，请尝试浏览器打开https://seatable.10an.fun！")
        


# 定义一个函数，用于重置所有条件
def reset_all():
    # 清空所有的entry框的内容
    filenum_var.set("")
    version_var.set("")
    issuer_var.set("")
    filename_var.set("")
    author_var.set("")
    author_date_var.set("")
    reviewer_var.set("")
    review_date_var.set("")
    publish_date_var.set("")
    implement_date_var.set("")
    content_var.set("")
    # 更新SQL语句
    update_sql()


# 创建一个Button组件，用于重置所有条件
reset_button = tk.Button(search_frame, text="重置", command=reset_all)
reset_button.grid(row=11, column=0,pady=5)

# 创建一个Button组件，用于执行SQL语句
sql_button = tk.Button(search_frame, text="执行", command=search,bootstyle="outline")
sql_button.grid(row=11, column=1,pady=5)

# 读取markdown字符串转html再转word
def md_word(str):
    html = mistune.markdown(str)
    soup = BeautifulSoup(html, 'html.parser')
    return soup.get_text()


# 选择要打开的Word模板文件
def open_template():
    global path
    template_path = askopenfilename(filetypes=[("Word模板文件", "*.docx")])
    template_entry.delete(0, 'end')
    template_entry.insert(0, template_path)
    path = template_path


# Create a Treeview widget
columns = ['filename',
           'version',
           'issuer',
           'file_number',
           'content',
           'author',
           'author_date',
           'reviewer',
           'review_date',
           'publish_date',
           'implement_date']
tree = tk.Treeview(tabel_frame, columns=columns, height=25, show="headings",bootstyle='primary')

# 滚动条初始化（scrollBar为垂直滚动条，scrollBarx为水平滚动条）
scrollBar = tk.Scrollbar(tabel_frame, orient="vertical", command=tree.yview,bootstyle='primary-round')
scrollBar.pack(side="right", fill="y")
scrollBarx = tk.Scrollbar(tabel_frame, orient="horizontal", command=tree.xview,bootstyle='primary-round')
scrollBarx.pack(side="bottom", fill="x")
# 绑定Scrollbar和Treeview
tree.configure(yscrollcommand=scrollBar.set, xscrollcommand=scrollBarx.set)
# Pack the treeview widget
tree.pack(fill="both", expand=True)
# Add column headings
tree.heading('#0', text='ID')
tree.heading('file_number', text='文件编号')
tree.heading('version', text='版本')
tree.heading('issuer', text='签发人')
tree.heading('filename', text='管理文件名称')
tree.heading('content', text='内容')
tree.heading('author', text='制订人')
tree.heading('author_date', text='制订日期')
tree.heading('reviewer', text='审核人')
tree.heading('review_date', text='审核日期')
tree.heading('publish_date', text='发布日期')
tree.heading('implement_date', text='实施日期')

# Config column width
tree.column('file_number', width=80)
tree.column('version', width=40)
tree.column('issuer', width=60)
tree.column('filename', width=120)
tree.column('content', width=200)
tree.column('author', width=60)
tree.column('author_date', width=90)
tree.column('reviewer', width=60)
tree.column('review_date', width=90)
tree.column('publish_date', width=90)
tree.column('implement_date', width=90)


#根据选中记录生成Word文档
def generate_word():
    selected_item = tree.selection()  # 获取选中的数据
    if selected_item:
        if path == None:
            messagebox.showerror("错误", "未选择模板文件！")
        else:
            template = path         
            for item in selected_item:
                text = tree.item(item, "values")  # 获取选中的数据的文本内容
                # 渲染变量
                try:
                    document = MailMerge(template)
                    document.merge(
                        file_number=text[3],
                        version=text[1],
                        issuer=text[2],
                        filename=text[0],
                        content=text[4],
                        author=text[5],
                        author_date=datetime.datetime.strptime(
                            text[6], '%Y-%m-%d').strftime('%Y年%m月%d日'),
                        reviewer=text[7],
                        review_date=datetime.datetime.strptime(
                            text[8], '%Y-%m-%d').strftime('%Y年%m月%d日'),
                        publish_date=datetime.datetime.strptime(
                            text[9], '%Y-%m-%d').strftime('%Y年%m月%d日'),
                        implement_date=datetime.datetime.strptime(
                            text[10], '%Y-%m-%d').strftime('%Y年%m月%d日')
                    )
                    print(text[4])
                # 保存新文件             
                    output = "WGWLZ" + text[3] + \
                        text[0] + "(第" + text[1] + "版）.docx"
                    document.write(output)
                    messagebox.showinfo("成功", "WGWLZ" + text[3] + \
                        text[0] + "(第" + text[1] + "版）.docx\n文件已生成！")
                except:
                    messagebox.showerror("错误", "保存文件时出错！")

    else:
        messagebox.showerror("错误", "未选中任何数据！")

# 创建一个选择模板文件的按钮和输入框
template_button = tk.Button(button_frame, text="选择模板文件", command=open_template)
template_button.pack(side="left",pady=5)
template_entry = tk.Entry(button_frame)
template_entry.pack(side="left",pady=5)
ToolTip(template_button, text="选择Word模板文件应该是\n制作好的邮件合并的模板的文件!", bootstyle=("danger","inverse"))


# 创建一个生成Word文档的按钮
generate_button = tk.Button(
    button_frame, text="生成Word文档",bootstyle="outline", command=generate_word)
generate_button.pack(side="right",pady=5)
ToolTip(generate_button, text="在执行文件同文件夹下\n批量生产管理文件!", bootstyle=("info","inverse"))

#创建一个程序说明
info = tk.LabelFrame(info_frame, text='程序说明', padding=5,bootstyle="primary")
info.pack(fill="both", expand=True)
info_label = tk.Label(
    info, text="本程序数据来源于https://seatable.10an.fun，原始数据请在SeaTable里面编辑作为数据库保存，这里只是查询数据根据模板更快生成Word文档，原始文档中的表格和图片不能很好的生成，请自行修改一下！但是本程序也完成了大部分的排版工作，而且利于查找归档。\n点击本段文字直达https://seatable.10an.fun",font=("仿宋", 14,"bold") ,wraplength=700)
info_label.pack()

# 此处必须注意，绑定的事件函数中必须要包含event参数
def open_url(event):
    webbrowser.open("https://seatable.10an.fun", new=0)
 
# 绑定label单击事件
info_label.bind("<Button-1>", open_url)

#创建一个表格处理说明
table_info = tk.LabelFrame(info_frame, text='表格处理说明', padding=5,bootstyle="warning")
table_info.pack(fill="both", expand=True)
table_info_label = tk.Label(
    table_info, text="关于内容中表格导出来是markdown形式的表格，可以粘贴https://tableconvert.com/zh-cn/markdown-to-excel中生成Excel表格文件，复制粘贴到任意Excel文件中，就可以复制粘贴到Word文档中，这个方法曲线救国，但也免去自己画表重新输入。\n点击本段文字直达https://tableconvert.com/zh-cn/markdown-to-excel",font=("仿宋", 14,"bold") ,wraplength=700)
table_info_label.pack()

# 此处必须注意，绑定的事件函数中必须要包含event参数
def open_tableinfo_url(event):
    webbrowser.open("https://tableconvert.com/zh-cn/markdown-to-excel", new=0)
 
# 绑定label单击事件
table_info_label.bind("<Button-1>", open_tableinfo_url)

#初始化搜索条件
update_sql()
root.mainloop()