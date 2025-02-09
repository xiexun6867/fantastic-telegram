
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import sqlite3
from datetime import datetime
import pandas as pd

# 连接数据库
conn = sqlite3.connect('t1.db')
cursor = conn.cursor()

# 如果表已经存在，先删除表
cursor.execute('''
DROP TABLE IF EXISTS t1
''')

# 创建 t1 表（如果不存在）
cursor.execute('''
CREATE TABLE t1 (
    档案号 INTEGER,
    姓名 TEXT,
    来校时间 TEXT DEFAULT CURRENT_DATE,
    工龄 REAL DEFAULT 0 -- 工龄字段
)
''')
conn.commit()

# 创建主窗口
root = tk.Tk()
root.title("教师工龄管理V1.0")
# 设置窗口大小
root.geometry("800x600")

# 禁用窗口最大化按钮，第一个参数 False 表示禁止水平方向的调整，第二个参数 False 表示禁止垂直方向的调整。
root.resizable(False, False)

# 设置窗口图标
try:
    root.iconbitmap('icon.ico')
except:
    pass

# 获取屏幕的宽度和高度
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# 获取窗口的宽度和高度
window_width = 800
window_height = 600

# 计算窗口的位置
x = (screen_width - window_width) // 2
y = (screen_height - window_height) // 2

# 设置窗口的位置
root.geometry(f"{window_width}x{window_height}+{x}+{y}")

# 创建标签和文本框用于输入日期
date_label = tk.Label(root, text="请输入日期（格式：YYYYMM）：")
date_label.pack(pady=5)
date_entry = tk.Entry(root, width=20)
date_entry.pack(pady=5)
comment_label = tk.Label(root, text="若文本框输入年月，则更新、查询按输入值计算，否则按现在时间计算！")
comment_label.pack(pady=5)

# 创建表格用于显示查询结果
columns = ('档案号', '姓名', '来校时间', '工龄')

# 增加横向、竖向滚动条
tree_frame = tk.Frame(root,width=400)
# 创建表格控件
tree_frame.pack(pady=50, fill=tk.BOTH, side=tk.LEFT)

# 修改：将滚动条置于表格控件内部
tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=15)
tree_y_scroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=tree.yview)
tree_x_scroll = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=tree.xview)
tree.configure(yscrollcommand=tree_y_scroll.set, xscrollcommand=tree_x_scroll.set)

tree_y_scroll.pack(side=tk.RIGHT, fill=tk.Y)
tree_x_scroll.pack(side=tk.BOTTOM, fill=tk.X)

# 修改：调整表格控件位置，距左右边各30
tree.pack(pady=20, padx=30, fill=tk.BOTH, expand=True)

for col in columns:
    tree.heading(col, text=col)
    tree.column(col, width=100)
    
input_date_str = date_entry.get()

# 定义查询工龄为x.2的行的函数，要求>1
def query_teachers():
    try:
        if date_entry.get()!='':
            # 将输入的日期字符串转换为日期对象，只取年月

            query='''
            select * from t1
            where(CAST(工龄 * 100 AS INTEGER) % 100 = 17)
            '''
        else:
            # 获取当前时间
            now = datetime.now()
            current_month = now.month

            if current_month == 9:
                # 现在查询时间是9月，查找来校时间是2016年9月之前的所有行，以及来校时间是2016年9月以后，工龄为x.17，并且大于1年的行
                query = """
                SELECT * FROM t1
                WHERE (来校时间 < '2016-09-01')
                OR (来校时间 >= '2016-09-01' AND 工龄 > 1 AND CAST(工龄 * 100 AS INTEGER) % 100 =17)
               """
            else:
               # 现在查询时间不是9月，只查找来校时间是2016年9月以后，工龄为x.2，并且大于1年的行
               query = """
                SELECT * FROM t1
                WHERE 来校时间 >= '2016-09-01' AND 工龄 > 1 AND CAST(工龄 * 100 AS INTEGER) % 100 = 17
                """
        
        cursor.execute(query)
        
        rows = cursor.fetchall()
      #  print(f"查询结果：{rows}")  # 添加调试信息

        # 创建新窗口
        new_window = tk.Toplevel(root)
        new_window.title("查询结果")
        new_window.geometry("600x400")

        # 让新窗口获得焦点
        new_window.focus_force()

        # 获取屏幕的宽度和高度
        screen_width = new_window.winfo_screenwidth()
        screen_height = new_window.winfo_screenheight()

        # 获取窗口的宽度和高度
        window_width = 600
        window_height = 400

        # 计算窗口的位置
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2

        # 设置窗口的位置
        new_window.geometry(f"{window_width}x{window_height}+{x}+{y}")


        # 创建表格用于显示查询结果
        columns = ('档案号', '姓名', '来校时间', '工龄')
        result_frame = tk.Frame(new_window)
        result_frame.pack(pady=20, fill=tk.BOTH, expand=True)

        # 修改：将滚动条置于表格控件内部
        result_tree = ttk.Treeview(result_frame, columns=columns, show='headings')
        result_y_scroll = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=result_tree.yview)
        result_x_scroll = ttk.Scrollbar(result_frame, orient=tk.HORIZONTAL, command=result_tree.xview)
        result_tree.configure(yscrollcommand=result_y_scroll.set, xscrollcommand=result_x_scroll.set)

        result_y_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        result_x_scroll.pack(side=tk.BOTTOM, fill=tk.X)

        # 修改：调整表格控件位置，距左右边各30
        result_tree.pack(pady=20, padx=30, fill=tk.BOTH)

        for col in columns:
            result_tree.heading(col, text=col)
            result_tree.column(col, width=100)

        for row in rows:
            result_tree.insert('', 'end', values=row)

        # 获取导入的数据行数，弹出消息框，已查询xxx条记录
        query_count = len(rows)  
        custom_messagebox(query_count)

        # 定义关闭窗口的函数
        def close_window(event=None):  # 增加 event=None 参数，使其可以在不同调用方式下正常工作
            if event is not None:
                event.widget.destroy()
            else:
                new_window.destroy()  # 这里假设 root 是要关闭的窗口，根据实际情况修改

        # 绑定 Escape 键到 close_window 函数
        new_window.bind("<Escape>", close_window)

        # 定义导出到Excel的函数
        def export_to_excel():
            df = pd.DataFrame(rows, columns=['档案号', '姓名', '来校时间', '工龄'])
            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if file_path:
                df.to_excel(file_path, index=False)

        # 创建一个框架用于放置按钮
        button_frame = tk.Frame(new_window)
        button_frame.pack(side=tk.BOTTOM, fill=tk.X)

        # 创建关闭按钮
        close_button = tk.Button(button_frame, text="关闭", command=close_window, width=20,font=("微软雅黑",12))
        close_button.pack(side=tk.LEFT, padx=40, pady=15)

        # 创建导出按钮
        export_button = tk.Button(button_frame, text="导出到Excel表格", command=export_to_excel, width=20,font=("微软雅黑",12))
        export_button.pack(side=tk.LEFT, padx=25, pady=15)

        # 居中按钮
        button_frame.pack_configure(anchor=tk.CENTER)       

    except sqlite3.Error as e:
        for item in tree.get_children():
            tree.delete(item)
        tree.insert('', 'end', values=("查询失败", str(e), "", ""))

# 定义导入 Excel 文件的函数
def import_excel():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        try:
            df = pd.read_excel(file_path)
            # 选择需要的列
            required_columns = ['档案号', '姓名', '来校时间']
            # 检查所需列是否都存在
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                raise ValueError(f"Excel 文件中缺少必要的列：{', '.join(missing_columns)}")
            # 只选取所需列
            df = df[required_columns]

            # 删除原表数据
            cursor.execute("DELETE FROM t1")
            conn.commit()  # 提交删除操作

            # 插入新数据
            df.to_sql('t1', conn, if_exists='append', index=False)
            conn.commit()  # 提交插入操作
            
            # 获取导入的数据行数
            imported_count = len(df)
             #    messagebox.showinfo("导入成功", f"已导入{imported_count}条数据")   #这是一条官方消息框
            custom_messagebox(imported_count)
        
            query_all_teachers()
        except Exception as e:
            for item in tree.get_children():
                tree.delete(item)
            tree.insert('', 'end', values=("导入失败", str(e), ""))
            print(f"导入失败: {e}")  # 打印错误信息

    query_all_teachers()

# 定义查询全部教师的函数
def query_all_teachers():
    for item in tree.get_children():
        tree.delete(item)
    cursor.execute("SELECT * FROM t1")
    rows = cursor.fetchall()
    for row in rows:
        tree.insert('', 'end', values=row)

# 定义删除全部数据的函数
def delete_all_data():
    # 删除数据库表中的数据
    cursor.execute("DELETE FROM t1")
    conn.commit()
    # 删除root窗口中表格控件里的内容
    for item in tree.get_children():
        tree.delete(item)

updated_rows = []  # 存储需要更新的行
# 定义更新工龄的函数
def update_工龄():
    try:
        # 获取所有教师的信息
        cursor.execute("SELECT 档案号, 来校时间 FROM t1")
        rows = cursor.fetchall()

        global updated_rows
        updated_rows = []  # 存储需要更新的行

        # 遍历每一行数据
        for row in rows:
            档案号 = row[0]
            来校时间_str = row[1]

            # 步骤1：将来校时间进行调整，将之补充为YYYY.MM.DD的结构
            parts = 来校时间_str.split('.')
            if len(parts) == 1:  # 只有年份，如2008
                new_来校时间 = f"{来校时间_str}.01.01"
                updated_rows.append((档案号, new_来校时间))
            elif len(parts) == 2:  # 有年份和月份，如2016.3
                month = parts[1].zfill(2)
                new_来校时间 = f"{parts[0]}.{month}.01"
                updated_rows.append((档案号, new_来校时间))
            else:
                new_来校时间 = 来校时间_str

            # 步骤2：更新来校时间字段
            cursor.execute("UPDATE t1 SET 来校时间 =? WHERE 档案号 =?", (new_来校时间, 档案号))
            
        conn.commit()  # 提交来校时间的更新操作

        # 重新查询所有教师的信息
        cursor.execute("SELECT 档案号, 姓名, 来校时间, 工龄 FROM t1")
        all_rows = cursor.fetchall()

        # 清空表格控件
        for item in tree.get_children():
            tree.delete(item)

        # 插入更新后的数据到表格控件
        for row in all_rows:
            档案号, 姓名, 来校时间, 工龄 = row
            tree.insert('', 'end', values=(档案号, 姓名, 来校时间, 工龄))

        # 步骤3：在调整过的来校时间基础上，提取hire_year和hire_month,重新计算工龄
        if date_entry.get()!='':          
            current_year = int(date_entry.get()[:4])
            current_month = int(date_entry.get()[5:6])
        else:
            current_date = datetime.now()
            current_year = current_date.year
            current_month = current_date.month

        for row in all_rows:
            档案号, _, 来校时间, _ = row
            parts = 来校时间.split('.')
            hire_year = int(parts[0])
            hire_month = int(parts[1])

            # 计算工龄
            if current_month >= hire_month:
                工龄 = current_year - hire_year
            else:
                工龄 = current_year - hire_year - 1
            # 计算小数部分（月份差）
            month_difference = (current_month - hire_month) if current_month >= hire_month else (12 + current_month - hire_month)
            decimal_part = round(month_difference / 12, 2)
            工龄 += decimal_part

            # 更新数据库中对应 档案号 的 工龄 字段
            cursor.execute("UPDATE t1 SET 工龄 =? WHERE 档案号 =?", (工龄, 档案号))

        # 提交事务
        conn.commit()

        # 重新查询并显示所有教师信息
        query_all_teachers()

    except sqlite3.Error as e:
        print(f"更新失败: {e}")
        # 回滚事务
        conn.rollback()

        #标黄调整过来校时间的单元格  
    tree.tag_configure('adjusted', background='yellow')
    for item in tree.get_children():
            values = tree.item(item, 'values')
            档案号 = values[0]
            for row in updated_rows:
                if str(row[0]) == str(档案号):
                    # 找到对应的行，标黄来校时间单元格
                    tree.item(item, tags=('adjusted',))                  
                    break

# 定义导出 t1 表到 Excel 的函数
def export_t1_to_excel():
    cursor.execute("SELECT * FROM t1")
    rows = cursor.fetchall()
    df = pd.DataFrame(rows, columns=['档案号', '姓名', '来校时间', '工龄'])
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        df.to_excel(file_path, index=False)

# 创建底部一个框架用于放置按钮
button_frame = tk.Frame(root)
#button_frame.pack(side=tk.BOTTOM, fill=tk.X)
button_frame.place(x=20, y=545, width=750, height=50)

# 创建查询全部按钮
query_all_button = tk.Button(button_frame, text="查询全部", command=query_all_teachers, width=10,font=("微软雅黑", 12))
query_all_button.pack(side=tk.LEFT, padx=20, pady=5)
# 绑定焦点事件
query_all_button.bind("<FocusIn>", lambda event: show_focus(event.widget))
query_all_button.bind("<FocusOut>", lambda event: hide_focus(event.widget))

# 创建删除按钮
delete_button = tk.Button(button_frame, text="清屏", command=delete_all_data, width=10,font=("微软雅黑", 12))
delete_button.pack(side=tk.LEFT, padx=20, pady=5)
# 绑定焦点事件
delete_button.bind("<FocusIn>", lambda event: show_focus(event.widget))
delete_button.bind("<FocusOut>", lambda event: hide_focus(event.widget))

# 创建导出按钮
export_button = tk.Button(button_frame, text="导出所有教师", command=export_t1_to_excel, width=10,font=("微软雅黑", 12))
export_button.pack(side=tk.LEFT, padx=20, pady=5)
# 绑定焦点事件
export_button.bind("<FocusIn>", lambda event: show_focus(event.widget))
export_button.bind("<FocusOut>", lambda event: hide_focus(event.widget))

# 创建关闭按钮
close_root_button = tk.Button(button_frame, text="关闭", command=root.destroy, width=10,font=("微软雅黑", 12))
close_root_button.pack(side=tk.LEFT, padx=20, pady=5)
# 绑定焦点事件
close_root_button.bind("<FocusIn>", lambda event: show_focus(event.widget))
close_root_button.bind("<FocusOut>", lambda event: hide_focus(event.widget))

author_label = tk.Label(button_frame, text="如有使用问题，\n请联系：新新")
author_label.pack(side=tk.LEFT,padx=5, pady=5)

# 创建右半边的框架用于放置按钮
button_frame = tk.Frame(root)
#button_frame.pack(pady=50, fill=tk.Y, side=tk.RIGHT)
button_frame.place(x=500, y=100, width=250, height=400)

step1_label = tk.Label(button_frame, text="第一步：请导入教师信息excel表格",font=("微软雅黑", 12))
step1_label.pack(pady=5, anchor=tk.CENTER)

# 创建导入按钮
import_button = tk.Button(button_frame, text="导入", command=import_excel, width=10,font=("微软雅黑", 12))
import_button.pack(pady=10,anchor=tk.CENTER)
# 绑定焦点事件
import_button.bind("<FocusIn>", lambda event: show_focus(event.widget))
import_button.bind("<FocusOut>", lambda event: hide_focus(event.widget))

# 绑定回车键到导入按钮的命令函数
root.bind('<Return>', lambda event: import_excel())

step2_label = tk.Label(button_frame, text="第二步：更新工龄",font=("微软雅黑", 12))
step2_label.pack(pady=5)

# 创建更新按钮
update_button = tk.Button(button_frame, text="更新", command=update_工龄, width=10,font=("微软雅黑", 12))
update_button.pack(pady=10)
# 绑定焦点事件
update_button.bind("<FocusIn>", lambda event: show_focus(event.widget))
update_button.bind("<FocusOut>", lambda event: hide_focus(event.widget))

step3_label = tk.Label(button_frame, text="第三步：查询当月需加工龄的老师",font=("微软雅黑", 12))
step3_label.pack(pady=5)

# 创建查询按钮
query_button = tk.Button(button_frame, text="查询", command=query_teachers, width=10,font=("微软雅黑", 12))
query_button.pack(pady=10)
# 绑定焦点事件
query_button.bind("<FocusIn>", lambda event: show_focus(event.widget))
query_button.bind("<FocusOut>", lambda event: hide_focus(event.widget))

step4_label = tk.Label(button_frame, text="备注：来校时间在2016年9月之前的，\n统一为每年9月调整工龄工资。\n标黄行为补充了来校时间的数据",font=("微软雅黑",10))
step4_label.pack(pady=5)

# 显示焦点矩形框
def show_focus(button):
    x, y, width, height = button.winfo_x() + 2, button.winfo_y() + 2, button.winfo_width() - 4, button.winfo_height() - 4
    focus_frame = tk.Frame(button, relief=tk.GROOVE, bd=1)
    focus_frame.place(x=x, y=y, width=width, height=height)
    button.focus_frame = focus_frame

# 隐藏焦点矩形框
def hide_focus(button):
    if hasattr(button, 'focus_frame'):
        button.focus_frame.destroy()

# 定义一个函数，用于处理回车键事件
def handle_return(event):
    focused_widget = root.focus_get()
    if focused_widget:
        if isinstance(focused_widget, tk.Button):
            focused_widget.invoke()

# 绑定回车键事件到 handle_return 函数
root.bind('<Return>', handle_return)


# 自定义消息框
def custom_messagebox(imported_count):
    # 创建一个顶级窗口
    top = tk.Toplevel(root)
    # 获取屏幕的宽度和高度
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    # 获取窗口的宽度和高度
    window_width = 300
    window_height = 150
    # 计算窗口的位置
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    # 设置窗口的位置和大小
    top.geometry(f"{window_width}x{window_height}+{x}+{y}")
    # 设置窗口标题
    top.title("数据信息")
    # 创建一个标签，显示导入的数据行数，设置字体大小为16
    label = tk.Label(top, text=f"共{imported_count}条数据", font=("微软雅黑", 16))
    label.pack(pady=30)
    # 创建一个按钮，用于关闭窗口
    button = tk.Button(top, text="确定", command=top.destroy, font=("微软雅黑", 14))
    button.pack()

# 运行主循环
root.mainloop()

# 关闭数据库连接
conn.close()
