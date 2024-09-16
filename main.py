import tkinter as tk
from tkinter import scrolledtext, filedialog, messagebox,simpledialog
from tkinter import ttk
import subprocess
import random
import os
import pandas as pd  # 需要安装 pandas 库
import openpyxl  # 需要安装 openpyxl 库
from datetime import datetime, timedelta

class Application(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("自动化评语@LZY")
        self.geometry("1200x600")

        self.style = ttk.Style(self)
        self.style.theme_use('clam')  # 选择一个主题，clam 是其中之一
        self.iconbitmap('a.ico')
        self.current_file = None
        self.name_vars = []
        self.create_widgets()
        self.load_mingdan_in()
        self.load_pingyu_in()
        self.load_config()

    def create_widgets(self):
        # Frame for the file list and checklist
        self.left_frame = ttk.Frame(self)
        self.left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10)

        # Listbox for names
        self.name_listbox = tk.Listbox(self.left_frame, selectmode=tk.SINGLE, height=13)
        self.name_listbox.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.name_listbox.bind('<<ListboxSelect>>', self.on_name_select)

        # Scrollbar for the listbox
        self.scrollbar = ttk.Scrollbar(self.left_frame, orient=tk.VERTICAL)
        self.scrollbar.config(command=self.name_listbox.yview)
        self.name_listbox.config(yscrollcommand=self.scrollbar.set)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Refresh Button
        self.refresh_button = ttk.Button(self.left_frame, text="刷新名单", command=self.load_mingdan_in)
        self.refresh_button.pack(side=tk.TOP, pady=5)

        # Checklist frame
        self.checklist_frame = ttk.Frame(self.left_frame)
        self.checklist_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # Right frame for the buttons and text area
        self.middle_frame = ttk.Frame(self)
        self.middle_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Button frame for middle buttons
        self.button_frame = ttk.Frame(self.middle_frame)
        self.button_frame.pack(pady=5)

        # Generate Comments Button
        self.generate_comments_button = ttk.Button(self.button_frame, text="自动生成反馈模板", command=self.generate_comments)
        self.generate_comments_button.grid(row=0, column=0, padx=5, pady=5)

        # Generate Excel seating chart button
        self.generate_excel_button = ttk.Button(self.button_frame, text="生成Excel姓名表", command=self.generate_excel)
        self.generate_excel_button.grid(row=0, column=1, padx=5, pady=5)

        # Load and update seating chart button
        self.load_seating_chart_button = ttk.Button(self.button_frame, text="加载excel评语", command=self.load_seating_chart)
        self.load_seating_chart_button.grid(row=0, column=2, padx=5, pady=5)

        # Open a.md File
        self.open_a_md_button = ttk.Button(self.button_frame, text="查看和当前反馈", command=self.open_a_md)
        self.open_a_md_button.grid(row=1, column=0, padx=5, pady=5)

        # Open neirong.in File
        self.open_neirong_in_button = ttk.Button(self.button_frame, text="查看今日内容", command=self.open_neirong_in)
        self.open_neirong_in_button.grid(row=1, column=1, padx=5, pady=5)

        # Open mingdan.in File
        self.open_mingdan_in_button = ttk.Button(self.button_frame, text="查看学生名单", command=self.open_mingdan_in)
        self.open_mingdan_in_button.grid(row=1, column=2, padx=5, pady=5)

        # Save File
        self.save_button = ttk.Button(self.button_frame, text="保存当前文件", command=self.save_file)
        self.save_button.grid(row=1, column=3, padx=5, pady=5)


        # Text Area
        self.text_area = scrolledtext.ScrolledText(self.middle_frame, wrap=tk.WORD, width=80, height=20)
        self.text_area.pack(pady=10)

        # Label to show the current file
        self.current_file_label = ttk.Label(self.middle_frame, text="")
        self.current_file_label.pack(pady=5)

        # Frame for the comments and buttons
        self.right_frame = ttk.Frame(self)
        self.right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Comment text area
        self.comment_text = scrolledtext.ScrolledText(self.right_frame, wrap=tk.WORD, width=40, height=20)
        self.comment_text.pack(pady=5)

        # Comment buttons (now in two rows and three columns)
        self.button_grid_frame = ttk.Frame(self.right_frame)
        self.button_grid_frame.pack()

        self.listen_carefully_button = ttk.Button(self.button_grid_frame, text="听讲认真", command=lambda: self.add_comment("听讲认真"))
        self.listen_carefully_button.grid(row=0, column=0, padx=5, pady=5)

        self.listen_carefully_button = ttk.Button(self.button_grid_frame, text="听讲不认真", command=lambda: self.add_comment("听讲不认真"))
        self.listen_carefully_button.grid(row=1, column=0, padx=5, pady=5)

        self.active_participation_button = ttk.Button(self.button_grid_frame, text="比较积极", command=lambda: self.add_comment("比较积极"))
        self.active_participation_button.grid(row=0, column=1, padx=5, pady=5)

        self.deep_questions_button = ttk.Button(self.button_grid_frame, text="性格较闷", command=lambda: self.add_comment("性格较闷"))
        self.deep_questions_button.grid(row=1, column=1, padx=5, pady=5)

        self.quick_response_button = ttk.Button(self.button_grid_frame, text="思维灵活", command=lambda: self.add_comment("思维灵活"))
        self.quick_response_button.grid(row=0, column=2, padx=5, pady=5)

        self.quick_response_button = ttk.Button(self.button_grid_frame, text="慢半拍子", command=lambda: self.add_comment("慢半拍子"))
        self.quick_response_button.grid(row=1, column=2, padx=5, pady=5)

        self.listen_carefully_button = ttk.Button(self.button_grid_frame, text="上课划水", command=lambda: self.add_comment("上课划水"))
        self.listen_carefully_button.grid(row=2, column=0, padx=5, pady=5)

        self.listen_carefully_button = ttk.Button(self.button_grid_frame, text="扩展知识", command=lambda: self.add_comment("扩展知识"))
        self.listen_carefully_button.grid(row=3, column=0, padx=5, pady=5)

        self.active_participation_button = ttk.Button(self.button_grid_frame, text="进度较快", command=lambda: self.add_comment("进度较快"))
        self.active_participation_button.grid(row=2, column=1, padx=5, pady=5)

        self.deep_questions_button = ttk.Button(self.button_grid_frame, text="进度较慢", command=lambda: self.add_comment("进度较慢"))
        self.deep_questions_button.grid(row=3, column=1, padx=5, pady=5)

        self.quick_response_button = ttk.Button(self.button_grid_frame, text="其他好", command=lambda: self.add_comment("其他好"))
        self.quick_response_button.grid(row=2, column=2, padx=5, pady=5)

        self.quick_response_button = ttk.Button(self.button_grid_frame, text="其他不好", command=lambda: self.add_comment("其他不好"))
        self.quick_response_button.grid(row=3, column=2, padx=5, pady=5)

        # Frame for date setting
        self.date_frame = ttk.Frame(self.right_frame)
        self.date_frame.pack(pady=10)

        self.start_date_label = ttk.Label(self.date_frame, text="设置起始日期 :")
        self.start_date_label.grid(row=0, column=0, padx=5, pady=5)

        self.start_date_entry = ttk.Entry(self.date_frame)
        self.start_date_entry.grid(row=0, column=1, padx=5, pady=5)

        self.modify_date_button = ttk.Button(self.date_frame, text="修改日期", command=self.modify_date)
        self.modify_date_button.grid(row=0, column=2, padx=5, pady=5)
        self.fenban_var=0
        # Create a Checkbutton
        self.fenban_checkbutton = tk.Checkbutton(self.date_frame, text="(假按钮，没用)", variable=self.fenban_var)
        self.fenban_checkbutton.grid(row=1, column=0, padx=5, pady=5)

    def load_config(self):
        try:
            with open("config.in", 'r', encoding='utf-8') as file:
                self.config = {}
                for line in file:
                    key, value = line.strip().split('=')
                    self.config[key.strip()] = value.strip()
            self.start_date_entry.insert(0, self.config.get("start_date", "2024-07-15"))  # 默认起始日期
            self.fenban_var = tk.BooleanVar(value=self.config.get("fenban_flag",1))
        except Exception as e:
            messagebox.showerror("错误", f"无法加载 config.in 文件：{e}")

    def load_seating_chart(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls *.xlsx")])
        if filepath:
            try:
                df = pd.read_excel(filepath,header=None)
                for i in range(df.shape[0]):  # Iterate over all rows
                    for j in range(df.shape[1]):  # Iterate over all columns
                        print(i,j,"excal")
                        cell_value = df.iloc[i, j]
                        if isinstance(cell_value, str):
                            parts = cell_value.split("\n")
                            if len(parts) >= 2:
                                student_name = parts[0].strip()
                                print(student_name)
                                student_performance = parts[1].strip()
                                self.update_student_comment(student_name, student_performance)
                            else:
                                pass
                                
                        else:
                            pass
                messagebox.showinfo("成功", "座位表处理完成并更新到 a.md 文件中。")
                self.open_file("a.md")
            except Exception as e:
                messagebox.showerror("错误", f"处理 Excel 文件时出错：{e}")

    def modify_date(self):
        new_date = self.start_date_entry.get()
        self.config['start_date'] = new_date
        try:
            with open("config.in", 'w', encoding='utf-8') as file:
                for key, value in self.config.items():
                    file.write(f"{key}={value}\n")
            messagebox.showinfo("成功", "日期已更新。")
        except Exception as e:
            messagebox.showerror("错误", f"更新 config.in 文件时出错：{e}")

    def confirm_run_program(self):
        result = messagebox.askokcancel("确认运行", "点击确认会覆盖原有评语模板，你确认要重新覆盖吗？")
        if result:
            self.run_program()
            self.open_file("a.md")

    def run_program(self):
        try:
            subprocess.run(["a.exe"], check=True)
            messagebox.showinfo("成功", "评语模板生成成功。")
        except subprocess.CalledProcessError as e:
            messagebox.showerror("错误", f"执行出错：{e}")

    def open_file(self, filepath):
        try:
            with open(filepath, 'r', encoding='utf-8') as file:
                content = file.read()
                self.text_area.delete(1.0, tk.END)
                self.text_area.insert(tk.INSERT, content)
                self.current_file = filepath
                self.current_file_label.config(text=f"当前文件: {self.current_file}")
        except Exception as e:
            messagebox.showerror("错误", f"无法打开文件：{e}")

    def open_a_md(self):
        self.open_file("a.md")

    def open_neirong_in(self):
        self.open_file("neirong.in")

    def open_mingdan_in(self):
        self.open_file("mingdan.in")

    def load_mingdan_in(self):
        self.name_listbox.delete(0, tk.END)
        for widget in self.checklist_frame.winfo_children():
            widget.destroy()
        self.name_vars.clear()

        try:
            with open("mingdan.in", 'r', encoding='utf-8') as file:
                ii=1
                name_List=[]
                for line in file:
                    name = line.strip()
                    name_List.append(name)
                    self.name_listbox.insert(tk.END, name)
                    
                    var = tk.IntVar()
                    chk = tk.Button(self.checklist_frame,  text=name,width=5, height=1, command=lambda index=name: self.edit_comment(index))
                    #chk.pack(row=(int)(ii/3), column=(int)(ii%3),anchor=tk.W)
                    chk.grid(row=(int)(ii/3), column=(int)(ii%3), padx=5, pady=5)
                    self.name_vars.append(var)
                    ii=ii+1
                load_name_button=tk.Button(self.checklist_frame,  text="载入评语",width=5, height=1, command=lambda:self.load_student_comment_from_txt(name_List))
                load_name_button.grid(row=0, column=0, padx=5, pady=5)
        except Exception as e:
            messagebox.showerror("错误", f"无法加载 mingdan.in 文件：{e}")
    # 定义从文件读取学生名字的函数
    
    def edit_comment(self,student):
        # 创建弹出窗口
        self.edit_window = tk.Toplevel(self)
        self.edit_window.title(f"编辑 {student} 的评语")
        self.edit_window.geometry("400x400")
        
        # 创建文本框
        self.text_editor = tk.Text(self.edit_window, wrap="word")
        self.text_editor.pack(expand=1, fill="both")
        
        # 创建底部信息标签
        self.info_label = tk.Label(self.edit_window, text="编辑的个人评语点击保存后会存在studen/学生名字.txt。\n本编辑与excel不互通，会重复叠加。", anchor="w")
        self.info_label.pack(fill="x")
        
        # 从文件中读取现有评语（如果存在）
        file_path = f"student/{student}.txt"
        print(file_path)
        if os.path.exists(file_path):
            with open(file_path, "r", encoding="utf-8") as file:
                self.text_editor.insert("1.0", file.read())
        self.save_button = tk.Button(self.edit_window, text="保存", command=lambda: self.save_comment(file_path))
        self.save_button.pack(side="right")
    
    # 创建保存按钮
    def save_comment(self,file_path):
        content = self.text_editor.get("1.0", "end-1c")
        if not os.path.exists("student"):
            os.makedirs("student")
        with open(file_path, "w", encoding="utf-8") as file:
            file.write(content)
        messagebox.showinfo("保存成功", f"评语已保存到 {file_path}")
        self.edit_window.destroy()

    
    def on_name_select(self, event):
        try:
            selection = event.widget.curselection()
            if selection:
                index = selection[0]
                name = self.name_listbox.get(index)
                content = self.text_area.get(1.0, tk.END)
                position = content.find(f"{name}-----------------")
                if position != -1:
                    self.text_area.see(f"1.0+{position}c")
                    self.text_area.tag_remove("highlight", "1.0", tk.END)
                    end_position = position + len(name)
                    self.text_area.tag_add("highlight", f"1.0+{position}c", f"1.0+{end_position}c")
                    self.text_area.tag_config("highlight", background="yellow")
                else:
                    messagebox.showinfo("提示", f"未找到名字: {name}")
        except Exception as e:
            messagebox.showerror("错误", f"处理选择时出错：{e}")

    def save_file(self):
        if self.current_file:
            try:
                with open(self.current_file, 'w', encoding='utf-8') as file:
                    content = self.text_area.get(1.0, tk.END)
                    file.write(content)
                    messagebox.showinfo("成功", f"文件 '{self.current_file}' 保存成功。")
            except Exception as e:
                messagebox.showerror("错误", f"无法保存文件：{e}")
        else:
            messagebox.showwarning("警告", "没有打开的文件可以保存。")

    def load_student_comment_from_txt(self, name_List):
        for student_name in name_List:
            try:
                txt_filepath = os.path.join("student", f"{student_name}.txt")
                if os.path.isfile(txt_filepath):
                    with open(txt_filepath, 'r', encoding='utf-8') as f:
                        comments = f.read().strip()
                        print(f"从 {student_name}.txt 文件中读取的评语: {comments}")
                        self.update_student_comment(student_name, comments)
                else:
                    print(f"{student_name}.txt 文件不存在。")
            except Exception as e:
                print(f"读取 {student_name}.txt 文件时出错：{e}")

    def load_pingyu_in(self):
        self.comments = {
            "听讲认真": [],
            "听讲不认真": [],
            "比较积极": [],
            "性格较闷": [],
            "思维灵活": [],
            "慢半拍子": [],
            "上课划水": [],
            "扩展知识": [],
            "进度较快": [],
            "进度较慢": [],
            "其他好": [],
            "其他不好": []
        }
        try:
            with open("pingyu.in", 'r', encoding='utf-8') as file:
                for line in file:
                    for key in self.comments.keys():
                        if line.startswith(key):
                            comment = line.split(':', 1)[1].strip()
                            self.comments[key].append(comment)
        except Exception as e:
            messagebox.showerror("错误", f"无法加载 pingyu.in 文件：{e}")

    def add_comment(self, category):
        if category in self.comments and self.comments[category]:
            comment = random.choice(self.comments[category])
            self.comment_text.insert(tk.END, comment + "\n")
        else:
            messagebox.showwarning("警告", f"没有找到 '{category}' 的评语。")
    def update_student_comment(self, student_name, student_performance):
        try:
            with open("a.md", 'r+', encoding='utf-8') as file:
                content = file.read()
                start_index = content.find(f"{student_name}--------------------------------------------------------")
                if start_index != -1:
                    insert_index = content.find("【教师寄语】", start_index) + len("【教师寄语】\n今天是我们学习的第*天,")
                    updated_content = content[:insert_index] + f"\n{student_performance}" + content[insert_index:]
                    file.seek(0)
                    file.write(updated_content)
                    file.truncate()
                else:
                    messagebox.showwarning("警告", f"未找到学生: {student_name} 的评语部分。")
        except Exception as e:
            messagebox.showerror("错误", f"更新学生评语时出错：{e}")

    def generate_excel(self):
        try:
            names = []
            with open("mingdan.in", 'r', encoding='utf-8') as file:
                for line in file:
                    names.append(line.strip())
            
            random.shuffle(names)
            data = [['' for _ in range(30)] for _ in range(30)]
            for i in range(36):
                row = i // 15
                col = i % 15
                data[row][col] = ""
            for i, name in enumerate(names[:25]):
                row = i // 5
                col = i % 5
                data[row][col] = name+"\n\n"

            df = pd.DataFrame(data)
            df.to_excel("a.xls", index=False, header=False)
            messagebox.showinfo("成功", "Excel 座位表生成成功，文件名为 a.xls")
        except Exception as e:
            messagebox.showerror("错误", f"生成 Excel 座位表时出错：{e}")

    def generate_comments(self):
            try:
                start_date_str = self.start_date_entry.get()
                start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
                current_date = datetime.now()
                day_difference = (current_date - start_date).days

                with open("neirong.in", 'r', encoding='utf-8') as neirong_file:
                    neirong_content = neirong_file.read().strip()

                with open("mingdan.in", 'r', encoding='utf-8') as mingdan_file:
                    students = [line.strip() for line in mingdan_file]

                with open("a.md", 'w', encoding='utf-8') as output_file:
                    for student in students:
                        comment = f"{student}-------------------------------------------------------------------\n"
                        comment += f"{student}家长晚上好！\n\n"
                        comment += f"我是今天的初赛营的助教老师。今天是学习的第 {day_difference} 天，感谢{student}同学的坚持努力和各位家长的持续关注。今日的教学反馈如下：\n\n"
                        comment += neirong_content
                        #comment += f"\n\n【教师寄语】\n今天是我们学习的第{day_difference}天,\n\n今天的分班测试{student}还在我们班级。希望{student}同学继续努力，继续进步！\n\n"
                        if day_difference==1:
                            comment += f"\n\n【教师寄语】\n今天是我们学习的第{day_difference}天,\n\n希望{student}同学继续努力，继续进步！\n\n"
                        else:
                            comment += f"\n\n【教师寄语】\n今天是我们学习的第{day_difference}天,\n\n今天的分班测试{student}还在我们班级。希望{student}同学继续努力，继续进步！\n\n"
                        comment += "\n-----------------------------------------------------------------\n"
                        output_file.write(comment)

                messagebox.showinfo("成功", "评语生成成功并保存至 a.md 文件。")
                self.open_file("a.md")
            except Exception as e:
                messagebox.showerror("错误", f"生成评语时出错：{e}")

if __name__ == "__main__":
    app = Application()
    app.mainloop()
