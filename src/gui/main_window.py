import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import sys
import threading
import queue

# 添加src目录到Python路径
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))

from src.main import SportsMeetManager

class OrderBookGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("秩序册生成器")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # 初始化变量
        self.excel_path = ""
        self.output_path = "秩序册.docx"
        self.is_running = False
        self.thread = None
        self.queue = queue.Queue()
        
        # 解析进度条变量
        self.parser_progress_var = tk.DoubleVar()
        
        # 创建主框架
        self.main_frame = ttk.Frame(root, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建标题
        self.title_label = ttk.Label(
            self.main_frame, 
            text="秩序册生成器", 
            font=("微软雅黑", 16, "bold")
        )
        self.title_label.pack(pady=10)
        
        # 创建主笔记本（三页面）
        self.main_notebook = ttk.Notebook(self.main_frame)
        self.main_notebook.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 创建Excel解析页面
        self.excel_parser_frame = ttk.Frame(self.main_notebook)
        self.main_notebook.add(self.excel_parser_frame, text="Excel解析")
        
        # 创建数据查询页面
        self.data_query_frame = ttk.Frame(self.main_notebook)
        self.main_notebook.add(self.data_query_frame, text="数据查询")
        
        # 创建秩序册生成页面
        self.order_book_frame = ttk.Frame(self.main_notebook)
        self.main_notebook.add(self.order_book_frame, text="秩序册生成")
        
        # 初始化Excel解析页面
        self.init_excel_parser_page()
        
        # 初始化数据查询页面
        self.init_data_query_page()
        
        # 初始化秩序册生成页面
        self.init_order_book_page()
        
        # 绑定队列处理
        self.root.after(100, self.process_queue)
    
    def init_excel_parser_page(self):
        """初始化Excel解析页面"""
        # 创建文件选择区域
        self.file_frame = ttk.LabelFrame(self.excel_parser_frame, text="Excel文件选择", padding="10")
        self.file_frame.pack(fill=tk.X, pady=10)
        
        self.file_path_var = tk.StringVar()
        self.file_entry = ttk.Entry(self.file_frame, textvariable=self.file_path_var, width=50)
        self.file_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        self.browse_button = ttk.Button(self.file_frame, text="浏览", command=self.browse_file)
        self.browse_button.pack(side=tk.RIGHT, padx=5)
        
        # 创建操作按钮区域
        self.parser_button_frame = ttk.Frame(self.excel_parser_frame)
        self.parser_button_frame.pack(fill=tk.X, pady=10)
        
        self.preview_button = ttk.Button(
            self.parser_button_frame, 
            text="预览运动项目", 
            command=self.preview_events
        )
        self.preview_button.pack(side=tk.LEFT, padx=5)
        
        self.parse_button = ttk.Button(
            self.parser_button_frame, 
            text="解析Excel文件", 
            command=self.start_parsing
        )
        self.parse_button.pack(side=tk.LEFT, padx=5)
        
        # 创建运动项目展示区域
        self.events_frame = ttk.LabelFrame(self.excel_parser_frame, text="运动项目预览", padding="10")
        self.events_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 创建笔记本（选项卡）
        self.events_notebook = ttk.Notebook(self.events_frame)
        self.events_notebook.pack(fill=tk.BOTH, expand=True)
        
        # 创建1-3年级选项卡
        self.lower_grades_frame = ttk.Frame(self.events_notebook)
        self.events_notebook.add(self.lower_grades_frame, text="1-3年级")
        
        # 创建4-6年级选项卡
        self.upper_grades_frame = ttk.Frame(self.events_notebook)
        self.events_notebook.add(self.upper_grades_frame, text="4-6年级")
        
        # 创建1-3年级项目列表
        self.lower_grades_scroll = ttk.Scrollbar(self.lower_grades_frame)
        self.lower_grades_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.lower_grades_list = tk.Listbox(
            self.lower_grades_frame, 
            yscrollcommand=self.lower_grades_scroll.set,
            width=50, 
            height=10
        )
        self.lower_grades_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.lower_grades_scroll.config(command=self.lower_grades_list.yview)
        
        # 创建4-6年级项目列表
        self.upper_grades_scroll = ttk.Scrollbar(self.upper_grades_frame)
        self.upper_grades_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.upper_grades_list = tk.Listbox(
            self.upper_grades_frame, 
            yscrollcommand=self.upper_grades_scroll.set,
            width=50, 
            height=10
        )
        self.upper_grades_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.upper_grades_scroll.config(command=self.upper_grades_list.yview)
        
        # 创建进度条
        self.parser_progress_frame = ttk.LabelFrame(self.excel_parser_frame, text="进度", padding="10")
        self.parser_progress_frame.pack(fill=tk.X, pady=10)
        
        self.parser_progress_bar = ttk.Progressbar(
            self.parser_progress_frame, 
            variable=self.parser_progress_var, 
            maximum=100
        )
        self.parser_progress_bar.pack(fill=tk.X, padx=5)
        
        # 创建状态文本区域
        self.parser_status_frame = ttk.LabelFrame(self.excel_parser_frame, text="解析状态", padding="10")
        self.parser_status_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.parser_status_text = tk.Text(self.parser_status_frame, height=8, wrap=tk.WORD)
        self.parser_status_text.pack(fill=tk.BOTH, expand=True, padx=5)
        self.parser_status_text.config(state=tk.DISABLED)
    
    def init_data_query_page(self):
        """初始化数据查询页面"""
        # 创建查询条件区域
        self.query_conditions_frame = ttk.LabelFrame(self.data_query_frame, text="查询条件", padding="10")
        self.query_conditions_frame.pack(fill=tk.X, pady=10)
        
        # 创建年级选择
        self.grade_frame = ttk.Frame(self.query_conditions_frame)
        self.grade_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(self.grade_frame, text="年级:").pack(side=tk.LEFT, padx=5)
        self.grade_var = tk.StringVar()
        self.grade_combobox = ttk.Combobox(
            self.grade_frame, 
            textvariable=self.grade_var,
            values=["全部", "1", "2", "3", "4", "5", "6"]
        )
        self.grade_combobox.current(0)
        self.grade_combobox.pack(side=tk.LEFT, padx=5)
        
        # 创建班级选择
        self.class_frame = ttk.Frame(self.query_conditions_frame)
        self.class_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(self.class_frame, text="班级:").pack(side=tk.LEFT, padx=5)
        self.class_var = tk.StringVar()
        self.class_combobox = ttk.Combobox(
            self.class_frame, 
            textvariable=self.class_var,
            values=["全部", "1班", "2班", "3班", "4班", "5班", "6班"]
        )
        self.class_combobox.current(0)
        self.class_combobox.pack(side=tk.LEFT, padx=5)
        
        # 创建性别选择
        self.gender_frame = ttk.Frame(self.query_conditions_frame)
        self.gender_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(self.gender_frame, text="性别:").pack(side=tk.LEFT, padx=5)
        self.gender_var = tk.StringVar()
        self.gender_combobox = ttk.Combobox(
            self.gender_frame, 
            textvariable=self.gender_var,
            values=["全部", "男", "女"]
        )
        self.gender_combobox.current(0)
        self.gender_combobox.pack(side=tk.LEFT, padx=5)
        
        # 创建项目选择
        self.event_frame = ttk.Frame(self.query_conditions_frame)
        self.event_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(self.event_frame, text="项目:").pack(side=tk.LEFT, padx=5)
        self.event_var = tk.StringVar()
        self.event_combobox = ttk.Combobox(
            self.event_frame, 
            textvariable=self.event_var,
            values=["全部"]
        )
        self.event_combobox.current(0)
        self.event_combobox.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # 加载项目列表
        self.load_events()
        
        # 创建查询按钮
        self.query_button_frame = ttk.Frame(self.data_query_frame)
        self.query_button_frame.pack(fill=tk.X, pady=10)
        
        self.query_button = ttk.Button(
            self.query_button_frame, 
            text="执行查询", 
            command=self.execute_query
        )
        self.query_button.pack(side=tk.LEFT, padx=5)
        
        # 创建查询结果区域
        self.query_result_frame = ttk.LabelFrame(self.data_query_frame, text="查询结果", padding="10")
        self.query_result_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 创建结果表格
        self.result_tree = ttk.Treeview(self.query_result_frame, columns=("number", "name", "grade", "class", "gender", "event"), show="headings")
        self.result_tree.heading("number", text="学号")
        self.result_tree.heading("name", text="姓名")
        self.result_tree.heading("grade", text="年级")
        self.result_tree.heading("class", text="班级")
        self.result_tree.heading("gender", text="性别")
        self.result_tree.heading("event", text="项目")
        
        # 设置列宽
        self.result_tree.column("number", width=80)
        self.result_tree.column("name", width=100)
        self.result_tree.column("grade", width=60)
        self.result_tree.column("class", width=80)
        self.result_tree.column("gender", width=60)
        self.result_tree.column("event", width=200)
        
        # 添加滚动条
        self.result_scrollbar = ttk.Scrollbar(self.query_result_frame, orient=tk.VERTICAL, command=self.result_tree.yview)
        self.result_tree.configure(yscrollcommand=self.result_scrollbar.set)
        
        self.result_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_tree.pack(fill=tk.BOTH, expand=True)
        
        # 创建结果统计区域
        self.result_stats_frame = ttk.LabelFrame(self.data_query_frame, text="结果统计", padding="10")
        self.result_stats_frame.pack(fill=tk.X, pady=10)
        
        self.stats_label = ttk.Label(self.result_stats_frame, text="共 0 条记录")
        self.stats_label.pack(side=tk.LEFT, padx=5)
    
    def init_order_book_page(self):
        """初始化秩序册生成页面"""
        # 创建操作按钮区域
        self.generator_button_frame = ttk.Frame(self.order_book_frame)
        self.generator_button_frame.pack(fill=tk.X, pady=10)
        
        self.generate_button = ttk.Button(
            self.generator_button_frame, 
            text="生成秩序册", 
            command=self.start_generation
        )
        self.generate_button.pack(side=tk.LEFT, padx=5)
        
        self.open_file_button = ttk.Button(
            self.generator_button_frame, 
            text="打开生成文件", 
            command=self.open_generated_file, 
            state=tk.DISABLED
        )
        self.open_file_button.pack(side=tk.LEFT, padx=5)
        
        self.open_folder_button = ttk.Button(
            self.generator_button_frame, 
            text="打开输出文件夹", 
            command=self.open_output_folder, 
            state=tk.DISABLED
        )
        self.open_folder_button.pack(side=tk.LEFT, padx=5)
        
        # 创建进度条
        self.progress_frame = ttk.LabelFrame(self.order_book_frame, text="进度", padding="10")
        self.progress_frame.pack(fill=tk.X, pady=10)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            self.progress_frame, 
            variable=self.progress_var, 
            maximum=100
        )
        self.progress_bar.pack(fill=tk.X, padx=5)
        
        # 创建状态文本区域
        self.generator_status_frame = ttk.LabelFrame(self.order_book_frame, text="生成状态", padding="10")
        self.generator_status_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.generator_status_text = tk.Text(self.generator_status_frame, height=8, wrap=tk.WORD)
        self.generator_status_text.pack(fill=tk.BOTH, expand=True, padx=5)
        self.generator_status_text.config(state=tk.DISABLED)
    
    def browse_file(self):
        """浏览选择Excel文件"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if file_path:
            self.excel_path = file_path
            self.file_path_var.set(file_path)
            self.update_parser_status(f"已选择文件: {os.path.basename(file_path)}")
    
    def start_generation(self):
        """开始生成秩序册"""
        # 禁用按钮
        self.generate_button.config(state=tk.DISABLED)
        self.open_file_button.config(state=tk.DISABLED)
        self.open_folder_button.config(state=tk.DISABLED)
        
        # 重置进度条
        self.progress_var.set(0)
        
        # 清空状态文本
        self.clear_generator_status()
        self.update_generator_status("开始生成秩序册...")
        
        # 启动线程执行生成任务
        self.is_running = True
        self.thread = threading.Thread(target=self.generate_order_book)
        self.thread.daemon = True
        self.thread.start()
    
    def generate_order_book(self):
        """生成秩序册的线程函数"""
        try:
            # 发送进度更新
            self.queue.put(("progress", 20))
            self.queue.put(("status", "初始化管理器..."))
            
            # 创建管理器
            # 注意：这里不再需要excel_path，因为数据已经在数据库中
            manager = SportsMeetManager('')
            
            self.queue.put(("progress", 40))
            self.queue.put(("status", "准备生成秩序册..."))
            
            # 执行生成
            success = manager.generator.generate_order_book(self.output_path)
            
            self.queue.put(("progress", 80))
            self.queue.put(("status", "生成秩序册..."))
            
            if success:
                self.queue.put(("progress", 100))
                self.queue.put(("status", f"秩序册生成成功！保存在: {self.output_path}"))
                self.queue.put(("success", True))
            else:
                self.queue.put(("progress", 100))
                self.queue.put(("status", "生成失败，请检查数据库是否有数据"))
                self.queue.put(("success", False))
                
        except Exception as e:
            self.queue.put(("progress", 100))
            self.queue.put(("status", f"错误: {str(e)}"))
            self.queue.put(("success", False))
    
    def process_queue(self):
        """处理队列中的消息"""
        try:
            while not self.queue.empty():
                message = self.queue.get_nowait()
                if message[0] == "progress":
                    self.progress_var.set(message[1])
                elif message[0] == "parser_progress":
                    self.parser_progress_var.set(message[1])
                elif message[0] == "status":
                    self.update_parser_status(message[1])
                elif message[0] == "success":
                    self.generation_completed(message[1])
                elif message[0] == "parse_success":
                    self.parsing_completed(message[1])
        except queue.Empty:
            pass
        finally:
            self.root.after(100, self.process_queue)
    
    def parsing_completed(self, success):
        """解析完成后的处理"""
        # 启用按钮
        self.parse_button.config(state=tk.NORMAL)
        self.browse_button.config(state=tk.NORMAL)
        self.preview_button.config(state=tk.NORMAL)
        
        if success:
            # 更新数据查询页面的项目列表
            self.load_events()
            # 显示成功消息
            messagebox.showinfo("成功", "Excel文件解析完成！数据已存入数据库")
        else:
            # 显示失败消息
            messagebox.showerror("失败", "Excel文件解析失败，请检查错误信息")
    
    def execute_query(self):
        """执行数据查询"""
        from database.db_manager import DatabaseManager
        
        # 获取查询条件
        grade = self.grade_var.get()
        class_name = self.class_var.get()
        gender = self.gender_var.get()
        event_name = self.event_var.get()
        
        # 转换年级为整数
        grade_int = None
        if grade != "全部":
            grade_int = int(grade)
        
        # 处理班级
        if class_name == "全部":
            class_name = None
        
        # 处理性别
        if gender == "全部":
            gender = None
        
        # 处理项目
        if event_name == "全部":
            event_name = None
        
        # 创建数据库管理器
        db_manager = DatabaseManager()
        
        # 执行查询
        try:
            # 清空结果表格
            for item in self.result_tree.get_children():
                self.result_tree.delete(item)
            
            # 执行查询
            students = db_manager.get_students_by_multiple_conditions(
                grade=grade_int, 
                class_name=class_name, 
                gender=gender, 
                event_name=event_name
            )
            
            # 显示结果
            for student in students:
                self.result_tree.insert("", tk.END, values=student)
            
            # 更新统计信息
            self.stats_label.config(text=f"共 {len(students)} 条记录")
            
        except Exception as e:
            messagebox.showerror("错误", f"查询失败: {str(e)}")
    
    def load_events(self):
        """加载所有可用的运动项目"""
        from database.db_manager import DatabaseManager
        
        try:
            # 创建数据库管理器
            db_manager = DatabaseManager()
            
            # 获取所有项目
            conn = db_manager._get_connection()
            cursor = conn.cursor()
            
            cursor.execute('SELECT name FROM events ORDER BY name')
            events = [row[0] for row in cursor.fetchall()]
            
            conn.close()
            
            # 更新项目下拉框
            self.event_combobox['values'] = ["全部"] + events
            
        except Exception as e:
            print(f"加载项目列表失败: {e}")
    
    def generation_completed(self, success):
        """生成完成后的处理"""
        # 启用按钮
        self.generate_button.config(state=tk.NORMAL)
        
        if success:
            # 启用打开文件按钮
            self.open_file_button.config(state=tk.NORMAL)
            self.open_folder_button.config(state=tk.NORMAL)
            
            # 显示成功消息
            messagebox.showinfo("成功", f"秩序册生成成功！保存在: {self.output_path}")
        else:
            # 显示失败消息
            messagebox.showerror("失败", "秩序册生成失败，请检查错误信息")
    
    def open_generated_file(self):
        """打开生成的文件"""
        if os.path.exists(self.output_path):
            try:
                os.startfile(self.output_path)
                self.update_status(f"已打开文件: {self.output_path}")
            except Exception as e:
                messagebox.showerror("错误", f"无法打开文件: {str(e)}")
        else:
            messagebox.showerror("错误", "文件不存在，请先生成秩序册")
    
    def open_output_folder(self):
        """打开输出文件夹"""
        folder_path = os.path.dirname(os.path.abspath(self.output_path))
        if os.path.exists(folder_path):
            try:
                os.startfile(folder_path)
                self.update_generator_status(f"已打开文件夹: {folder_path}")
            except Exception as e:
                messagebox.showerror("错误", f"无法打开文件夹: {str(e)}")
        else:
            messagebox.showerror("错误", "文件夹不存在")
    
    def start_parsing(self):
        """开始解析Excel文件"""
        if not self.excel_path:
            messagebox.showerror("错误", "请先选择Excel文件")
            return
        
        if not os.path.exists(self.excel_path):
            messagebox.showerror("错误", "选择的文件不存在")
            return
        
        # 禁用按钮
        self.parse_button.config(state=tk.DISABLED)
        self.browse_button.config(state=tk.DISABLED)
        self.preview_button.config(state=tk.DISABLED)
        
        # 重置进度条
        self.parser_progress_var.set(0)
        
        # 清空状态文本
        self.clear_parser_status()
        self.update_parser_status("开始解析Excel文件...")
        
        # 启动线程执行解析任务
        self.is_running = True
        self.thread = threading.Thread(target=self.parse_excel)
        self.thread.daemon = True
        self.thread.start()
    
    def parse_excel(self):
        """解析Excel文件的线程函数"""
        try:
            # 发送初始进度
            self.queue.put(("parser_progress", 0))
            self.queue.put(("status", "开始解析Excel文件..."))
            
            # 创建管理器
            manager = SportsMeetManager(self.excel_path)
            
            # 发送进度更新
            self.queue.put(("parser_progress", 10))
            self.queue.put(("status", "初始化管理器..."))
            
            # 清空数据库
            manager.db_manager.clear_database()
            self.update_parser_status("已清空数据库")
            
            # 发送进度更新
            self.queue.put(("parser_progress", 20))
            self.queue.put(("status", "已清空数据库"))
            
            # 获取年级工作表
            grade_sheets = manager.parser.get_grade_sheets()
            if not grade_sheets:
                self.update_parser_status("未找到年级工作表")
                self.queue.put(("parser_progress", 100))
                self.queue.put(("parse_success", False))
                return
            
            self.update_parser_status(f"找到 {len(grade_sheets)} 个年级工作表")
            
            # 发送进度更新
            self.queue.put(("parser_progress", 30))
            self.queue.put(("status", f"找到 {len(grade_sheets)} 个年级工作表"))
            
            # 解析每个工作表
            total_sheets = len(grade_sheets)
            for i, (grade, sheet_name) in enumerate(grade_sheets.items()):
                progress = int((i + 1) / total_sheets * 60) + 30
                
                self.update_parser_status(f"解析工作表: {sheet_name} ({grade}年级)")
                self.queue.put(("parser_progress", progress))
                self.queue.put(("status", f"解析工作表: {sheet_name} ({grade}年级)"))
                
                df = manager.parser.read_sheet(sheet_name)
                if df is None:
                    self.update_parser_status(f"无法读取工作表: {sheet_name}")
                    continue
                
                # 提取班级信息
                class_name, gender = manager.extractor.extract_class_info(df)
                if not class_name:
                    class_name = f"{grade}年级 1 班"
                    gender = "男"
                    self.update_parser_status(f"使用默认班级信息: {class_name}, {gender}")
                else:
                    self.update_parser_status(f"提取到班级信息: {class_name}, {gender}")
                
                # 提取学生信息
                students = manager.extractor.extract_students(df, grade, class_name, gender)
                self.update_parser_status(f"提取到 {len(students)} 名学生")
                
                # 提取学生参赛项目
                student_events = manager.extractor.extract_student_events(df)
                self.update_parser_status(f"提取到 {len(student_events)} 条参赛记录")
                
                # 存储数据到数据库
                for student in students:
                    student_id = manager.db_manager.add_student(student)
                    
                    # 添加学生参赛项目
                    if student_id and student.number in student_events:
                        for event_name in student_events[student.number]:
                            event_id = manager.db_manager.add_event(event_name)
                            manager.db_manager.add_student_event(student_id, event_id)
                
                self.update_parser_status(f"{grade}年级数据处理完成")
            
            # 发送完成进度
            self.queue.put(("parser_progress", 90))
            self.queue.put(("status", "Excel文件解析完成！"))
            
            self.update_parser_status("Excel文件解析完成！")
            self.queue.put(("parser_progress", 100))
            self.queue.put(("parse_success", True))
            
        except Exception as e:
            self.update_parser_status(f"错误: {str(e)}")
            self.queue.put(("parser_progress", 100))
            self.queue.put(("parse_success", False))
    
    def update_parser_status(self, message):
        """更新Excel解析页面的状态文本"""
        self.parser_status_text.config(state=tk.NORMAL)
        self.parser_status_text.insert(tk.END, message + "\n")
        self.parser_status_text.see(tk.END)
        self.parser_status_text.config(state=tk.DISABLED)
    
    def clear_parser_status(self):
        """清空Excel解析页面的状态文本"""
        self.parser_status_text.config(state=tk.NORMAL)
        self.parser_status_text.delete(1.0, tk.END)
        self.parser_status_text.config(state=tk.DISABLED)
    
    def update_generator_status(self, message):
        """更新秩序册生成页面的状态文本"""
        self.generator_status_text.config(state=tk.NORMAL)
        self.generator_status_text.insert(tk.END, message + "\n")
        self.generator_status_text.see(tk.END)
        self.generator_status_text.config(state=tk.DISABLED)
    
    def clear_generator_status(self):
        """清空秩序册生成页面的状态文本"""
        self.generator_status_text.config(state=tk.NORMAL)
        self.generator_status_text.delete(1.0, tk.END)
        self.generator_status_text.config(state=tk.DISABLED)
    
    def preview_events(self):
        """预览运动项目"""
        if not self.excel_path:
            messagebox.showerror("错误", "请先选择Excel文件")
            return
        
        if not os.path.exists(self.excel_path):
            messagebox.showerror("错误", "选择的文件不存在")
            return
        
        # 清空状态文本
        self.clear_parser_status()
        self.update_parser_status("开始预览运动项目...")
        
        try:
            # 创建管理器
            manager = SportsMeetManager(self.excel_path)
            
            # 提取运动项目
            events = manager.extract_grade_events()
            
            if events is None:
                messagebox.showerror("错误", "无法提取运动项目，请检查Excel文件")
                return
            
            # 清空列表
            self.lower_grades_list.delete(0, tk.END)
            self.upper_grades_list.delete(0, tk.END)
            
            # 添加1-3年级项目
            lower_grades = events.get('lower_grades', [])
            if lower_grades:
                for event in lower_grades:
                    self.lower_grades_list.insert(tk.END, event)
                self.update_parser_status(f"1-3年级共有 {len(lower_grades)} 个运动项目")
            else:
                self.lower_grades_list.insert(tk.END, "无运动项目")
                self.update_parser_status("1-3年级无运动项目")
            
            # 添加4-6年级项目
            upper_grades = events.get('upper_grades', [])
            if upper_grades:
                for event in upper_grades:
                    self.upper_grades_list.insert(tk.END, event)
                self.update_parser_status(f"4-6年级共有 {len(upper_grades)} 个运动项目")
            else:
                self.upper_grades_list.insert(tk.END, "无运动项目")
                self.update_parser_status("4-6年级无运动项目")
            
            # 显示成功消息
            messagebox.showinfo("成功", "运动项目预览完成")
            
        except Exception as e:
            messagebox.showerror("错误", f"预览运动项目时出错: {str(e)}")
            self.update_parser_status(f"错误: {str(e)}")

def main():
    """主函数"""
    print("启动GUI界面...")
    root = tk.Tk()
    print("创建主窗口成功")
    app = OrderBookGUI(root)
    print("GUI应用初始化成功")
    root.mainloop()

if __name__ == "__main__":
    main()