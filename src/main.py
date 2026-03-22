import os
import sys

# 添加src目录到Python路径
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from parser.excel_parser import ExcelParser
from parser.data_extractor import DataExtractor
from database.db_manager import DatabaseManager
from database.models import Student
from generator.order_book_generator import OrderBookGenerator

class SportsMeetManager:
    def __init__(self, excel_path, db_path='sports_meet.db'):
        self.excel_path = excel_path
        self.db_path = db_path
        self.parser = ExcelParser(excel_path) if excel_path else None
        self.extractor = DataExtractor()
        self.db_manager = DatabaseManager(db_path)
        self.generator = OrderBookGenerator(db_path)
    
    def extract_grade_events(self):
        """提取并分类年级运动项目"""
        print("提取年级运动项目...")
        
        if not self.parser:
            print("未初始化Excel解析器")
            return None
        
        # 获取年级工作表
        grade_sheets = self.parser.get_grade_sheets()
        if not grade_sheets:
            print("未找到年级工作表")
            return None
        
        # 存储各年级的运动项目
        grade_events = {}
        
        # 解析每个工作表
        for grade, sheet_name in grade_sheets.items():
            print(f"解析工作表: {sheet_name}")
            df = self.parser.read_sheet(sheet_name)
            if df is None:
                print(f"无法读取工作表: {sheet_name}")
                continue
            
            # 提取运动项目
            events = self.extractor.extract_events(df)
            if events:
                grade_events[grade] = events
                print(f"{grade}年级提取到 {len(events)} 个运动项目")
        
        # 分类运动项目
        lower_grade_events = []  # 1-3年级
        upper_grade_events = []  # 4-6年级
        
        for grade, events in grade_events.items():
            if 1 <= grade <= 3:
                lower_grade_events.extend(events)
            elif 4 <= grade <= 6:
                upper_grade_events.extend(events)
        
        # 去重
        lower_grade_events = list(set(lower_grade_events))
        upper_grade_events = list(set(upper_grade_events))
        
        # 排序
        lower_grade_events.sort()
        upper_grade_events.sort()
        
        return {
            'lower_grades': lower_grade_events,  # 1-3年级
            'upper_grades': upper_grade_events   # 4-6年级
        }
    
    def parse_and_generate(self, output_path='秩序册.docx'):
        """解析Excel并生成秩序册"""
        print("开始解析Excel文件...")
        
        if not self.parser:
            print("未初始化Excel解析器")
            return False
        
        # 清理数据库
        self.db_manager.clear_database()
        
        # 获取年级工作表
        grade_sheets = self.parser.get_grade_sheets()
        if not grade_sheets:
            print("未找到年级工作表")
            return False
        
        print(f"找到 {len(grade_sheets)} 个年级工作表")
        
        # 解析每个工作表
        for grade, sheet_name in grade_sheets.items():
            print(f"解析工作表: {sheet_name}")
            df = self.parser.read_sheet(sheet_name)
            if df is None:
                print(f"无法读取工作表: {sheet_name}")
                continue
            
            # 提取班级信息
            class_name, gender = self.extractor.extract_class_info(df)
            if not class_name:
                class_name = f"{grade}年级 1 班"
                gender = "男"
                print(f"使用默认班级信息: {class_name}, {gender}")
            
            # 提取学生信息
            students = self.extractor.extract_students(df, grade, class_name, gender)
            print(f"提取到 {len(students)} 名学生")
            
            # 提取学生参赛项目
            student_events = self.extractor.extract_student_events(df)
            print(f"提取到 {len(student_events)} 条参赛记录")
            
            # 存储数据到数据库
            for student in students:
                student_id = self.db_manager.add_student(student)
                
                # 添加学生参赛项目
                if student.number in student_events:
                    for event_name in student_events[student.number]:
                        event_id = self.db_manager.add_event(event_name)
                        self.db_manager.add_student_event(student_id, event_id)
        
        # 生成秩序册
        print("生成秩序册...")
        self.generator.generate_order_book(output_path)
        
        return True

def main():
    # 获取Excel文件路径
    if len(sys.argv) > 1:
        excel_path = sys.argv[1]
    else:
        excel_path = 'data/origin_data.xlsx'
    
    # 检查文件是否存在
    if not os.path.exists(excel_path):
        print(f"Excel文件不存在: {excel_path}")
        sys.exit(1)
    
    # 创建管理器并执行
    manager = SportsMeetManager(excel_path)
    manager.parse_and_generate()

if __name__ == '__main__':
    main()
