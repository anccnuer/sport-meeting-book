import sys
import os

# 添加src目录到Python路径
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'src'))

from parser.excel_parser import ExcelParser
from parser.data_extractor import DataExtractor

def test_excel_parser():
    """测试Excel解析器"""
    print("测试Excel解析器...")
    excel_path = 'data/origin_data.xlsx'
    
    parser = ExcelParser(excel_path)
    
    # 测试打开Excel文件
    assert parser.open_excel(), "打开Excel文件失败"
    print("Excel文件打开成功")
    
    # 测试获取工作表名称
    sheet_names = parser.get_sheet_names()
    assert len(sheet_names) > 0, "获取工作表名称失败"
    print(f"获取到 {len(sheet_names)} 个工作表")
    for sheet_name in sheet_names:
        print(f"  - {sheet_name}")
    
    # 测试获取年级工作表
    grade_sheets = parser.get_grade_sheets()
    assert len(grade_sheets) > 0, "获取年级工作表失败"
    print(f"获取到 {len(grade_sheets)} 个年级工作表")
    for grade, sheet_name in grade_sheets.items():
        print(f"  - 年级 {grade}: {sheet_name}")
    
    # 测试读取工作表
    for grade, sheet_name in grade_sheets.items():
        df = parser.read_sheet(sheet_name)
        assert df is not None, f"读取工作表 {sheet_name} 失败"
        print(f"工作表 {sheet_name} 读取成功，共 {len(df)} 行")

def test_data_extractor():
    """测试数据提取器"""
    print("测试数据提取器...")
    excel_path = 'data/origin_data.xlsx'
    
    parser = ExcelParser(excel_path)
    extractor = DataExtractor()
    
    grade_sheets = parser.get_grade_sheets()
    
    for grade, sheet_name in grade_sheets.items():
        df = parser.read_sheet(sheet_name)
        
        # 测试提取班级信息
        class_name, gender = extractor.extract_class_info(df)
        if not class_name:
            class_name = f"{grade}年级 1 班"
            gender = "男"
            print(f"{sheet_name} - 班级信息提取失败，使用默认值: {class_name}, 性别: {gender}")
        else:
            print(f"{sheet_name} - 班级: {class_name}, 性别: {gender}")
        
        # 测试提取比赛项目
        events = extractor.extract_events(df)
        assert len(events) > 0, "提取比赛项目失败"
        print(f"{sheet_name} - 比赛项目: {', '.join(events)}")
        
        # 测试提取学生信息
        students = extractor.extract_students(df, grade, class_name, gender)
        assert len(students) > 0, "提取学生信息失败"
        print(f"{sheet_name} - 学生数量: {len(students)}")
        
        # 测试提取学生参赛项目
        student_events = extractor.extract_student_events(df)
        assert len(student_events) > 0, "提取学生参赛项目失败"
        print(f"{sheet_name} - 学生参赛项目: {len(student_events)} 条记录")

if __name__ == "__main__":
    test_excel_parser()
    test_data_extractor()
    print("所有Excel解析测试通过！")
