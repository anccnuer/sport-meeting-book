import os
import sys

# 添加src目录到Python路径
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from parser.excel_parser import ExcelParser
from parser.data_extractor import DataExtractor
from src.main import SportsMeetManager

# 测试Excel文件路径
excel_path = 'data/origin_data.xlsx'

print("测试1-3年级运动项目解析修复...")
print(f"使用Excel文件: {excel_path}")

# 测试1: 测试DataExtractor.extract_events方法
print("\n测试1: 测试DataExtractor.extract_events方法")
parser = ExcelParser(excel_path)
grade_sheets = parser.get_grade_sheets()
print(f"找到的年级工作表: {grade_sheets}")

extractor = DataExtractor()

for grade, sheet_name in grade_sheets.items():
    print(f"\n解析工作表: {sheet_name} (年级: {grade})")
    df = parser.read_sheet(sheet_name)
    if df is not None:
        events = extractor.extract_events(df)
        print(f"提取到的运动项目: {events}")
        print(f"项目数量: {len(events)}")
    else:
        print("无法读取工作表")

# 测试2: 测试SportsMeetManager.extract_grade_events方法
print("\n测试2: 测试SportsMeetManager.extract_grade_events方法")
manager = SportsMeetManager(excel_path)
events = manager.extract_grade_events()

if events:
    print("\n1-3年级运动项目:")
    lower_grades_events = events.get('lower_grades', [])
    for event in lower_grades_events:
        print(f"- {event}")
    print(f"1-3年级项目数量: {len(lower_grades_events)}")
    
    print("\n4-6年级运动项目:")
    upper_grades_events = events.get('upper_grades', [])
    for event in upper_grades_events:
        print(f"- {event}")
    print(f"4-6年级项目数量: {len(upper_grades_events)}")
else:
    print("无法提取运动项目")

print("\n测试完成！")