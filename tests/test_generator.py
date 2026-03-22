import sys
import os

# 添加src目录到Python路径
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'src'))

from generator.order_book_generator import OrderBookGenerator
from database.db_manager import DatabaseManager
from database.models import Student

def test_order_book_generator():
    """测试秩序册生成器"""
    print("测试秩序册生成器...")
    
    # 准备测试数据
    db_path = 'test_sports_meet.db'
    
    # 清理测试数据库
    if os.path.exists(db_path):
        os.remove(db_path)
    
    # 初始化数据库
    db_manager = DatabaseManager(db_path)
    
    # 添加测试数据
    student1 = Student(grade=1, class_name="一年级 1 班", gender="男", number="1101", name="王廷睿")
    student2 = Student(grade=1, class_name="一年级 1 班", gender="男", number="1102", name="李承儒")
    student3 = Student(grade=2, class_name="二年级 1 班", gender="男", number="2101", name="张彧褀")
    
    student_id1 = db_manager.add_student(student1)
    student_id2 = db_manager.add_student(student2)
    student_id3 = db_manager.add_student(student3)
    
    event_id1 = db_manager.add_event("50米")
    event_id2 = db_manager.add_event("150米")
    
    db_manager.add_student_event(student_id1, event_id1)
    db_manager.add_student_event(student_id2, event_id1)
    db_manager.add_student_event(student_id3, event_id2)
    
    # 生成秩序册
    generator = OrderBookGenerator(db_path)
    output_path = 'test_秩序册.docx'
    generator.generate_order_book(output_path)
    
    # 验证秩序册是否生成
    assert os.path.exists(output_path), "秩序册生成失败"
    print(f"秩序册生成成功: {output_path}")
    
    # 清理测试文件
    if os.path.exists(db_path):
        os.remove(db_path)
    if os.path.exists(output_path):
        os.remove(output_path)

if __name__ == "__main__":
    test_order_book_generator()
    print("秩序册生成测试通过！")
