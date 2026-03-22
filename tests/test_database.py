import sys
import os
import sqlite3

# 添加src目录到Python路径
sys.path.append(os.path.join(os.path.dirname(__file__), '..', 'src'))

from database.db_manager import DatabaseManager
from database.models import Student

def test_database_initialization():
    """测试数据库初始化"""
    print("测试数据库初始化...")
    db_path = 'test_sports_meet.db'
    
    # 清理测试数据库
    if os.path.exists(db_path):
        os.remove(db_path)
    
    db_manager = DatabaseManager(db_path)  # 使用文件数据库进行测试
    
    # 检查表是否创建成功
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='students'")
    assert cursor.fetchone() is not None, "学生表未创建"
    
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='events'")
    assert cursor.fetchone() is not None, "事件表未创建"
    
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='student_events'")
    assert cursor.fetchone() is not None, "学生事件关系表未创建"
    
    conn.close()
    
    # 清理测试数据库
    if os.path.exists(db_path):
        os.remove(db_path)
    
    print("数据库初始化测试通过！")

def test_add_student():
    """测试添加学生"""
    print("测试添加学生...")
    db_path = 'test_sports_meet.db'
    
    # 清理测试数据库
    if os.path.exists(db_path):
        os.remove(db_path)
    
    db_manager = DatabaseManager(db_path)
    
    student = Student(grade=1, class_name="一年级 1 班", gender="男", number="1101", name="王廷睿")
    student_id = db_manager.add_student(student)
    
    assert student_id is not None, "学生添加失败"
    print(f"学生添加成功，ID: {student_id}")
    
    # 清理测试数据库
    if os.path.exists(db_path):
        os.remove(db_path)

def test_add_event():
    """测试添加事件"""
    print("测试添加事件...")
    db_path = 'test_sports_meet.db'
    
    # 清理测试数据库
    if os.path.exists(db_path):
        os.remove(db_path)
    
    db_manager = DatabaseManager(db_path)
    
    event_id = db_manager.add_event("50米")
    assert event_id is not None, "事件添加失败"
    print(f"事件添加成功，ID: {event_id}")
    
    # 清理测试数据库
    if os.path.exists(db_path):
        os.remove(db_path)

def test_add_student_event():
    """测试添加学生事件关系"""
    print("测试添加学生事件关系...")
    db_path = 'test_sports_meet.db'
    
    # 清理测试数据库
    if os.path.exists(db_path):
        os.remove(db_path)
    
    db_manager = DatabaseManager(db_path)
    
    # 添加学生
    student = Student(grade=1, class_name="一年级 1 班", gender="男", number="1101", name="王廷睿")
    student_id = db_manager.add_student(student)
    
    # 添加事件
    event_id = db_manager.add_event("50米")
    
    # 添加学生事件关系
    db_manager.add_student_event(student_id, event_id)
    print("学生事件关系添加成功！")
    
    # 清理测试数据库
    if os.path.exists(db_path):
        os.remove(db_path)

def test_get_students_by_grade_and_event():
    """测试按年级和事件查询学生"""
    print("测试按年级和事件查询学生...")
    db_path = 'test_sports_meet.db'
    
    # 清理测试数据库
    if os.path.exists(db_path):
        os.remove(db_path)
    
    db_manager = DatabaseManager(db_path)
    
    # 添加学生
    student1 = Student(grade=1, class_name="一年级 1 班", gender="男", number="1101", name="王廷睿")
    student2 = Student(grade=1, class_name="一年级 1 班", gender="男", number="1102", name="李承儒")
    student_id1 = db_manager.add_student(student1)
    student_id2 = db_manager.add_student(student2)
    
    # 添加事件
    event_id = db_manager.add_event("50米")
    
    # 添加学生事件关系
    db_manager.add_student_event(student_id1, event_id)
    db_manager.add_student_event(student_id2, event_id)
    
    # 查询学生
    students = db_manager.get_students_by_grade_and_event(1, "50米")
    assert len(students) == 2, "查询学生失败"
    print(f"查询到 {len(students)} 名学生")
    for student in students:
        print(f"  {student[0]} - {student[1]} - {student[2]} - {student[3]}")
    
    # 清理测试数据库
    if os.path.exists(db_path):
        os.remove(db_path)

if __name__ == "__main__":
    test_database_initialization()
    test_add_student()
    test_add_event()
    test_add_student_event()
    test_get_students_by_grade_and_event()
    print("所有数据库测试通过！")
