import sqlite3
from .models import Student, Event, StudentEvent

class DatabaseManager:
    def __init__(self, db_path='sports_meet.db'):
        self.db_path = db_path
        self.initialize_database()
    
    def initialize_database(self):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # 创建学生表
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS students (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            grade INTEGER,
            class_name TEXT,
            gender TEXT,
            number TEXT,
            name TEXT
        )
        ''')
        
        # 创建事件表
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS events (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE
        )
        ''')
        
        # 创建学生事件关系表
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS student_events (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id INTEGER,
            event_id INTEGER,
            FOREIGN KEY (student_id) REFERENCES students (id),
            FOREIGN KEY (event_id) REFERENCES events (id)
        )
        ''')
        
        conn.commit()
        conn.close()
    
    def add_student(self, student):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
        INSERT OR IGNORE INTO students (grade, class_name, gender, number, name)
        VALUES (?, ?, ?, ?, ?)
        ''', (student.grade, student.class_name, student.gender, student.number, student.name))
        
        conn.commit()
        
        # 获取学生ID
        cursor.execute('SELECT id FROM students WHERE number = ?', (student.number,))
        result = cursor.fetchone()
        if result:
            student_id = result[0]
        else:
            # 如果没有找到，可能是因为INSERT OR IGNORE没有插入新记录
            # 尝试再次查询，确保获取到ID
            cursor.execute('SELECT id FROM students WHERE number = ?', (student.number,))
            result = cursor.fetchone()
            if result:
                student_id = result[0]
            else:
                # 如果仍然没有找到，返回None或抛出异常
                student_id = None
        conn.close()
        
        return student_id
    
    def add_event(self, event_name):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('INSERT OR IGNORE INTO events (name) VALUES (?)', (event_name,))
        conn.commit()
        
        # 获取事件ID
        cursor.execute('SELECT id FROM events WHERE name = ?', (event_name,))
        event_id = cursor.fetchone()[0]
        conn.close()
        
        return event_id
    
    def add_student_event(self, student_id, event_id):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
        INSERT OR IGNORE INTO student_events (student_id, event_id)
        VALUES (?, ?)
        ''', (student_id, event_id))
        
        conn.commit()
        conn.close()
    
    def get_students_by_grade_and_event(self, grade, event_name):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
        SELECT s.number, s.name, s.class_name, s.gender
        FROM students s
        JOIN student_events se ON s.id = se.student_id
        JOIN events e ON se.event_id = e.id
        WHERE s.grade = ? AND e.name = ?
        ORDER BY s.number
        ''', (grade, event_name))
        
        students = cursor.fetchall()
        conn.close()
        
        return students
    
    def get_events_by_grade(self, grade):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
        SELECT DISTINCT e.name
        FROM events e
        JOIN student_events se ON e.id = se.event_id
        JOIN students s ON se.student_id = s.id
        WHERE s.grade = ?
        ''', (grade,))
        
        events = [row[0] for row in cursor.fetchall()]
        conn.close()
        
        return events
    
    def clear_database(self):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('DELETE FROM student_events')
        cursor.execute('DELETE FROM events')
        cursor.execute('DELETE FROM students')
        
        conn.commit()
        conn.close()
    
    def get_students_by_class(self, grade, class_name):
        """按年级和班级查询学生"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
        SELECT s.number, s.name, s.class_name, s.gender
        FROM students s
        WHERE s.grade = ? AND s.class_name = ?
        ORDER BY s.number
        ''', (grade, class_name))
        
        students = cursor.fetchall()
        conn.close()
        
        return students
    
    def get_students_by_gender(self, gender):
        """按性别查询学生"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
        SELECT s.number, s.name, s.class_name, s.gender
        FROM students s
        WHERE s.gender = ?
        ORDER BY s.grade, s.class_name, s.number
        ''', (gender,))
        
        students = cursor.fetchall()
        conn.close()
        
        return students
    
    def get_students_by_event(self, event_name):
        """按项目查询学生"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
        SELECT s.number, s.name, s.class_name, s.gender
        FROM students s
        JOIN student_events se ON s.id = se.student_id
        JOIN events e ON se.event_id = e.id
        WHERE e.name = ?
        ORDER BY s.grade, s.class_name, s.number
        ''', (event_name,))
        
        students = cursor.fetchall()
        conn.close()
        
        return students
    
    def get_students_by_multiple_conditions(self, grade=None, class_name=None, gender=None, event_name=None):
        """按多个条件查询学生"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # 构建查询语句
        query = '''
        SELECT DISTINCT s.number, s.name, s.grade, s.class_name, s.gender, e.name as event
        FROM students s
        LEFT JOIN student_events se ON s.id = se.student_id
        LEFT JOIN events e ON se.event_id = e.id
        WHERE 1=1
        '''
        
        params = []
        
        if grade is not None:
            query += ' AND s.grade = ?'
            params.append(grade)
        
        if class_name is not None:
            query += ' AND s.class_name = ?'
            params.append(class_name)
        
        if gender is not None:
            query += ' AND s.gender = ?'
            params.append(gender)
        
        if event_name is not None:
            query += ' AND e.name = ?'
            params.append(event_name)
        
        query += ' ORDER BY s.grade, s.class_name, s.number, e.name'
        
        cursor.execute(query, params)
        students = cursor.fetchall()
        conn.close()
        
        return students
    
    def _get_connection(self):
        """获取数据库连接"""
        return sqlite3.connect(self.db_path)
