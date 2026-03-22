from .word_writer import WordWriter
from database.db_manager import DatabaseManager

class OrderBookGenerator:
    def __init__(self, db_path='sports_meet.db'):
        self.db_manager = DatabaseManager(db_path)
        self.writer = WordWriter()
    
    def generate_order_book(self, output_path='秩序册.docx'):
        """生成秩序册"""
        # 设置字体
        self.writer.set_font()
        
        # 添加标题
        self.writer.add_title('运动会秩序册')
        self.writer.add_paragraph()
        
        # 按年级生成秩序
        grades = self._get_all_grades()
        for grade in sorted(grades):
            self._generate_grade_section(grade)
        
        # 保存文档
        self.writer.save(output_path)
        print(f'秩序册已生成：{output_path}')
        return True
    
    def _get_all_grades(self):
        """获取所有年级"""
        conn = self.db_manager._get_connection()
        cursor = conn.cursor()
        
        cursor.execute('SELECT DISTINCT grade FROM students ORDER BY grade')
        grades = [row[0] for row in cursor.fetchall()]
        
        conn.close()
        return grades
    
    def _generate_grade_section(self, grade):
        """生成年级部分"""
        # 添加年级标题
        self.writer.add_title(f'{grade}年级', level=1)
        
        # 获取该年级的所有比赛项目
        events = self.db_manager.get_events_by_grade(grade)
        
        for event_name in events:
            self._generate_event_section(grade, event_name)
    
    def _generate_event_section(self, grade, event_name):
        """生成比赛项目部分"""
        # 添加项目标题
        self.writer.add_title(event_name, level=2)
        
        # 获取参加该项目的学生
        students = self.db_manager.get_students_by_grade_and_event(grade, event_name)
        
        if students:
            # 创建表格
            table = self.writer.add_table(rows=1, cols=4)
            hdr_cells = table.rows[0].cells
            hdr_cells[0].text = '号码'
            hdr_cells[1].text = '姓名'
            hdr_cells[2].text = '班级'
            hdr_cells[3].text = '性别'
            
            # 填充表格
            for student in students:
                row_cells = table.add_row().cells
                row_cells[0].text = student[0]
                row_cells[1].text = student[1]
                row_cells[2].text = student[2]
                row_cells[3].text = student[3]
        
        # 添加空行
        self.writer.add_paragraph()
