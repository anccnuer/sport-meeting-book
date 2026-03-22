import pandas as pd
from database.models import Student

class DataExtractor:
    def __init__(self):
        pass
    
    def extract_class_info(self, df):
        """从工作表中提取班级和性别信息"""
        class_name = ''
        gender = ''
        
        # 尝试从前几行中查找班级信息
        for i in range(min(5, len(df))):
            try:
                class_info = str(df.iloc[i, 0]).strip()
                if class_info:
                    # 提取班级信息
                    if '班级' in class_info:
                        # 尝试不同的分隔符
                        if '班级：' in class_info:
                            class_part = class_info.split('班级：')[1]
                        elif '班级:' in class_info:
                            class_part = class_info.split('班级:')[1]
                        else:
                            class_part = class_info
                        
                        # 提取班级名称
                        if '参赛组别' in class_part:
                            class_name = class_part.split('参赛组别')[0].strip()
                        else:
                            class_name = class_part.strip()
                    
                    # 提取性别信息
                    if '参赛组别' in class_info:
                        if '参赛组别：' in class_info:
                            gender_part = class_info.split('参赛组别：')[1]
                        elif '参赛组别:' in class_info:
                            gender_part = class_info.split('参赛组别:')[1]
                        else:
                            gender_part = ''
                        gender = gender_part.strip()
                    
                    # 如果都找到了，就退出循环
                    if class_name and gender:
                        break
            except Exception:
                continue
        
        # 处理班级名称，分离年级和班级
        if class_name:
            # 汉字数字映射
            chinese_nums = {
                '一': 1,
                '二': 2,
                '三': 3,
                '四': 4,
                '五': 5,
                '六': 6
            }
            
            # 检查是否包含年级信息
            import re
            
            # 匹配汉字数字+年级+数字+班的格式，如"一年级1班"
            chinese_match = re.search(r'(一|二|三|四|五|六)年级(\d+班)', class_name)
            if chinese_match:
                # 只保留班级部分
                class_name = chinese_match.group(2)
            
            # 匹配阿拉伯数字+年级+数字+班的格式，如"1年级1班"
            arabic_match = re.search(r'(\d+)年级(\d+班)', class_name)
            if arabic_match:
                # 只保留班级部分
                class_name = arabic_match.group(2)
        
        return class_name, gender
    
    def extract_events(self, df):
        """从工作表中提取比赛项目"""
        events = []
        
        # 无效关键词列表
        invalid_keywords = ['参赛项目', '班级', '参赛组别', '号码', '姓名', '性别', 'nan']
        
        # 典型运动项目关键词
        sport_keywords = ['米', '跑', '跳', '球', '接力', '掷', '投', '远', '高', '往返']
        
        # 方法1: 传统方式 - 找到表头行后的下一行
        header_row = -1
        for i, row in df.iterrows():
            if isinstance(row.iloc[0], str) and '号码' in row.iloc[0]:
                header_row = i
                break
        
        if header_row != -1:
            # 尝试从表头行的下一行提取
            event_row = header_row + 1
            if event_row < len(df):
                for col in df.columns[2:]:
                    event_name = str(df.iloc[event_row, df.columns.get_loc(col)]).strip()
                    if self._is_valid_event(event_name, invalid_keywords, sport_keywords):
                        events.append(event_name)
        
        # 如果方法1没有提取到项目，尝试方法2: 从表头行本身提取
        if not events and header_row != -1:
            for col in df.columns[2:]:
                event_name = str(df.iloc[header_row, df.columns.get_loc(col)]).strip()
                if self._is_valid_event(event_name, invalid_keywords, sport_keywords):
                    events.append(event_name)
        
        # 如果方法2也没有提取到项目，尝试方法3: 搜索包含典型运动项目关键词的单元格
        if not events:
            for i, row in df.iterrows():
                for j, cell in enumerate(row):
                    cell_str = str(cell).strip()
                    if self._is_valid_event(cell_str, invalid_keywords, sport_keywords):
                        events.append(cell_str)
        
        # 去重并排序
        events = list(set(events))
        events.sort()
        
        return events
    
    def _is_valid_event(self, event_name, invalid_keywords, sport_keywords):
        """判断是否为有效的运动项目名称"""
        # 检查是否为空或太短
        if not event_name or len(event_name) < 2:
            return False
        
        # 检查是否包含无效关键词
        for keyword in invalid_keywords:
            if keyword in event_name:
                return False
        
        # 检查是否包含运动项目关键词
        has_sport_keyword = any(keyword in event_name for keyword in sport_keywords)
        
        # 对于特殊情况，如"50*20"这样的项目，也应该被视为有效
        if not has_sport_keyword:
            # 检查是否包含数字和特殊字符，可能是运动项目
            import re
            if re.search(r'\d+[*×]\d+', event_name):
                return True
            if re.search(r'\d+米', event_name):
                return True
        
        return has_sport_keyword
    
    def extract_students(self, df, grade, class_name, gender):
        """从工作表中提取学生信息"""
        # 找到表头行
        header_row = 0
        for i, row in df.iterrows():
            if isinstance(row.iloc[0], str) and '号码' in row.iloc[0]:
                header_row = i
                break
        
        event_row = header_row + 1
        students = []
        
        for i in range(event_row + 1, len(df)):
            row = df.iloc[i]
            number = str(row.iloc[0]).strip()
            name = str(row.iloc[1]).strip()
            
            if not number or number == 'nan' or not name or name == 'nan':
                continue
            
            student = Student(
                grade=grade,
                class_name=class_name,
                gender=gender,
                number=number,
                name=name
            )
            students.append(student)
        
        return students
    
    def extract_student_events(self, df):
        """提取学生的参赛项目"""
        # 找到表头行
        header_row = 0
        for i, row in df.iterrows():
            if isinstance(row.iloc[0], str) and '号码' in row.iloc[0]:
                header_row = i
                break
        
        event_row = header_row + 1
        events = self.extract_events(df)
        
        student_events = {}
        
        for i in range(event_row + 1, len(df)):
            row = df.iloc[i]
            number = str(row.iloc[0]).strip()
            
            if not number or number == 'nan':
                continue
            
            student_events[number] = []
            
            for j, event_name in enumerate(events):
                col_index = 2 + j
                if col_index < len(df.columns):
                    value = str(row.iloc[col_index]).strip()
                    if value and value != 'nan':
                        student_events[number].append(event_name)
        
        return student_events
