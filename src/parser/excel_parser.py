import pandas as pd

class ExcelParser:
    def __init__(self, excel_path):
        self.excel_path = excel_path
        self.xls = None
    
    def open_excel(self):
        """打开Excel文件"""
        try:
            self.xls = pd.ExcelFile(self.excel_path)
            return True
        except Exception as e:
            print(f"打开Excel文件失败: {e}")
            return False
    
    def get_sheet_names(self):
        """获取所有工作表名称"""
        if not self.xls:
            if not self.open_excel():
                return []
        return self.xls.sheet_names
    
    def get_grade_sheets(self):
        """获取包含年级信息的工作表"""
        sheet_names = self.get_sheet_names()
        grade_sheets = {}
        
        for sheet_name in sheet_names:
            if '年级' in sheet_name:
                grade = self._extract_grade(sheet_name)
                if grade:
                    grade_sheets[grade] = sheet_name
                else:
                    print(f"无法从工作表名称 '{sheet_name}' 中提取年级")
        
        return grade_sheets
    
    def _extract_grade(self, sheet_name):
        """从工作表名称中提取年级"""
        # 汉字数字映射
        chinese_nums = {
            '一': 1,
            '二': 2,
            '三': 3,
            '四': 4,
            '五': 5,
            '六': 6
        }
        
        # 查找汉字数字
        for char in sheet_name:
            if char in chinese_nums:
                return chinese_nums[char]
        
        # 查找阿拉伯数字
        import re
        match = re.search(r'\d+', sheet_name)
        if match:
            return int(match.group())
        
        return None
    
    def read_sheet(self, sheet_name):
        """读取指定工作表"""
        if not self.xls:
            if not self.open_excel():
                return None
        
        try:
            df = pd.read_excel(self.excel_path, sheet_name=sheet_name)
            return df
        except Exception as e:
            print(f"读取工作表失败: {e}")
            return None
