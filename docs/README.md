# 运动会秩序册生成器

## 项目简介

运动会秩序册生成器是一个用于自动解析Excel格式的运动会报名表，并生成Word格式秩序册的工具。该工具能够智能识别不同年级的比赛项目，处理各种类型的报名标记，并按照年级和项目组织秩序册内容。

## 功能特性

- **智能Excel解析**：自动识别包含年级信息的工作表，提取学生信息和比赛项目
- **灵活的项目识别**：动态解析每个年级的比赛项目，无需硬编码
- **多格式支持**：支持不同类型的勾选标记（通过非空判断）
- **数据存储**：使用SQLite数据库存储报名数据，便于查询和管理
- **自动生成**：一键生成结构清晰的Word格式秩序册
- **模块化设计**：采用软件工程最佳实践，代码结构清晰，易于维护和扩展

## 项目结构

```
秩序册生成器/
├── src/                    # 源代码目录
│   ├── parser/             # Excel解析模块
│   │   ├── __init__.py
│   │   ├── excel_parser.py # Excel文件解析
│   │   └── data_extractor.py # 数据提取逻辑
│   ├── database/           # 数据库模块
│   │   ├── __init__.py
│   │   ├── db_manager.py   # 数据库管理
│   │   └── models.py       # 数据模型定义
│   ├── generator/          # 生成器模块
│   │   ├── __init__.py
│   │   ├── order_book_generator.py # 秩序册生成
│   │   └── word_writer.py  # Word文档写入
│   ├── utils/              # 工具模块
│   │   ├── __init__.py
│   │   └── helpers.py      # 辅助函数
│   └── main.py             # 主入口文件
├── data/                   # 数据目录
│   └── origin_data.xlsx    # 原始Excel数据
├── config/                 # 配置目录
│   └── config.py           # 配置文件
├── tests/                  # 测试目录
│   ├── __init__.py
│   ├── test_parser.py      # 解析模块测试
│   ├── test_database.py    # 数据库模块测试
│   └── test_generator.py   # 生成器模块测试
├── docs/                   # 文档目录
│   └── README.md           # 项目说明
├── requirements.txt        # 依赖包列表
└── setup.py                # 项目安装脚本
```

## 安装说明

### 环境要求

- Python 3.8+
- pip 20.0+

### 安装依赖

```bash
pip install -r requirements.txt
```

### 安装项目

```bash
pip install -e .
```

## 使用方法

### 命令行使用

```bash
# 使用默认Excel文件
python src/main.py

# 使用指定Excel文件
python src/main.py path/to/your/excel/file.xlsx
```

### 作为模块使用

```python
from src.main import SportsMeetManager

# 创建管理器
manager = SportsMeetManager('data/origin_data.xlsx')

# 解析Excel并生成秩序册
manager.parse_and_generate('秩序册.docx')
```

## Excel文件格式要求

1. **工作表命名**：工作表名称应包含年级信息，如"一年级用表"、"二年级用表"等
2. **数据结构**：
   - 第一行：标题信息
   - 第二行：班级和性别信息（格式："班级：X年级X班 参赛组别：男/女"）
   - 第三行：表头（包含"号码"、"姓名"、"参赛项目"等）
   - 第四行：具体比赛项目名称
   - 后续行：学生信息和报名项目（用勾或其他标记表示）

## 输出结果

生成的秩序册将按照以下结构组织：
- 标题：运动会秩序册
- 按年级分组
- 每个年级下按比赛项目分组
- 每个项目下列出参赛学生信息（号码、姓名、班级、性别）

## 测试

运行测试用例：

```bash
python -m pytest tests/ -v
```

## 许可证

本项目采用MIT许可证。

## 贡献

欢迎提交Issue和Pull Request！

## 联系方式

- 作者：Your Name
- 邮箱：your.email@example.com
