# 秩序册生成器项目 - 目录结构优化方案

## 项目背景
当前项目结构较为简单，所有功能都集中在单个main.py文件中，不利于代码维护和扩展。需要按照软件工程最佳实践重新组织项目结构，实现功能解耦。

## 目标
创建一个模块化、可维护、可扩展的项目结构，将不同功能分离到独立模块中。

## 新目录结构设计

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

## 任务分解与优先级

### [x] 任务1: 创建基本目录结构
- **Priority**: P0
- **Depends On**: None
- **Description**: 创建项目所需的目录结构，包括src、config、tests、docs等目录
- **Success Criteria**: 所有目录结构创建完成
- **Test Requirements**:
  - `programmatic` TR-1.1: 检查所有目录是否存在
  - `human-judgement` TR-1.2: 目录结构清晰合理

### [x] 任务2: 实现数据模型和数据库管理
- **Priority**: P0
- **Depends On**: 任务1
- **Description**: 创建数据库模型和数据库管理模块，负责数据的存储和查询
- **Success Criteria**:
  - 数据库表结构创建完成
  - 数据库操作功能实现
- **Test Requirements**:
  - `programmatic` TR-2.1: 数据库表创建成功
  - `programmatic` TR-2.2: 数据插入和查询功能正常

### [x] 任务3: 实现Excel解析模块
- **Priority**: P0
- **Depends On**: 任务1
- **Description**: 创建Excel解析模块，负责读取和解析Excel文件中的数据
- **Success Criteria**:
  - 能够正确读取Excel文件
  - 能够解析不同年级的比赛项目
  - 能够处理不同类型的勾选标记
- **Test Requirements**:
  - `programmatic` TR-3.1: 正确读取Excel文件
  - `programmatic` TR-3.2: 正确解析比赛项目
  - `programmatic` TR-3.3: 正确处理勾选标记

### [x] 任务4: 实现秩序册生成模块
- **Priority**: P0
- **Depends On**: 任务1
- **Description**: 创建秩序册生成模块，负责将数据转换为Word格式的秩序册
- **Success Criteria**:
  - 能够生成Word格式的秩序册
  - 秩序册内容结构清晰
  - 支持按年级和项目组织内容
- **Test Requirements**:
  - `programmatic` TR-4.1: 生成Word文件成功
  - `human-judgement` TR-4.2: 秩序册内容结构清晰

### [x] 任务5: 实现主入口文件
- **Priority**: P0
- **Depends On**: 任务2, 任务3, 任务4
- **Description**: 创建主入口文件，整合各个模块的功能
- **Success Criteria**:
  - 主入口文件能够协调各个模块的工作
  - 能够完成从Excel解析到秩序册生成的完整流程
- **Test Requirements**:
  - `programmatic` TR-5.1: 完整流程执行成功
  - `programmatic` TR-5.2: 生成秩序册文件

### [x] 任务6: 创建配置文件和依赖管理
- **Priority**: P1
- **Depends On**: 任务1
- **Description**: 创建配置文件和依赖管理文件
- **Success Criteria**:
  - 配置文件创建完成
  - 依赖包列表创建完成
- **Test Requirements**:
  - `programmatic` TR-6.1: 配置文件存在
  - `programmatic` TR-6.2: 依赖包列表存在

### [x] 任务7: 编写测试用例
- **Priority**: P1
- **Depends On**: 任务2, 任务3, 任务4
- **Description**: 为各个模块编写测试用例
- **Success Criteria**:
  - 测试用例覆盖主要功能
  - 测试通过
- **Test Requirements**:
  - `programmatic` TR-7.1: 测试用例执行成功
  - `programmatic` TR-7.2: 测试覆盖率达到80%以上

### [x] 任务8: 编写项目文档
- **Priority**: P2
- **Depends On**: 任务1
- **Description**: 编写项目说明文档
- **Success Criteria**:
  - 项目文档创建完成
  - 文档内容完整清晰
- **Test Requirements**:
  - `human-judgement` TR-8.1: 文档内容完整
  - `human-judgement` TR-8.2: 文档结构清晰

## 实施步骤
1. 首先创建基本目录结构
2. 实现数据库模块和数据模型
3. 实现Excel解析模块
4. 实现秩序册生成模块
5. 实现主入口文件
6. 创建配置文件和依赖管理
7. 编写测试用例
8. 编写项目文档

## 技术栈
- Python 3.13
- pandas: 用于Excel文件解析
- openpyxl: 用于Excel文件读取
- python-docx: 用于生成Word文档
- SQLite: 用于数据存储

## 预期成果
- 模块化的项目结构
- 清晰的功能分离
- 可维护和可扩展的代码
- 完整的测试覆盖
- 详细的项目文档