# 秩序册生成器 - 三页面GUI实现计划

## 1. 项目分析

### 当前状态
- 项目使用tkinter实现了基本的GUI界面
- 目前只有一个页面，包含文件选择、操作按钮、运动项目预览和状态显示
- 数据库功能已实现基本操作，但需要扩展查询功能
- 秩序册生成功能已实现

### 目标
将GUI界面分成三个独立页面：
1. **Excel解析页面**：选择并解析excel报名表，显示处理过程中的打印内容
2. **数据查询页面**：展示数据库中的各条报名数据，支持按年级、班级、性别、项目查询
3. **秩序册生成页面**：根据数据库内容生成秩序册

## 2. 实现计划

### [ ] 任务1：重构GUI主框架，实现三页面切换
- **Priority**: P0
- **Depends On**: None
- **Description**:
  - 修改main_window.py，使用ttk.Notebook创建三个选项卡页面
  - 设计页面布局和导航结构
  - 实现页面间的切换功能
- **Success Criteria**:
  - GUI界面显示三个选项卡页面
  - 页面切换功能正常
- **Test Requirements**:
  - `programmatic` TR-1.1: 启动GUI后能看到三个选项卡页面
  - `human-judgement` TR-1.2: 页面布局合理，切换流畅

### [ ] 任务2：实现Excel解析页面
- **Priority**: P0
- **Depends On**: 任务1
- **Description**:
  - 实现Excel文件选择功能
  - 实现解析Excel文件的功能
  - 显示解析过程中的详细信息
  - 实现解析完成后的状态反馈
- **Success Criteria**:
  - 能够选择Excel文件并解析
  - 解析过程中显示详细的处理信息
  - 解析完成后显示成功或失败的状态
- **Test Requirements**:
  - `programmatic` TR-2.1: 选择Excel文件后能正确解析
  - `programmatic` TR-2.2: 解析过程中显示实时处理信息
  - `human-judgement` TR-2.3: 界面布局清晰，信息展示完整

### [ ] 任务3：扩展数据库查询功能
- **Priority**: P1
- **Depends On**: None
- **Description**:
  - 在db_manager.py中添加按年级、班级、性别、项目查询的功能
  - 实现复合条件查询
  - 优化查询性能
- **Success Criteria**:
  - 能够按年级查询学生数据
  - 能够按班级查询学生数据
  - 能够按性别查询学生数据
  - 能够按项目查询学生数据
  - 能够组合多个条件进行查询
- **Test Requirements**:
  - `programmatic` TR-3.1: 按年级查询返回正确结果
  - `programmatic` TR-3.2: 按班级查询返回正确结果
  - `programmatic` TR-3.3: 按性别查询返回正确结果
  - `programmatic` TR-3.4: 按项目查询返回正确结果
  - `programmatic` TR-3.5: 复合条件查询返回正确结果

### [ ] 任务4：实现数据查询页面
- **Priority**: P0
- **Depends On**: 任务1, 任务3
- **Description**:
  - 设计查询条件输入界面
  - 实现查询结果展示功能
  - 添加查询按钮和结果刷新功能
  - 实现查询结果的导出功能（可选）
- **Success Criteria**:
  - 能够输入查询条件
  - 能够执行查询并显示结果
  - 查询结果显示清晰、格式正确
- **Test Requirements**:
  - `programmatic` TR-4.1: 输入查询条件后能正确执行查询
  - `programmatic` TR-4.2: 查询结果显示正确
  - `human-judgement` TR-4.3: 界面布局合理，操作便捷

### [ ] 任务5：实现秩序册生成页面
- **Priority**: P0
- **Depends On**: 任务1
- **Description**:
  - 实现秩序册生成的配置选项
  - 实现生成按钮和进度显示
  - 实现生成完成后的文件打开功能
  - 添加生成历史记录（可选）
- **Success Criteria**:
  - 能够配置秩序册生成选项
  - 能够启动生成过程并显示进度
  - 生成完成后能够打开生成的文件
- **Test Requirements**:
  - `programmatic` TR-5.1: 点击生成按钮后能正确生成秩序册
  - `programmatic` TR-5.2: 生成过程中显示实时进度
  - `human-judgement` TR-5.3: 界面布局清晰，操作便捷

### [ ] 任务6：优化整体GUI体验
- **Priority**: P2
- **Depends On**: 任务1-5
- **Description**:
  - 优化界面布局和响应式设计
  - 添加错误处理和用户提示
  - 优化操作流程，提高用户体验
  - 添加键盘快捷键（可选）
- **Success Criteria**:
  - GUI界面美观、易用
  - 操作流程顺畅
  - 错误处理合理，用户提示清晰
- **Test Requirements**:
  - `human-judgement` TR-6.1: 界面美观，布局合理
  - `human-judgement` TR-6.2: 操作流程顺畅，响应及时
  - `human-judgement` TR-6.3: 错误处理合理，用户提示清晰

## 3. 技术实现细节

### 页面结构
- 使用ttk.Notebook创建三个选项卡：
  1. "Excel解析"
  2. "数据查询"
  3. "秩序册生成"

### 数据库查询功能扩展
- 在DatabaseManager类中添加以下方法：
  - get_students_by_class(grade, class_name)
  - get_students_by_gender(gender)
  - get_students_by_event(event_name)
  - get_students_by_multiple_conditions(grade=None, class_name=None, gender=None, event_name=None)

### GUI组件
- **Excel解析页面**：
  - 文件选择按钮和路径显示
  - 解析按钮
  - 处理过程信息显示文本框
  - 解析状态和结果显示

- **数据查询页面**：
  - 年级选择下拉框
  - 班级选择下拉框
  - 性别选择下拉框
  - 项目选择下拉框
  - 查询按钮
  - 查询结果表格
  - 结果统计信息

- **秩序册生成页面**：
  - 生成配置选项（如包含哪些年级、项目等）
  - 生成按钮
  - 进度条
  - 生成状态显示
  - 打开文件按钮
  - 打开文件夹按钮

## 4. 测试计划

### 功能测试
- 测试Excel解析功能
- 测试数据查询功能
- 测试秩序册生成功能
- 测试页面切换功能

### 边界测试
- 测试空Excel文件
- 测试格式错误的Excel文件
- 测试无数据的查询条件
- 测试大数据量的查询

### 性能测试
- 测试解析大型Excel文件的性能
- 测试生成大型秩序册的性能
- 测试复杂查询的性能

## 5. 预期完成时间

- 任务1-2：1天
- 任务3-4：1天
- 任务5-6：1天
- 总计：3天

## 6. 风险评估

- **风险1**：Excel文件格式变化可能导致解析失败
  - 缓解措施：添加更健壮的错误处理和格式检查

- **风险2**：数据库查询性能问题
  - 缓解措施：优化SQL查询，添加索引

- **风险3**：GUI响应速度问题
  - 缓解措施：使用多线程处理耗时操作，避免阻塞主线程

- **风险4**：用户操作错误
  - 缓解措施：添加合理的用户提示和错误处理