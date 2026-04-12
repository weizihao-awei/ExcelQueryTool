# Excel 筛选处理工具

一款基于 Python + Tkinter 开发的图形界面 Excel 数据处理工具，支持多维度筛选、实时预览和多格式导出。

## 功能特性

### 核心功能

- **Excel 文件导入**
  - 支持 `.xlsx` 和 `.xls` 格式
  - 自动识别所有列和数据
  - 完整展示表格内容

- **多维度精准筛选**
  - 每列独立筛选控件
  - 下拉选择框：自动加载该列所有不重复值
  - 关键字搜索框：支持模糊匹配
  - 多列筛选条件同时生效（交集筛选）
  - 实时刷新筛选结果

- **结果展示与操作**
  - 表格形式实时显示筛选结果
  - 支持多行选择
  - 状态栏显示数据统计

- **多格式导出**
  - 导出全部筛选结果
  - 导出选中的行
  - 支持 Excel (`.xlsx`) 格式
  - 支持 Markdown (`.md`) 表格格式

### 界面特点

- 简洁直观的操作界面
- 响应快速的筛选操作
- 完整的中文支持
- 操作流程：打开 → 筛选 → 查看 → 导出

## 环境要求

- Python 3.7+
- 依赖库：
  - pandas >= 1.3.0
  - openpyxl >= 3.0.0

## 安装使用

### 1. 克隆或下载项目

```bash
git clone <repository-url>
cd ExcelQueryTool
```

### 2. 安装依赖

```bash
pip install -r requirements.txt
```

### 3. 运行程序

```bash
python excel_filter_tool.py
```

## 使用说明

### 基本操作流程

1. **打开文件**
   - 点击「📂 打开 Excel 文件」按钮
   - 选择本地的 `.xlsx` 或 `.xls` 文件

2. **设置筛选条件**
   - 每列下方有两个筛选控件：
     - **下拉框**：选择该列的特定值
     - **输入框**：输入关键字进行模糊搜索
   - 多列条件自动取交集

3. **查看结果**
   - 筛选结果实时显示在下方的表格中
   - 状态栏显示当前显示的行数

4. **导出数据**
   - 「📥 导出全部结果」：导出所有筛选后的数据
   - 「📥 导出选中行」：先选中行，再导出选中部分
   - 支持导出为 Excel 或 Markdown 格式

### 快捷键

- `Ctrl + O`：打开文件
- `Ctrl + E`：导出数据
- `Delete`：清除筛选条件

## 打包为可执行文件

使用 PyInstaller 打包成独立的 exe 文件：

```bash
# 安装 PyInstaller
pip install pyinstaller

# 打包（单文件模式）
pyinstaller --onefile --windowed --name "Excel筛选工具" excel_filter_tool.py

# 或打包（包含依赖）
pyinstaller --windowed --name "Excel筛选工具" excel_filter_tool.py
```

打包后的文件位于 `dist` 目录下。

## 项目结构

```
ExcelQueryTool/
├── excel_filter_tool.py    # 主程序文件
├── requirements.txt        # 依赖列表
└── README.md              # 项目说明文档
```

## 技术说明

- **GUI 框架**：Tkinter（Python 内置，无需额外安装）
- **数据处理**：Pandas
- **Excel 读写**：openpyxl
- **性能优化**：大数据集采用分批加载，最多显示 1000 行预览

## 注意事项

1. 筛选下拉框最多显示每列前 50 个唯一值
2. 数据预览最多显示 1000 行，导出时不受此限制
3. 建议处理数据量不超过 10 万行的 Excel 文件

## 许可证

MIT License

## 更新日志

### v1.0.0 (2024-XX-XX)
- 初始版本发布
- 实现基本的 Excel 导入、筛选、导出功能
- 支持 Excel 和 Markdown 两种导出格式
