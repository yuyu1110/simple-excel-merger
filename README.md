# Excel 表格合并

一个简单的 Python 办公自动化命令行工具。

## 快速开始

需要 Python 3.10 或以上。在项目目录打开终端，建议先建立虚拟环境：

```bash
python -m venv .venv
```

Windows：`.venv\Scripts\activate`；macOS/Linux：`source .venv/bin/activate`。然后运行：

```bash
python -m pip install -r requirements.txt
python examples/create_demo.py
python merge_excel.py demo merged.xlsx --dedupe
```

示例生成器只创建虚构数据，并在 demo 目录已存在时退出，避免覆盖。

## 功能与边界

只支持 .xlsx，读取每个文件的第一张工作表，第一行为表头，名称与顺序必须完全一致且不能空白或重复。跳过空行及 Excel 临时文件；--dedupe 按整行值去重。输出已有时拒绝覆盖。仅合并数据值，不保留样式、图表、宏或其他工作表；数据行含公式时拒绝合并，请先粘贴为值。适用于小型办公表格，数据行存储于内存。示例的两个表共四行，去重后输出三行。

查看全部参数：`python merge_excel.py --help`。

## 测试

```bash
python -m unittest discover -s tests -v
```

测试使用临时目录，不修改个人文件。

## 项目结构

- `merge_excel.py`：核心功能及命令行入口
- `examples/create_demo.py`：生成示例输入
- `tests/test_tool.py`：功能及异常场景测试

## 简历表述参考

使用 openpyxl 实现多工作簿数据合并、表头校验及整行去重，并为错误输入编写测试。

这是学习与练习项目；请在理解实现后按实际参与情况描述，不包含生产使用或性能提升声明。
