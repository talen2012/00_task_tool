# 小工具使用指南

1. 配置环境
    - 安装python3.9或更高版本  
    - 本工程根目录（00_task_tool）下创建虚拟环境  
    ```python -m venv .venv```
    - 安装依赖  
    ```pip install -r requirements.txt```
2. 所有工具的工作目录默认为本工程的根目录，即00_task_tool，注意将代码中文件路径修改为相对此目录
3. 使用小工具前先查看其文件夹内部的README文件（如有）
4. 请根据自己的实际文件名修改代码中的相应位置
5. 运行小工具时无需cd到小工具的文件夹内部，直接在工程根目录运行
    例如：```python .\01_company_info_by_selenium\company_info_collector.py```
