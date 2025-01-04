# Excel to Database Importer

这是一个用于将Excel文件中的数据导入到MySQL或PostgreSQL数据库的Python脚本。它支持从Excel文件中读取多个sheet，并基于sheet名称创建或更新相应的数据库表。

## 功能

- 从Excel文件读取数据。
- 支持MySQL和PostgreSQL数据库。
- 自动检测并创建或更新数据库表。
- 提供处理进度显示。

## 使用方法

1. **安装依赖**：
   ```bash
   pip install pandas psycopg2-binary mysql-connector-python tqdm

2.配置数据库：   
[database]
type = postgresql  # 或者 mysql

[postgresql]
host = your_postgresql_host
port = 5432
dbname = your_db_name
user = your_username
password = your_password

[mysql]
host = your_mysql_host
port = 3306
dbname = your_db_name
user = your_username
password = your_password
3.运行脚本：
python script_name.py path_to_your_excel_file.xlsx [postgresql|mysql]
代码结构
    DatabaseHandler 类：处理数据库连接、表检测、创建和数据插入。
    ExcelProcessor 类：处理Excel文件，读取sheet，并协调数据导入到数据库。


贡献
欢迎改进和贡献！如果你有任何建议或发现任何问题，请通过Issues提交。

许可证
MIT (LICENSE)  # 你可以根据需要选择其他合适的许可证
