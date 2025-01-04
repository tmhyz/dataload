import pandas as pd
from psycopg2 import connect as pg_connect
import mysql.connector
import argparse
from tqdm import tqdm
import configparser

class DatabaseHandler:
    def __init__(self, db_type, db_config):
        self.db_type = db_type.lower()
        self.db_config = db_config
        self.connection = None

    def connect(self):
        if self.db_type == 'postgresql':
            conn_string = f"host={self.db_config['postgresql']['host']} port={self.db_config['postgresql']['port']} dbname={self.db_config['postgresql']['dbname']} user={self.db_config['postgresql']['user']} password={self.db_config['postgresql']['password']}"
            self.connection = pg_connect(conn_string)
        elif self.db_type == 'mysql':
            self.connection = mysql.connector.connect(
                host=self.db_config['mysql']['host'],
                port=int(self.db_config['mysql']['port']),
                database=self.db_config['mysql']['dbname'],
                user=self.db_config['mysql']['user'],
                password=self.db_config['mysql']['password']
            )
        else:
            raise ValueError("Unsupported database type")

    def close(self):
        if self.connection:
            self.connection.close()

    def check_table_exists(self, table_name):
        with self.connection.cursor() as cursor:
            if self.db_type == 'postgresql':
                cursor.execute("SELECT to_regclass(%s)", (table_name,))
                return cursor.fetchone()[0] is not None
            elif self.db_type == 'mysql':
                cursor.execute("SHOW TABLES LIKE %s", (table_name,))
                return cursor.fetchone() is not None

    def create_table(self, table_name, columns):
        with self.connection.cursor() as cursor:
            if self.db_type == 'postgresql':
                columns_defs = ', '.join([f"{col} VARCHAR" for col in columns])
            elif self.db_type == 'mysql':
                columns_defs = ', '.join([f"{col} VARCHAR(255)" for col in columns])
            create_table_sql = f"CREATE TABLE {table_name} ({columns_defs})"
            cursor.execute(create_table_sql)

    def insert_data(self, table_name, columns, data):
        with self.connection.cursor() as cursor:
            placeholders = ', '.join(['%s' for _ in columns])
            insert_sql = f"INSERT INTO {table_name} ({', '.join(columns)}) VALUES ({placeholders})"
            cursor.executemany(insert_sql, data)
            self.connection.commit()

class ExcelProcessor:
    def __init__(self, excel_file):
        self.excel_file = excel_file
        self.xls = pd.ExcelFile(self.excel_file)
        self.sheets = self.xls.sheet_names

    def process_sheets(self, db_handler):
        # 总体进度
        with tqdm(total=len(self.sheets), desc="Processing Sheets", unit="sheet") as total_progress:
            for sheet in self.sheets:
                df = pd.read_excel(self.xls, sheet_name=sheet)
                columns = df.columns.tolist()
                data = [tuple(row) for _, row in df.iterrows()]

                # 单表进度
                with tqdm(total=len(df), desc=f"Inserting into {sheet}", unit="row") as single_progress:
                    if db_handler.check_table_exists(sheet):
                        db_handler.insert_data(sheet, columns, data)
                    else:
                        db_handler.create_table(sheet, columns)
                        db_handler.insert_data(sheet, columns, data)
                    single_progress.update(len(df))
                total_progress.update(1)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Process Excel file and import into database.')
    parser.add_argument('excel_file', help='Path to the Excel file')
    parser.add_argument('db_type', help='Type of database (postgresql or mysql)')
    
    args = parser.parse_args()

    # 读取配置文件
    config = configparser.ConfigParser()
    config.read('dbconf.ini')
    
    db_handler = DatabaseHandler(args.db_type, config)
    db_handler.connect()

    try:
        excel_processor = ExcelProcessor(args.excel_file)
        excel_processor.process_sheets(db_handler)
    except Exception as error:
        print(f"An error occurred: {error}")
    finally:
        db_handler.close()