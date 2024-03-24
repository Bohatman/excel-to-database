from openpyxl import load_workbook
import pandas as pd
from openpyxl.cell.rich_text import CellRichText,TextBlock
import sqlite3
from my_parser import read_and_remove_strikethrough_content,pandas_type_to_sqlite_type

df = read_and_remove_strikethrough_content('strikethrough.xlsx')
columns_str = ', '.join([f"{col} {pandas_type_to_sqlite_type(str(df[col].dtype))}" for col in df.columns])

print(columns_str)