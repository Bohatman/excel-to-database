import sqlite3
from my_parser import pandas_type_to_sqlite_type, read_and_remove_strikethrough_content

df = read_and_remove_strikethrough_content('strikethrough.xlsx')
conn = sqlite3.connect('database.db')
cur = conn.cursor()

columns_str = ', '.join([f"{col} {pandas_type_to_sqlite_type(str(df[col].dtype))}" for col in df.columns])
create_table_sql = f"CREATE TABLE IF NOT EXISTS my_table ({columns_str})"
cur.execute(create_table_sql)
conn.commit()
df = df.iloc[1:, :]
df.to_sql('my_table', conn, if_exists='append', index=False)
conn.close()