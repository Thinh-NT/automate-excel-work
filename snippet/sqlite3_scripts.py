import sqlite3
import pandas as pd

dbname = 'students_performance'
conn = sqlite3.connect(dbname + '.sqlite')
chunksize = 10000
for chunk in pd.read_csv('students_performance.csv', chunksize=chunksize):
    # replacing spaces with underscores for column names
    chunk.columns = chunk.columns.str.replace(' ', '_')
    try:
        chunk.to_sql(name='Performance', con=conn)
    except ValueError:
        chunk.to_sql(name='Performance', con=conn, if_exists='append')

cur = conn.cursor()
cur.execute('SELECT * FROM Performance')

for row in cur:
    print(row)

cur.close()
