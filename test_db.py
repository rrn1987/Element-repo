import sqlite3
import pandas as pd
import os


def select_all_tasks(conn):
    cur = conn.cursor()
    cur.execute("SELECT * FROM datatable")
    rows = cur.fetchall()
    for row in rows:
        print(row)


if __name__ == '__main__':
    conn = sqlite3.connect('mydb.db')
    wb = pd.read_excel(os.path.join(os.getcwd(), 'TMO_NPT_ATE_PQA_Q3_2023_Release.xlsx'), sheet_name='Data')
    wb.to_sql(name='datatable', con=conn, if_exists='replace', index=True)
    conn.commit()
    select_all_tasks(conn)
    conn.close()
