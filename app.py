import pyodbc
import pandas as pd
from pathlib import Path
import sys

BASE_DIR = Path(__file__).resolve().parent

def main(MDB:Path, EXCEL:Path, sheet_name:str):
    print(f"MDB: {MDB}, EXCEL: {EXCEL}")
    # specify your MDB file path
    # MDB = 'path\\to\\your\\file.mdb'
    DRV = '{Microsoft Access Driver (*.mdb, *.accdb)}'
    PWD = 'pw'  # if your MDB file is password protected, replace 'pw' with your password

    # establish a connection to the MDB file
    con = pyodbc.connect('DRIVER={};DBQ={};PWD={}'.format(DRV, MDB, PWD))

    # SQL query to fetch all data from the table
    SQL = f'SELECT * FROM {sheet_name}'  # replace 'your_table' with your table's name

    # fetch all rows as a pandas DataFrame
    df = pd.read_sql_query(SQL, con)

    # specify your Excel file path
    # EXCEL = 'path\\to\\your\\output.xlsx'

    # write the DataFrame to an Excel file
    df.to_excel(EXCEL, index=False)


if __name__ == "__main__":
    try:
        main(Path(sys.argv[1]).resolve(), Path(sys.argv[2]).resolve(),sys.argv[3])        
    except IndexError:
        print('Error: missing arguments\n\nExample:\npython app.py <MDB Path> <XLSX Path> <Target Sheet Name>')