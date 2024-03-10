
import pyodbc # connection to ODS
import csv
import pandas as pd

conn = pyodbc.connect(
    Driver='SQL Server'
    ,Server='10.100.17.147,34503'
    ,Database='cst2200group7'
    ,UID = 'group7sa'
    ,PWD = 'nga2IeH5k9uh3cAdmJ')
cur = conn.cursor()


def doit(sql):
    cur.execute(sql)
    conn.commit()
    # print(sql)
print('start')
# Path to your Excel file
excel_file_path = 'D:\DatabaseManagement\Project_Etl\FinalSheets\StreetTable.xlsx'
print('done')
# Load Excel data into a pandas DataFrame
excel_data = pd.read_excel(excel_file_path)

# Display the loaded data
print(excel_data.head())  # Adjust the number of rows to display as needed

table_name = 'ab_street_data'

for index, row in excel_data.iterrows():
    insert_query = f"INSERT INTO {table_name} (st_id, str_name_pcs, str_dir, city_id, st_type_id) VALUES (?,?,?,?,?)"  # Adjust columns and query structure
    values = (row['Street_id'], row['Street_Name'], row['str_dir'], row['city_id'], row['str_type_id'])  # Replace with your column names
    
    cur.execute(insert_query, values)
    conn.commit()

# Close the cursor and the connection
cur.close()
conn.close()
