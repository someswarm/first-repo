import pyodbc
import pandas as pd
# Connection details
server = r'BOTD-360SQL21\QA'
database = 'AMAINNPRD1'
username = 'sa'
password = 'k@?aDTim+$L#^yx'

# Connection string
conn = pyodbc.connect(
    'DRIVER={ODBC Driver 18 for SQL Server};'
    f'SERVER={server};'
    f'DATABASE={database};'
    f'UID={username};'
    f'PWD={password};'
    'Trusted_Connection=yes;'
    'TrustServerCertificate=yes;'
)
# conn = pyodbc.connect(
#     'DRIVER={ODBC Driver 17 for SQL Server};'
#     f'SERVER={server};'
#     'DATABASE=YourDatabase;'
#     'Trusted_Connection=yes;'
# )
cursor = conn.cursor()

# Execute query
#cursor.execute("SELECT * FROM AFW_Customer where CustNo='1234'")
query="SELECT * FROM AFW_Customer where CustNo='2105'"
# Fetch all rows
# rows = cursor.fetchall()

# for row in rows:
#     print(row)
df = pd.read_sql(query, conn)

# Export to Excel
df.to_excel("output.xlsx", index=False)

conn.close()

print("Excel file created successfully!")

# Close connection
# cursor.close()
# conn.close()