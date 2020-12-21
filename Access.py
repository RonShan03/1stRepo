import pyodbc
import pandas as pd


conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\ronsh\Access\Database21.accdb;')
cursor = conn.cursor()

data = pd.read_csv(r'C:Users\ronsh\Excel\QCOM2020.csv')
df = pd.DataFrame(data, columns=['Date', 'Open', 'High', 'Low', 'Close', 'Adj Close', 'Volume'])
print(df)


#conn1 = pyodbc.connect(r'Driver={Microsoft Access Text Driver (*.txt, *.csv)};DBQ=C:\Users\ronsh\Excel\QCOM2020.csv;')
#cursor1 = conn1.cursor()

#print(pyodbc.drivers())

#[x for x in pyodbc.drivers() if x.startswith('Microsoft Excel Driver')]
#Do the same thing as you did for conn but with conn1 and cursor1 for xlcel spreadsheet.
#Make a loop that goes through every row/column of the .csv file and copies/cuts that to the access table


#cursor1.execute('select * from QCOM')

#for row in cursor1.fetchall():
#    print(row)



sql = '''
        INSERT INTO Stock_quotes (ID,Symbol,Day,ShortVol,ShortExVol,TotalVol)
        VALUES(7,'TEST','02/08/2014',1234,'789','12098765')
       '''

print(sql)
cursor.execute(sql)
conn.commit()
