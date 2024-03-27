
Warning: uc_driver update needed. Getting it now:


cursor.execute("INSERT INTO C0001 VALUES(?,3,?,null)",

pyodbc.IntegrityError: ('23000', "[23000] [Microsoft][ODBC SQL Server Driver][SQL Server]Violation of PRIMARY KEY constraint 'PK_c0001'. Cannot insert duplicate key in object 'dbo.c0001'. The duplicate key value is (2024-04-03, 3). (2627) (SQLExecDirectW); [23000] [Microsoft][ODBC SQL Server Driver][SQL Server]The statement has been terminated. (3621)")
