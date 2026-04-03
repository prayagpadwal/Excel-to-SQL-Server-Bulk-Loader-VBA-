# Excel-to-SQL-Server-Bulk-Loader-VBA-
What it does: Reads data from the first populated sheet in your workbook and bulk-inserts every row into a target SQL Server table. It auto-detects your column headers, maps them directly to table columns, and shows progress in the Excel status bar for large loads. When done, it tells you how many rows were inserted and how long it took.
How to use it:

Open your Excel file with the data you want to load.
Make sure your first row contains column headers that exactly match the destination table's column names.
Update the three placeholders marked with TODO in the code: your SQL Server name, database name, and target table name.
Open the VBA editor (Alt + F11), paste the code into a module, and run LoadDataToSQL.

Requirements: The machine running this must have the MSOLEDBSQL driver installed and Windows Integrated Authentication access to the SQL Server.
