# Crystal Reports OLEDB Connection Updater
Microsoft OLE DB Provider for SQL Server (SQLOLEDB) is deprecated and some servers are dropping TLS and TLS 1.1 support, making some reports using SQLOLEDB unusable.

This simple CLI application updates all report files to the new Crystal Reports to the new format and updates the connection string of each report and sub report to use the new Microsoft OLE DB Driver for SQL Server (MSOLEDBSQL).
