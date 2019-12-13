echo off
cd "C:\Program Files\EM - Ventas y Compras\BACKUPDATABASE"

set DBFILENAME=%CD%\EM_FULL_%date:~-4,4%%date:~-10,2%%date:~-7,2%_%time:~0,2%%time:~3,2%.bak
set SERVERNAME=WIN-9K023S4JOWB\SQLEXPRESS

echo %DBFILENAME%
echo %SERVERNAME%


SqlCmd -E -S %SERVERNAME% -Q "BACKUP DATABASE [EM] TO DISK='%DBFILENAME%' WITH NOFORMAT, INIT,  NAME = N'EM-Full Database Backup', SKIP, NOREWIND, NOUNLOAD,  STATS = 10"

