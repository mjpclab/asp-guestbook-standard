<%
if IsSqlServer then
	sql_compact_dbfile="DECLARE @curr_dbname NVARCHAR(255); SET @curr_dbname=DB_NAME(0); DBCC SHRINKDATABASE(@curr_dbname,8)"
	sql_compact_dblog="DECLARE @curr_dbname NVARCHAR(255); SET @curr_dbname=DB_NAME(0); BACKUP LOG @curr_dbname WITH NO_LOG"
end if
%>