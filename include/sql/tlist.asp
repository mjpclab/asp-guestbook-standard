<%
if IsAccess then
	sql_tlist=		"SELECT TOP {0} id,title FROM " &table_main& " WHERE parent_id<=0 AND (guestflag mod 64)\8=0 ORDER BY parent_id ASC,lastupdated DESC"
elseif IsSqlServer then
	sql_tlist=		"SELECT TOP {0} id,title FROM " &table_main& " WHERE parent_id<=0 AND guestflag & 56=0 ORDER BY parent_id ASC,lastupdated DESC"
end if
%>