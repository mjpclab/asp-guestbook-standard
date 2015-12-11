<%
sql_admindelreply_delete="DELETE FROM " &table_reply& " WHERE articleid="
if IsAccess then
	sql_admindelreply_update="UPDATE " &table_main& " SET replied=replied-(replied mod 2) WHERE id="
elseif IsSqlServer then
	sql_admindelreply_update="UPDATE " &table_main& " SET replied=replied-(replied & 1) WHERE id="
end if
%>