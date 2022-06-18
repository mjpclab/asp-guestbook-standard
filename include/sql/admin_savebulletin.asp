<%
if IsAccess then
	sql_adminsavebulletin="UPDATE " &table_supervisor& " SET declareflag={0},[declare]='{1}'"
elseif IsSqlServer then
	sql_adminsavebulletin="UPDATE " &table_supervisor& " SET declareflag={0},[declare]=N'{1}'"
end if
%>