<%
if IsAccess then
	sql_adminmpass="UPDATE " &table_main& " SET guestflag=guestflag-16 WHERE ((guestflag mod 32)\16)<>0 AND id IN ({0})"
elseif IsSqlServer then
	sql_adminmpass="UPDATE " &table_main& " SET guestflag=guestflag-16 WHERE guestflag & 16<>0 AND id IN ({0})"
end if
%>