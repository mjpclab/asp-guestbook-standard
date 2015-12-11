<%
if IsAccess then
	sql_condition_hidehidden=	" AND (((guestflag mod 16)\8)=0 OR replied<>0)"
	sql_condition_hideaudit=	" AND (((guestflag mod 32)\16)=0 OR replied<>0)"
	sql_condition_hidewhisper=	" AND (((guestflag mod 64)\32)=0 OR replied<>0)"
elseif IsSqlServer then
	sql_condition_hidehidden=	" AND (guestflag & 8=0 OR replied<>0)"
	sql_condition_hideaudit=	" AND (guestflag & 16=0 OR replied<>0)"
	sql_condition_hidewhisper=	" AND (guestflag & 32=0 OR replied<>0)"
end if
%>