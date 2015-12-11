<%
IsAccess=false
IsSqlServer=false

if dbtype>=1 and dbtype<=3 then
	IsAccess=true
	table_config			="[" & dbprefix & "config]"
	table_filterconfig		="[" & dbprefix & "filterconfig]"
	table_floodconfig		="[" & dbprefix & "floodconfig]"
	table_ipv4config		="[" & dbprefix & "ipv4config]"
	table_ipv6config		="[" & dbprefix & "ipv6config]"
	table_main				="[" & dbprefix & "main]"
	table_reply				="[" & dbprefix & "reply]"
	table_stats				="[" & dbprefix & "stats]"
	table_stats_clientinfo	="[" & dbprefix & "stats_clientinfo]"
	table_style				="[" & dbprefix & "style]"
	table_supervisor		="[" & dbprefix & "supervisor]"
elseif dbtype=10 then
	IsSqlServer=true
	if dbschema<>"" then
		schema="[" &dbschema& "]."
	else
		schema=""
	end if
	table_config			=schema & "[" & dbprefix & "config]"
	table_filterconfig		=schema & "[" & dbprefix & "filterconfig]"
	table_floodconfig		=schema & "[" & dbprefix & "floodconfig]"
	table_ipv4config		=schema & "[" & dbprefix & "ipv4config]"
	table_ipv6config		=schema & "[" & dbprefix & "ipv6config]"
	table_main				=schema & "[" & dbprefix & "main]"
	table_reply				=schema & "[" & dbprefix & "reply]"
	table_stats				=schema & "[" & dbprefix & "stats]"
	table_stats_clientinfo	=schema & "[" & dbprefix & "stats_clientinfo]"
	table_style				=schema & "[" & dbprefix & "style]"
	table_supervisor		=schema & "[" & dbprefix & "supervisor]"
end if
%>