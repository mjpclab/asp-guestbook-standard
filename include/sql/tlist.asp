<%
'sql_tlist_itemsperpage=	"SELECT TOP 1 itemsperpage FROM " &table_config
sql_tlist_maxtime=		"SELECT TOP 1 lastupdated FROM " &table_main& " WHERE parent_id<=0 {0} ORDER BY lastupdated DESC"
if IsAccess then
	sql_tlist_mintime=	"SELECT TOP 1 lastupdated FROM (SELECT TOP {0} parent_id,lastupdated FROM " &table_main& " WHERE parent_id<=0 AND (guestflag mod 32)\16=0 AND (guestflag mod 64)\32=0 {1} ORDER BY parent_id ASC,lastupdated DESC) ORDER BY lastupdated ASC"
	sql_tlist=			"SELECT guestflag,title FROM " &table_main& " WHERE parent_id<=0 AND lastupdated>=#{0}# AND lastupdated<=#{1}# {2} ORDER BY parent_id ASC,lastupdated DESC"
elseif IsSqlServer then
	sql_tlist_mintime=	"SELECT TOP 1 lastupdated FROM (SELECT TOP {0} parent_id,lastupdated FROM " &table_main& " WHERE parent_id<=0 AND guestflag & 48=0 {1} ORDER BY parent_id ASC,lastupdated DESC) AS __temp1 ORDER BY lastupdated ASC"
	sql_tlist=			"SELECT guestflag,title FROM " &table_main& " WHERE parent_id<=0 AND lastupdated>='{0}' AND lastupdated<='{1}' {2} ORDER BY parent_id ASC,lastupdated DESC"
end if
%>