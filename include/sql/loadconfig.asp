<%
sql_loadconfig_config=		"SELECT TOP 1 * FROM " & table_config
sql_loadconfig_style=		"SELECT * FROM " &table_style& " WHERE styleid={0}"
sql_loadconfig_top1style=	"SELECT Top 1 * FROM " &table_style
sql_loadconfig_floodconfig=	"SELECT TOP 1 * FROM " &table_floodconfig
%>