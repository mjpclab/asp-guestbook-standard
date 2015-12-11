<%
sql_adminfilter_null=	"SELECT * FROM " &table_filterconfig& " WHERE filterid IS NULL"
sql_adminfilter_max=	"SELECT MAX(filterid) FROM " &table_filterconfig
sql_adminfilter_update=	"UPDATE " &table_filterconfig& " SET filtersort={0} WHERE filterid={0}"
%>