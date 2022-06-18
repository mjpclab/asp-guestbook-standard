<%
sql_adminsaveipconfig_update=	"UPDATE " &table_config& " SET ipconstatus="
sql_adminsaveipv4config_delete=	"DELETE FROM " &table_ipv4config& " WHERE listid IN ({0})"
sql_adminsaveipv4config_insert1=	"INSERT INTO " &table_ipv4config& "(cfgtype,ipfrom,ipto) VALUES (1,'{0}','{1}')"
sql_adminsaveipv4config_insert2=	"INSERT INTO " &table_ipv4config& "(cfgtype,ipfrom,ipto) VALUES (2,'{0}','{1}')"
sql_adminsaveipv6config_delete=	"DELETE FROM " &table_ipv6config& " WHERE listid IN ({0})"
sql_adminsaveipv6config_insert1=	"INSERT INTO " &table_ipv6config& "(cfgtype,ipfrom,ipto) VALUES (1,'{0}','{1}')"
sql_adminsaveipv6config_insert2=	"INSERT INTO " &table_ipv6config& "(cfgtype,ipfrom,ipto) VALUES (2,'{0}','{1}')"
%>