<%
sql_adminipv4config_status1=	"SELECT listid,ipfrom,ipto FROM " &table_ipv4config& " WHERE cfgtype=1"
sql_adminipv4config_status2=	"SELECT listid,ipfrom,ipto FROM " &table_ipv4config& " WHERE cfgtype=2"
sql_adminipv6config_status1=	"SELECT listid,ipfrom,ipto FROM " &table_ipv6config& " WHERE cfgtype=1"
sql_adminipv6config_status2=	"SELECT listid,ipfrom,ipto FROM " &table_ipv6config& " WHERE cfgtype=2"
%>