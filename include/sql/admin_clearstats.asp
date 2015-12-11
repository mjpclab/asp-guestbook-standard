<%
sql_adminclearstats_startdate=	"UPDATE " &table_stats& " SET startdate='{0}',[view]=0,[search]=0,[leaveword]=0,[written]=0,[filtered]=0,[banned]=0,[login]=0,[loginfailed]=0"
sql_adminclearstats_client=		"DELETE FROM " &table_stats_clientinfo
%>