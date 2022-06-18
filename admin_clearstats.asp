<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/sql/admin_clearstats.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->
<%
Response.Expires=-1

set cn=server.CreateObject("ADODB.Connection")
Call CreateConn(cn)

tnow=now()
cn.Execute Replace(sql_adminclearstats_startdate,"{0}",now()),,129
cn.Execute sql_adminclearstats_client,,129

cn.close
set cn=nothing

response.redirect "admin_stats.asp"
%>