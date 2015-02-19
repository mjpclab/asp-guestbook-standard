<!-- #include file="config.asp" -->
<!-- #include file="admin_verify.asp" -->

<%
Response.Expires=-1

set cn=server.CreateObject("ADODB.Connection")
CreateConn cn,dbtype

tnow=now()
cn.Execute Replace(sql_adminclearstats_startdate,"{0}",now()),,1
cn.Execute sql_adminclearstats_client,,1

cn.close
set cn=nothing

response.redirect "admin_stats.asp"
%>