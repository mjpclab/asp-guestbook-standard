<!-- #include file="config.asp" -->
<!-- #include file="admin_verify.asp" -->

<%
Response.Expires=-1
if isnumeric(Request.QueryString("id"))=false or Request.QueryString("id")="" then
	Response.Redirect "admin.asp"
	Response.End
end if

set cn=server.CreateObject("ADODB.Connection")
CreateConn cn,dbtype

cn.Execute Replace(Replace(sql_adminbring2top,"{0}",now()),"{1}",Request.QueryString("id")),,1
cn.close : set cn=nothing
%>
<!-- #include file="include/admin_traceback.inc" -->