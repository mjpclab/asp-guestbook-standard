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

cn.BeginTrans
	cn.Execute sql_admindelreply_delete & Request.QueryString("id"),,1
	cn.Execute sql_admindelreply_update & Request.QueryString("id"),,1
cn.CommitTrans

cn.Close : set cn=nothing
%>
<!-- #include file="include/admin_traceback.inc" -->