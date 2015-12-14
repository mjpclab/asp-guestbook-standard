<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/sql/admin_delreply.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->
<%
Response.Expires=-1
if isnumeric(Request.QueryString("id"))=false or Request.QueryString("id")="" then
	Response.Redirect "admin.asp"
	Response.End 
end if

set cn=server.CreateObject("ADODB.Connection")
Call CreateConn(cn)

cn.BeginTrans
	cn.Execute sql_admindelreply_delete & Request.QueryString("id"),,1
	cn.Execute sql_admindelreply_update & Request.QueryString("id"),,1
cn.CommitTrans

cn.Close : set cn=nothing
%>
<!-- #include file="include/template/admin_traceback.inc" -->