<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/sql/admin_messageaction.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->
<%
Response.Expires=-1
if isnumeric(Request.QueryString("id"))=false or Request.QueryString("id")="" then
	Response.Write "admin.asp"
	Response.End
end if

set cn=server.CreateObject("ADODB.Connection")
Call CreateConn(cn)

cn.Execute sql_adminpubwhisper & Request.QueryString("id"),,129
cn.close : set cn=nothing
%>
<!-- #include file="include/template/admin_traceback.inc" -->