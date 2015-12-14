<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/sql/admin_mpassaudit.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->
<%
Response.Expires=-1
if Request.Form("seltodel")="" then
	if isnumeric(Request.Form("page")) and Request.Form("page")<>"" then
		Response.Redirect "admin.asp?page=" & Request.Form("page")
	else
		Response.Redirect "admin.asp"
	end if
end if

dim ids
ids=FilterSql(Request.Form("seltodel"))
if Left(ids,1) = "," then
	ids=Mid(ids,2)
end if

set cn=server.CreateObject("ADODB.Connection")
Call CreateConn(cn)

cn.Execute Replace(sql_adminmpass,"{0}",ids),,1
cn.Close : set cn=nothing
%>
<!-- #include file="include/template/admin_traceback.inc" -->