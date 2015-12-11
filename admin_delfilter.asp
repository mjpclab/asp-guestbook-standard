<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/sql/admin_delfilter.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->
<%
Response.Expires=-1

if isnumeric(Request.Form("filterid")) and Request.Form("filterid")<>"" then
	 tfilterid=clng(Request.Form("filterid"))

	set cn=server.CreateObject("ADODB.Connection")
	CreateConn cn,dbtype

	cn.Execute sql_admindelfilter & tfilterid,,1
	
	cn.Close
	set cn=nothing
end if

Response.Redirect "admin_filter.asp"
%>