<!-- #include file="config.asp" -->
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