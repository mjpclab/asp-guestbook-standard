<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/sql/admin_saveedit.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->
<%
Response.Expires=-1
if isnumeric(request.form("mainid"))=false or request.form("mainid")="" then
	Response.Redirect "admin.asp"
	Response.End
end if

if Not IsEmpty(Request.Form) then
	dim tlimit
	tlimit=0
	if Request.Form("html1")="1" then tlimit=tlimit+1
	if Request.Form("ubb1")="1" then tlimit=tlimit+2
	if Request.Form("newline1")="1" then tlimit=tlimit+4

	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")
	Call CreateConn(cn)
	rs.Open sql_adminsaveedit_open & Request.Form("mainid"),cn,0,3,1
	if Not rs.EOF then		'ÁôÑÔ´æÔÚ
		rs.Fields("guestflag")=cbyte(rs.Fields("guestflag") and 248) + tlimit
		if Request.Form("etitle")<>"" then rs.Fields("title")=HtmlEncode(Request.Form("etitle"))
		rs.Fields("article")=Request.Form("econtent")
		rs.Update

		rs.close : cn.close : set rs=nothing : set cn=nothing
	else
		rs.close : cn.close : set rs=nothing : set cn=nothing
	end if
end if
%><!-- #include file="include/template/admin_traceback.inc" -->