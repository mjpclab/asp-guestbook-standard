<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/sql/admin_savebulletin.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/book.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->
<%
Response.Expires=-1
if Not IsEmpty(Request.Form) then
	dim tlimit
	tlimit=0
	if Request.Form("html2")="1" then tlimit=tlimit OR 1
	if Request.Form("ubb2")="1" then tlimit=tlimit OR 2
	if Request.Form("newline2")="1" then tlimit=tlimit OR 4

	dim tbul
	tbul=Request.Form("abulletin")
	tbul=replace(tbul,"'","''")
	tbul=replace(tbul,"<%","< %")

	set cn=server.CreateObject("ADODB.Connection")
	Call CreateConn(cn)
	cn.Execute Replace(Replace(sql_adminsavebulletin,"{0}",tlimit),"{1}",tbul),,129

	cn.close : set cn=nothing
	Response.Redirect "admin.asp"
else
	Response.Redirect "admin_setbulletin.asp"
end if
%>