<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/md5.asp" -->
<!-- #include file="include/utility/string.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="tips.asp" -->
<%
Response.Expires=-1

if StatusStatistics then call addstat("login")

Dim referrer
referrer=Request.Form("referrer")

if VcodeCount>0 and (Request.Form("ivcode")<>Session(InstanceName & "_vcode") or Session(InstanceName & "_vcode")="") then
	Session(InstanceName & "_vcode")=""
	if StatusStatistics then call addstat("loginfailed")
	Call TipsPage("��֤�����","admin_login.asp?referrer=" & Server.UrlEncode(referrer))

	set rs=nothing
	set cn=nothing
	Response.End
else
	Session(InstanceName & "_vcode")=""
end if

Dim cn,rs,refPass
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
Call CreateConn(cn)
rs.Open sql_adminverify,cn,0,1,1
if Not rs.EOF then
	refPass=rs.Fields(0)
end if
rs.Close : cn.Close : set rs=nothing : set cn=nothing

Session(InstanceName & "_adminpass")=md5(request("iadminpass"),32)
if Session(InstanceName & "_adminpass")=refPass then
	session.Timeout=AdminTimeOut
	if referrer<>"" then
		Response.Redirect referrer
	else
		Response.Redirect "admin.asp"
	end if
else
	if StatusStatistics then call addstat("loginfailed")
	Call TipsPage("���벻��ȷ��","admin_login.asp?referrer=" & Server.UrlEncode(referrer))
end if
%>