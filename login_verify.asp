<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/md5.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="tips.asp" -->
<%
Response.Expires=-1

if StatusStatistics then call addstat("login")

Dim referrer
referrer=Request.Form("referrer")

if VcodeCount>0 and (Request.Form("ivcode")<>Session("vcode") or Session("vcode")="") then
	Session("vcode")=""
	if StatusStatistics then call addstat("loginfailed")
	Call TipsPage("验证码错误。","admin_login.asp?referrer=" & Server.UrlEncode(referrer))

	set rs=nothing
	set cn=nothing
	Response.End
else
	Session("vcode")=""
end if

Dim cn,rs
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
Call CreateConn(cn)
rs.Open sql_adminverify,cn,0,1,1

Session(InstanceName & "_adminpass")=md5(request("iadminpass"),32)
if Not rs.EOF then
	if Session(InstanceName & "_adminpass")=rs(0) then
		session.Timeout=clng(AdminTimeOut)
		if referrer<>"" then
			Response.Redirect referrer
		else
			Response.Redirect "admin.asp"
		end if
	else
		if StatusStatistics then call addstat("loginfailed")
		Call TipsPage("密码不正确。","admin_login.asp?referrer=" & Server.UrlEncode(referrer))

		rs.Close : cn.Close : set rs=nothing : set cn=nothing
		Response.End
	end if
else
	if StatusStatistics then call addstat("loginfailed")
	Call TipsPage("密码验证失败。","admin_login.asp?referrer=" & Server.UrlEncode(referrer))

	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	Response.End
end if
rs.Close : cn.Close : set rs=nothing : set cn=nothing
%>