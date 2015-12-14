<!-- #include file="include/template/page_instruction.inc" -->
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

if VcodeCount>0 and (Request.Form("ivcode")<>session("vcode") or session("vcode")="") then
	session("vcode")=""
	if StatusStatistics then call addstat("loginfailed")
	Call TipsPage("验证码错误。","admin_login.asp")

	set rs=nothing
	set cn=nothing
	Response.End
else
	session("vcode")=""
end if

Dim cn,rs
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")

Call CreateConn(cn)
rs.Open sql_adminverify,cn,0,1,1

session.Contents(InstanceName & "_adminpass")=md5(request("iadminpass"),32)
if rs.EOF=false then
	if session.Contents(InstanceName & "_adminpass")=rs(0) then
		session.Timeout=clng(AdminTimeOut)
		Response.Redirect "admin.asp"
	else
		if StatusStatistics then call addstat("loginfailed")
		Call TipsPage("密码不正确。","admin_login.asp")

		rs.Close : cn.Close : set rs=nothing : set cn=nothing
		Response.End
	end if
else
	if StatusStatistics then call addstat("loginfailed")
	Call TipsPage("密码验证失败。","admin_login.asp")

	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	Response.End
end if
rs.Close : cn.Close : set rs=nothing : set cn=nothing
%>