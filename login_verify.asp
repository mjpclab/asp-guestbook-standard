<!-- #include file="config.asp" -->
<link rel="stylesheet" type="text/css" href="css/style.css"/>
<!-- #include file="css/style.asp" -->
<!-- #include file="md5.asp" -->

<%
Response.Expires=-1
if StatusStatistics then call addstat("login")

if VcodeCount>0 and (Request.Form("ivcode")<>session("vcode") or session("vcode")="") then
	session("vcode")=""
	if StatusStatistics then call addstat("loginfailed")
	Call MessagePage("��֤�����","admin_login.asp")

	set rs=nothing
	set cn=nothing
	Response.End
else
	session("vcode")=""
end if

Dim cn,rs
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")

CreateConn cn,dbtype
rs.Open sql_loginverify,cn,0,1,1

session.Contents(InstanceName & "_adminpass")=md5(request("iadminpass"),32)
if rs.EOF=false then
	if session.Contents(InstanceName & "_adminpass")=rs(0) then
		session.Timeout=clng(AdminTimeOut)
		Response.Redirect "admin.asp"
	else
		if StatusStatistics then call addstat("loginfailed")
		Call MessagePage("���벻��ȷ��","admin_login.asp")

		rs.Close : cn.Close : set rs=nothing : set cn=nothing
		Response.End
	end if
else
	if StatusStatistics then call addstat("loginfailed")
	Call MessagePage("������֤ʧ�ܡ�","admin_login.asp")

	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	Response.End
end if
rs.Close : cn.Close : set rs=nothing : set cn=nothing
%>