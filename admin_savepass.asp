<!-- #include file="config.asp" -->
<link rel="stylesheet" type="text/css" href="style.css"/>
<!-- #include file="style.asp" -->
<!-- #include file="md5.asp" -->
<!-- #include file="common.asp" -->

<%
Response.Expires=-1
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype
rs.Open sql_adminsavepass_select,cn,0,1,1

if request.Form("inewpass1")<> request.Form("inewpass2") then
	Call MessagePage("新密码不一致，请检查。","admin_chpass.asp")
	Response.End
elseif request.Form("inewpass1")="" then
	Call MessagePage("密码不能为空，请重新输入。","admin_chpass.asp")
	Response.End
else
	if rs.EOF=false then
		if md5(request.Form("ioldpass"),32)=rs(0) then
			pwd=md5(request.Form("inewpass1"),32)
			
			cn.Execute Replace(sql_adminsavepass_update,"{0}",pwd),,1
			session.Contents(InstanceName & "_adminpass")=pwd

			Response.Redirect "admin.asp"
		else
			Call MessagePage("原密码错误，请检查。","admin_chpass.asp")

			rs.Close
			cn.Close
			set rs=nothing
			set cn=nothing
			Response.End
		end if
	else
		Call MessagePage("密码验证失败，请检查。","admin_chpass.asp")

		rs.Close
		cn.Close
		set rs=nothing
		set cn=nothing
		Response.End
	end if
end if
rs.Close
cn.Close
set rs=nothing
set cn=nothing
%>