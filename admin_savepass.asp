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
	Call MessagePage("�����벻һ�£����顣","admin_chpass.asp")
	Response.End
elseif request.Form("inewpass1")="" then
	Call MessagePage("���벻��Ϊ�գ����������롣","admin_chpass.asp")
	Response.End
else
	if rs.EOF=false then
		if md5(request.Form("ioldpass"),32)=rs(0) then
			pwd=md5(request.Form("inewpass1"),32)
			
			cn.Execute Replace(sql_adminsavepass_update,"{0}",pwd),,1
			session.Contents(InstanceName & "_adminpass")=pwd

			Response.Redirect "admin.asp"
		else
			Call MessagePage("ԭ����������顣","admin_chpass.asp")

			rs.Close
			cn.Close
			set rs=nothing
			set cn=nothing
			Response.End
		end if
	else
		Call MessagePage("������֤ʧ�ܣ����顣","admin_chpass.asp")

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