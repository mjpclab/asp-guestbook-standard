<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/sql/admin_savepass.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/md5.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/book.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->
<!-- #include file="tips.asp" -->
<%
Response.Expires=-1
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
Call CreateConn(cn)
rs.Open sql_adminsavepass_select,cn,0,1,1

if request.Form("inewpass1")<> request.Form("inewpass2") then
	Call TipsPage("�����벻һ�£����顣","admin_chpass.asp")
	Response.End
elseif request.Form("inewpass1")="" then
	Call TipsPage("���벻��Ϊ�գ����������롣","admin_chpass.asp")
	Response.End
else
	if rs.EOF=false then
		if md5(request.Form("ioldpass"),32)=rs(0) then
			pwd=md5(request.Form("inewpass1"),32)
			
			cn.Execute Replace(sql_adminsavepass_update,"{0}",pwd),,1
			session.Contents(InstanceName & "_adminpass")=pwd

			Response.Redirect "admin.asp"
		else
			Call TipsPage("ԭ����������顣","admin_chpass.asp")

			rs.Close
			cn.Close
			set rs=nothing
			set cn=nothing
			Response.End
		end if
	else
		Call TipsPage("������֤ʧ�ܣ����顣","admin_chpass.asp")

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