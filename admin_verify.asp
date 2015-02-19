<!-- #include file="common.asp" -->

<%
Response.Expires=-1
Dim cn,rs
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype
rs.Open sql_adminverify,cn,0,1,1

if rs.EOF=false then
	if session.Contents(InstanceName & "_adminpass")<>rs(0) then
		rs.Close : cn.Close : set rs=nothing : set cn=nothing
		Response.Redirect "admin_login.asp"
		Response.End
	end if
else
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	Response.Redirect "admin_login.asp"
	Response.End
end if

rs.Close : cn.Close : set rs=nothing : set cn=nothing
%>