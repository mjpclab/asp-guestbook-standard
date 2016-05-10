<%
Function GetReferrer
	Dim QueryString
	QueryString=Request.QueryString
	if QueryString="" then
		GetReferrer = Request.ServerVariables("URL")
	else
		GetReferrer = Request.ServerVariables("URL") & "?" & Request.QueryString
	end if
End Function

Response.Expires=-1
Dim cn,rs,refPass
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
Call CreateConn(cn)
rs.Open sql_adminverify,cn,0,1,1
if Not rs.EOF then
	refPass=rs.Fields(0)
end if
rs.Close : cn.Close : set rs=nothing : set cn=nothing

if refPass="" Or Session(InstanceName & "_adminpass")<>refPass then
	Response.Redirect "admin_login.asp?referrer=" & Server.UrlEncode(GetReferrer())
	Response.End
end if
%>