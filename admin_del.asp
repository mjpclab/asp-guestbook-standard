<!-- #include file="config.asp" -->
<!-- #include file="admin_verify.asp" -->

<%
Response.Expires=-1
if isnumeric(Request.QueryString("id"))=false or Request.QueryString("id")="" then
	Response.Redirect "admin.asp"
	Response.End 
end if

set cn=server.CreateObject("ADODB.Connection")
CreateConn cn,dbtype

cn.BeginTrans
cn.Execute Replace(sql_global_noguestreply_flag,"{0}",Request.QueryString("id")),,1
cn.Execute Replace(sql_admindel_reply,"{0}",Request.QueryString("id")),,1
cn.Execute Replace(sql_admindel_main,"{0}",Request.QueryString("id")),,1
cn.CommitTrans

cn.close : set cn=nothing
%>
<!-- #include file="admin_traceback.inc" -->