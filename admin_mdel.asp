<!-- #include file="config.asp" -->
<!-- #include file="admin_verify.asp" -->

<%
Response.Expires=-1
if Request.Form("seltodel")="" then
	if isnumeric(Request.Form("page")) and Request.Form("page")<>"" then
		Response.Redirect "admin.asp?page=" & Request.Form("page")
	else
		Response.Redirect "admin.asp"
	end if
end if

dim ids,iids
ids=split(Request.Form("seltodel"),",")
for each iids in ids
	if isnumeric(iids)=false or iids="" then
		Response.Redirect "admin.asp"
		Response.End
	end if
next

set cn=server.CreateObject("ADODB.Connection")
CreateConn cn,dbtype

cn.BeginTrans
cn.Execute Replace(sql_global_noguestreply_flag,"{0}",Request.Form("seltodel")),,1
cn.Execute Replace(sql_adminmdel_reply,"{0}",Request.Form("seltodel")),,1
cn.Execute Replace(sql_adminmdel_main,"{0}",Request.Form("seltodel")),,1
cn.CommitTrans

cn.Close : set cn=nothing
%>
<!-- #include file="admin_traceback.inc" -->