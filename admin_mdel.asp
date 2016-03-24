<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/sql/admin_noreplyflag.asp" -->
<!-- #include file="include/sql/admin_mdel.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->
<%
Response.Expires=-1
if Request.Form("seltodel")<>"" then
	dim ids
	ids=FilterSql(Request.Form("seltodel"))
	if Left(ids,1) = "," then
		ids=Mid(ids,2)
	end if

	set cn=server.CreateObject("ADODB.Connection")
	Call CreateConn(cn)

	cn.BeginTrans
	cn.Execute Replace(sql_noguestreply_flag,"{0}",ids),,1
	cn.Execute Replace(sql_adminmdel_reply,"{0}",ids),,1
	cn.Execute Replace(sql_adminmdel_main,"{0}",ids),,1
	cn.CommitTrans

	cn.Close : set cn=nothing
end if
%>
<!-- #include file="include/template/admin_traceback.inc" -->