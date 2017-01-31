<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/sql/admin_noreplyflag.asp" -->
<!-- #include file="include/sql/admin_delmessage.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->
<%
Response.Expires=-1
if isnumeric(Request.QueryString("id"))=false or Request.QueryString("id")="" then
	Response.Redirect "admin.asp"
	Response.End 
end if

set cn=server.CreateObject("ADODB.Connection")
Call CreateConn(cn)

cn.BeginTrans
cn.Execute Replace(sql_noguestreply_flag,"{0}",Request.QueryString("id")),,129
cn.Execute Replace(sql_admindelmessage_reply,"{0}",Request.QueryString("id")),,129
cn.Execute Replace(sql_admindelmessage_main,"{0}",Request.QueryString("id")),,129
cn.CommitTrans

cn.close : set cn=nothing
%>
<!-- #include file="include/template/admin_traceback.inc" -->