<!-- #include file="config.asp" -->
<!-- #include file="admin_verify.asp" -->

<%
Response.Expires=-1
set cn=server.CreateObject("ADODB.Connection")
CreateConn cn,dbtype

dim tlimit
tlimit=0
if Request.Form("html2")="1" then tlimit=tlimit+1
if Request.Form("ubb2")="1" then tlimit=tlimit+2
if Request.Form("newline2")="1" then tlimit=tlimit+4

dim tbul
tbul=Request.Form("abulletin")
tbul=replace(tbul,"'","''")
tbul=replace(tbul,"<%","< %")

cn.Execute Replace(Replace(sql_adminsavebulletin,"{0}",tlimit),"{1}",tbul),,1

cn.close
set cn=nothing
		
Response.Redirect "admin.asp"
%>