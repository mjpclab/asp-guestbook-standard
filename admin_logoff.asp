<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<%
Session.Contents.Remove(InstanceName & "_adminpass")
Response.Redirect "admin_login.asp"
%>