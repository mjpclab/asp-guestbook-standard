<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<%
Session.Contents(InstanceName & "_adminpass")=empty
Response.Redirect "index.asp"
%>