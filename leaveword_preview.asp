<!-- #include file="config.asp" -->

<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"

dim PreviewContent
PreviewContent=Request.Form("icontent")
convertstr PreviewContent,guestlimit,false

Response.Write PreviewContent
%>