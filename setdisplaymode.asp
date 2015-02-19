<%@ CodePage="936" Language="VBScript" %>
<!-- #include file="common.asp" -->
<%
if Request.QueryString("modeflag")="guest" then
	SetTimelessCookie InstanceName & "_DisplayMode",Request.QueryString("mode")
elseif Request.QueryString("modeflag")="admin" then
	SetTimelessCookie InstanceName & "_AdminDisplayMode",Request.QueryString("mode")
end if

if Request("rpage")<>"" then
	dim requests
	requests=HtmlDecode(GetRequests)
	if requests<>"" then requests="?" & mid(requests,2)

	Response.Redirect Request("rpage") & requests
end if
%>