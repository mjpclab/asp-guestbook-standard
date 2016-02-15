<!-- #include file="include/utility/captcha.asp" -->
<%
if Request.QueryString("type")="write" then
	Call OutputCaptcha(Session("vcode_write"))
else
	Call OutputCaptcha(Session("vcode"))
end if
%>