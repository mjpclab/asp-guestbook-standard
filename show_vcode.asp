<!-- #include file="config/system.asp" -->
<!-- #include file="include/utility/captcha.asp" -->
<%
if Request.QueryString("type")="write" then
	Call OutputCaptcha(Session(InstanceName & "_vcode_write"))
else
	Call OutputCaptcha(Session(InstanceName & "_vcode"))
end if
%>