<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="loadconfig.asp" -->
<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"

if VcodeCount>0 then Session("vcode")=getvcode(VcodeCount)
%>

<!-- #include file="include/template/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=HomeName%> 留言本 管理登录</title>
	<!-- #include file="inc_admin_stylesheet.asp" -->
	
	<script type="text/javascript">
	function submitCheck(obj)
	{
		if(obj.iadminpass.value.length===0)
		{
			alert('请输入密码。');
			obj.iadminpass.focus();
			return false;
		}
		else if(obj.ivcode && obj.ivcode.value.length===0)
		{
			alert('请输入验证码。');
			obj.ivcode.focus();
			return false;
		}
		else
		{
			obj.submit1.disabled=true;
			return true;
		}
	}
	</script>
</head>

<body onload="if(form5.iadminpass.value.length===0)form5.iadminpass.focus()">

<div id="mainborder" class="mainborder narrow-mainborder">
<div class="region form-region narrow-region">
	<h3 class="title">管理员登录</h3>
	<div class="content">
		<form method="post" action="login_verify.asp" name="form5" onsubmit="return submitCheck(this);">
			<input type="hidden" name="referrer" value="<%=Request.QueryString("referrer")%>"/>
			<div class="field">
				<span class="label">密码：</span>
				<span class="value"><input type="password" name="iadminpass" maxlength="32" autofocus="autofocus" /></span>
			</div>
			<%if VcodeCount>0 then%>
			<div class="field">
				<span class="label">验证码：</span>
				<span class="value"><input type="text" name="ivcode" autocomplete="off" /><img id="captcha" class="captcha" src="show_vcode.asp?t=0" /></span>
			</div>
			<%end if%>
			<div class="command">
				<input value="登录" type="submit" name="submit1" />
			</div>
			</form>
	</div>
</div>
</div>
<script type="text/javascript">
	<!-- #include file="asset/js/refresh-captcha.js" -->
</script>
</body>
</html>