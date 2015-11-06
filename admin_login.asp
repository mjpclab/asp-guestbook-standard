<!-- #include file="config.asp" -->
<%
Response.Expires = -1
Response.AddHeader "Pragma","no-cache"
Response.AddHeader "cache-control","no-cache, must-revalidate"

if VcodeCount>0 then session("vcode")=getvcode(VcodeCount)
%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
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

<div class="region form-region narrow-region">
	<h3 class="title">管理员登录</h3>
	<div class="content">
		<form method="post" action="login_verify.asp" name="form5" onsubmit="return submitCheck(this);">
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
<script type="text/javascript">
	<!-- #include file="js/refresh-captcha.js" -->
</script>
</body>
</html>