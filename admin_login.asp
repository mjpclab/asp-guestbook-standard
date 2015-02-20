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
	<link rel="stylesheet" type="text/css" href="style.css"/>
	<!-- #include file="style.asp" -->
	
	<script type="text/javascript">
	//<![CDATA[
	function submitCheck(obj)
	{
		if(obj.iadminpass.value=='')
		{
			alert('请输入密码。');
			obj.iadminpass.focus();
			return false;
		}
		else if(obj.ivcode && obj.ivcode.value=='')
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
	//]]>
	</script>
</head>

<body onload="if(form5.iadminpass.value=='')form5.iadminpass.focus()" style="text-align:center;">

<br/>
<table border="1" cellpadding="2" style="width:300px; border:solid 1px <%=TableBorderColor%>; border-collapse:collapse; margin-left:auto; margin-right:auto;">
	<tr>
		<td class="centertitle">管理员登录</td>
	</tr>
	<tr>
		<td class="wordscontent" style="text-align:center; padding:20px 0px;">
			<form method="post" action="login_verify.asp" name="form5" onsubmit="return submitCheck(this);">
			<table style="border-width:0px; margin-left:auto; margin-right:auto;" cellpadding="2" cellspacing="0">
				<tr>
					<td>密　码：</td>
					<td><input type="password" name="iadminpass" size="26" maxlength="32" /></td>
				</tr>
				<%if VcodeCount>0 then%>
				<tr style="height:10px;"><td></td></tr>
				<tr>
					<td>验证码：</td>
					<td><input type="text" name="ivcode" size="10" /> <img alt="" src="show_vcode.asp" class="vcode" onclick="this.src=this.src" /></td>
				</tr>
				<%end if%>
				<tr style="height:15px;"><td></td></tr>
				<tr>
					<td style="width:100%; text-align:center;" colspan="2">
						<input value="登录" type="submit" name="submit1" />
					</td>
				</tr>
			</table>
			</form>
		</td>
	</tr>
</table>
</body>
</html>