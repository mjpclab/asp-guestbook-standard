<!-- #include file="config.asp" -->
<!-- #include file="admin_verify.asp" -->

<%Response.Expires=-1%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=HomeName%> 留言本 修改管理员密码</title>
	<link rel="stylesheet" type="text/css" href="css/style.css"/>
	<link rel="stylesheet" type="text/css" href="css/adminstyle.css"/>
	<!-- #include file="css/style.asp" -->
	<!-- #include file="css/adminstyle.asp" -->

	<script type="text/javascript">
	function checkpass(cobject)
	{
		if (cobject.ioldpass.value.length===0) {alert('请输入原密码。'); cobject.ioldpass.focus(); return(false);}
		if (cobject.inewpass1.value.length===0) {alert('请输入新密码。'); cobject.inewpass1.focus(); return(false);}
		if (cobject.inewpass2.value.length===0) {alert('请输入确认密码。'); cobject.inewpass2.focus(); return(false);}
		if (cobject.inewpass1.value!==cobject.inewpass2.value) {alert('新密码与确认密码不同，请重新输入。'); cobject.inewpass1.focus(); return(false);}
		cobject.submit1.disabled=true;
		return (true);
	}
	</script>
</head>

<body<%=bodylimit%> onload="form4.ioldpass.focus();<%=framecheck%>">

<div id="outerborder" class="outerborder">

	<%show_book_title 3,"管理"%>
	<!-- #include file="admincontrols.inc" -->

	<div class="region form-region">
		<h3 class="title">修改密码</h3>
		<div class="content">
			<form method="post" action="admin_savepass.asp" onsubmit="return checkpass(this)" name="form4">
			<div class="field">
				<span class="label">原密码：</span>
				<span class="value"><input type="password" name="ioldpass" size="<%=SetInfoTextWidth%>" maxlength="32" /></span>
			</div>
			<div class="field">
				<span class="label">新密码：</span>
				<span class="value"><input type="password" name="inewpass1" size="<%=SetInfoTextWidth%>" maxlength="32" /></span>
			</div>
			<div class="field">
				<span class="label">确认密码：</span>
				<span class="value"><input type="password" name="inewpass2" size="<%=SetInfoTextWidth%>" maxlength="32" /></span>
			</div>
			<div class="command"><input value="更新数据" type="submit" name="submit1" /></div>
			</form>
		</div>
	</div>

</div>

<!-- #include file="bottom.asp" -->
</body>
</html>