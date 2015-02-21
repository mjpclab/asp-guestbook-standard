<!-- #include file="config.asp" -->
<!-- #include file="admin_verify.asp" -->

<%Response.Expires=-1%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=HomeName%> 留言本 高级删除</title>
	<link rel="stylesheet" type="text/css" href="style.css"/>
	<link rel="stylesheet" type="text/css" href="adminstyle.css"/>
	<!-- #include file="style.asp" -->
	<!-- #include file="adminstyle.asp" -->
</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<div id="outerborder" class="outerborder">

	<%if ShowTitle=true then show_book_title 3,"管理"%>
	<!-- #include file="admincontrols.inc" -->

	<div class="region form-region">
		<h3 class="title">高级删除</h3>
		<div class="content">
			<form method="post" action="admin_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				删除指定日期和时间前的留言，包括此日期/时间。<br/>
				<input type="hidden" name="option" value="1" />
				<input type="text" name="iparam" size="<%=AdvDelTextWidth%>" maxlength="20" />
				<input type="submit" value="执行" name="submit1"<%if DelAdvTip=true then Response.Write " onclick=""return confirm('确实要执行删除操作吗？');"""%> />(时间默认为0:0:0)
			</form>
			<form method="post" action="admin_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				删除指定日期和时间后的留言，包括此日期/时间。<br/>
				<input type="hidden" name="option" value="2" />
				<input type="text" name="iparam" size="<%=AdvDelTextWidth%>" maxlength="20" />
				<input type="submit" value="执行" name="submit1"<%if DelAdvTip=true then Response.Write " onclick=""return confirm('确实要执行删除操作吗？');"""%> />(时间默认为0:0:0)
			</form>
			<form method="post" action="admin_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				删除最靠前(最老)的n条留言，请输入n的值：<br/>
				<input type="hidden" name="option" value="3" />
				<input type="text" name="iparam" size="<%=AdvDelTextWidth%>" maxlength="10" />
				<input type="submit" value="执行" name="submit1"<%if DelAdvTip=true then Response.Write " onclick=""return confirm('确实要执行删除操作吗？');"""%> />
			</form>
			<form method="post" action="admin_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				删除最靠后(最新)的n条留言，请输入n的值：<br/>
				<input type="hidden" name="option" value="4" />
				<input type="text" name="iparam" size="<%=AdvDelTextWidth%>" maxlength="10" />
				<input type="submit" value="执行" name="submit1"<%if DelAdvTip=true then Response.Write " onclick=""return confirm('确实要执行删除操作吗？');"""%> />
			</form>
			<form method="post" action="admin_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				删除称呼中包含特定字符串的留言：<br/>
				<input type="hidden" name="option" value="5" />
				<input type="text" name="iparam" size="<%=AdvDelTextWidth%>" maxlength="64" />
				<input type="submit" value="执行" name="submit1"<%if DelAdvTip=true then Response.Write " onclick=""return confirm('确实要执行删除操作吗？');"""%> />("%"为多个字符，"_"为一个字符)
			</form>
			<form method="post" action="admin_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				删除标题中包含特定字符串的留言：<br/>
				<input type="hidden" name="option" value="6" />
				<input type="text" name="iparam" size="<%=AdvDelTextWidth%>" maxlength="64" />
				<input type="submit" value="执行" name="submit1"<%if DelAdvTip=true then Response.Write " onclick=""return confirm('确实要执行删除操作吗？');"""%> />("%"为多个字符，"_"为一个字符)
			</form>
			<form method="post" action="admin_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				删除留言正文中包含特定字符串的留言：<br/>
				<input type="hidden" name="option" value="7" />
				<input type="text" name="iparam" size="<%=AdvDelTextWidth%>" maxlength="64" />
				<input type="submit" value="执行" name="submit1"<%if DelAdvTip=true then Response.Write " onclick=""return confirm('确实要执行删除操作吗？');"""%> />("%"为多个字符，"_"为一个字符)
			</form>
			<form method="post" action="admin_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				删除全部留言，清空留言本，请谨慎使用：
				<input type="hidden" name="option" value="8" />
				<input type="hidden" name="iparam" value="DEL_ALL" />
				<input type="submit" value="执行" name="submit1"<%if DelAdvTip=true then Response.Write " onclick=""return confirm('确实要执行删除操作吗？');"""%> />
			</form>
		</div>
	</div>

</div>

<!-- #include file="bottom.asp" -->
</body>
</html>