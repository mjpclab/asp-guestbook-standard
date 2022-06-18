<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/book.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->
<%Response.Expires=-1%>

<!-- #include file="include/template/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=HomeName%> 留言本 高级删除</title>
	<!-- #include file="inc_admin_stylesheet.asp" -->
</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<div id="outerborder" class="outerborder">

	<%if ShowTitle then%><%Call InitHeaderData("管理")%><!-- #include file="include/template/header.inc" --><%end if%>
	<div id="mainborder" class="mainborder">
	<!-- #include file="include/template/admin_mainmenu.inc" -->
	<div class="region form-region">
		<h3 class="title">高级删除</h3>
		<div class="content">
			<form method="post" action="admin_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				删除指定日期和时间前的留言，包括此日期/时间。<br/>
				<input type="hidden" name="option" value="1" />
				<input type="text" name="iparam" maxlength="20" />
				<input type="submit" value="执行" name="submit1"<%if DelAdvTip then Response.Write " onclick=""return confirm('确实要执行删除操作吗？');"""%> />(时间默认为0:0:0)
			</form>
			<form method="post" action="admin_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				删除指定日期和时间后的留言，包括此日期/时间。<br/>
				<input type="hidden" name="option" value="2" />
				<input type="text" name="iparam" maxlength="20" />
				<input type="submit" value="执行" name="submit1"<%if DelAdvTip then Response.Write " onclick=""return confirm('确实要执行删除操作吗？');"""%> />(时间默认为0:0:0)
			</form>
			<form method="post" action="admin_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				删除最靠前(最老)的n条留言，请输入n的值：<br/>
				<input type="hidden" name="option" value="3" />
				<input type="text" name="iparam" maxlength="10" />
				<input type="submit" value="执行" name="submit1"<%if DelAdvTip then Response.Write " onclick=""return confirm('确实要执行删除操作吗？');"""%> />
			</form>
			<form method="post" action="admin_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				删除最靠后(最新)的n条留言，请输入n的值：<br/>
				<input type="hidden" name="option" value="4" />
				<input type="text" name="iparam" maxlength="10" />
				<input type="submit" value="执行" name="submit1"<%if DelAdvTip then Response.Write " onclick=""return confirm('确实要执行删除操作吗？');"""%> />
			</form>
			<form method="post" action="admin_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				删除称呼中包含特定字符串的留言：<br/>
				<input type="hidden" name="option" value="5" />
				<input type="text" name="iparam" maxlength="64" />
				<input type="submit" value="执行" name="submit1"<%if DelAdvTip then Response.Write " onclick=""return confirm('确实要执行删除操作吗？');"""%> />("%"为多个字符，"_"为一个字符)
			</form>
			<form method="post" action="admin_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				删除标题中包含特定字符串的留言：<br/>
				<input type="hidden" name="option" value="6" />
				<input type="text" name="iparam" maxlength="64" />
				<input type="submit" value="执行" name="submit1"<%if DelAdvTip then Response.Write " onclick=""return confirm('确实要执行删除操作吗？');"""%> />("%"为多个字符，"_"为一个字符)
			</form>
			<form method="post" action="admin_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				删除留言正文中包含特定字符串的留言：<br/>
				<input type="hidden" name="option" value="7" />
				<input type="text" name="iparam" maxlength="64" />
				<input type="submit" value="执行" name="submit1"<%if DelAdvTip then Response.Write " onclick=""return confirm('确实要执行删除操作吗？');"""%> />("%"为多个字符，"_"为一个字符)
			</form>
			<form method="post" action="admin_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				删除全部留言，清空留言本，请谨慎使用：
				<input type="hidden" name="option" value="8" />
				<input type="hidden" name="iparam" value="DEL_ALL" />
				<input type="submit" value="执行" name="submit1"<%if DelAdvTip then Response.Write " onclick=""return confirm('确实要执行删除操作吗？');"""%> />
			</form>
		</div>
	</div>
	</div>

	<!-- #include file="include/template/footer.inc" -->
</div>
</body>
</html>