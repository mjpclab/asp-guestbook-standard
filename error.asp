<%
sub ErrorPage(errorCode)
	Response.Status="403 Forbidden"

	dim errmsg
	select case errorCode
	case 1
		errmsg="对不起，您的访问已被禁止。"
	case 2
		errmsg="抱歉，留言本已关闭，请稍候再试。"
	case 3
		errmsg="抱歉，留言权限已关闭，请稍候再试。"
	case 4
		errmsg="对不起，您的留言中含有禁止出现的内容。"
	case 5
		errmsg="抱歉，搜索权限已关闭，请稍候再试。"
	case 6
		errmsg="对不起，您的发言速度太快了，请休息一下。"
	case 7
		errmsg="对不起，请不要发送重复内容。"
	case else
		errmsg="未知错误，请联系管理员。"
	end select
	%>
	<!-- #include file="include/template/dtd.inc" -->
	<html>
	<head>
		<!-- #include file="include/template/metatag.inc" -->
		<title><%=HomeName%> 留言本 错误</title>
		<!-- #include file="inc_stylesheet.asp" -->
	</head>

	<body<%=bodylimit%> onload="<%=framecheck%>">

	<div id="outerborder" class="outerborder">

		<%if ShowTitle=true then show_book_title 3,"错误"%>

		<div id="mainborder" class="mainborder">
		<div class="guest-functions">
			<div class="aside">
				<a class="function" href="admin.asp">管理</a>
			</div>
		</div>

		<p style="margin-bottom: 3em;">　　■　<%=errmsg%></p>
		</div>

		<!-- #include file="include/template/footer.inc" -->
	</div>
	</body>
	</html>
	<%
end sub
%>
