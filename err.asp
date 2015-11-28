<!-- #include file="config.asp" -->

<%Response.Expires=-1%>

<!-- #include file="include/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/metatag.inc" -->
	<title><%=HomeName%> 留言本 错误</title>
	<!-- #include file="inc_stylesheet.asp" -->
</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<div id="outerborder" class="outerborder">

	<%if ShowTitle=true then show_book_title 3,"错误"%>

	<span style="float:right; clear:all; width:100%; text-align:right; padding:4px; color:<%=LinkNormal%>;">
			| <a href="admin_login.asp">管理</a>
	</span>
	<%
	dim errmsg
	select case Request("number")
	case "1"
		errmsg="对不起，您的访问已被禁止。"
	case "2"
		errmsg="抱歉，留言本已关闭，请稍候再试。"
	case "3"
		errmsg="抱歉，留言权限已关闭，请稍候再试。"
	case "4"
		errmsg="对不起，您的留言中含有禁止出现的内容。"
	case "5"
		errmsg="抱歉，搜索权限已关闭，请稍候再试。"
	case "6"
		errmsg="对不起，您的发言速度太快了，请休息一下。"
	case "7"
		errmsg="对不起，请不要发送重复内容。"
	case else
		errmsg="未知错误，请联系管理员。"
	end select
	
	Response.Write "<br/><br/><span style=""width:100%; text-align:left;"">　　■　" &errmsg& "</span><br/><br/><br/>"
	%>
</div>

<!-- #include file="include/footer.inc" -->
</body>
</html>