<%
sub TipsPage(strTips,backPage)
	%>
	<!-- #include file="include/template/dtd.inc" -->
	<html lang="zh-CN">
	<head>
		<!-- #include file="include/template/metatag.inc" -->
		<title><%=HomeName%> 留言本</title>
		<!-- #include file="inc_stylesheet.asp" -->
	</head>
	<body>
		<div id="mainborder" class="mainborder">
		<p><%=HtmlEncode(strTips)%></p>
		<p><a href="<%=backPage%>">[返回]</a></p>
		</div>
		<script type="text/javascript" defer="defer" async="async">
			alert('<%=Replace(strTips,"'","\'")%>');window.location.replace('<%=backPage%>');
		</script>
	</body>
	</html>
	<%
end sub
%>
