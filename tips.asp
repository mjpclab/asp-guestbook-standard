<%
sub TipsPage(strTips,backPage)
	%>
	<!-- #include file="include/template/dtd.inc" -->
	<html>
	<head>
		<!-- #include file="include/template/metatag.inc" -->
		<title><%=HomeName%> ���Ա�</title>
		<!-- #include file="inc_stylesheet.asp" -->
	</head>
	<body>
		<p><%=server.HtmlEncode(strTips)%></p>
		<p><a href="<%=backPage%>">[����]</a></p>
		<script type="text/javascript" defer="defer" async="async">
			alert('<%=Replace(strTips,"'","\'")%>');window.location.replace('<%=backPage%>');
		</script>
	</body>
	</html>
	<%
end sub
%>
