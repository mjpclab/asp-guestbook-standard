<%
sub TipsPage(strTips,backPage)
	%>
	<!-- #include file="include/template/dtd.inc" -->
	<html>
	<head>
		<!-- #include file="include/template/metatag.inc" -->
		<title><%=HomeName%> ¡Ù—‘±æ</title>
		<!-- #include file="inc_stylesheet.asp" -->
	</head>
	<body>
		<div id="mainborder" class="mainborder">
		<p><%=HtmlEncode(strTips)%></p>
		<p><a href="<%=backPage%>">[∑µªÿ]</a></p>
		</div>
		<script type="text/javascript" defer="defer" async="async">
			alert('<%=Replace(strTips,"'","\'")%>');window.location.replace('<%=backPage%>');
		</script>
	</body>
	</html>
	<%
end sub
%>
