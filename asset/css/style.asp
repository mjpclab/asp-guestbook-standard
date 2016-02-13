<style type="text/css">
body {
	line-height:<%=CssLineHeight%>;
}
body, input, textarea, select, h1, h2, h3, h4, h5, h6 {
	font-family:<%=CssOptionalSeparator(CssFontFamily,",")%>sans-serif;
	<%=CssOptionalSize("font-size",CssFontSize)%>
}

.mainborder {
	margin:<%=WindowSpace%>px auto;
	padding:<%=WindowSpace%>px;
	<%=CssOptionalSize("max-width",TableWidth)%>
}
*html .mainborder {
	<%=CssOptionalSize("width",TableWidth)%>
}

.guest-functions {
	margin:<%=WindowSpace%>px 0;
}

.topic {
	margin:<%=WindowSpace%>px 0;
}
.topic .info {
	width:<%=TableLeftWidth%>px;
}

.topic-list {
	margin:<%=WindowSpace%>px 0;
}

.region {
	margin-top:<%=WindowSpace%>px;
}

.footer {
	<%=CssOptionalSize("max-width",TableWidth)%>
}
*html .footer {
	<%=CssOptionalSize("width",TableWidth)%>
}
</style>
