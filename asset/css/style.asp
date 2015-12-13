<style type="text/css">
body {
	line-height:<%=CssLineHeight%>;
}
body, input, textarea, h1, h2, h3, h4, h5, h6 {
	font-family:'<%=CssFontFamily%>', sans-serif;
	<%=CssOptionalSize("font-size",CssFontSize)%>
}

.outerborder {
	<%=CssOptionalSize("max-width",TableWidth)%>
	<%if TableAlign="left" then Response.Write "margin-right:auto;"%>
	<%if TableAlign="right" then Response.Write "margin-left:auto;"%>
	<%if TableAlign="center" then Response.Write "margin-left:auto; margin-right:auto;"%>
	<%if ShowBorder=false then%>border:0 none transparent;<%end if%>
}
*html .outerborder {
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
	<%if TableAlign="left" then Response.Write "margin-right:auto;"%>
	<%if TableAlign="right" then Response.Write "margin-left:auto;"%>
	<%if TableAlign="center" then Response.Write "margin-left:auto; margin-right:auto;"%>
}
*html .footer {
	<%=CssOptionalSize("width",TableWidth)%>
}
</style>
