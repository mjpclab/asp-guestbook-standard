<style type="text/css">
body {
	color:<%=TableContentColor%>;
	<%=CssBgColor(PageBackColor)%>
	<%=CssBgImg(PageBackImage)%>
	line-height:<%=CssLineHeight%>;
}

body, input, textarea, h1, h2, h3, h4, h5, h6 {
	font-family:'<%=CssFontFamily%>', sans-serif;
	<%=CssOptionalSize("font-size",CssFontSize)%>
}

p {
	margin:0 0 1em 0;
}


a:link {color:<%=LinkNormal%>;}
a:visited {color:<%=LinkVisited%>;}
a:active {color:<%=LinkActive%>;}
a:hover {color:<%=LinkHover%>;}

textarea,input,select {
	border-color:<%=TableBorderColor%>;
	color:<%=FormColor%>;
	<%=CssBgColor(FormBGC)%>
}

.outerborder {
	<%=CssOptionalSize("max-width",TableWidth)%>
	border-width:<%=TableBorderWidth%>px;
	border-color:<%=TableBorderColor%>;
	padding:<%=TableBorderPadding%>px;
	<%if TableAlign="left" then Response.Write "margin-right:auto;"%>
	<%if TableAlign="right" then Response.Write "margin-left:auto;"%>
	<%if TableAlign="center" then Response.Write "margin-left:auto; margin-right:auto;"%>
	<%=CssBgColor(TableBGC)%>
	<%=CssBgImg(TablePic)%>
}
*html .outerborder {
	<%=CssOptionalSize("width",TableWidth)%>
}


.header {
	color:<%=TitleColor%>;
	<%=CssBgColor(TitleBGC)%>
	<%=CssBgImg(TitlePic)%>
}

.guest-functions {
	margin:<%=WindowSpace%>px 0;
}
.guest-functions .aside .function {
	border-color:<%=TableContentColor%>;
}


.topic {
	margin:<%=WindowSpace%>px 0;
	border-color:<%=TableBorderColor%>;
	<%=CssBgColor(TableContentBGC)%>
}

.topic .message {
	border-color:<%=TableBorderColor%>;
	color:<%=TableContentColor%>;
}

.topic .info {
	width:<%=TableLeftWidth%>px;
	color:<%=TableGuestInfoColor%>;
	border-color:<%=TableBorderColor%>;
	<%=CssBgColor(TableGuestInfoBGC)%>
	<%=CssBgImg(TableGuestInfoPic)%>
}

.topic .detail {
	border-color:<%=TableBorderColor%>;
}

.topic .title {
	border-color:<%=TableBorderColor%>;
	color:<%=TableTitleColor%>;
	<%=CssBgColor(TableTitleBGC)%>
	<%=CssBgImg(TableTitlePic)%>
}

.topic .inner-message {
	border-color:<%=TableBorderColor%>;
	<%=CssBgColor(BlockBGC)%>
}

.topic .inner-message .summary {
	color:<%=TableTitleColor%>;
	border-color:<%=TableBorderColor%>;
	font-weight: bold;
}

.topic .admin-message .title,
.topic .admin-message .summary {
	color:<%=TableReplyColor%>;
}

.topic .admin-message .words {
	color:<%=TableReplyContentColor%>;
}

.topic .outer-hint {
	border-color:<%=TableBorderColor%>;
}

.topic .admin-tools {
	border-color:<%=TableBorderColor%>;
}

.topic-list {
	margin:<%=WindowSpace%>px 0;
	border-color:<%=TableBorderColor%>;
	<%=CssBgColor(TableContentBGC)%>
}

.topic-list tbody tr:hover td {
	background:<%=BlockBGC%>;
}

.topic-list th,
.topic-list td {
	border-color:<%=TableBorderColor%>;
}
.topic-list th {
	color:<%=TableTitleColor%>;
	<%=CssBgColor(TableTitleBGC)%>
	<%=CssBgImg(TableTitlePic)%>
}



.region {
	margin-top:<%=WindowSpace%>px;
	border-color:<%=TableBorderColor%>;
	color:<%=TableContentColor%>;
	<%=CssBgColor(TableContentBGC)%>
}

.region .title {
	border-color:<%=TableBorderColor%>;
	color:<%=TableTitleColor%>;
	<%=CssBgColor(TableTitleBGC)%>
	<%=CssBgImg(TableTitlePic)%>
}

.region .content table {
	border-color:<%=TableBorderColor%>;
}

.region .content table th,
.region .content table td {
	border-color:<%=TableBorderColor%>;
}

.page-list .nav a {
	border-color:<%=TableContentBGC%>;
}
.page-list .nav a:hover {
	border-color:<%=TableBorderColor%>;
	<%=CssBgColor(BlockBGC)%>
}

.page-list .pagenum,
.page-list .pagenum:link,
.page-list .pagenum:visited,
.page-list .pagenum:active,
.page-list .pagenum:hover {
	border-color:<%=TableContentBGC%>;
    color:<%=PageNumColor_Normal%>;
}
.page-list .pagenum:hover {
	border-color:<%=TableBorderColor%>;
	<%=CssBgColor(BlockBGC)%>
}

.page-list .pagenum-current,
.page-list .pagenum-current:link,
.page-list .pagenum-current:visited,
.page-list .pagenum-current:active,
.page-list .pagenum-current:hover {
	color:<%=PageNumColor_Curr%>;
}


.tab-outer-container .tab-title-container {
	border-color:<%=TableBorderColor%>;
}
.tab-outer-container .tab-title {
	border-color:transparent;
	border-bottom-color:<%=TableBorderColor%>;
}
*html .tab-outer-container .tab-title {
	border-color:<%=TableContentBGC%>;
	border-bottom-color:<%=TableBorderColor%>;
}


.tab-outer-container .tab-title-selected {
	border-color:<%=TableBorderColor%>;
	border-bottom-color:transparent;
	<%=CssBgColor(TableContentBGC)%>
}
*html .tab-outer-container .tab-title-selected {
	border-color:<%=TableBorderColor%>;
	border-bottom-color:<%=TableContentBGC%>;
}
*+html .tab-outer-container .tab-title-selected {
	border-bottom-color:<%=TableContentBGC%>;
}


.ubbtoolbar {
	<%=CssOptionalSize("width",UbbToolWidth)%>
	<%=CssOptionalSize("height",UbbToolHeight)%>
}

.ubbtoolbar .ubbbutton,
.ubbtoolbar .ubbface {
	border-color:<%=TableContentBGC%>;
}
.ubbtoolbar .ubbbutton:hover,
.ubbtoolbar .ubbface:hover {
	border-color:<%=TableBorderColor%>;
	<%=CssBgColor(BlockBGC)%>
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
.footer, .footer a {
	color:#888;
}
.footer a:hover {
	color:#000;
}

<%=Additional_Css%>
</style>
