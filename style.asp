<style type="text/css">
body {
	color:<%=TableContentColor%>;
	<%=CssBgColor(PageBackColor)%>
	<%=CssBgImg(PageBackImage)%>
}

body, input, textarea {
	font-family:<%=CssFontFamily%>, sans-serif;
	font-size:<%=CssFontSize%>;
}

.outerborder {
	width:<%=TableWidth%><%if right(cstr(TableWidth),1)<>"%" then response.write "px"%>;
	border-width:<%=TableBorderWidth%>px;
	border-color:<%=TableBorderColor%>;
	padding:<%=TableBorderPadding%>px;
	<%if TableAlign="left" then Response.Write "margin-right:auto;"%>
	<%if TableAlign="right" then Response.Write "margin-left:auto;"%>
	<%if TableAlign="center" then Response.Write "margin-left:auto; margin-right:auto;"%>
	<%=CssBgColor(TableBGC)%>
	<%=CssBgImg(TablePic)%>
}

th,td,p {
	line-height:<%=CssLineHeight%>;
}
textarea,input,select {
	border-color:<%=TableBorderColor%>;
	color:<%=FormColor%>;
	<%=CssBgColor(FormBGC)%>
}
hr {
	color:<%=TableBorderColor%>;
}
div {
	line-height:<%=CssLineHeight%>;
}

a:link {color:<%=LinkNormal%>;}
a:visited {color:<%=LinkVisited%>;}
a:active {color:<%=LinkActive%>;}
a:hover {color:<%=LinkHover%>;}

.booktitle			{color:<%=TitleColor%>;}
a.booktitle {color:<%=TitleColor%>;}

img.pageicon	{border-color:<%=TableContentBGC%>;}
img.ubbbutton	{border-color:<%=TableContentBGC%>;}
img.smallface	{border-color:<%=TableContentBGC%>;}

.generalwindow {
	border-color:<%=TableBorderColor%>;
	margin-top:<%=WindowSpace%>px;
}

table.grid {
	border-color:<%=TableBorderColor%>;
	color:<%=TableContentColor%>;
	<%=CssBgColor(TableContentBGC)%>
}
.grid th, .grid td {border-color:<%=TableBorderColor%>;}

table.onetopic		{border-color:<%=TableBorderColor%>; margin-top:<%=WindowSpace%>px;}
td.centertitle		{border-color:<%=TableBorderColor%>; color:<%=TableTitleColor%>; <%=CssBgColor(TableTitleBGC)%> <%=CssBgImg(TableTitlePic)%>}
td.wordstitle		{border-bottom-color:<%=TableBorderColor%>; color:<%=TableTitleColor%>; <%=CssBgColor(TableTitleBGC)%> <%=CssBgImg(TableTitlePic)%>}
td.bulletintitle	{border-bottom-color:<%=TableBorderColor%>; color:<%=TableBulletinTitleColor%>; <%=CssBgColor(TableBulletinTitleBGC)%> <%=CssBgImg(TableBulletinTitlePic)%>}
td.replytitle		{border-color:<%=TableBorderColor%>; color:<%=TableReplyColor%>; <%=CssBgColor(TableReplyBGC)%> <%=CssBgImg(TableReplyPic)%>}
td.tableleft		{width:<%=TableLeftWidth%>px; border-color:<%=TableBorderColor%>; color:<%=TableGuestInfoColor%>; <%=CssBgColor(TableGuestInfoBGC)%> <%=CssBgImg(TableGuestInfoPic)%>}
td.tableright		{border-color:<%=TableBorderColor%>; background-color:<%=TableContentBGC%>;}
td.wordscontent		{border-color:<%=TableBorderColor%>; color:<%=TableContentColor%>; <%=CssBgColor(TableContentBGC)%>}
td.vguestpass		{border-color:<%=TableBorderColor%>; color:<%=TableContentColor%>; <%=CssBgColor(TableContentBGC)%>}
tr.header			{color:<%=TableTitleColor%>; <%=CssBgColor(TableTitleBGC)%> <%=CssBgImg(TableTitlePic)%>}

td.embedbox			{color:<%=TableContentColor%>; <%=CssBgColor(TableContentBGC)%>}
div.embedbox		{border-color:<%=TableBorderColor%>; <%=CssBgColor(BlockBGC)%>}

.tabTitle			{border:solid 1px <%=TableContentBGC%>;}
.tabTitleOver		{border-color:<%=TableBorderColor%>; <%=CssBgColor(BlockBGC)%>}
.tabTitleSelected	{<%=CssBgColor(TableTitleBGC)%> <%=CssBgImg(TableTitlePic)%>}

.pagenum_curr, .pagenum_curr:link, .pagenum_curr:visited, .pagenum_curr:active, .pagenum_curr:hover	 {color:<%=PageNumColor_Curr%>;}
.pagenum_normal, .pagenum_normal:link, .pagenum_normal:visited, .pagenum_normal:active, .pagenum_normal:hover {color:<%=PageNumColor_Normal%>;}
<%=Additional_Css%>
</style>
