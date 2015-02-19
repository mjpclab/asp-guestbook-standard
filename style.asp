<style type="text/css">
* {font-family:<%=CssFontFamily%>, sans-serif; font-size:<%=CssFontSize%>;}
body {
	margin:0px;
	padding:10px;
	text-align:<%=TableAlign%>;
	font-style:normal;
	color:<%=TableContentColor%>;
	<%=CssBgColor(PageBackColor)%>
	<%=CssBgImg(PageBackImage)%>
}

.outerborder {
	width:<%=TableWidth%><%if right(cstr(TableWidth),1)<>"%" then response.write "px"%>;
	overflow:visible;
	text-align:left;
	border:solid <%=TableBorderWidth%>px <%=TableBorderColor%>;
	padding:<%=TableBorderPadding%>px;
	margin:0px;
	<%if TableAlign="left" then Response.Write "margin-right:auto;"%>
	<%if TableAlign="right" then Response.Write "margin-left:auto;"%>
	<%if TableAlign="center" then Response.Write "margin-left:auto; margin-right:auto;"%>
	<%=CssBgColor(TableBGC)%>
	<%=CssBgImg(TablePic)%>
}

th,td,p					{line-height:<%=CssLineHeight%>;}
textarea,input,select	{border:solid 1px <%=TableBorderColor%>; color:<%=FormColor%>; <%=CssBgColor(FormBGC)%>}
textarea				{display:block;}
hr						{width:100%; height:1px; color:<%=TableBorderColor%>;}
div						{line-height:<%=CssLineHeight%>;}
.centertext				{border-style:none; border-width:0px; width:100%; text-align:center; color:#808080; display:block;}

a:link		{color:<%=LinkNormal%>; text-decoration: none;}
a:visited	{color:<%=LinkVisited%>; text-decoration: none;}
a:active	{color:<%=LinkActive%>; text-decoration: none;}
a:hover		{color:<%=LinkHover%>; text-decoration: underline;}

.booktitle			{color:<%=TitleColor%>;}
a.booktitle:link	{color:<%=TitleColor%>; text-decoration: none;}
a.booktitle:visited {color:<%=TitleColor%>; text-decoration: none;}
a.booktitle:active	{color:<%=TitleColor%>; text-decoration: none;}
a.booktitle:hover	{color:<%=TitleColor%>; text-decoration: underline;}

form			{margin:0px; padding:0px; border-width:0px; border-style:none;}

img				{border-style:none; border-width:0px;}
img.imgicon		{width:16px; height:16px; vertical-align:text-bottom; border-style:none; border-width:0px;}
img.ipicon		{width:13px; height:15px; vertical-align:text-bottom; border-style:none; border-width:0px;}
img.forumicon	{vertical-align:middle; margin:2px;}
img.pageicon	{width:16px; height:16px; vertical-align:text-bottom; border:solid 1px <%=TableContentBGC%>;}
img.ubbbutton	{cursor:pointer; border:solid 1px <%=TableContentBGC%>; margin:1px; width:16px; height:16px;}
img.smallface	{cursor:pointer; border:solid 1px <%=TableContentBGC%>; margin:1px;}
div.colorsquare	{cursor:pointer; border:solid 1px #000000; margin:0px; padding:0px; width:11px; height:11px; float:left;}
img.face		{border-style:none; border-width:0px;}
img.listface	{cursor:pointer; border-style:none; border-width:0px;}
img.innerface	{border-style:none; border-width:0px; float:left; margin:0px 5px 5px 0px;}
img.vcode		{vertical-align:text-bottom;}

.generalwindow		{width:100%; border:solid 1px <%=TableBorderColor%>; border-collapse:collapse; margin-top:<%=WindowSpace%>px;}
.noborder			{border-style:none; border-width:0px;}

table.grid			{width:100%; border:solid 1px <%=TableBorderColor%>; border-collapse:collapse; color:<%=TableContentColor%>; <%=CssBgColor(TableContentBGC)%>}
.grid th, .grid td	{border:solid 1px <%=TableBorderColor%>;}

table.onetopic		{width:100%; overflow:hidden; border:solid 1px <%=TableBorderColor%>; border-collapse:collapse; table-layout:fixed; margin-top:<%=WindowSpace%>px;}
td.centertitle		{width:100%; height:25px; border:solid 1px <%=TableBorderColor%>; text-align:center; vertical-align:middle; font-weight:bold; color:<%=TableTitleColor%>; <%=CssBgColor(TableTitleBGC)%> <%=CssBgImg(TableTitlePic)%>}
td.wordstitle		{height:25px; text-align:left; vertical-align:middle; border-bottom:solid 1px <%=TableBorderColor%>; font-weight:bold; color:<%=TableTitleColor%>; <%=CssBgColor(TableTitleBGC)%> <%=CssBgImg(TableTitlePic)%>}
td.bulletintitle	{height:25px; text-align:left; vertical-align:middle; border-bottom:solid 1px <%=TableBorderColor%>; font-weight:bold; color:<%=TableBulletinTitleColor%>; <%=CssBgColor(TableBulletinTitleBGC)%> <%=CssBgImg(TableBulletinTitlePic)%>}
td.replytitle		{height:25px; text-align:left; vertical-align:middle; border:solid 1px <%=TableBorderColor%>; font-weight:bold; color:<%=TableReplyColor%>; <%=CssBgColor(TableReplyBGC)%> <%=CssBgImg(TableReplyPic)%>}
td.tableleft		{width:<%=TableLeftWidth%>px; text-align:center; vertical-align:top; border:solid 1px <%=TableBorderColor%>; padding:3px 0px; overflow:hidden; color:<%=TableGuestInfoColor%>; <%=CssBgColor(TableGuestInfoBGC)%> <%=CssBgImg(TableGuestInfoPic)%>}
td.tableright		{vertical-align:top; border:solid 1px <%=TableBorderColor%>; background-color:<%=TableContentBGC%>;}
td.wordscontent		{overflow:hidden; vertical-align:top; border-color:<%=TableBorderColor%>; margin:0px; padding:10px; color:<%=TableContentColor%>; <%=CssBgColor(TableContentBGC)%>}
td.vguestpass		{width:100%; height:25px; text-align:right; border:solid 1px <%=TableBorderColor%>; margin:0px; padding:2px; color:<%=TableContentColor%>; <%=CssBgColor(TableContentBGC)%>}
tr.header			{font-weight:bold; color:<%=TableTitleColor%>; <%=CssBgColor(TableTitleBGC)%> <%=CssBgImg(TableTitlePic)%>}

td.embedbox			{vertical-align:top; margin:0px; padding:12px 24px 12px 24px; color:<%=TableContentColor%>; <%=CssBgColor(TableContentBGC)%>}
div.embedbox		{margin:0px; padding:10px; border:solid 1px <%=TableBorderColor%>; <%=CssBgColor(BlockBGC)%>}
table.embedbox		{overflow:hidden; border-style:none; border-width:0px; border-collapse:collapse; margin:0px; padding:0px; table-layout:fixed;}

.tabOuterContainer	{width:auto; height:400px;}
.tabTitle			{text-align:center; width:110px; height:25px; line-height:25px; border:solid 1px <%=TableContentBGC%>;}
.tabTitleOver		{border:solid 1px <%=TableBorderColor%>; <%=CssBgColor(BlockBGC)%>}
.tabTitleSelected	{<%=CssBgColor(TableTitleBGC)%> <%=CssBgImg(TableTitlePic)%>}
.tabPageContainer	{width:100%;}
.tabPage			{padding:5px;}

.pagenum_curr, .pagenum_curr:link, .pagenum_curr:visited, .pagenum_curr:active, .pagenum_curr:hover				{color:<%=PageNumColor_Curr%>;}
.pagenum_normal, .pagenum_normal:link, .pagenum_normal:visited, .pagenum_normal:active, .pagenum_normal:hover	{color:<%=PageNumColor_Normal%>;}
<%=Additional_Css%>
</style>
