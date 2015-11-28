<!-- #include file="config.asp" -->
<!-- #include file="admin_verify.asp" -->
<!-- #include file="common2.asp" -->
<%
Response.Expires=-1
Response.AddHeader "cache-control","private"
%>

<!-- #include file="include/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/metatag.inc" -->
	<title><%=HomeName%> 留言本 管理首页</title>
	<!-- #include file="inc_admin_stylesheet.asp" -->
</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<%
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype

Dim WordsPerPage
if AdminDisplayMode()="book" then
	WordsPerPage=ItemsPerPage
elseif AdminDisplayMode()="forum" then
	WordsPerPage=TitlesPerPage
end if

Dim ItemsCount,PagesCount,CurrentItemsCount,ipage
get_divided_page cn,rs,sql_pk_main,sql_admin_words_count,sql_admin_words_query,"parent_id INC,lastupdated DEC,id DEC",Request.QueryString("page"),WordsPerPage,ItemsCount,PagesCount,CurrentItemsCount,ipage
%>

<div id="outerborder" class="outerborder">

	<%if ShowTitle=true then show_book_title 3,"管理"%>
	<!-- #include file="include/admin_mainmenu.inc" -->
	<!-- #include file="include/topbulletin.inc" -->
	<%if PagesCount>1 and ShowTopPageList then show_page_list ipage,PagesCount,"admin.asp","[留言分页]",""%>

	<form method="post" action="admin_mdel.asp" name="form7">
		<%RPage="admin.asp"%><!-- #include file="include/admin_func.inc" -->
		<%
		if ItemsCount=0 then
			Response.Write "<br/><br/><div class=""centertext"">目前尚无留言。</div><br/><br/>"
		else
			dim pagename
			pagename="admin"
			if AdminDisplayMode()="book" then
				%><!-- #include file="include/admin_listword.inc" --><%
			elseif AdminDisplayMode()="forum" then
				%><!-- #include file="include/admin_listtitle.inc" --><%
			end if
			rs.Close
		end if
		cn.Close : set rs=nothing : set cn=nothing%>

		<input type="hidden" name="page" value="<%=Request.QueryString("page")%>" />
		<!-- #include file="include/admin_func.inc" -->
	</form>

	<%if PagesCount>1 and ShowBottomPageList then show_page_list ipage,PagesCount,"admin.asp","[留言分页]",""%>
</div>

<!-- #include file="include/footer.inc" -->
</body>
</html>