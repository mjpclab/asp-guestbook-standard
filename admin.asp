<!-- #include file="config.asp" -->
<!-- #include file="admin_verify.asp" -->

<%
Response.Expires=-1
Response.AddHeader "cache-control","private"
%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=HomeName%> ���Ա� ������ҳ</title>
	
	<!-- #include file="style.asp" -->
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

	<%if ShowTitle=true then show_book_title 3,"����"%>
	<!-- #include file="admintool.inc" -->
	<!-- #include file="topbulletin.inc" -->
	<%if PagesCount>1 and ShowTopPageList then show_page_list ipage,PagesCount,"admin.asp","[���Է�ҳ]","center",""%>

	<form method="post" action="admin_mdel.asp" name="form7">
		<%RPage="admin.asp"%><!-- #include file="func_admin.inc" -->
		<%
		if ItemsCount=0 then
			Response.Write "<br/><br/><div class=""centertext"">Ŀǰ�������ԡ�</div><br/><br/>"
		else
			dim pagename
			pagename="admin"
			if AdminDisplayMode()="book" then
				%><!-- #include file="listword_admin.inc" --><%
			elseif AdminDisplayMode()="forum" then
				%><!-- #include file="listtitle_admin.inc" --><%
			end if
			rs.Close
		end if
		cn.Close : set rs=nothing : set cn=nothing%>

		<input type="hidden" name="page" value="<%=Request.QueryString("page")%>" />
		<!-- #include file="func_admin.inc" -->
	</form>

	<%if PagesCount>1 and ShowBottomPageList then show_page_list ipage,PagesCount,"admin.asp","[���Է�ҳ]","center",""%>
</div>

<!-- #include file="bottom.asp" -->
</body>
</html>