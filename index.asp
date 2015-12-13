<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/const.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/common2.asp" -->
<!-- #include file="include/sql/topbulletin.asp" -->
<!-- #include file="include/sql/index.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/ip.asp" -->
<!-- #include file="include/utility/ubbcode.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/book.asp" -->
<!-- #include file="include/utility/message.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="error.asp" -->
<%
Response.Expires=-1
Response.AddHeader "cache-control","private"

if checkIsBannedIP() then
	Call ErrorPage(1)
	Response.End
elseif StatusOpen=false then
	Call ErrorPage(2)
	Response.End
end if
if StatusStatistics then call addstat("view")
%>

<!-- #include file="include/template/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=HomeName%> 留言本</title>
	<!-- #include file="inc_stylesheet.asp" -->
</head>

<body<%=bodylimit%> onload="<%=framecheck%>">
<%
Dim cn,rs
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")

CreateConn cn,dbtype

Dim WordsPerPage
if GuestDisplayMode()="book" then
	WordsPerPage=ItemsPerPage
elseif GuestDisplayMode()="forum" then
	WordsPerPage=TitlesPerPage
end if

Dim ItemsCount,PagesCount,CurrentItemsCount,ipage

dim local_sql_count,local_sql_query
local_sql_count=sql_index_words_count & GetHiddenWordCondition()
local_sql_query=sql_index_words_query & GetHiddenWordCondition()
get_divided_page cn,rs,sql_pk_main,local_sql_count,local_sql_query,"parent_id INC,lastupdated DEC,id DEC",Request.QueryString("page"),WordsPerPage,ItemsCount,PagesCount,CurrentItemsCount,ipage
%>

<div id="outerborder" class="outerborder">
	<%if ShowTitle=true then show_book_title 2,""%>

	<%RPage="index.asp"%><!-- #include file="include/template/guest_func.inc" -->
	<!-- #include file="include/template/topbulletin.inc" -->
	<!-- #include file="include/template/guest_tiphidden.inc" -->
	<%if PagesCount>1 and ShowTopPageList then show_page_list ipage,PagesCount,"index.asp","[留言分页]",""%>
	<%if ItemsCount>0 and StatusSearch and ShowTopSearchBox then%><!-- #include file="include/template/guest_searchbox.inc" --><%end if%>

	<%
	if ItemsCount=0 then
		Response.Write "<br/><br/><div class=""centertext"">目前尚无留言，请点击“签写留言”。</div><br/><br/>"
	else
		dim pagename
		pagename="index"
		if GuestDisplayMode()="book" then
			%><!-- #include file="include/template/guest_listword.inc" --><%
		elseif GuestDisplayMode()="forum" then
			%><!-- #include file="include/template/guest_listtitle.inc" --><%
		end if
		rs.Close
	end if
	cn.Close : set rs=nothing : set cn=nothing%>
	
	<!-- #include file="include/template/guest_func.inc" -->

	<%if PagesCount>1 and ShowBottomPageList then show_page_list ipage,PagesCount,"index.asp","[留言分页]",""%>
	<%if ItemsCount>0 and StatusSearch and ShowBottomSearchBox then%><!-- #include file="include/template/guest_searchbox.inc" --><%end if%>
</div>

<!-- #include file="include/template/footer.inc" -->
<!-- #include file="include/template/getclientinfo.inc" -->
</body>
</html>