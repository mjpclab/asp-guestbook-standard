<!-- #include file="config.asp" -->
<!-- #include file="common2.asp" -->
<%
Response.Expires=-1
Response.AddHeader "cache-control","private"

if checkIsBannedIP then
	Response.Redirect "err.asp?number=1"
	Response.End
elseif StatusOpen=false then
	Response.Redirect "err.asp?number=2"
	Response.End
end if
if StatusStatistics then call addstat("view")
%>

<!-- #include file="include/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/metatag.inc" -->
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

	<%RPage="index.asp"%><!-- #include file="include/guest_func.inc" -->
	<!-- #include file="include/topbulletin.inc" -->
	<!-- #include file="include/guest_tiphidden.inc" -->
	<%if PagesCount>1 and ShowTopPageList then show_page_list ipage,PagesCount,"index.asp","[留言分页]",""%>
	<%if ItemsCount>0 and StatusSearch and ShowTopSearchBox then%><!-- #include file="include/guest_searchbox.inc" --><%end if%>

	<%
	if ItemsCount=0 then
		Response.Write "<br/><br/><div class=""centertext"">目前尚无留言，请点击“签写留言”。</div><br/><br/>"
	else
		dim pagename
		pagename="index"
		if GuestDisplayMode()="book" then
			%><!-- #include file="include/guest_listword.inc" --><%
		elseif GuestDisplayMode()="forum" then
			%><!-- #include file="include/guest_listtitle.inc" --><%
		end if
		rs.Close
	end if
	cn.Close : set rs=nothing : set cn=nothing%>
	
	<!-- #include file="include/guest_func.inc" -->

	<%if PagesCount>1 and ShowBottomPageList then show_page_list ipage,PagesCount,"index.asp","[留言分页]",""%>
	<%if ItemsCount>0 and StatusSearch and ShowBottomSearchBox then%><!-- #include file="include/guest_searchbox.inc" --><%end if%>
</div>

<!-- #include file="include/footer.inc" -->
<%if StatusStatistics and session("gotclientinfo")<>true then%><script type="text/javascript" src="getclientinfo.asp" defer="defer" async="async"></script><%end if%>
</body>
</html>