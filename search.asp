<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/const.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/common2.asp" -->
<!-- #include file="include/sql/topbulletin.asp" -->
<!-- #include file="include/sql/search.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/ip.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" -->
<!-- #include file="include/utility/string.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/ubbcode.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/book.asp" -->
<!-- #include file="include/utility/message.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="error.asp" -->
<%
Response.Expires=-1
if checkIsBannedIP() then
	Call ErrorPage(1)
	Response.End
elseif Not StatusOpen then
	Call ErrorPage(2)
	Response.End
elseif Not StatusSearch then
	Call ErrorPage(5)
	Response.End
end if
if StatusStatistics then call addstat("search")

%>

<!-- #include file="include/template/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=HomeName%> 留言本 搜索结果</title>
	<!-- #include file="inc_stylesheet.asp" -->
</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<%
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
Call CreateConn(cn)

tparam=FilterGuestLike(Request("searchtxt"))

dim sql_condition,sql_count,sql_full
if (request("type")="reply" or request("type")="name" or request("type")="title" or request("type")="article") and tparam<>"" then CanOpenDB=true
if CanOpenDB then
	if request("type")="reply" then
		sql_condition=Replace(sql_search_condition_reply,"{0}",tparam)
	elseif request("type")="name" then
		sql_condition=Replace(sql_search_condition_name,"{0}",tparam)
	elseif request("type")="title" then
		sql_condition=Replace(sql_search_condition_title,"{0}",HtmlEncode(tparam))
	elseif request("type")="article" then
		sql_condition=Replace(sql_search_condition_article,"{0}",tparam)
	end if

	sql_condition=sql_condition & GetHiddenWordCondition()

	sql_count=sql_search_count_inner & sql_condition
	sql_count=Replace(sql_search_count,"{0}",sql_count)
	sql_full=sql_search_full_inner & sql_condition
	sql_full=Replace(sql_search_full,"{0}",sql_full)
	Dim ItemsCount,PagesCount,CurrentItemsCount,ipage
	get_divided_page cn,rs,sql_pksearch_main,sql_count,sql_full,"parent_id INC,lastupdated DEC,id DEC",Request("page"),ItemsPerPage,ItemsCount,PagesCount,CurrentItemsCount,ipage
end if
%>

<div id="outerborder" class="outerborder">
	<%if ShowTitle then%><%Call InitHeaderData("搜索结果")%><!-- #include file="include/template/header.inc" --><%end if%>
	<div id="mainborder" class="mainborder">
	<%RPage="search.asp"%><!-- #include file="include/template/guest_func.inc" -->
	<!-- #include file="include/template/topbulletin.inc" -->
	<%if CanOpenDB and PagesCount>1 and ShowTopPageList then show_page_list ipage,PagesCount,"search.asp","[搜索结果分页]","type,searchtxt"%>
	<%if ShowTopSearchBox then%><!-- #include file="include/template/guest_searchbox.inc" --><%end if%>

	<%
	if CanOpenDB then
		if ItemsCount=0 then
			Response.Write "<br/><br/><div class=""centertext"">没有找到符合条件的留言。</div><br/><br/>"
		else
			dim pagename, inAdminPage
			pagename="search"
			inAdminPage=false
			if GuestDisplayMode()="book" then
				%><!-- #include file="include/template/guest_listword.inc" --><%
			elseif GuestDisplayMode()="forum" then
				%><!-- #include file="include/template/guest_listtitle.inc" --><%
			end if
			rs.Close
		end if
	end if
	%>

	<!-- #include file="include/template/guest_func.inc" -->

	<%if CanOpenDB and PagesCount>1 and ShowBottomPageList then show_page_list ipage,PagesCount,"search.asp","[搜索结果分页]","type,searchtxt"%>
	<%if ShowBottomSearchBox then%><!-- #include file="include/template/guest_searchbox.inc" --><%end if%>
	</div>

	<!-- #include file="include/template/footer.inc" -->
</div>
<!-- #include file="include/template/getclientinfo.inc" -->
</body>
</html>
<%
if CanOpenDB then
	cn.Close
	set rs=nothing
	set cn=nothing
end if
%>