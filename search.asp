<!-- #include file="config.asp" -->

<%
Response.Expires=-1
if isbanip(Request.ServerVariables("REMOTE_ADDR"))=true or isbanip(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))=true then
	Response.Redirect "err.asp?number=1"
	Response.End
elseif StatusOpen=false then
	Response.Redirect "err.asp?number=2"
	Response.End
elseif StatusSearch=false then
	Response.Redirect "err.asp?number=5"
	Response.End
end if
call addstat("search")

%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=HomeName%> 留言本 搜索结果</title>
	<link rel="stylesheet" type="text/css" href="style.css"/>
	<!-- #include file="style.asp" -->
	<!-- #include file="getclientinfo.inc" -->
</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<%
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype

tparam=FilterGuestLike(Request("searchtxt"))

dim sql_condition,sql_count,sql_full
if (request("type")="reply" or request("type")="name" or request("type")="title" or request("type")="article") and tparam<>"" then CanOpenDB=true
if CanOpenDB=true then
	if request("type")="reply" then
		sql_condition=Replace(sql_search_condition_reply,"{0}",tparam)
	elseif request("type")="name" then
		sql_condition=Replace(sql_search_condition_name,"{0}",tparam)
	elseif request("type")="title" then
		sql_condition=Replace(sql_search_condition_title,"{0}",Server.HTMLEncode(tparam))
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
	<%if ShowTitle=true then show_book_title 3,"搜索结果"%>

	<%RPage="search.asp"%><!-- #include file="func_guest.inc" -->
	<!-- #include file="topbulletin.inc" -->
	<%if CanOpenDB=true and PagesCount>1 and ShowTopPageList then show_page_list ipage,PagesCount,"search.asp","[搜索结果分页]","center","type=" &request("type")& "&searchtxt=" &server.URLEncode(request("searchtxt"))%>
	<%if ShowTopSearchBox then%><!-- #include file="searchbox_guest.inc" --><%end if%>

	<%
	if CanOpenDB=true then
		if ItemsCount=0 then
			Response.Write "<br/><br/><div class=""centertext"">没有找到符合条件的留言。</div><br/><br/>"
		else
			dim pagename
			pagename="search"
			if GuestDisplayMode()="book" then
				%><!-- #include file="listword_guest.inc" --><%
			elseif GuestDisplayMode()="forum" then
				%><!-- #include file="listtitle_guest.inc" --><%
			end if
			rs.Close
		end if
	end if
	%>

	<!-- #include file="func_guest.inc" -->

	<%if CanOpenDB=true and PagesCount>1 and ShowBottomPageList then show_page_list ipage,PagesCount,"search.asp","[搜索结果分页]","center","type=" &request("type")& "&searchtxt=" &server.URLEncode(request("searchtxt"))%>
	<%if ShowBottomSearchBox then%><!-- #include file="searchbox_guest.inc" --><%end if%>
</div>

<%
if CanOpenDB=true then
	cn.Close
	set rs=nothing
	set cn=nothing
end if
%>

<!-- #include file="bottom.asp" -->
</body>
</html>