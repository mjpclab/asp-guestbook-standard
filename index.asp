<!-- #include file="config.asp" -->
<%
Response.Expires=-1
Response.AddHeader "cache-control","private"

if isbanip(Request.ServerVariables("REMOTE_ADDR"))=true or isbanip(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))=true then
	Response.Redirect "err.asp?number=1"
	Response.End
elseif StatusOpen=false then
	Response.Redirect "err.asp?number=2"
	Response.End
end if
call addstat("view")
%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=HomeName%> 留言本</title>
	<link rel="stylesheet" type="text/css" href="style.css"/>
	<!-- #include file="style.asp" -->
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

	<%RPage="index.asp"%><!-- #include file="func_guest.inc" -->
	<!-- #include file="topbulletin.inc" -->
	<!-- #include file="hidetip.inc" -->
	<%if PagesCount>1 and ShowTopPageList then show_page_list ipage,PagesCount,"index.asp","[留言分页]","center",""%>
	<%if ItemsCount>0 and StatusSearch and ShowTopSearchBox then%><!-- #include file="searchbox_guest.inc" --><%end if%>

	<%
	if ItemsCount=0 then
		Response.Write "<br/><br/><div class=""centertext"">目前尚无留言，请点击“签写留言”。</div><br/><br/>"
	else
		dim pagename
		pagename="index"
		if GuestDisplayMode()="book" then
			%><!-- #include file="listword_guest.inc" --><%
		elseif GuestDisplayMode()="forum" then
			%><!-- #include file="listtitle_guest.inc" --><%
		end if
		rs.Close
	end if
	cn.Close : set rs=nothing : set cn=nothing%>
	
	<!-- #include file="func_guest.inc" -->

	<%if PagesCount>1 and ShowBottomPageList then show_page_list ipage,PagesCount,"index.asp","[留言分页]","center",""%>
	<%if ItemsCount>0 and StatusSearch and ShowBottomSearchBox then%><!-- #include file="searchbox_guest.inc" --><%end if%>
</div>

<!-- #include file="bottom.asp" -->
<script type="text/javascript" src="getclientinfo.asp" defer="defer" async="async"></script>
</body>
</html>