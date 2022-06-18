<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/const.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/common2.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/sql/admin_search.asp" -->
<!-- #include file="include/sql/admin_listword.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/ip.asp" -->
<!-- #include file="include/utility/sqlfilter.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/string.asp" -->
<!-- #include file="include/utility/ubbcode.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/book.asp" -->
<!-- #include file="include/utility/message.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->
<!-- #include file="tips.asp" -->
<%
Response.Expires=-1
Response.AddHeader "cache-control","private"
%>

<!-- #include file="include/template/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=HomeName%> 留言本 搜索留言</title>
	<!-- #include file="inc_admin_stylesheet.asp" -->

	<script type="text/javascript" src="asset/js/jquery-1.x-min.js"></script>
	<script type="text/javascript">
	function submitcheck()
	{
		if (form1.searchtxt.value.length===0) {
			alert('请输入搜索内容。');
			form1.searchtxt.focus();
			return false;
		}
		form1.searchsubmit.disabled=true;
		return true;
	}
	</script>
</head>

<body<%=bodylimit%> onload="if(form1.searchtxt.value.length===0)form1.searchtxt.focus();<%=framecheck%>">

<%
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
Call CreateConn(cn)

tparam=FilterAdminLike(Request("searchtxt"))
while left(tparam,1)="%" or left(tparam,1)="_"
	tparam=right(tparam,len(tparam)-1)
wend
while right(tparam,1)="%" or right(tparam,1)="_"
	tparam=left(tparam,len(tparam)-1)
wend
if Request("type")<>"" and tparam="" then
	Call TipsPage("不能输入空字符串或全部为通配符。","admin_search.asp")
	Response.End
end if

dim sql_condition,sql_count,sql_full
if Request("type")<>"" and tparam<>"" then CanOpenDB=true
if CanOpenDB then
	if Request("type")="audit" then
		if tparam<>"0" and tparam<>"1" then tparam="1"
		sql_condition=sql_adminsearch_condition_audit & tparam
	elseif Request("type")="reply" then
		sql_condition=Replace(sql_adminsearch_condition_reply,"{0}",tparam)
	else
		sql_condition=Replace(Replace(sql_adminsearch_condition_else,"{0}",FilterSql(Request("type"))),"{1}",tparam)
	end if

	sql_count=sql_adminsearch_count_inner & sql_condition
	sql_count=Replace(sql_adminsearch_count,"{0}",sql_count)
	sql_full=sql_adminsearch_full_inner & sql_condition
	sql_full=Replace(sql_adminsearch_full,"{0}",sql_full)
	Dim ItemsCount,PagesCount,CurrentItemsCount,ipage
	get_divided_page cn,rs,sql_pksearch_main,sql_count,sql_full,"parent_id INC,lastupdated DEC,id DEC",Request("page"),ItemsPerPage,ItemsCount,PagesCount,CurrentItemsCount,ipage
end if
%>

<div id="outerborder" class="outerborder">

	<%if ShowTitle then%><%Call InitHeaderData("管理")%><!-- #include file="include/template/header.inc" --><%end if%>
	<div id="mainborder" class="mainborder">
	<!-- #include file="include/template/admin_mainmenu.inc" -->
	<div class="region form-region center-region">
		<h3 class="title">搜索留言</h3>
		<div class="content">
			<form method="post" action="admin_search.asp" id="form1" name="form1" onsubmit="return submitcheck();">
			搜索：<input type="text" name="searchtxt" size="<%=SearchTextWidth%>" value="<%=request("searchtxt")%>" />
			<input type="submit" value="搜索" name="searchsubmit" />
			<select name="type" size="1" onchange="searchtxt.focus();">
				<option value="name" <%=seled(Request("type")="name")%>>按姓名搜索</option>
				<option value="title" <%=seled(Request("type")="title")%>>按标题搜索</option>
				<option value="article" <%=seled(Request("type")="article" or Request("type")="")%>>按留言内容搜索</option>
				<option value="email" <%=seled(Request("type")="email")%>>按邮件地址搜索</option>
				<option value="qqid" <%=seled(Request("type")="qqid")%>>按QQ号码搜索</option>
				<option value="msnid" <%=seled(Request("type")="msnid")%>>按MSN搜索</option>
				<option value="homepage" <%=seled(Request("type")="homepage")%>>按主页地址搜索</option>
				<option value="ipv4addr" <%=seled(Request("type")="ipv4addr")%>>按IPv4地址搜索</option>
				<option value="originalipv4" <%=seled(Request("type")="originalipv4")%>>按原始IPv4地址搜索</option>
				<option value="ipv6addr" <%=seled(Request("type")="ipv6addr")%>>按IPv6地址搜索</option>
				<option value="originalipv6" <%=seled(Request("type")="originalipv6")%>>按原始IPv6地址搜索</option>
				<option value="reply" <%=seled(Request("type")="reply")%>>按版主回复搜索</option>
				<option value="audit" <%=seled(Request("type")="audit")%>>待审核(1:是 0:否)</option>
			</select><br/>("%"代表任意个字符，"_"代表一个字符)
			</form>
		</div>
	</div>

	<%if CanOpenDB and PagesCount>1 and ShowTopPageList then show_page_list ipage,PagesCount,"[搜索结果分页]","type,searchtxt"%>
	<form method="post" action="admin_mdel.asp" name="form7">
	<%RPage="admin_search.asp"%><!-- #include file="include/template/admin_func.inc" -->
	<%
	if CanOpenDB then
		if ItemsCount=0 then
			Response.Write "<br/><br/><div class=""centertext"">没有找到符合条件的留言。</div><br/><br/>"
		else
			dim pagename, inAdminPage
			pagename="admin_search"
			inAdminPage=true
			if AdminDisplayMode()="book" then
				%><!-- #include file="include/template/admin_listword.inc" --><%
			elseif AdminDisplayMode()="forum" then
				%><!-- #include file="include/template/admin_listtitle.inc" --><%
			end if
			rs.Close
		end if
	end if
	%>

	<input type="hidden" name="page" value="<%=request("page")%>" />
	<input type="hidden" name="type" value="<%=Request("type")%>" />
	<input type="hidden" name="searchtxt" value="<%=request("searchtxt")%>" />
	<!-- #include file="include/template/admin_func.inc" -->
	</form>

	<%if CanOpenDB and PagesCount>1 and ShowBottomPageList then show_page_list ipage,PagesCount,"[搜索结果分页]","type,searchtxt"%>
	</div>

	<!-- #include file="include/template/footer.inc" -->
</div>
</body>
</html>
<%
if CanOpenDB then
	cn.Close : set rs=nothing : set cn=nothing
end if
%>
