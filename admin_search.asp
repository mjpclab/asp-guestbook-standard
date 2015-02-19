<!-- #include file="config.asp" -->
<!-- #include file="admin_verify.asp" -->

<%
Response.Expires=-1
Response.AddHeader "cache-control","private"
%>

<!-- #include file="inc_dtd.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="zh-cn">
<head>
	<title><%=HomeName%> 留言本 搜索留言</title>

	<script type="text/javascript">
	//<![CDATA[
	
	function submitcheck()
	{
		if (form1.searchtxt.value=='') {alert('请输入搜索内容。');form1.searchtxt.focus();return false;}
		form1.searchsubmit.disabled=true;
		return true;
	}
	
	//]]>
	</script>
	
	<!-- #include file="style.asp" -->
</head>

<body<%=bodylimit%> onload="if(form1.searchtxt.value=='')form1.searchtxt.focus();<%=framecheck%>">

<%
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
CreateConn cn,dbtype

tparam=FilterAdminLike(Request("searchtxt"))
while left(tparam,1)="%" or left(tparam,1)="_"
	tparam=right(tparam,len(tparam)-1)
wend
while right(tparam,1)="%" or right(tparam,1)="_"
	tparam=left(tparam,len(tparam)-1)
wend
if request("type")<>"" and tparam="" then
	Call MessagePage("不能输入空字符串或全部为通配符。","admin_search.asp")
	Response.End
end if

dim sql_condition,sql_count,sql_full
if request("type")<>"" and tparam<>"" then CanOpenDB=true
if CanOpenDB=true then
	if request("type")="audit" then
		if tparam<>"0" and tparam<>"1" then tparam="1"
		sql_condition=sql_adminsearch_condition_audit & tparam
	elseif request("type")="reply" then
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

	<%if ShowTitle=true then show_book_title 3,"管理"%>
	<!-- #include file="admintool.inc" -->

	<table border="1" cellpadding="2" class="generalwindow">
		<tr>
			<td class="centertitle">搜索留言</td>
		</tr>
		<tr>
			<td class="wordscontent" style="text-align:center; padding:20px 0px;">
			<form method="post" action="admin_search.asp" id="form1" name="form1" onsubmit="return submitcheck();">
			搜索：<input type="text" name="searchtxt" size="<%=SearchTextWidth%>" value="<%=request("searchtxt")%>" />
			<input type="submit" value="搜索" name="searchsubmit" />
			<select name="type" size="1" onchange="searchtxt.focus();">
				<option value="name" <%=seled(request("type")="name")%>>按姓名搜索</option>
				<option value="title" <%=seled(request("type")="title")%>>按标题搜索</option>
				<option value="article" <%=seled(request("type")="article" or request("type")="")%>>按留言内容搜索</option>
				<option value="email" <%=seled(request("type")="email")%>>按邮件地址搜索</option>
				<option value="qqid" <%=seled(request("type")="qqid")%>>按QQ号码搜索</option>
				<option value="msnid" <%=seled(request("type")="msnid")%>>按MSN搜索</option>
				<option value="homepage" <%=seled(request("type")="homepage")%>>按主页地址搜索</option>
				<option value="ipaddr" <%=seled(request("type")="ipaddr")%>>按IP地址搜索</option>
				<option value="originalip" <%=seled(request("type")="originalip")%>>按原始IP地址搜索</option>
				<option value="reply" <%=seled(request("type")="reply")%>>按版主回复搜索</option>
				<option value="audit" <%=seled(request("type")="audit")%>>待审核(1:是 0:否)</option>
			</select><br/>("%"代表任意个字符，"_"代表一个字符)
			</form>
			</td>
		</tr>
	</table>
	
	<%if CanOpenDB and PagesCount>1 and ShowTopPageList then show_page_list ipage,PagesCount,"admin_search.asp","[搜索结果分页]","center","type=" &request("type")& "&searchtxt=" &server.URLEncode(request("searchtxt"))%>
	<form method="post" action="admin_mdel.asp" name="form7">
	<%RPage="admin_search.asp"%><!-- #include file="func_admin.inc" -->
	<%
	if CanOpenDB=true then
		if ItemsCount=0 then
			Response.Write "<br/><br/><div class=""centertext"">没有找到符合条件的留言。</div><br/><br/>"
		else
			dim pagename
			pagename="admin_search"
			if AdminDisplayMode()="book" then
				%><!-- #include file="listword_admin.inc" --><%
			elseif AdminDisplayMode()="forum" then
				%><!-- #include file="listtitle_admin.inc" --><%
			end if
			rs.Close
		end if
	end if
	%>

	<input type="hidden" name="page" value="<%=request("page")%>" />
	<input type="hidden" name="type" value="<%=request("type")%>" />
	<input type="hidden" name="searchtxt" value="<%=request("searchtxt")%>" />
	<!-- #include file="func_admin.inc" -->
	</form>

	<%if CanOpenDB and PagesCount>1 and ShowBottomPageList then show_page_list ipage,PagesCount,"admin_search.asp","[搜索结果分页]","center","type=" &request("type")& "&searchtxt=" &server.URLEncode(request("searchtxt"))%>
</div>

<%
if CanOpenDB=true then
	cn.Close : set rs=nothing : set cn=nothing
end if
%>

<!-- #include file="bottom.asp" -->
</body>
</html>