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
	<title><%=HomeName%> ���Ա� ��������</title>
	<!-- #include file="inc_admin_stylesheet.asp" -->

	<script type="text/javascript" src="asset/js/jquery-1.x-min.js"></script>
	<script type="text/javascript">
	function submitcheck()
	{
		if (form1.searchtxt.value.length===0) {
			alert('�������������ݡ�');
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
	Call TipsPage("����������ַ�����ȫ��Ϊͨ�����","admin_search.asp")
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

	<%if ShowTitle then%><%Call InitHeaderData("����")%><!-- #include file="include/template/header.inc" --><%end if%>
	<div id="mainborder" class="mainborder">
	<!-- #include file="include/template/admin_mainmenu.inc" -->
	<div class="region form-region center-region">
		<h3 class="title">��������</h3>
		<div class="content">
			<form method="post" action="admin_search.asp" id="form1" name="form1" onsubmit="return submitcheck();">
			������<input type="text" name="searchtxt" size="<%=SearchTextWidth%>" value="<%=request("searchtxt")%>" />
			<input type="submit" value="����" name="searchsubmit" />
			<select name="type" size="1" onchange="searchtxt.focus();">
				<option value="name" <%=seled(Request("type")="name")%>>����������</option>
				<option value="title" <%=seled(Request("type")="title")%>>����������</option>
				<option value="article" <%=seled(Request("type")="article" or Request("type")="")%>>��������������</option>
				<option value="email" <%=seled(Request("type")="email")%>>���ʼ���ַ����</option>
				<option value="qqid" <%=seled(Request("type")="qqid")%>>��QQ��������</option>
				<option value="msnid" <%=seled(Request("type")="msnid")%>>��MSN����</option>
				<option value="homepage" <%=seled(Request("type")="homepage")%>>����ҳ��ַ����</option>
				<option value="ipv4addr" <%=seled(Request("type")="ipv4addr")%>>��IPv4��ַ����</option>
				<option value="originalipv4" <%=seled(Request("type")="originalipv4")%>>��ԭʼIPv4��ַ����</option>
				<option value="ipv6addr" <%=seled(Request("type")="ipv6addr")%>>��IPv6��ַ����</option>
				<option value="originalipv6" <%=seled(Request("type")="originalipv6")%>>��ԭʼIPv6��ַ����</option>
				<option value="reply" <%=seled(Request("type")="reply")%>>�������ظ�����</option>
				<option value="audit" <%=seled(Request("type")="audit")%>>�����(1:�� 0:��)</option>
			</select><br/>("%"����������ַ���"_"����һ���ַ�)
			</form>
		</div>
	</div>

	<%if CanOpenDB and PagesCount>1 and ShowTopPageList then show_page_list ipage,PagesCount,"admin_search.asp","[���������ҳ]","type,searchtxt"%>
	<form method="post" action="admin_mdel.asp" name="form7">
	<%RPage="admin_search.asp"%><!-- #include file="include/template/admin_func.inc" -->
	<%
	if CanOpenDB then
		if ItemsCount=0 then
			Response.Write "<br/><br/><div class=""centertext"">û���ҵ��������������ԡ�</div><br/><br/>"
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

	<%if CanOpenDB and PagesCount>1 and ShowBottomPageList then show_page_list ipage,PagesCount,"admin_search.asp","[���������ҳ]","type,searchtxt"%>
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
