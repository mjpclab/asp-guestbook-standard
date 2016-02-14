<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/const.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/common2.asp" -->
<!-- #include file="include/sql/topbulletin.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/sql/admin_showword.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/ip.asp" -->
<!-- #include file="include/utility/ubbcode.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/book.asp" -->
<!-- #include file="include/utility/message.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->
<%
Response.Expires=-1
Response.AddHeader "cache-control","private"
%>

<!-- #include file="include/template/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=HomeName%> 留言本 管理首页</title>
	<!-- #include file="inc_admin_stylesheet.asp" -->
	<script type="text/javascript" src="asset/js/jquery-1.x-min.js"></script>
</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<%
dim id,ipage
ipage=Request("page")
if isnumeric(Request.QueryString("id")) and Request.QueryString("id")<>"" then
	id=Request.QueryString("id")
else
	id=0
end if

set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
Call CreateConn(cn)
rs.Open sql_admin_showword & id,cn,0,1,1
if rs.EOF then		'留言不存在，退回主界面
	rs.Close : cn.Close : set rs=nothing : set cn=nothing
	if Request("type")<>"" and Request("searchtxt")<>"" then
		Response.Redirect "admin_search.asp?page=" & Request("page") & "&type=" & server.URLEncode(Request("type")) & "&searchtxt=" & server.URLEncode(Request("searchtxt"))
	else
		Response.Redirect "admin.asp?page=" & Request("page")
	end if
	Response.End
end if
%>

<div id="outerborder" class="outerborder">

	<%if ShowTitle then show_book_title 3,"管理"%>
	<div id="mainborder" class="mainborder">
	<!-- #include file="include/template/admin_mainmenu.inc" -->
	<!-- #include file="include/template/topbulletin.inc" -->

	<form method="post" action="admin_mdel.asp" name="form7">
		<!-- #include file="include/template/admin_func.inc" -->
		<%
			dim pagename
			pagename="admin_showword"
			%><!-- #include file="include/template/admin_listword.inc" --><%
			rs.Close : cn.Close : set rs=nothing : set cn=nothing
		%>

		<input type="hidden" name="rootid" value="<%=request("id")%>" />
		<input type="hidden" name="page" value="<%=Request.QueryString("page")%>" />
		<input type="hidden" name="type" value="<%=Request.QueryString("type")%>" />
		<input type="hidden" name="searchtxt" value="<%=Request.QueryString("searchtxt")%>" />
		<!-- #include file="include/template/admin_func.inc" -->
	</form>
	</div>

	<!-- #include file="include/template/footer.inc" -->
</div>
</body>
</html>