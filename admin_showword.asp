<!-- #include file="config.asp" -->
<!-- #include file="admin_verify.asp" -->

<%
Response.Expires=-1
Response.AddHeader "cache-control","private"
%>

<!-- #include file="inc_dtd.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="zh-cn">
<head>
	<title><%=HomeName%> ���Ա� ������ҳ</title>
	
	<!-- #include file="style.asp" -->
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
CreateConn cn,dbtype
rs.Open sql_admin_showword & id,cn,0,1,1
if rs.EOF then		'���Բ����ڣ��˻�������
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

	<%if ShowTitle=true then show_book_title 3,"����"%>
	<!-- #include file="admintool.inc" -->
	<!-- #include file="topbulletin.inc" -->

	<form method="post" action="admin_mdel.asp" name="form7">
		<!-- #include file="func_admin.inc" -->
		<br/>
		<%
			dim pagename
			pagename="admin_showword"
			%><!-- #include file="listword_admin.inc" --><%
			rs.Close : cn.Close : set rs=nothing : set cn=nothing
		%>

		<input type="hidden" name="rootid" value="<%=request("id")%>" />
		<input type="hidden" name="page" value="<%=Request.QueryString("page")%>" />
		<input type="hidden" name="type" value="<%=Request.QueryString("type")%>" />
		<input type="hidden" name="searchtxt" value="<%=Request.QueryString("searchtxt")%>" />
		<!-- #include file="func_admin.inc" -->
	</form>

</div>

<!-- #include file="bottom.asp" -->
</body>
</html>