<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/common.asp" -->
<!-- #include file="include/sql/tlist.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/ip.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="loadconfig.asp" -->
<%
Response.Expires=-1
if checkIsBannedIP() then
	Response.End
elseif Not StatusOpen then
	Response.End
end if

function jsfilter(byref jsstr)
	jsfilter=replace(replace(replace(jsstr,"\","\\"),"'","\'"),"""","\""")
end function

dim cn,rs,t,max_len,str_title,n
n=Trim(Request.QueryString("n"))
if isnumeric(n) and n<>"" then
	n=abs(clng(Request.QueryString("n")))
	if n=0 then n=10
else
	n=10
end if
max_len=0
if isnumeric(Request.QueryString("len")) and Request.QueryString("len")<>"" then max_len=abs(clng(Request.QueryString("len")))

t=Request.QueryString("target")
pre=Request.QueryString("prefix")
mode=Request.QueryString("mode")

set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
Call CreateConn(cn)

rs.Open Replace(sql_tlist,"{0}",n),cn,0,1,1

if mode="json" then
	Dim firstDone
	firstDone=false
	Response.ContentType="application/json; Charset=gbk"
	Response.write "{""messages"":["
	do while not rs.EOF
		if firstDone then
			Response.Write ","
		else
			firstDone=true
		end if
		if max_len=0 then str_title=rs.Fields("title") else str_title=left(rs.Fields("title"),max_len) end if
		Response.write "{""title"":""" & jsfilter(str_title) & """,""fulltitle"":""" & jsfilter(rs.Fields("title")) & """,""url"":""" & geturlpath & "showword.asp?id=" & rs.Fields("id") & """}"

		rs.MoveNext
	loop
	Response.write "]"
	if Not IsEmpty(pre) then Response.Write ",""prefix"":""" & pre & """"
	if Not IsEmpty(t) then Response.Write ",""target"":""" & t & """"
	Response.Write ",""more"":""" & geturlpath & "index.asp" & """}"
elseif mode="js" then
	if t="" then t="_self"
	%>document.write('<ul><%do while not rs.EOF
		if max_len=0 then str_title=rs.Fields("title") else str_title=left(rs.Fields("title"),max_len) end if
		%><li><a href="<%=geturlpath & "showword.asp?id=" & rs.Fields("id")%>" target="<%=t%>" title="<%=jsfilter(rs.Fields("title"))%>"><%=jsfilter(pre & str_title)%></a></li><%

		rs.MoveNext
	loop
	%></ul><p><a title="更多留言..." href="<%=geturlpath & "index.asp"%>" target="<%=t%>">更多留言...</a></p>');
<%elseif mode="iframe" then%>
<%if t="" then t="_parent"%>
<!-- #include file="include/template/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=HomeName%> 留言本 留言列表</title>
	<!-- #include file="inc_stylesheet.asp" -->
</head>
<body class="tlist"<%=bodylimit%>>
	<ul>
	<%do while not rs.EOF
		if max_len=0 then str_title=rs.Fields("title") else str_title=left(rs.Fields("title"),max_len) end if%><li><a href="<%=geturlpath & "showword.asp?id=" & rs.Fields("id")%>" target="<%=t%>" title="<%=rs.Fields("title")%>"><%=pre & str_title%></a></li><%
		rs.MoveNext
	loop%>
	</ul>
	<p><a title="更多留言..." href="<%=geturlpath & "index.asp"%>" target="<%=t%>">更多留言...</a></p>
</body>
</html>
<%end if%>

<%rs.Close : cn.Close : set rs=nothing : set cn=nothing%>