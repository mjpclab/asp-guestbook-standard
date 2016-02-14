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
<!-- #include file="include/utility/message.asp" -->
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

dim cn,rs,t,CountPerPage,page,i,max_time,min_time,max_len,str_title,n,output_counter
n=Trim(Request.QueryString("n"))
if isnumeric(n) and n<>"" then
	n=abs(clng(Request.QueryString("n")))
	if n=0 then n=10
else
	n=10
end if
max_len=0
output_counter=0
if isnumeric(Request.QueryString("len")) and Request.QueryString("len")<>"" then max_len=abs(clng(Request.QueryString("len")))

t=Request.QueryString("target")
pre=Request.QueryString("prefix")
mode=Request.QueryString("mode")
if GuestDisplayMode()="book" then CountPerPage=ItemsPerPage else CountPerPage=TitlesPerPage
i=0
page=1

set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")
Call CreateConn(cn)

rs.Open Replace(sql_tlist_maxtime,"{0}",GetHiddenWordCondition()),cn,0,1,1
if not rs.EOF then max_time=rs.Fields(0) else max_time=now() end if
rs.Close

rs.Open Replace(Replace(sql_tlist_mintime,"{0}",n),"{1}",GetHiddenWordCondition()),cn,0,1,1
if not rs.EOF then min_time=rs.Fields(0) else min_time=now() end if
rs.Close

rs.Open Replace(Replace(Replace(sql_tlist,"{0}",DateTimeStr(min_time)),"{1}",DateTimeStr(max_time)),"{2}",GetHiddenWordCondition()),cn,0,1,1

if mode="json" then
	Response.ContentType="application/json; Charset=gbk"
	Response.write "{""messages"":["
	do while not rs.EOF
		if Not CBool(rs.Fields("guestflag") and 48) then
			output_counter=output_counter+1
			if output_counter>1 then Response.Write ","
			if max_len=0 then str_title=rs.Fields("title") else str_title=left(rs.Fields("title"),max_len) end if
			Response.write "{""title"":""" & jsfilter(str_title) & """,""fulltitle"":""" & jsfilter(rs.Fields("title")) & """,""url"":""" & geturlpath & "index.asp?page=" & page & "#a" & (((i-1) mod CountPerPage)+1) & """}"
		end if
		if i mod CountPerPage=0 then page=page+1
		if output_counter=n then exit do
		rs.MoveNext
	loop
	Response.write "]"
	if Not IsEmpty(pre) then Response.Write ",""prefix"":""" & pre & """"
	if Not IsEmpty(t) then Response.Write ",""target"":""" & t & """"
	Response.Write ",""more"":""" & geturlpath & "index.asp" & """}"
elseif mode="js" then
	if t="" then t="_self"
	%>document.write('<ul><%do while not rs.EOF
		i=i+1
		if Not CBool(rs.Fields("guestflag") and 48) then
			output_counter=output_counter+1
			if max_len=0 then str_title=rs.Fields("title") else str_title=left(rs.Fields("title"),max_len) end if
			%><li><a href="<%=geturlpath & "index.asp?page=" & page & "#a" & (((i-1) mod CountPerPage)+1)%>" target="<%=t%>" title="<%=jsfilter(rs.Fields("title"))%>"><%=jsfilter(pre & str_title)%></a></li><%
		end if
		if i mod CountPerPage=0 then page=page+1
		if output_counter=n then exit do
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
		i=i+1
		if Not CBool(rs.Fields("guestflag") and 48) then
			output_counter=output_counter+1
			if max_len=0 then str_title=rs.Fields("title") else str_title=left(rs.Fields("title"),max_len) end if%><li><a href="<%=geturlpath & "index.asp?page=" & page & "#a" & (((i-1) mod CountPerPage)+1)%>" target="<%=t%>" title="<%=rs.Fields("title")%>"><%=pre & str_title%></a></li>
		<%end if
		if i mod CountPerPage=0 then page=page+1
		if output_counter=n then exit do
		rs.MoveNext
	loop%>
	</ul>
	<p><a title="更多留言..." href="<%=geturlpath & "index.asp"%>" target="<%=t%>">更多留言...</a></p>
</body>
</html>
<%end if%>

<%rs.Close : cn.Close : set rs=nothing : set cn=nothing%>