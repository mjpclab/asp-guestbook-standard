<!-- #include file="config.asp" -->
<!-- #include file="common2.asp" -->
<%
Response.Expires=-1
if isbanip(Request.ServerVariables("REMOTE_ADDR"))=true or isbanip(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))=true then
	Response.End
elseif StatusOpen=false then
	Response.End
end if

function jsfilter(byref jsstr)
	jsfilter=replace(replace(jsstr,"\","\\"),"'","\'")
end function

if isnumeric(Request.QueryString("n")) and trim(Request.QueryString("n"))<>"" then
	dim cn,rs,t,CountPerPage,page,i,max_time,min_time,max_len,str_title,n,output_counter
	n=abs(clng(Request.QueryString("n")))
	if n=0 then n=10
	max_len=0
	output_counter=0
	if isnumeric(Request.QueryString("len")) and Request.QueryString("len")<>"" then max_len=abs(clng(Request.QueryString("len")))
	
	t=Request.QueryString("target")
	pre=Request.QueryString("prefix")
	if Request.QueryString("nobr")="yes" then br="" else br="<br/>"
	if GuestDisplayMode()="book" then CountPerPage=ItemsPerPage else CountPerPage=TitlesPerPage
	i=0
	page=1
	
	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")
	CreateConn cn,dbtype
	
	rs.Open Replace(sql_tlist_maxtime,"{0}",GetHiddenWordCondition()),cn,0,1,1
	if not rs.EOF then max_time=rs.Fields(0) else max_time=now() end if
	rs.Close
	
	rs.Open Replace(Replace(sql_tlist_mintime,"{0}",n),"{1}",GetHiddenWordCondition()),cn,0,1,1
	if not rs.EOF then min_time=rs.Fields(0) else min_time=now() end if
	rs.Close
	
	rs.Open Replace(Replace(Replace(sql_tlist,"{0}",DateTimeStr(min_time)),"{1}",DateTimeStr(max_time)),"{2}",GetHiddenWordCondition()),cn,0,1,1
%>
	<%if lcase(trim(Request.QueryString("js")))="yes" then%>
		<%if t="" then t="_self"%>
		<%do while not rs.EOF
			i=i+1
			if clng(rs.Fields("guestflag") and 48)=0 then%><%output_counter=output_counter+1%><%if max_len=0 then str_title=rs.Fields("title") else str_title=left(rs.Fields("title"),max_len) end if%>document.write('<a href="<%=geturlpath & "index.asp?page=" & page & "#a" & (((i-1) mod CountPerPage)+1)%>" target="<%=t%>" title="<%=jsfilter(rs.Fields("title"))%>"><%=jsfilter(pre & str_title)%></a><%=br%>');<%end if
			if i mod CountPerPage=0 then
				page=page+1
			end if
			if output_counter=n then exit do
			rs.MoveNext
		loop%>
		document.write('<%=br%><a style="margin-top:2ex;" title="更多留言..." href="<%=geturlpath & "index.asp"%>" target="<%=t%>">更多留言...</a>');
	<%else%>
	<%if t="" then t="_parent"%>
	<!-- #include file="inc_dtd.asp" -->
	<html>
	<head>
		<!-- #include file="inc_metatag.asp" -->
		<title><%=HomeName%> 留言本 留言列表</title>
		
		<!-- #include file="style.asp" -->
	</head>
	
	<body<%=bodylimit%> style="text-align:left;overflow:hidden;">
		<nobr/>
			<%do while not rs.EOF
				i=i+1
				if clng(rs.Fields("guestflag") and 48)=0 then%><%output_counter=output_counter+1%><%if max_len=0 then str_title=rs.Fields("title") else str_title=left(rs.Fields("title"),max_len) end if%><a href="<%=geturlpath & "index.asp?page=" & page & "#a" & (((i-1) mod CountPerPage)+1)%>" target="<%=t%>" title="<%=rs.Fields("title")%>"><%=pre & str_title%></a><%=br%><%end if
				if i mod CountPerPage=0 then
					page=page+1
				end if
				if output_counter=n then exit do
				rs.MoveNext
			loop%>
			<%=br%><a style="margin-top:2ex;" title="更多留言..." href="<%=geturlpath & "index.asp"%>" target="<%=t%>">更多留言...</a></span>
		</nobr>
	</body>
	</html>
	<%end if%>
	
	<%rs.Close : cn.Close : set rs=nothing : set cn=nothing%>
<%end if%>