<!-- #include file="config.asp" -->
<!-- #include file="admin_verify.asp" -->

<%Response.Expires=-1%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=HomeName%> 留言本 访问统计</title>
	<link rel="stylesheet" type="text/css" href="css/style.css"/>
	<link rel="stylesheet" type="text/css" href="css/adminstyle.css"/>
	<!-- #include file="css/style.asp" -->
	<!-- #include file="css/adminstyle.asp" -->
</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<div id="outerborder" class="outerborder">

	<%if ShowTitle=true then show_book_title 3,"管理"%>
	<!-- #include file="admincontrols.inc" -->

<%
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")

CreateConn cn,dbtype
rs.open sql_adminstats_startdate,cn,0,3,1
if rs.EOF then
	cn.Execute Replace(sql_adminstats_insert,"{0}",now()),,1
else
	if isdate(rs(0))=false then
		rs(0)=now()
		rs.Update
	end if
end if
rs.Close

rs.Open sql_adminstats,cn,,,1

tstartdate=rs("startdate")
tview=rs("view")
tsearch=rs("search")
tleaveword=rs("leaveword")
twritten=rs("written")
tfiltered=rs("filtered")
tbanned=rs("banned")
tlogin=rs("login")
tloginfailed=rs("loginfailed")
tnow=now()

rs.Close

on error resume next
%>

<div class="region form-region">
	<h3 class="title">访问统计</h3>
	<div class="content">
		<h4>开始统计日期</h4>
		<blockquote>
			<p><%=Year(tstartdate) & "-" & Month(tstartdate) & "-" & Day(tstartdate)%></p>
		</blockquote>

		<div id="div_outer"></div>

		<div id="div_words">
			<h4>留言列表页面</h4>
			<blockquote>
				<table>
					<tr><td style="width:120px;">访问次数：</td><td><%=tview%></td></tr>
					<tr><td style="width:120px;">平均月访问次数：</td><td><%=formatnumber(tview/((datediff("m",tstartdate,tnow)+1)),2,true,false,false)%></td></tr>
					<tr><td style="width:120px;">平均周访问次数：</td><td><%=formatnumber(tview/((datediff("ww",tstartdate,tnow)+1)),2,true,false,false)%></td></tr>
					<tr><td style="width:120px;">平均日访问次数：</td><td><%=formatnumber(tview/((datediff("d",tstartdate,tnow)+1)),2,true,false,false)%></td></tr>
				</table>
			</blockquote>
		</div>

		<div id="div_search">
			<h4>搜索页面</h4>
			<blockquote>
				<table>
					<tr><td style="width:120px;">搜索次数：</td><td><%=tsearch%></td></tr>
					<tr><td style="width:120px;">平均月搜索次数：</td><td><%=formatnumber(tsearch/((datediff("m",tstartdate,tnow)+1)),2,true,false,false)%></td></tr>
					<tr><td style="width:120px;">平均周搜索次数：</td><td><%=formatnumber(tsearch/((datediff("ww",tstartdate,tnow)+1)),2,true,false,false)%></td></tr>
					<tr><td style="width:120px;">平均日搜索次数：</td><td><%=formatnumber(tsearch/((datediff("d",tstartdate,tnow)+1)),2,true,false,false)%></td></tr>
				</table>
			</blockquote>
		</div>

		<div id="div_leaveword">
			<h4>留言/回复页面</h4>
			<blockquote>
				<table>
					<tr><td style="width:120px;">访问次数：</td><td><%=tleaveword%></td></tr>
					<tr><td style="width:120px;">成功留言次数：</td><td><%=twritten%></td></tr>
					<tr><td style="width:120px;">放弃留言率：</td><td><%if tleaveword=0 then Response.Write "/" else Response.Write formatpercent((tleaveword-twritten)/tleaveword,2,true)%></td></tr>
					<tr><td style="width:120px;">留言被过滤次数：</td><td><%=tfiltered%></td></tr>
					<tr><td style="width:120px;">过滤率：</td><td><%if twritten+tfiltered+tbanned=0 then Response.Write "/" else Response.Write formatpercent(tfiltered/(twritten+tfiltered+tbanned),2,true)%></td></tr>
					<tr><td style="width:120px;">留言被拒绝次数：</td><td><%=tbanned%></td></tr>
					<tr><td style="width:120px;">拒绝率：</td><td><%if twritten+tfiltered+tbanned=0 then Response.Write "/" else Response.Write formatpercent(tbanned/(twritten+tfiltered+tbanned),2,true)%></td></tr>
				</table>
			</blockquote>
		</div>

		<div id="div_login">
			<h4>管理登录页面</h4>
			<blockquote>
				<table>
					<tr><td style="width:120px;">登录次数：</td><td><%=tlogin%></td></tr>
					<tr><td style="width:120px;">登录失败次数：</td><td><%=tloginfailed%></td></tr>
					<tr><td style="width:120px;">登录失败率：</td><td><%if tlogin=0 then Response.Write "/" else Response.Write formatpercent(tloginfailed/tlogin,2,true)%></td></tr>
				</table>
			</blockquote>
		</div>

		<%
		on error goto 0

		rs.Open sql_adminstats_client_count,cn,0,1,1
		if rs.EOF=false then tclientcount=rs.Fields(0) else tclientcount=0

		if tclientcount>0 then
			'客户端操作系统
			rs.Close
			rs.Open sql_adminstats_client_os,cn,,,1

			Response.Write "<div id=""div_os"">"
			Response.Write "<h4>客户端操作系统</h4>"
			Response.Write "<blockquote><table>"
			while rs.EOF=false
				Response.Write "<tr><td style=""width:180px;"">"
				Response.Write server.HTMLEncode(rs.Fields("os")) & "："
				Response.Write "</td><td>"
				Response.Write rs.Fields(1) & "(" & formatpercent(rs.Fields(1)/tclientcount,2,true) & ")"
				Response.Write "</td></tr>"
				rs.MoveNext
			wend
			Response.Write "</table></blockquote></div>"

			'客户端浏览器
			rs.Close
			rs.Open sql_adminstats_client_browser,cn,0,1,1

			Response.Write "<div id=""div_browser"">"
			Response.Write "<h4>客户端浏览器</h4>"
			Response.Write "<blockquote><table>"
			while rs.EOF=false
				Response.Write "<tr><td style=""width:120px;"">"
				Response.Write server.HTMLEncode(rs.Fields("browser")) & "："
				Response.Write "</td><td>"
				Response.Write rs.Fields(1) & "(" & formatpercent(rs.Fields(1)/tclientcount,2,true) & ")"
				Response.Write "</td></tr>"
				rs.MoveNext
			wend
			Response.Write "</table></blockquote></div>"

			'客户端屏幕分辨率
			rs.Close
			rs.Open sql_adminstats_client_screen,cn,0,1,1

			Response.Write "<div id=""div_screen"">"
			Response.Write "<h4>客户端屏幕分辨率</h4>"
			Response.Write "<blockquote><table>"
			while rs.EOF=false
				Response.Write "<tr><td style=""width:120px;"">"
				if rs.Fields("screenwh")<>"0*0" then
					Response.Write server.HTMLEncode(rs.Fields("screenwh")) & "："
				else
					Response.Write "未知："
				end if
				Response.Write "</td><td>"
				Response.Write rs.Fields(1) & "(" & formatpercent(rs.Fields(1)/tclientcount,2,true) & ")"
				Response.Write "</td></tr>"
				rs.MoveNext
			wend
			Response.Write "</table></blockquote></div>"

			'访问时段
			rs.Close
			rs.Open sql_adminstats_client_timesect,cn,0,1,1

			Response.Write "<div id=""div_timesect"">"
			Response.Write "<h4>访问时段</h4>"
			Response.Write "<blockquote><table>"
			while rs.EOF=false
				Response.Write "<tr><td style=""width:120px;"">"
				Response.Write server.HTMLEncode(rs.Fields(0) & ":00～" & rs.Fields(0) & ":59") & "："
				Response.Write "</td><td>"
				Response.Write rs.Fields(1) & "(" & formatpercent(rs.Fields(1)/tclientcount,2,true) & ")"
				Response.Write "</td></tr>"
				rs.MoveNext
			wend
			Response.Write "</table></blockquote></div>"

			'访问周期
			dim weeklist
			weeklist=array("星期日","星期一","星期二","星期三","星期四","星期五","星期六")

			rs.Close
			rs.Open sql_adminstats_client_week,cn,0,1,1

			Response.Write "<div id=""div_week"">"
			Response.Write "<h4>访问周期</h4>"
			Response.Write "<blockquote><table>"
			while rs.EOF=false
				Response.Write "<tr><td style=""width:120px;"">"
				'Response.Write server.HTMLEncode(weekdayname(rs.Fields(0),false,1))
				Response.Write server.HTMLEncode(weeklist(rs.Fields(0)-1)) & "："
				Response.Write "</td><td>"
				Response.Write rs.Fields(1) & "(" & formatpercent(rs.Fields(1)/tclientcount,2,true) & ")"
				Response.Write "</td></tr>"
				rs.MoveNext
			wend
			Response.Write "</table></blockquote></div>"

			'最近30天访问量
			rs.Close
			rs.Open sql_adminstats_client_30day,cn,0,1,1

			Response.Write "<div id=""div_30day"">"
			Response.Write "<h4>最近30天访问量</h4>"
			Response.Write "<blockquote><table>"
			while rs.EOF=false
				Response.Write "<tr><td style=""width:120px;"">"
				Response.Write server.HTMLEncode(rs.Fields("datesect")) & "："
				Response.Write "</td><td>"
				Response.Write rs.Fields(1)
				Response.Write "</td></tr>"
				rs.MoveNext
			wend
			Response.Write "</table></blockquote></div>"

			'访问来源
			rs.Close
			rs.Open sql_adminstats_client_source,cn,0,1,1

			Response.Write "<div id=""div_source"">"
			Response.Write "<h4>访问来源</h4>"
			Response.Write "<blockquote><table>"
			while rs.EOF=false
				Response.Write "<tr><td style=""width:150px;"">"
				Response.Write server.HTMLEncode(rs.Fields("sourceaddr")) & "："
				Response.Write "</td><td>"
				Response.Write rs.Fields(1) & "(" & formatpercent(rs.Fields(1)/tclientcount,2,true) & ")"
				Response.Write "</td></tr>"
				rs.MoveNext
			wend
			Response.Write "</table></blockquote></div>"

		end if

		rs.Close : cn.Close : set rs=nothing : set cn=nothing
		%>

	</div>
	<form method="post" action="admin_clearstats.asp" onsubmit="if (confirm('确实要将统计清零吗？')){submit1.disabled=true;return true;} else return false;">
		<div class="command"><input type="submit" value="统计清零" name="submit1" /></div>
	</form>
</div>
</div>

<script type="text/javascript" src="js/tabcontrol.js"></script>
<script type="text/javascript">
	tab=new TabControl('div_outer');

	tab.setOuterContainerCssClass('tab-outer-container');
	tab.setTitleContainerCssClass('tab-title-container');
	tab.setTitleCssClass('tab-title');
	tab.setTitleSelectedCssClass('tab-title-selected');
	tab.setPageContainerCssClass('tab-page-container');
	tab.setPageCssClass('tab-page');

	tab.addPage('div_words','留言列表页面');
	tab.addPage('div_search','搜索页面');
	tab.addPage('div_leaveword','留言/回复页面');
	tab.addPage('div_login','管理登录页面');
	tab.addPage('div_os','客户端操作系统');
	tab.addPage('div_browser','客户端浏览器');
	tab.addPage('div_screen','客户端屏幕分辨率');
	tab.addPage('div_timesect','访问时段');
	tab.addPage('div_week','访问周期');
	tab.addPage('div_30day','最近30天访问量');
	tab.addPage('div_source','访问来源');
</script>
	
<!-- #include file="bottom.asp" -->
</body>
</html>