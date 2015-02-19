<!-- #include file="config.asp" -->
<!-- #include file="admin_verify.asp" -->

<%Response.Expires=-1%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=HomeName%> 留言本 IP屏蔽策略</title>
	
	<!-- #include file="style.asp" -->
</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<div id="outerborder" class="outerborder">

	<%if ShowTitle=true then show_book_title 3,"管理"%>
	<!-- #include file="admintool.inc" -->

	<%
	function hextoip(byref valip)
		dim vi,strip
		strip=""
		for vi=0 to 3
			strip=strip & cstr(cint("&H" & mid(valip,vi*2+1,2)))
			if vi<>3 then strip=strip & "."
		next
		hextoip=strip
	end function

	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")

	CreateConn cn,dbtype
	rs.Open sql_adminipconfig_status,cn,,,1
	tipconstatus=rs("ipconstatus")
	rs.Close
	%>

<style type="text/css">
	#ipconfigtable {width:440px; text-align:left; margin:0px auto; border:solid 1px <%=TableBorderColor%>; border-collapse:collapse;}
	#ipconfigtable td {border:solid 1px <%=TableBorderColor%>;}
</style>

<table border="1" cellpadding="2" class="generalwindow">
	<tr>
		<td class="centertitle">IP屏蔽策略</td>
	</tr>
	<tr>
		<td class="wordscontent" style="text-align:center; padding:20px 0px;">
		<form method="post" action="admin_saveipconfig.asp" name="ipconfigform" onsubmit="submit1.disabled=true;">
		<table border="1" id="ipconfigtable" cellspacing="0" cellpadding="2">
			<tr>
				<td colspan="2" style="vertical-align:middle"><br/><input type="radio" name="ipconstatus" value="0" id="r1"<%=cked(tipconstatus=0)%> /><label for="r1">不使用IP屏蔽策略</label><br/><br/></td>
			</tr>
			<tr>
				<td style="width:50%; vertical-align:top;">
					<br/><input type="radio" name="ipconstatus" value="1" id="r2"<%=cked(tipconstatus=1)%> /><label for="r2">只屏蔽以下IP段，其余放行</label><br/>
					<br/>添加新IP段,格式:"起始IP-终止IP"
					<br/><textarea name="txt1" rows="6" cols="32"></textarea><br/><br/>
					<%rs.Open sql_adminipconfig_status1,cn,,,1
					if rs.EOF=false then
						Response.Write "选择要删除的IP段：<br/>"
						while rs.EOF=false
							tlistid=rs("listid")
							tstartip=rs("startip")
							tendip=rs("endip")
							Response.Write "<input type=""checkbox"" name=""savediplist1"" value=""" &tlistid& """ id=""" &tlistid& """ /><label for=""" &tlistid& """>" & hextoip(tstartip) & "-" & hextoip(tendip) &"</label><br/>"
							rs.MoveNext
						wend
						Response.Write "<br/>"
					end if
					rs.Close
					%>
				</td>
				<td style="width:50%; vertical-align:top;">
					<br/><input type="radio" name="ipconstatus" value="2" id="r3"<%=cked(tipconstatus=2)%> /><label for="r3">只允许以下IP段，其余均不放行</label><br/>
					<br/>添加新IP段,格式:"起始IP-终止IP"
					<br/><textarea name="txt2" rows="6" cols="32"></textarea><br/><br/>
					<%rs.Open sql_adminipconfig_status2,cn,,,1
					if rs.EOF=false then
						Response.Write "选择要删除的IP段：<br/>"
						while rs.EOF=false
							tlistid=rs("listid")
							tstartip=rs("startip")
							tendip=rs("endip")
							Response.Write "<input type=""checkbox"" name=""savediplist2"" value=""" &tlistid& """ id=""" &tlistid& """ /><label for=""" &tlistid& """>" & hextoip(tstartip) & "-" & hextoip(tendip) &"</label><br/>"
							rs.MoveNext
						wend
						Response.Write "<br/>"
					end if
					rs.Close
					%>
				</td>
			</tr>
			<tr>
				<td colspan="2" style="width:100%; text-align:center;"><br/><input type="submit" name="submit1" value="更新数据" /><br/><br/></td>
			</tr>
		</table>
		</form>
		</td>
	</tr>
</table>
<%cn.Close : set rs=nothing : set cn=nothing%>
</div>

<!-- #include file="bottom.asp" -->
</body>
</html>