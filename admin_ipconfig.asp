<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/sql/admin_ipconfig.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/ip.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/book.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->
<%Response.Expires=-1%>

<!-- #include file="include/template/dtd.inc" -->
<html lang="zh-CN">
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=HomeName%> 留言本 IP屏蔽策略</title>
	<!-- #include file="inc_admin_stylesheet.asp" -->
</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<div id="outerborder" class="outerborder">

	<%if ShowTitle then%><%Call InitHeaderData("管理")%><!-- #include file="include/template/header.inc" --><%end if%>
	<div id="mainborder" class="mainborder">
	<!-- #include file="include/template/admin_mainmenu.inc" -->
	<%
	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")
	Call CreateConn(cn)
	%>


<div class="region form-region">
	<h3 class="title">IP屏蔽策略</h3>
	<div class="content">
		<form method="post" action="admin_saveipconfig.asp" name="ipconfigform" onsubmit="submit1.disabled=true;">
		<input type="hidden" name="tabIndex" id="tabIndex" value="<%=Request.QueryString("tabIndex")%>" />
		<div id="tabContainer">
			<div id="div-ipv4">
				<h4>IPv4</h4>
				<table cellpadding="10">
					<tr>
						<td colspan="2"><input type="radio" name="ipv4constatus" value="0" id="ipv4constatus-0"<%=cked(IPv4ConStatus=0)%> /><label for="ipv4constatus-0">不使用IP屏蔽策略</label></td>
					</tr>
					<tr>
						<td style="width:50%; vertical-align:top;">
							<p class="row"><input type="radio" name="ipv4constatus" value="1" id="ipv4constatus-1"<%=cked(IPv4ConStatus=1)%> /><label for="ipv4constatus-1">只屏蔽以下IP段，其余放行</label></p>
							<p class="row">添加新IP段,格式:"起始IP-终止IP"</p>
							<p class="row ipadd"><textarea name="newipv4status1" rows="6"></textarea></p>
							<p class="row">选择要删除的IP段：</p>
							<%rs.Open sql_adminipv4config_status1,cn,,,1
							if Not rs.EOF then
								while Not rs.EOF
									tlistid=rs("listid")
									tipfrom=rs("ipfrom")
									tipto=rs("ipto")%>
									<span class="row iplist">
									<input type="checkbox" name="savedipv4status1" value="<%=tlistid%>" id="ipv4-<%=tlistid%>" /><label for="ipv4-<%=tlistid%>"><%=hexToIPv4(tipfrom) & "-" & hexToIPv4(tipto)%></label>
									</span>
									<%rs.MoveNext
								wend
							end if
							rs.Close
							%>
						</td>
						<td style="width:50%; vertical-align:top;">
							<p class="row"><input type="radio" name="ipv4constatus" value="2" id="ipv4constatus-2"<%=cked(IPv4ConStatus=2)%> /><label for="ipv4constatus-2">只允许以下IP段，其余均不放行</label></p>
							<p class="row">添加新IP段,格式:"起始IP-终止IP"</p>
							<p class="row ipadd"><textarea name="newipv4status2" rows="6"></textarea></p>
							<p class="row">选择要删除的IP段：</p>
							<%rs.Open sql_adminipv4config_status2,cn,,,1
							if Not rs.EOF then
								while Not rs.EOF
									tlistid=rs("listid")
									tipfrom=rs("ipfrom")
									tipto=rs("ipto")%>
									<span class="row iplist">
									<input type="checkbox" name="savedipv4status2" value="<%=tlistid%>" id="ipv4-<%=tlistid%>" /><label for="ipv4-<%=tlistid%>"><%=hexToIPv4(tipfrom) & "-" & hexToIPv4(tipto)%></label>
									</span>
									<%rs.MoveNext
								wend
							end if
							rs.Close
							%>
						</td>
					</tr>
				</table>
			</div>
			<div id="div-ipv6">
				<h4>IPv6</h4>
				<table cellpadding="10">
					<tr>
						<td colspan="2"><input type="radio" name="ipv6constatus" value="0" id="ipv6constatus-0"<%=cked(IPv6ConStatus=0)%> /><label for="ipv6constatus-0">不使用IP屏蔽策略</label></td>
					</tr>
					<tr>
						<td style="width:50%; vertical-align:top;">
							<p class="row"><input type="radio" name="ipv6constatus" value="1" id="ipv6constatus-1"<%=cked(IPv6ConStatus=1)%> /><label for="ipv6constatus-1">只屏蔽以下IP段，其余放行</label></p>
							<p class="row">添加新IP段,格式:"起始IP-终止IP"</p>
							<p class="row ipadd"><textarea name="newipv6status1" rows="6"></textarea></p>
							<p class="row">选择要删除的IP段：</p>
							<%rs.Open sql_adminipv6config_status1,cn,,,1
							if Not rs.EOF then
								while Not rs.EOF
									tlistid=rs("listid")
									tipfrom=rs("ipfrom")
									tipto=rs("ipto")%>
									<span class="row iplist">
									<input type="checkbox" name="savedipv6status1" value="<%=tlistid%>" id="ipv6-<%=tlistid%>" /><label for="ipv6-<%=tlistid%>"><%=hexToIPv6(tipfrom) & "-" & hexToIPv6(tipto)%></label>
									</span>
									<%rs.MoveNext
								wend
							end if
							rs.Close
							%>
						</td>
						<td style="width:50%; vertical-align:top;">
							<p class="row"><input type="radio" name="ipv6constatus" value="2" id="ipv6constatus-2"<%=cked(IPv6ConStatus=2)%> /><label for="ipv6constatus-2">只允许以下IP段，其余均不放行</label></p>
							<p class="row">添加新IP段,格式:"起始IP-终止IP"</p>
							<p class="row ipadd"><textarea name="newipv6status2" rows="6"></textarea></p>
							<p class="row">选择要删除的IP段：</p>
							<%rs.Open sql_adminipv6config_status2,cn,,,1
							if Not rs.EOF then
								while Not rs.EOF
									tlistid=rs("listid")
									tipfrom=rs("ipfrom")
									tipto=rs("ipto")%>
									<span class="row iplist">
									<input type="checkbox" name="savedipv6status2" value="<%=tlistid%>" id="ipv6-<%=tlistid%>" /><label for="ipv6-<%=tlistid%>"><%=hexToIPv6(tipfrom) & "-" & hexToIPv6(tipto)%></label>
									</span>
									<%rs.MoveNext
								wend
							end if
							rs.Close
							%>
						</td>
					</tr>
				</table>
			</div>

			<script type="text/javascript" src="asset/js/tabcontrol.js"></script>
			<script type="text/javascript">
				var tab=new TabControl('tabContainer');

				tab.addPage('div-ipv4','IPv4');
				tab.addPage('div-ipv6','IPv6');

				tab.restoreFromField('tabIndex');
			</script>
		</div>
		<div class="command"><input type="submit" name="submit1" value="更新数据" /></div>
		</form>
	</div>
</div>

	</div>

	<!-- #include file="include/template/footer.inc" -->
</div>
</body>
</html>
<%cn.Close : set rs=nothing : set cn=nothing%>