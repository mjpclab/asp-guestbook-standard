<!-- #include file="include/template/page_instruction.inc" -->
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
<html>
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=HomeName%> ���Ա� IP���β���</title>
	<!-- #include file="inc_admin_stylesheet.asp" -->
</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<div id="outerborder" class="outerborder">

	<%if ShowTitle=true then show_book_title 3,"����"%>
	<!-- #include file="include/template/admin_mainmenu.inc" -->

	<%
	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")
	Call CreateConn(cn)
	%>


<div class="region form-region">
	<h3 class="title">IP���β���</h3>
	<div class="content">
		<form method="post" action="admin_saveipconfig.asp" name="ipconfigform" onsubmit="submit1.disabled=true;">
		<input type="hidden" name="tabIndex" id="tabIndex" value="<%=Request.QueryString("tabIndex")%>" />
		<div id="tabContainer">
			<div id="div-ipv4">
				<h4>IPv4</h4>
				<table cellpadding="10">
					<tr>
						<td colspan="2"><input type="radio" name="ipv4constatus" value="0" id="ipv4constatus-0"<%=cked(IPv4ConStatus=0)%> /><label for="ipv4constatus-0">��ʹ��IP���β���</label></td>
					</tr>
					<tr>
						<td style="width:50%; vertical-align:top;">
							<p class="row"><input type="radio" name="ipv4constatus" value="1" id="ipv4constatus-1"<%=cked(IPv4ConStatus=1)%> /><label for="ipv4constatus-1">ֻ��������IP�Σ��������</label></p>
							<p class="row">�����IP��,��ʽ:"��ʼIP-��ֹIP"</p>
							<p class="row"><textarea name="newipv4status1" rows="6"></textarea></p>
							<p class="row">ѡ��Ҫɾ����IP�Σ�</p>
							<%rs.Open sql_adminipv4config_status1,cn,,,1
							if rs.EOF=false then
								while rs.EOF=false
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
							<p class="row"><input type="radio" name="ipv4constatus" value="2" id="ipv4constatus-2"<%=cked(IPv4ConStatus=2)%> /><label for="ipv4constatus-2">ֻ��������IP�Σ������������</label></p>
							<p class="row">�����IP��,��ʽ:"��ʼIP-��ֹIP"</p>
							<p class="row"><textarea name="newipv4status2" rows="6"></textarea></p>
							<p class="row">ѡ��Ҫɾ����IP�Σ�</p>
							<%rs.Open sql_adminipv4config_status2,cn,,,1
							if rs.EOF=false then
								while rs.EOF=false
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
						<td colspan="2"><input type="radio" name="ipv6constatus" value="0" id="ipv6constatus-0"<%=cked(IPv6ConStatus=0)%> /><label for="ipv6constatus-0">��ʹ��IP���β���</label></td>
					</tr>
					<tr>
						<td style="width:50%; vertical-align:top;">
							<p class="row"><input type="radio" name="ipv6constatus" value="1" id="ipv6constatus-1"<%=cked(IPv6ConStatus=1)%> /><label for="ipv6constatus-1">ֻ��������IP�Σ��������</label></p>
							<p class="row">�����IP��,��ʽ:"��ʼIP-��ֹIP"</p>
							<p class="row"><textarea name="newipv6status1" rows="6"></textarea></p>
							<p class="row">ѡ��Ҫɾ����IP�Σ�</p>
							<%rs.Open sql_adminipv6config_status1,cn,,,1
							if rs.EOF=false then
								while rs.EOF=false
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
							<p class="row"><input type="radio" name="ipv6constatus" value="2" id="ipv6constatus-2"<%=cked(IPv6ConStatus=2)%> /><label for="ipv6constatus-2">ֻ��������IP�Σ������������</label></p>
							<p class="row">�����IP��,��ʽ:"��ʼIP-��ֹIP"</p>
							<p class="row"><textarea name="newipv6status2" rows="6"></textarea></p>
							<p class="row">ѡ��Ҫɾ����IP�Σ�</p>
							<%rs.Open sql_adminipv6config_status2,cn,,,1
							if rs.EOF=false then
								while rs.EOF=false
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
				tab=new TabControl('tabContainer');
				tab.savingFieldId='tabIndex';
				var prevIndex=tab.loadPageIndex();

				tab.setOuterContainerCssClass('tab-outer-container');
				tab.setTitleContainerCssClass('tab-title-container');
				tab.setTitleCssClass('tab-title');
				tab.setTitleSelectedCssClass('tab-title-selected');
				tab.setPageContainerCssClass('tab-page-container');
				tab.setPageCssClass('tab-page');

				tab.addPage('div-ipv4','IPv4');
				tab.addPage('div-ipv6','IPv6');
				isFinite(prevIndex) && tab.selectPage(prevIndex);
			</script>
		</div>
		<div class="command"><input type="submit" name="submit1" value="��������" /></div>
		</form>
	</div>
</div>

<%cn.Close : set rs=nothing : set cn=nothing%>
</div>

<!-- #include file="include/template/footer.inc" -->
</body>
</html>