<!-- #include file="config.asp" -->
<!-- #include file="admin_verify.asp" -->

<%Response.Expires=-1%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=HomeName%> ���Ա� ����ˮ����</title>
	<!-- #include file="inc_admin_stylesheet.asp" -->
</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<div id="outerborder" class="outerborder">

	<%if ShowTitle=true then show_book_title 3,"����"%>
	<!-- #include file="admincontrols.inc" -->

	<%
	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")

	CreateConn cn,dbtype
	rs.Open sql_adminfloodconfig,cn,,,1
	%>

	<div class="region form-region">
		<h3 class="title">����ˮ����</h3>
		<div class="content">
			<form method="post" action="admin_savefloodconfig.asp" name="configform" onsubmit="return check();">
			<p>ͬһ�û���С����ʱ������<input type="text" name="minwait" size="10" maxlength="10" value="<%=flood_minwait%>" />�� (0=����)</p>

			<p>����<input type="text" name="searchrange" size="10" maxlength="10" value="<%=flood_searchrange%>" />��(0=����)
			<input type="checkbox" name="flag_newword" id="flag_newword" value="1"<%=cked(flood_sfnewword)%> /><label for="flag_newword">������</label>
			<input type="checkbox" name="flag_newreply" id="flag_newreply" value="1"<%=cked(flood_sfnewreply)%> /><label for="flag_newreply">�ÿͻظ�</label>
			<br/>������
			<input type="radio" name="flag_include_equal" id="flag_include" value="1"<%=cked(flood_include)%> /><label for="flag_include">����</label>
			<input type="radio" name="flag_include_equal" id="flag_equal" value="2"<%=cked(flood_equal)%> /><label for="flag_equal">����</label>
			<br/>��ͬ��
			<input type="checkbox" name="flag_title" id="flag_title" value="1"<%=cked(flood_sititle)%> /><label for="flag_title">����</label>
			<input type="checkbox" name="flag_content" id="flag_content" value="1"<%=cked(flood_sicontent)%> /><label for="flag_content">����</label>
			</p>

			<div class="command"><input value="��������" type="submit" name="submit1" /></div>
			</form>
		</div>
	</div>

<%rs.Close : cn.Close : set rs=nothing : set cn=nothing%>

</div>

<script type="text/javascript" defer="defer">
function check()
{
	document.configform.submit1.disabled=true;
	return true;
}
</script>

<!-- #include file="bottom.asp" -->
</body>
</html>