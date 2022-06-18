<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/sql/admin_floodconfig.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/book.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->
<%Response.Expires=-1%>

<!-- #include file="include/template/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=HomeName%> 留言本 防灌水策略</title>
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
	rs.Open sql_adminfloodconfig,cn,,,1
	%>

	<div class="region form-region">
		<h3 class="title">防灌水策略</h3>
		<div class="content">
			<form method="post" action="admin_savefloodconfig.asp" name="configform" onsubmit="return check();">
			<p>同一用户最小发言时间间隔：<input type="text" name="minwait" size="10" maxlength="10" value="<%=flood_minwait%>" />秒 (0=不限)</p>

			<p>最新<input type="text" name="searchrange" size="10" maxlength="10" value="<%=flood_searchrange%>" />条(0=不限)
			<input type="checkbox" name="flag_newword" id="flag_newword" value="1"<%=cked(flood_sfnewword)%> /><label for="flag_newword">新留言</label>
			<input type="checkbox" name="flag_newreply" id="flag_newreply" value="1"<%=cked(flood_sfnewreply)%> /><label for="flag_newreply">访客回复</label>
			<br/>不允许
			<input type="radio" name="flag_include_equal" id="flag_include" value="1"<%=cked(flood_include)%> /><label for="flag_include">含有</label>
			<input type="radio" name="flag_include_equal" id="flag_equal" value="2"<%=cked(flood_equal)%> /><label for="flag_equal">具有</label>
			<br/>相同的
			<input type="checkbox" name="flag_title" id="flag_title" value="1"<%=cked(flood_sititle)%> /><label for="flag_title">标题</label>
			<input type="checkbox" name="flag_content" id="flag_content" value="1"<%=cked(flood_sicontent)%> /><label for="flag_content">内容</label>
			</p>

			<div class="command"><input value="更新数据" type="submit" name="submit1" /></div>
			</form>
		</div>
	</div>
	</div>

	<!-- #include file="include/template/footer.inc" -->
</div>
<script type="text/javascript" defer="defer">
function check()
{
	document.configform.submit1.disabled=true;
	return true;
}
</script>
</body>
</html>
<%rs.Close : cn.Close : set rs=nothing : set cn=nothing%>