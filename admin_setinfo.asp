<!-- #include file="config.asp" -->
<!-- #include file="admin_verify.asp" -->

<%Response.Expires=-1%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=HomeName%> 留言本 修改版主资料</title>
	<!-- #include file="inc_admin_stylesheet.asp" -->
</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<div id="outerborder" class="outerborder">

	<%if ShowTitle=true then show_book_title 3,"管理"%>
	<!-- #include file="admincontrols.inc" -->

	<%
	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")

	CreateConn cn,dbtype
	rs.Open sql_adminsetinfo,cn,,,1
	tfaceid=rs("faceid")
	%>

	<div class="region form-region region-longtext">
		<h3 class="title">修改版主资料</h3>
		<div class="content">
			<form method="post" action="admin_saveinfo.asp" name="form1" onsubmit="form1.submit1.disabled=true;">
			<div class="field">
				<span class="label">昵称：</span>
				<span class="value"><input type="text" name="aname" maxlength="20" value="<%="" & rs("name") & ""%>" /></span>
			</div>
			<div class="field">
				<span class="label">邮件：</span>
				<span class="value"><input type="text" name="aemail" maxlength="50" value="<%="" & rs("email") & ""%>" /></span>
			</div>
			<div class="field">
				<span class="label">QQ号：</span>
				<span class="value"><input type="text" name="aqqid" maxlength="16" value="<%="" & rs("qqid") & ""%>" /></span>
			</div>
			<div class="field">
				<span class="label">Skype：</span>
				<span class="value"><input type="text" name="amsnid" maxlength="50" value="<%="" & rs("msnid") & ""%>" /></span>
			</div>
			<div class="field">
				<span class="label">主页：</span>
				<span class="value"><input type="text" name="ahomepage" maxlength="127" value="<%="" & rs("homepage") & ""%>" /></span>
			</div>
			<div class="field">
				<span class="label">头像编号：</span>
				<span class="value"><input type="text" name="afaceid" maxlength="3" value="<%=tfaceid%>" title="填写头像编号时URL必须清空" /></span>
			</div>
			<div class="field">
				<span class="label">或URL：</span>
				<span class="value"><input type="text" name="afaceurl" maxlength="127" value="<%="" & rs("faceurl") & ""%>" title="填写URL时忽略头像编号" /></span>
			</div>
			<div class="field">
				<%rs.Close : cn.Close : set rs=nothing : set cn=nothing
				dim listfacecount,defaultindex
				listfacecount=FrequentFaceCount
				defaultindex=tfaceid%>
				<!-- #include file="listface.inc" -->
			</div>
			<div class="command"><input value="更新数据" type="submit" name="submit1" /></div>
			</form>
		</div>
	</div>

</div>

<!-- #include file="bottom.asp" -->
</body>
</html>