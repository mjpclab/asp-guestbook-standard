<!-- #include file="config.asp" -->
<!-- #include file="admin_verify.asp" -->

<%Response.Expires=-1%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=HomeName%> 留言本 发布置顶公告</title>
	<link rel="stylesheet" type="text/css" href="style.css"/>
	<!-- #include file="style.asp" -->

	<script type="text/javascript">
	function sfocus()
	{
		if (!form6.abulletin.modified)
		{
			form6.abulletin.focus();
			form6.abulletin.select();
		}
	}
	</script>
</head>

<body<%=bodylimit%> onload="sfocus();<%=framecheck%>">

<div id="outerborder" class="outerborder">

	<%if ShowTitle=true then show_book_title 3,"管理"%>
	<!-- #include file="admintool.inc" -->

	<%
	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")

	CreateConn cn,dbtype
	rs.Open sql_adminsetbulletin,cn,,,1

	dim t_html
	if rs("declare")<>"" then
		t_html=rs("declareflag")
	else
		t_html=adminlimit
	end if
	%>

	<table border="1" cellpadding="2" bordercolor="<%=TableBorderColor%>" class="generalwindow">
		<tr>
			<td class="centertitle">发布置顶公告</td>
		</tr>
		<tr>
			<td class="wordscontent" style="text-align:center; padding:20px 2px;">
			<form method="post" action="admin_savebulletin.asp" name="form6" onsubmit="form6.submit1.disabled=true;">
			公告内容：<br/><textarea name="abulletin" id="abulletin" onkeydown="if(!this.modified)this.modified=true; var e=event?event:arguments[0]; if(e && e.ctrlKey && e.keyCode==13 && this.form.submit1)this.form.submit1.click();" cols="<%=ReplyTextWidth%>" rows="<%=ReplyTextHeight%>"><%=replace(server.htmlEncode("" & rs("declare") & ""),chr(13)&chr(10),"&#13;&#10;")%></textarea>
			<!-- #include file="ubbtoolbar.inc" -->
			<%ShowUbbToolBar(true)%>
			<p style="text-align:left;">
			<input type="checkbox" name="html2" id="html2" value="1"<%if cint(t_html and 1)<>0 then Response.Write " checked=""checked"""%> /><label for="html2">支持HTML标记</label><br/>
			<input type="checkbox" name="ubb2" id="ubb2" value="1"<%if cint(t_html and 2)<>0 then Response.Write " checked=""checked"""%> /><label for="ubb2">支持UBB标记</label><br/>
			<input type="checkbox" name="newline2" id="newline2" value="1"<%if cint(t_html and 4)<>0 then Response.Write " checked=""checked"""%> /><label for="newline2">不支持HTML和UBB标记时允许回车换行</label>
			</p>
			<input value="更新数据" type="submit" name="submit1" id="submit1" />
			</form>
			</td>
		</tr>
	</table>

	<%rs.Close : cn.Close : set rs=nothing : set cn=nothing%>

</div>

<!-- #include file="bottom.asp" -->
</body>
</html>