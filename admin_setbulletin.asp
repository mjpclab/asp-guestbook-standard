<!-- #include file="config.asp" -->
<!-- #include file="admin_verify.asp" -->

<%Response.Expires=-1%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=HomeName%> ���Ա� �����ö�����</title>
	<link rel="stylesheet" type="text/css" href="css/style.css"/>
	<link rel="stylesheet" type="text/css" href="css/adminstyle.css"/>
	<!-- #include file="css/style.asp" -->
	<!-- #include file="css/adminstyle.asp" -->

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

	<%if ShowTitle=true then show_book_title 3,"����"%>
	<!-- #include file="admincontrols.inc" -->

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

		<div class="region">
			<h3 class="title">�����ö�����</h3>
			<div class="content">
				<form method="post" action="admin_savebulletin.asp" name="form6" onsubmit="form6.submit1.disabled=true;">
				�������ݣ�<br/><textarea name="abulletin" id="abulletin" onkeydown="if(!this.modified)this.modified=true; var e=event?event:arguments[0]; if(e && e.ctrlKey && e.keyCode==13 && this.form.submit1)this.form.submit1.click();" rows="<%=ReplyTextHeight%>"><%=replace(server.htmlEncode("" & rs("declare") & ""),chr(13)&chr(10),"&#13;&#10;")%></textarea>
				<!-- #include file="ubbtoolbar.inc" -->
				<%ShowUbbToolBar(true)%>
				<p>
					<input type="checkbox" name="html2" id="html2" value="1"<%if cint(t_html and 1)<>0 then Response.Write " checked=""checked"""%> /><label for="html2">֧��HTML���</label><br/>
					<input type="checkbox" name="ubb2" id="ubb2" value="1"<%if cint(t_html and 2)<>0 then Response.Write " checked=""checked"""%> /><label for="ubb2">֧��UBB���</label><br/>
					<input type="checkbox" name="newline2" id="newline2" value="1"<%if cint(t_html and 4)<>0 then Response.Write " checked=""checked"""%> /><label for="newline2">��֧��HTML��UBB���ʱ����س�����</label>
				</p>
				<div class="command"><input value="��������" type="submit" name="submit1" id="submit1" /></div>
				</form>
			</div>
		</div>

	<%rs.Close : cn.Close : set rs=nothing : set cn=nothing%>

</div>

<!-- #include file="bottom.asp" -->
</body>
</html>