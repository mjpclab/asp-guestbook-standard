<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
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
	<title><%=HomeName%> ���Ա� �޸Ĺ���Ա����</title>
	<!-- #include file="inc_admin_stylesheet.asp" -->

	<script type="text/javascript">
	function checkpass(cobject)
	{
		if (cobject.ioldpass.value.length===0) {alert('������ԭ���롣'); cobject.ioldpass.focus(); return(false);}
		if (cobject.inewpass1.value.length===0) {alert('�����������롣'); cobject.inewpass1.focus(); return(false);}
		if (cobject.inewpass2.value.length===0) {alert('������ȷ�����롣'); cobject.inewpass2.focus(); return(false);}
		if (cobject.inewpass1.value!==cobject.inewpass2.value) {alert('��������ȷ�����벻ͬ�����������롣'); cobject.inewpass1.focus(); return(false);}
		cobject.submit1.disabled=true;
		return (true);
	}
	</script>
</head>

<body<%=bodylimit%> onload="form4.ioldpass.focus();<%=framecheck%>">

<div id="outerborder" class="outerborder">

	<%if ShowTitle then%><%Call InitHeaderData("����")%><!-- #include file="include/template/header.inc" --><%end if%>
	<div id="mainborder" class="mainborder">
	<!-- #include file="include/template/admin_mainmenu.inc" -->
	<div class="region form-region region-longtext">
		<h3 class="title">�޸�����</h3>
		<div class="content">
			<form method="post" action="admin_savepass.asp" onsubmit="return checkpass(this)" name="form4">
			<div class="field">
				<span class="label">ԭ���룺</span>
				<span class="value"><input type="password" name="ioldpass" maxlength="32" /></span>
			</div>
			<div class="field">
				<span class="label">�����룺</span>
				<span class="value"><input type="password" name="inewpass1" maxlength="32" /></span>
			</div>
			<div class="field">
				<span class="label">ȷ�����룺</span>
				<span class="value"><input type="password" name="inewpass2" maxlength="32" /></span>
			</div>
			<div class="command"><input value="��������" type="submit" name="submit1" /></div>
			</form>
		</div>
	</div>
	</div>

	<!-- #include file="include/template/footer.inc" -->
</div>
</body>
</html>