<!-- #include file="config.asp" -->
<!-- #include file="admin_verify.asp" -->

<%Response.Expires=-1%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=HomeName%> ���Ա� �޸Ĺ���Ա����</title>
	<link rel="stylesheet" type="text/css" href="style.css"/>
	<!-- #include file="style.asp" -->

	<script type="text/javascript">
	function checkpass(cobject)
	{
		if (cobject.ioldpass.value=="") {alert('������ԭ���롣'); cobject.ioldpass.focus(); return(false);}
		if (cobject.inewpass1.value=="") {alert('�����������롣'); cobject.inewpass1.focus(); return(false);}
		if (cobject.inewpass2.value=="") {alert('������ȷ�����롣'); cobject.inewpass2.focus(); return(false);}
		if (cobject.inewpass1.value!=cobject.inewpass2.value) {alert('��������ȷ�����벻ͬ�����������롣'); cobject.inewpass1.focus(); return(false);}
		cobject.submit1.disabled=true;
		return (true);
	}
	</script>
</head>

<body<%=bodylimit%> onload="form4.ioldpass.focus();<%=framecheck%>">

<div id="outerborder" class="outerborder">

	<%show_book_title 3,"����"%>
	<!-- #include file="admintool.inc" -->

	<table border="1" bordercolor="<%=TableBorderColor%>" cellpadding="2" class="generalwindow">
		<tr>
			<td class="centertitle">�޸�����</td>
		</tr>
		<tr>
			<td class="wordscontent" style="text-align:center; padding:20px 0px;">
			<form method="post" action="admin_savepass.asp" onsubmit="return checkpass(this)" name="form4">
			��ԭ���룺<input type="password" name="ioldpass" size="<%=SetInfoTextWidth%>" maxlength="32" /><br/>
			�������룺<input type="password" name="inewpass1" size="<%=SetInfoTextWidth%>" maxlength="32" /><br/>
			ȷ�����룺<input type="password" name="inewpass2" size="<%=SetInfoTextWidth%>" maxlength="32" /><br/><br/>
			<input value="��������" type="submit" name="submit1" />
			</form>
			</td>
		</tr>
	</table>
</div>

<!-- #include file="bottom.asp" -->
</body>
</html>