<!-- #include file="config.asp" -->
<!-- #include file="admin_verify.asp" -->

<%Response.Expires=-1%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=HomeName%> ���Ա� �߼�ɾ��</title>
	<link rel="stylesheet" type="text/css" href="style.css"/>
	<link rel="stylesheet" type="text/css" href="adminstyle.css"/>
	<!-- #include file="style.asp" -->
	<!-- #include file="adminstyle.asp" -->
</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<div id="outerborder" class="outerborder">

	<%if ShowTitle=true then show_book_title 3,"����"%>
	<!-- #include file="admincontrols.inc" -->

	<div class="region form-region">
		<h3 class="title">�߼�ɾ��</h3>
		<div class="content">
			<form method="post" action="admin_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				ɾ��ָ�����ں�ʱ��ǰ�����ԣ�����������/ʱ�䡣<br/>
				<input type="hidden" name="option" value="1" />
				<input type="text" name="iparam" size="<%=AdvDelTextWidth%>" maxlength="20" />
				<input type="submit" value="ִ��" name="submit1"<%if DelAdvTip=true then Response.Write " onclick=""return confirm('ȷʵҪִ��ɾ��������');"""%> />(ʱ��Ĭ��Ϊ0:0:0)
			</form>
			<form method="post" action="admin_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				ɾ��ָ�����ں�ʱ�������ԣ�����������/ʱ�䡣<br/>
				<input type="hidden" name="option" value="2" />
				<input type="text" name="iparam" size="<%=AdvDelTextWidth%>" maxlength="20" />
				<input type="submit" value="ִ��" name="submit1"<%if DelAdvTip=true then Response.Write " onclick=""return confirm('ȷʵҪִ��ɾ��������');"""%> />(ʱ��Ĭ��Ϊ0:0:0)
			</form>
			<form method="post" action="admin_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				ɾ���ǰ(����)��n�����ԣ�������n��ֵ��<br/>
				<input type="hidden" name="option" value="3" />
				<input type="text" name="iparam" size="<%=AdvDelTextWidth%>" maxlength="10" />
				<input type="submit" value="ִ��" name="submit1"<%if DelAdvTip=true then Response.Write " onclick=""return confirm('ȷʵҪִ��ɾ��������');"""%> />
			</form>
			<form method="post" action="admin_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				ɾ�����(����)��n�����ԣ�������n��ֵ��<br/>
				<input type="hidden" name="option" value="4" />
				<input type="text" name="iparam" size="<%=AdvDelTextWidth%>" maxlength="10" />
				<input type="submit" value="ִ��" name="submit1"<%if DelAdvTip=true then Response.Write " onclick=""return confirm('ȷʵҪִ��ɾ��������');"""%> />
			</form>
			<form method="post" action="admin_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				ɾ���ƺ��а����ض��ַ��������ԣ�<br/>
				<input type="hidden" name="option" value="5" />
				<input type="text" name="iparam" size="<%=AdvDelTextWidth%>" maxlength="64" />
				<input type="submit" value="ִ��" name="submit1"<%if DelAdvTip=true then Response.Write " onclick=""return confirm('ȷʵҪִ��ɾ��������');"""%> />("%"Ϊ����ַ���"_"Ϊһ���ַ�)
			</form>
			<form method="post" action="admin_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				ɾ�������а����ض��ַ��������ԣ�<br/>
				<input type="hidden" name="option" value="6" />
				<input type="text" name="iparam" size="<%=AdvDelTextWidth%>" maxlength="64" />
				<input type="submit" value="ִ��" name="submit1"<%if DelAdvTip=true then Response.Write " onclick=""return confirm('ȷʵҪִ��ɾ��������');"""%> />("%"Ϊ����ַ���"_"Ϊһ���ַ�)
			</form>
			<form method="post" action="admin_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				ɾ�����������а����ض��ַ��������ԣ�<br/>
				<input type="hidden" name="option" value="7" />
				<input type="text" name="iparam" size="<%=AdvDelTextWidth%>" maxlength="64" />
				<input type="submit" value="ִ��" name="submit1"<%if DelAdvTip=true then Response.Write " onclick=""return confirm('ȷʵҪִ��ɾ��������');"""%> />("%"Ϊ����ַ���"_"Ϊһ���ַ�)
			</form>
			<form method="post" action="admin_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				ɾ��ȫ�����ԣ�������Ա��������ʹ�ã�
				<input type="hidden" name="option" value="8" />
				<input type="hidden" name="iparam" value="DEL_ALL" />
				<input type="submit" value="ִ��" name="submit1"<%if DelAdvTip=true then Response.Write " onclick=""return confirm('ȷʵҪִ��ɾ��������');"""%> />
			</form>
		</div>
	</div>

</div>

<!-- #include file="bottom.asp" -->
</body>
</html>