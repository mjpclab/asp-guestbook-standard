<!-- #include file="config.asp" -->
<!-- #include file="admin_verify.asp" -->

<%Response.Expires=-1%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312"/>
	<title><%=HomeName%> ���Ա� �߼�ɾ��</title>
	
	<!-- #include file="style.asp" -->
	<style type="text/css">
	form {margin:20px 0px;}
	</style>
</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<div id="outerborder" class="outerborder">

	<%if ShowTitle=true then show_book_title 3,"����"%>
	<!-- #include file="admintool.inc" -->

	<table border="1" cellpadding="2" class="generalwindow">
		<tr>
			<td class="centertitle">�߼�ɾ��</td>
		</tr>
		<tr>
			<td class="wordscontent" style="padding:2px;">
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
		
			</td>
		</tr>
	</table>
</div>

<!-- #include file="bottom.asp" -->
</body>
</html>