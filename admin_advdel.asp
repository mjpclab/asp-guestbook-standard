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
	<title><%=HomeName%> ���Ա� �߼�ɾ��</title>
	<!-- #include file="inc_admin_stylesheet.asp" -->
</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<div id="outerborder" class="outerborder">

	<%if ShowTitle then%><%Call InitHeaderData("����")%><!-- #include file="include/template/header.inc" --><%end if%>
	<div id="mainborder" class="mainborder">
	<!-- #include file="include/template/admin_mainmenu.inc" -->
	<div class="region form-region">
		<h3 class="title">�߼�ɾ��</h3>
		<div class="content">
			<form method="post" action="admin_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				ɾ��ָ�����ں�ʱ��ǰ�����ԣ�����������/ʱ�䡣<br/>
				<input type="hidden" name="option" value="1" />
				<input type="text" name="iparam" maxlength="20" />
				<input type="submit" value="ִ��" name="submit1"<%if DelAdvTip then Response.Write " onclick=""return confirm('ȷʵҪִ��ɾ��������');"""%> />(ʱ��Ĭ��Ϊ0:0:0)
			</form>
			<form method="post" action="admin_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				ɾ��ָ�����ں�ʱ�������ԣ�����������/ʱ�䡣<br/>
				<input type="hidden" name="option" value="2" />
				<input type="text" name="iparam" maxlength="20" />
				<input type="submit" value="ִ��" name="submit1"<%if DelAdvTip then Response.Write " onclick=""return confirm('ȷʵҪִ��ɾ��������');"""%> />(ʱ��Ĭ��Ϊ0:0:0)
			</form>
			<form method="post" action="admin_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				ɾ���ǰ(����)��n�����ԣ�������n��ֵ��<br/>
				<input type="hidden" name="option" value="3" />
				<input type="text" name="iparam" maxlength="10" />
				<input type="submit" value="ִ��" name="submit1"<%if DelAdvTip then Response.Write " onclick=""return confirm('ȷʵҪִ��ɾ��������');"""%> />
			</form>
			<form method="post" action="admin_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				ɾ�����(����)��n�����ԣ�������n��ֵ��<br/>
				<input type="hidden" name="option" value="4" />
				<input type="text" name="iparam" maxlength="10" />
				<input type="submit" value="ִ��" name="submit1"<%if DelAdvTip then Response.Write " onclick=""return confirm('ȷʵҪִ��ɾ��������');"""%> />
			</form>
			<form method="post" action="admin_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				ɾ���ƺ��а����ض��ַ��������ԣ�<br/>
				<input type="hidden" name="option" value="5" />
				<input type="text" name="iparam" maxlength="64" />
				<input type="submit" value="ִ��" name="submit1"<%if DelAdvTip then Response.Write " onclick=""return confirm('ȷʵҪִ��ɾ��������');"""%> />("%"Ϊ����ַ���"_"Ϊһ���ַ�)
			</form>
			<form method="post" action="admin_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				ɾ�������а����ض��ַ��������ԣ�<br/>
				<input type="hidden" name="option" value="6" />
				<input type="text" name="iparam" maxlength="64" />
				<input type="submit" value="ִ��" name="submit1"<%if DelAdvTip then Response.Write " onclick=""return confirm('ȷʵҪִ��ɾ��������');"""%> />("%"Ϊ����ַ���"_"Ϊһ���ַ�)
			</form>
			<form method="post" action="admin_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				ɾ�����������а����ض��ַ��������ԣ�<br/>
				<input type="hidden" name="option" value="7" />
				<input type="text" name="iparam" maxlength="64" />
				<input type="submit" value="ִ��" name="submit1"<%if DelAdvTip then Response.Write " onclick=""return confirm('ȷʵҪִ��ɾ��������');"""%> />("%"Ϊ����ַ���"_"Ϊһ���ַ�)
			</form>
			<form method="post" action="admin_doadvdel.asp" onsubmit="this.submit1.disabled=true;">
				ɾ��ȫ�����ԣ�������Ա��������ʹ�ã�
				<input type="hidden" name="option" value="8" />
				<input type="hidden" name="iparam" value="DEL_ALL" />
				<input type="submit" value="ִ��" name="submit1"<%if DelAdvTip then Response.Write " onclick=""return confirm('ȷʵҪִ��ɾ��������');"""%> />
			</form>
		</div>
	</div>
	</div>

	<!-- #include file="include/template/footer.inc" -->
</div>
</body>
</html>