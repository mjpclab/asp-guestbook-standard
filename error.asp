<%
sub ErrorPage(errorCode)
	Response.Status="403 Forbidden"

	dim errmsg
	select case errorCode
	case 1
		errmsg="�Բ������ķ����ѱ���ֹ��"
	case 2
		errmsg="��Ǹ�����Ա��ѹرգ����Ժ����ԡ�"
	case 3
		errmsg="��Ǹ������Ȩ���ѹرգ����Ժ����ԡ�"
	case 4
		errmsg="�Բ������������к��н�ֹ���ֵ����ݡ�"
	case 5
		errmsg="��Ǹ������Ȩ���ѹرգ����Ժ����ԡ�"
	case 6
		errmsg="�Բ������ķ����ٶ�̫���ˣ�����Ϣһ�¡�"
	case 7
		errmsg="�Բ����벻Ҫ�����ظ����ݡ�"
	case else
		errmsg="δ֪��������ϵ����Ա��"
	end select
	%>
	<!-- #include file="include/template/dtd.inc" -->
	<html>
	<head>
		<!-- #include file="include/template/metatag.inc" -->
		<title><%=HomeName%> ���Ա� ����</title>
		<!-- #include file="inc_stylesheet.asp" -->
	</head>

	<body<%=bodylimit%> onload="<%=framecheck%>">

	<div id="outerborder" class="outerborder">

		<%if ShowTitle=true then show_book_title 3,"����"%>

		<div id="mainborder" class="mainborder">
		<div class="guest-functions">
			<div class="aside">
				<a class="function" href="admin.asp">����</a>
			</div>
		</div>

		<p style="margin-bottom: 3em;">��������<%=errmsg%></p>
		</div>

		<!-- #include file="include/template/footer.inc" -->
	</div>
	</body>
	</html>
	<%
end sub
%>
