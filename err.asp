<!-- #include file="config.asp" -->

<%Response.Expires=-1%>

<!-- #include file="include/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/metatag.inc" -->
	<title><%=HomeName%> ���Ա� ����</title>
	<!-- #include file="inc_stylesheet.asp" -->
</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<div id="outerborder" class="outerborder">

	<%if ShowTitle=true then show_book_title 3,"����"%>

	<span style="float:right; clear:all; width:100%; text-align:right; padding:4px; color:<%=LinkNormal%>;">
			| <a href="admin_login.asp">����</a>
	</span>
	<%
	dim errmsg
	select case Request("number")
	case "1"
		errmsg="�Բ������ķ����ѱ���ֹ��"
	case "2"
		errmsg="��Ǹ�����Ա��ѹرգ����Ժ����ԡ�"
	case "3"
		errmsg="��Ǹ������Ȩ���ѹرգ����Ժ����ԡ�"
	case "4"
		errmsg="�Բ������������к��н�ֹ���ֵ����ݡ�"
	case "5"
		errmsg="��Ǹ������Ȩ���ѹرգ����Ժ����ԡ�"
	case "6"
		errmsg="�Բ������ķ����ٶ�̫���ˣ�����Ϣһ�¡�"
	case "7"
		errmsg="�Բ����벻Ҫ�����ظ����ݡ�"
	case else
		errmsg="δ֪��������ϵ����Ա��"
	end select
	
	Response.Write "<br/><br/><span style=""width:100%; text-align:left;"">��������" &errmsg& "</span><br/><br/><br/>"
	%>
</div>

<!-- #include file="include/footer.inc" -->
</body>
</html>