<!-- #include file="config.asp" -->
<!-- #include file="admin_verify.asp" -->

<%Response.Expires=-1%>

<!-- #include file="inc_dtd.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="zh-cn">
<head>
	<title><%=HomeName%> ���Ա� �޸İ�������</title>
	<!-- #include file="style.asp" -->
</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<div id="outerborder" class="outerborder">

	<%if ShowTitle=true then show_book_title 3,"����"%>
	<!-- #include file="admintool.inc" -->

	<%
	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")

	CreateConn cn,dbtype
	rs.Open sql_adminsetinfo,cn,,,1
	tfaceid=rs("faceid")
	%>

	<table border="1" cellpadding="2" class="generalwindow">
		<tr>
			<td class="centertitle">�޸İ�������</td>
		</tr>
		<tr>
		<td class="wordscontent" style="padding:20px 10px;">
			<form method="post" action="admin_saveinfo.asp" name="form1" onsubmit="form1.submit1.disabled=true;">
			�ǳƣ�����<input type="text" name="aname" size="<%=SetInfoTextWidth%>" maxlength="20" value="<%="" & rs("name") & ""%>" /><br/>
			�ʼ�������<input type="text" name="aemail" size="<%=SetInfoTextWidth%>" maxlength="50" value="<%="" & rs("email") & ""%>" /><br/>
			QQ�ţ�����<input type="text" name="aqqid" size="<%=SetInfoTextWidth%>" maxlength="16" value="<%="" & rs("qqid") & ""%>" /><br/>
			MSN ������<input type="text" name="amsnid" size="<%=SetInfoTextWidth%>" maxlength="50" value="<%="" & rs("msnid") & ""%>" /><br/>
			��ҳ������<input type="text" name="ahomepage" size="<%=SetInfoTextWidth%>" maxlength="127" value="<%="" & rs("homepage") & ""%>" /><br/><br/>
			ͷ���ţ�<input type="text" name="afaceid" size="<%=SetInfoTextWidth%>" maxlength="3" value="<%=tfaceid%>" title="��дͷ����ʱURL�������" /><br/>
			��URL ����<input type="text" name="afaceurl" size="<%=SetInfoTextWidth%>" maxlength="127" value="<%="" & rs("faceurl") & ""%>" title="��дURLʱ����ͷ����" /><br/>
			
			<%
			rs.Close : cn.Close : set rs=nothing : set cn=nothing

			dim listfacecount,defaultindex
			listfacecount=FrequentFaceCount
			defaultindex=tfaceid%>
			<!-- #include file="listface.inc" -->
			
			<br/><input value="��������" type="submit" name="submit1" />
			</form>
		</td>
		</tr>
	</table>
</div>

<!-- #include file="bottom.asp" -->
</body>
</html>