<!-- #include file="config.asp" -->
<!-- #include file="admin_verify.asp" -->

<%Response.Expires=-1%>

<!-- #include file="inc_dtd.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="zh-cn">
<head>
	<title><%=HomeName%> 留言本 修改版主资料</title>
	<!-- #include file="style.asp" -->
</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<div id="outerborder" class="outerborder">

	<%if ShowTitle=true then show_book_title 3,"管理"%>
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
			<td class="centertitle">修改版主资料</td>
		</tr>
		<tr>
		<td class="wordscontent" style="padding:20px 10px;">
			<form method="post" action="admin_saveinfo.asp" name="form1" onsubmit="form1.submit1.disabled=true;">
			昵称：　　<input type="text" name="aname" size="<%=SetInfoTextWidth%>" maxlength="20" value="<%="" & rs("name") & ""%>" /><br/>
			邮件：　　<input type="text" name="aemail" size="<%=SetInfoTextWidth%>" maxlength="50" value="<%="" & rs("email") & ""%>" /><br/>
			QQ号：　　<input type="text" name="aqqid" size="<%=SetInfoTextWidth%>" maxlength="16" value="<%="" & rs("qqid") & ""%>" /><br/>
			MSN ：　　<input type="text" name="amsnid" size="<%=SetInfoTextWidth%>" maxlength="50" value="<%="" & rs("msnid") & ""%>" /><br/>
			主页：　　<input type="text" name="ahomepage" size="<%=SetInfoTextWidth%>" maxlength="127" value="<%="" & rs("homepage") & ""%>" /><br/><br/>
			头像编号：<input type="text" name="afaceid" size="<%=SetInfoTextWidth%>" maxlength="3" value="<%=tfaceid%>" title="填写头像编号时URL必须清空" /><br/>
			或URL ：　<input type="text" name="afaceurl" size="<%=SetInfoTextWidth%>" maxlength="127" value="<%="" & rs("faceurl") & ""%>" title="填写URL时忽略头像编号" /><br/>
			
			<%
			rs.Close : cn.Close : set rs=nothing : set cn=nothing

			dim listfacecount,defaultindex
			listfacecount=FrequentFaceCount
			defaultindex=tfaceid%>
			<!-- #include file="listface.inc" -->
			
			<br/><input value="更新数据" type="submit" name="submit1" />
			</form>
		</td>
		</tr>
	</table>
</div>

<!-- #include file="bottom.asp" -->
</body>
</html>