<!-- #include file="config.asp" -->
<!-- #include file="admin_verify.asp" -->

<%Response.Expires=-1%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=HomeName%> ���Ա� IP���β���</title>
	<link rel="stylesheet" type="text/css" href="css/style.css"/>
	<link rel="stylesheet" type="text/css" href="css/adminstyle.css"/>
	<!-- #include file="css/style.asp" -->
	<!-- #include file="css/adminstyle.asp" -->
</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<div id="outerborder" class="outerborder">

	<%if ShowTitle=true then show_book_title 3,"����"%>
	<!-- #include file="admincontrols.inc" -->

	<%
	function hextoip(byref valip)
		dim vi,strip
		strip=""
		for vi=0 to 3
			strip=strip & cstr(cint("&H" & mid(valip,vi*2+1,2)))
			if vi<>3 then strip=strip & "."
		next
		hextoip=strip
	end function

	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")

	CreateConn cn,dbtype
	rs.Open sql_adminipconfig_status,cn,,,1
	tipconstatus=rs("ipconstatus")
	rs.Close
	%>


<div class="region form-region">
	<h3 class="title">IP���β���</h3>
	<div class="content">
		<form method="post" action="admin_saveipconfig.asp" name="ipconfigform" onsubmit="submit1.disabled=true;">
		<table cellpadding="10">
			<tr>
				<td colspan="2"><input type="radio" name="ipconstatus" value="0" id="r1"<%=cked(tipconstatus=0)%> /><label for="r1">��ʹ��IP���β���</label></td>
			</tr>
			<tr>
				<td style="width:50%; vertical-align:top;">
					<p class="row"><input type="radio" name="ipconstatus" value="1" id="r2"<%=cked(tipconstatus=1)%> /><label for="r2">ֻ��������IP�Σ��������</label></p>
					<p class="row">�����IP��,��ʽ:"��ʼIP-��ֹIP"</p>
					<p class="row"><textarea name="txt1" rows="6" style="width:100%"></textarea></p>
					<p class="row">
					<%rs.Open sql_adminipconfig_status1,cn,,,1
					if rs.EOF=false then
						Response.Write "ѡ��Ҫɾ����IP�Σ�"
						while rs.EOF=false
							tlistid=rs("listid")
							tstartip=rs("startip")
							tendip=rs("endip")
							Response.Write "<input type=""checkbox"" name=""savediplist1"" value=""" &tlistid& """ id=""" &tlistid& """ /><label for=""" &tlistid& """>" & hextoip(tstartip) & "-" & hextoip(tendip) &"</label>"
							rs.MoveNext
						wend
						Response.Write ""
					end if
					rs.Close
					%>
					</p>
				</td>
				<td style="width:50%; vertical-align:top;">
					<p><input type="radio" name="ipconstatus" value="2" id="r3"<%=cked(tipconstatus=2)%> /><label for="r3">ֻ��������IP�Σ������������</label></p>
					<p>�����IP��,��ʽ:"��ʼIP-��ֹIP"</p>
					<p><textarea name="txt2" rows="6" style="width:100%"></textarea></p>
					<p>
					<%rs.Open sql_adminipconfig_status2,cn,,,1
					if rs.EOF=false then
						Response.Write "ѡ��Ҫɾ����IP�Σ�"
						while rs.EOF=false
							tlistid=rs("listid")
							tstartip=rs("startip")
							tendip=rs("endip")
							Response.Write "<input type=""checkbox"" name=""savediplist2"" value=""" &tlistid& """ id=""" &tlistid& """ /><label for=""" &tlistid& """>" & hextoip(tstartip) & "-" & hextoip(tendip) &"</label>"
							rs.MoveNext
						wend
						Response.Write ""
					end if
					rs.Close
					%>
					</p>
				</td>
			</tr>
		</table>
		<div class="command"><input type="submit" name="submit1" value="��������" /></div>
		</form>
	</div>
</div>

<%cn.Close : set rs=nothing : set cn=nothing%>
</div>

<!-- #include file="bottom.asp" -->
</body>
</html>