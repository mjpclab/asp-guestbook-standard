<!-- #include file="config.asp" -->
<!-- #include file="admin_verify.asp" -->

<%Response.Expires=-1%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=HomeName%> ���Ա� ����ͳ��</title>
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
set cn=server.CreateObject("ADODB.Connection")
set rs=server.CreateObject("ADODB.Recordset")

CreateConn cn,dbtype
rs.open sql_adminstats_startdate,cn,0,3,1
if rs.EOF then
	cn.Execute Replace(sql_adminstats_insert,"{0}",now()),,1
else
	if isdate(rs(0))=false then
		rs(0)=now()
		rs.Update
	end if
end if
rs.Close

rs.Open sql_adminstats,cn,,,1

tstartdate=rs("startdate")
tview=rs("view")
tsearch=rs("search")
tleaveword=rs("leaveword")
twritten=rs("written")
tfiltered=rs("filtered")
tbanned=rs("banned")
tlogin=rs("login")
tloginfailed=rs("loginfailed")
tnow=now()

rs.Close

on error resume next
%>

<div class="region form-region">
	<h3 class="title">����ͳ��</h3>
	<div class="content">
		<h4>��ʼͳ������</h4>
		<blockquote>
			<p><%=Year(tstartdate) & "-" & Month(tstartdate) & "-" & Day(tstartdate)%></p>
		</blockquote>

		<div id="div_outer"></div>

		<div id="div_words">
			<h4>�����б�ҳ��</h4>
			<blockquote>
				<table>
					<tr><td style="width:120px;">���ʴ�����</td><td><%=tview%></td></tr>
					<tr><td style="width:120px;">ƽ���·��ʴ�����</td><td><%=formatnumber(tview/((datediff("m",tstartdate,tnow)+1)),2,true,false,false)%></td></tr>
					<tr><td style="width:120px;">ƽ���ܷ��ʴ�����</td><td><%=formatnumber(tview/((datediff("ww",tstartdate,tnow)+1)),2,true,false,false)%></td></tr>
					<tr><td style="width:120px;">ƽ���շ��ʴ�����</td><td><%=formatnumber(tview/((datediff("d",tstartdate,tnow)+1)),2,true,false,false)%></td></tr>
				</table>
			</blockquote>
		</div>

		<div id="div_search">
			<h4>����ҳ��</h4>
			<blockquote>
				<table>
					<tr><td style="width:120px;">����������</td><td><%=tsearch%></td></tr>
					<tr><td style="width:120px;">ƽ��������������</td><td><%=formatnumber(tsearch/((datediff("m",tstartdate,tnow)+1)),2,true,false,false)%></td></tr>
					<tr><td style="width:120px;">ƽ��������������</td><td><%=formatnumber(tsearch/((datediff("ww",tstartdate,tnow)+1)),2,true,false,false)%></td></tr>
					<tr><td style="width:120px;">ƽ��������������</td><td><%=formatnumber(tsearch/((datediff("d",tstartdate,tnow)+1)),2,true,false,false)%></td></tr>
				</table>
			</blockquote>
		</div>

		<div id="div_leaveword">
			<h4>����/�ظ�ҳ��</h4>
			<blockquote>
				<table>
					<tr><td style="width:120px;">���ʴ�����</td><td><%=tleaveword%></td></tr>
					<tr><td style="width:120px;">�ɹ����Դ�����</td><td><%=twritten%></td></tr>
					<tr><td style="width:120px;">���������ʣ�</td><td><%if tleaveword=0 then Response.Write "/" else Response.Write formatpercent((tleaveword-twritten)/tleaveword,2,true)%></td></tr>
					<tr><td style="width:120px;">���Ա����˴�����</td><td><%=tfiltered%></td></tr>
					<tr><td style="width:120px;">�����ʣ�</td><td><%if twritten+tfiltered+tbanned=0 then Response.Write "/" else Response.Write formatpercent(tfiltered/(twritten+tfiltered+tbanned),2,true)%></td></tr>
					<tr><td style="width:120px;">���Ա��ܾ�������</td><td><%=tbanned%></td></tr>
					<tr><td style="width:120px;">�ܾ��ʣ�</td><td><%if twritten+tfiltered+tbanned=0 then Response.Write "/" else Response.Write formatpercent(tbanned/(twritten+tfiltered+tbanned),2,true)%></td></tr>
				</table>
			</blockquote>
		</div>

		<div id="div_login">
			<h4>�����¼ҳ��</h4>
			<blockquote>
				<table>
					<tr><td style="width:120px;">��¼������</td><td><%=tlogin%></td></tr>
					<tr><td style="width:120px;">��¼ʧ�ܴ�����</td><td><%=tloginfailed%></td></tr>
					<tr><td style="width:120px;">��¼ʧ���ʣ�</td><td><%if tlogin=0 then Response.Write "/" else Response.Write formatpercent(tloginfailed/tlogin,2,true)%></td></tr>
				</table>
			</blockquote>
		</div>

		<%
		on error goto 0

		rs.Open sql_adminstats_client_count,cn,0,1,1
		if rs.EOF=false then tclientcount=rs.Fields(0) else tclientcount=0

		if tclientcount>0 then
			'�ͻ��˲���ϵͳ
			rs.Close
			rs.Open sql_adminstats_client_os,cn,,,1

			Response.Write "<div id=""div_os"">"
			Response.Write "<h4>�ͻ��˲���ϵͳ</h4>"
			Response.Write "<blockquote><table>"
			while rs.EOF=false
				Response.Write "<tr><td style=""width:180px;"">"
				Response.Write server.HTMLEncode(rs.Fields("os")) & "��"
				Response.Write "</td><td>"
				Response.Write rs.Fields(1) & "(" & formatpercent(rs.Fields(1)/tclientcount,2,true) & ")"
				Response.Write "</td></tr>"
				rs.MoveNext
			wend
			Response.Write "</table></blockquote></div>"

			'�ͻ��������
			rs.Close
			rs.Open sql_adminstats_client_browser,cn,0,1,1

			Response.Write "<div id=""div_browser"">"
			Response.Write "<h4>�ͻ��������</h4>"
			Response.Write "<blockquote><table>"
			while rs.EOF=false
				Response.Write "<tr><td style=""width:120px;"">"
				Response.Write server.HTMLEncode(rs.Fields("browser")) & "��"
				Response.Write "</td><td>"
				Response.Write rs.Fields(1) & "(" & formatpercent(rs.Fields(1)/tclientcount,2,true) & ")"
				Response.Write "</td></tr>"
				rs.MoveNext
			wend
			Response.Write "</table></blockquote></div>"

			'�ͻ�����Ļ�ֱ���
			rs.Close
			rs.Open sql_adminstats_client_screen,cn,0,1,1

			Response.Write "<div id=""div_screen"">"
			Response.Write "<h4>�ͻ�����Ļ�ֱ���</h4>"
			Response.Write "<blockquote><table>"
			while rs.EOF=false
				Response.Write "<tr><td style=""width:120px;"">"
				if rs.Fields("screenwh")<>"0*0" then
					Response.Write server.HTMLEncode(rs.Fields("screenwh")) & "��"
				else
					Response.Write "δ֪��"
				end if
				Response.Write "</td><td>"
				Response.Write rs.Fields(1) & "(" & formatpercent(rs.Fields(1)/tclientcount,2,true) & ")"
				Response.Write "</td></tr>"
				rs.MoveNext
			wend
			Response.Write "</table></blockquote></div>"

			'����ʱ��
			rs.Close
			rs.Open sql_adminstats_client_timesect,cn,0,1,1

			Response.Write "<div id=""div_timesect"">"
			Response.Write "<h4>����ʱ��</h4>"
			Response.Write "<blockquote><table>"
			while rs.EOF=false
				Response.Write "<tr><td style=""width:120px;"">"
				Response.Write server.HTMLEncode(rs.Fields(0) & ":00��" & rs.Fields(0) & ":59") & "��"
				Response.Write "</td><td>"
				Response.Write rs.Fields(1) & "(" & formatpercent(rs.Fields(1)/tclientcount,2,true) & ")"
				Response.Write "</td></tr>"
				rs.MoveNext
			wend
			Response.Write "</table></blockquote></div>"

			'��������
			dim weeklist
			weeklist=array("������","����һ","���ڶ�","������","������","������","������")

			rs.Close
			rs.Open sql_adminstats_client_week,cn,0,1,1

			Response.Write "<div id=""div_week"">"
			Response.Write "<h4>��������</h4>"
			Response.Write "<blockquote><table>"
			while rs.EOF=false
				Response.Write "<tr><td style=""width:120px;"">"
				'Response.Write server.HTMLEncode(weekdayname(rs.Fields(0),false,1))
				Response.Write server.HTMLEncode(weeklist(rs.Fields(0)-1)) & "��"
				Response.Write "</td><td>"
				Response.Write rs.Fields(1) & "(" & formatpercent(rs.Fields(1)/tclientcount,2,true) & ")"
				Response.Write "</td></tr>"
				rs.MoveNext
			wend
			Response.Write "</table></blockquote></div>"

			'���30�������
			rs.Close
			rs.Open sql_adminstats_client_30day,cn,0,1,1

			Response.Write "<div id=""div_30day"">"
			Response.Write "<h4>���30�������</h4>"
			Response.Write "<blockquote><table>"
			while rs.EOF=false
				Response.Write "<tr><td style=""width:120px;"">"
				Response.Write server.HTMLEncode(rs.Fields("datesect")) & "��"
				Response.Write "</td><td>"
				Response.Write rs.Fields(1)
				Response.Write "</td></tr>"
				rs.MoveNext
			wend
			Response.Write "</table></blockquote></div>"

			'������Դ
			rs.Close
			rs.Open sql_adminstats_client_source,cn,0,1,1

			Response.Write "<div id=""div_source"">"
			Response.Write "<h4>������Դ</h4>"
			Response.Write "<blockquote><table>"
			while rs.EOF=false
				Response.Write "<tr><td style=""width:150px;"">"
				Response.Write server.HTMLEncode(rs.Fields("sourceaddr")) & "��"
				Response.Write "</td><td>"
				Response.Write rs.Fields(1) & "(" & formatpercent(rs.Fields(1)/tclientcount,2,true) & ")"
				Response.Write "</td></tr>"
				rs.MoveNext
			wend
			Response.Write "</table></blockquote></div>"

		end if

		rs.Close : cn.Close : set rs=nothing : set cn=nothing
		%>

	</div>
	<form method="post" action="admin_clearstats.asp" onsubmit="if (confirm('ȷʵҪ��ͳ��������')){submit1.disabled=true;return true;} else return false;">
		<div class="command"><input type="submit" value="ͳ������" name="submit1" /></div>
	</form>
</div>
</div>

<script type="text/javascript" src="js/tabcontrol.js"></script>
<script type="text/javascript">
	tab=new TabControl('div_outer');

	tab.setOuterContainerCssClass('tab-outer-container');
	tab.setTitleContainerCssClass('tab-title-container');
	tab.setTitleCssClass('tab-title');
	tab.setTitleSelectedCssClass('tab-title-selected');
	tab.setPageContainerCssClass('tab-page-container');
	tab.setPageCssClass('tab-page');

	tab.addPage('div_words','�����б�ҳ��');
	tab.addPage('div_search','����ҳ��');
	tab.addPage('div_leaveword','����/�ظ�ҳ��');
	tab.addPage('div_login','�����¼ҳ��');
	tab.addPage('div_os','�ͻ��˲���ϵͳ');
	tab.addPage('div_browser','�ͻ��������');
	tab.addPage('div_screen','�ͻ�����Ļ�ֱ���');
	tab.addPage('div_timesect','����ʱ��');
	tab.addPage('div_week','��������');
	tab.addPage('div_30day','���30�������');
	tab.addPage('div_source','������Դ');
</script>
	
<!-- #include file="bottom.asp" -->
</body>
</html>