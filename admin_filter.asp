<!-- #include file="config.asp" -->
<!-- #include file="admin_verify.asp" -->

<%Response.Expires=-1%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=HomeName%> ���Ա� ���ݹ��˲���</title>
	<link rel="stylesheet" type="text/css" href="style.css"/>
	<link rel="stylesheet" type="text/css" href="adminstyle.css"/>
	<!-- #include file="style.asp" -->
	<!-- #include file="adminstyle.asp" -->
</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<div id="outerborder" class="outerborder">

	<%if ShowTitle=true then show_book_title 3,"����"%>
	<!-- #include file="admincontrols.inc" -->

	<div class="region form-region region-filter">
		<h3 class="title">���ݹ��˲���</h3>
		<div class="content">
			<form method="post" name="newfilter" action="admin_appendfilter.asp" class="detail-item" onsubmit="if(findexp.value.length===0){alert('������������ݡ�');findexp.focus();return false;}submit1.disabled=true;">
			<h4>����¹��˲��ԣ�</h4>
			<p>��������(����������ʽ,������˴ʼ��á�|���ָ�)<br/>
			<input type="text" size="<%=FilterTextWidth%>" name="findexp" /><br/>
			<input type="checkbox" name="matchcase" id="matchcase" value="8192" /><label for="matchcase">���ִ�Сд</label>
			<input type="checkbox" name="multiline" id="multiline" value="2048" /><label for="multiline">�������ģʽ</label>
			</p>
			<p>���ҷ�Χ<br/>
			<input type="checkbox" name="findrange" id="findname" value="1" checked="checked" /><label for="findname">�ƺ�</label>
			<input type="checkbox" name="findrange" id="findmail" value="2" checked="checked" /><label for="findmail">�ʼ�</label>
			<input type="checkbox" name="findrange" id="findqq" value="4" checked="checked" /><label for="findqq">QQ��</label>
			<input type="checkbox" name="findrange" id="findmsn" value="8" checked="checked" /><label for="findmsn">Skype</label>
			<input type="checkbox" name="findrange" id="findhome" value="16" checked="checked" /><label for="findhome">��ҳ</label>
			<input type="checkbox" name="findrange" id="findtitle" value="32" checked="checked" /><label for="findtitle">����</label>
			<input type="checkbox" name="findrange" id="findcontent" value="64" checked="checked" /><label for="findcontent">����</label>
			</p>
			<p>����ʽ<br/>
			<input type="radio" name="filtermethod" id="filtermethod" value="0" checked="checked" onclick="if(typeof(newfilter.replacetxt.disabled)!=='undefined')newfilter.replacetxt.disabled=false;" /><label for="filtermethod">�滻Ϊ������ı�</label>
			<input type="radio" name="filtermethod" id="filtermethod2" value="4096" onclick="if(typeof(newfilter.replacetxt.disabled)!=='undefined')newfilter.replacetxt.disabled=true;" /><label for="filtermethod2">�ȴ����</label>
			<input type="radio" name="filtermethod" id="filtermethod3" value="16384" onclick="if(typeof(newfilter.replacetxt.disabled)!=='undefined')newfilter.replacetxt.disabled=true;" /><label for="filtermethod3">�ܾ�����</label><br/>
			<input type="text" size="<%=FilterTextWidth%>" name="replacetxt" />
			</p>
			<p>��ע<br/>
			<input type="text" size="<%=FilterTextWidth%>" name="memo" maxlength="25" />
			</p>
			<div class="field-command"><input type="submit" value="��ӹ��˲���" name="submit1" /></div>
			</form>

			<%
			set cn=server.CreateObject("ADODB.Connection")
			set rs=server.CreateObject("ADODB.Recordset")

			CreateConn cn,dbtype
			rs.Open sql_adminfilter,cn,,,1

			while rs.EOF=false%>
			<form method="post" action="admin_updatefilter.asp" class="detail-item">
				<%tfilterid=rs("filterid")%>
				<input type="hidden" name="filterid" value="<%=tfilterid%>" />
				<%tfiltermode=clng(rs("filtermode"))%>
				<p>��������<br/>
				<input type="text" size="<%=FilterTextWidth%>" name="findexp" value="<%=rs("regexp")%>" /><br/>
				<input type="checkbox" name="matchcase" id="matchcase<%=tfilterid%>" value="8192"<%=cked(clng(tfiltermode and 8192)<>0)%> /><label for="matchcase<%=tfilterid%>">���ִ�Сд</label>
				<input type="checkbox" name="multiline" id="multiline<%=tfilterid%>" value="2048"<%=cked(clng(tfiltermode and 2048)<>0)%> /><label for="multiline<%=tfilterid%>">�������ģʽ</label>
				</p>
				<p>���ҷ�Χ<br/>
				<input type="checkbox" name="findrange" id="findname<%=tfilterid%>" value="1"<%=cked(clng(tfiltermode and 1)<>0)%> /><label for="findname<%=tfilterid%>">�ƺ�</label>
				<input type="checkbox" name="findrange" id="findmail<%=tfilterid%>" value="2"<%=cked(clng(tfiltermode and 2)<>0)%> /><label for="findmail<%=tfilterid%>">�ʼ�</label>
				<input type="checkbox" name="findrange" id="findqq<%=tfilterid%>" value="4"<%=cked(clng(tfiltermode and 4)<>0)%> /><label for="findqq<%=tfilterid%>">QQ��</label>
				<input type="checkbox" name="findrange" id="findmsn<%=tfilterid%>" value="8"<%=cked(clng(tfiltermode and 8)<>0)%> /><label for="findmsn<%=tfilterid%>">Skype</label>
				<input type="checkbox" name="findrange" id="findhome<%=tfilterid%>" value="16"<%=cked(clng(tfiltermode and 16)<>0)%> /><label for="findhome<%=tfilterid%>">��ҳ</label>
				<input type="checkbox" name="findrange" id="findtitle<%=tfilterid%>" value="32"<%=cked(clng(tfiltermode and 32)<>0)%> /><label for="findtitle<%=tfilterid%>">����</label>
				<input type="checkbox" name="findrange" id="findcontent<%=tfilterid%>" value="64"<%=cked(clng(tfiltermode and 64)<>0)%> /><label for="findcontent<%=tfilterid%>">����</label>
				</p>
				<p>����ʽ<br/>
				<input type="radio" name="filtermethod" id="filtermethoda<%=tfilterid%>" value="0"<%=cked(clng(tfiltermode and 16384)=0)%> onclick="if(typeof(this.form.replacetxt.disabled)!='undefined')this.form.replacetxt.disabled=false;" /><label for="filtermethoda<%=tfilterid%>">�滻Ϊ������ı�</label>
				<input type="radio" name="filtermethod" id="filtermethodb<%=tfilterid%>" value="4096"<%=cked(clng(tfiltermode and 4096)<>0)%> onclick="if(typeof(this.form.replacetxt.disabled)!='undefined')this.form.replacetxt.disabled=true;" /><label for="filtermethodb<%=tfilterid%>">�ȴ����</label>
				<input type="radio" name="filtermethod" id="filtermethodc<%=tfilterid%>" value="16384"<%=cked(clng(tfiltermode and 16384)<>0)%> onclick="if(typeof(this.form.replacetxt.disabled)!='undefined')this.form.replacetxt.disabled=true;" /><label for="filtermethodc<%=tfilterid%>">�ܾ�����</label><br/>
				<input type="text" size="<%=FilterTextWidth%>" name="replacetxt" value="<%=rs("replacestr")%>"<%=dised(clng(tfiltermode and 16384+4096)<>0)%> />
				</p>
				<p>��ע<br/>
				<input type="text" size="<%=FilterTextWidth%>" name="memo" maxlength="25" value="<%=rs("memo")%>" />
				</p>
				<div class="field-command">
				<input type="submit" value="����" onclick="if (this.form.findexp.value.length===0) {alert('������������ݡ�');this.form.findexp.focus();return false;}" />
				<input type="button" value="ɾ��" onclick="if (confirm('ȷʵҪɾ������������')) {this.form.action='admin_delfilter.asp';this.form.submit();}" />
				<input type="button" value="����" onclick="{this.form.movedirection.value='up';this.form.action='admin_movefilter.asp';this.form.submit();}" />
				<input type="button" value="����" onclick="{this.form.movedirection.value='down';this.form.action='admin_movefilter.asp';this.form.submit();}" />
				<input type="button" value="��������" onclick="{this.form.movedirection.value='top';this.form.action='admin_movefilter.asp';this.form.submit();}" />
				<input type="button" value="�����ײ�" onclick="{this.form.movedirection.value='bottom';this.form.action='admin_movefilter.asp';this.form.submit();}" />
				<input type="hidden" name="movedirection" value="" />
				</div>
			</form>
			<%
			rs.MoveNext
			wend

			rs.Close : cn.Close : set rs=nothing : set cn=nothing
			%>
		</div>
	</div>

</div>

<!-- #include file="bottom.asp" -->
</body>
</html>