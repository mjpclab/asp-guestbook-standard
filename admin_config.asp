<!-- #include file="config.asp" -->
<!-- #include file="admin_verify.asp" -->

<%Response.Expires=-1%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=HomeName%> ���Ա� ���Ա�����</title>
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
	rs.Open sql_adminconfig_config,cn,,,1
	%>

	<div class="region region-config admin-tools">
		<h3 class="title">���Ա���������</h3>
		<div class="content flex-box">
			<ul>
				<li><a href="admin_config.asp?page=1">��������</a></li>
				<li><a href="admin_config.asp?page=2">�ʼ�֪ͨ</a></li>
				<li><a href="admin_config.asp?page=4">����ߴ�</a></li>
				<li><a href="admin_config.asp?page=8">��������</a></li>
				<li><a href="admin_config.asp">ȫ������</a></li>
			</ul>

			<form method="post" action="admin_saveconfig.asp" name="configform" onsubmit="return check();">

			<%
			showpage=15
			if isnumeric(Request.QueryString("page")) and Request.QueryString("page")<>"" then showpage=clng(Request.QueryString("page"))

			if clng(showpage and 1)<> 0 then
			%>
				<h4>��������</h4>
				<%tstatus=rs("status")%>
				<div class="field">
					<span class="label">���Ա�״̬��</span>
					<span class="value"><input type="radio" name="status1" value="1" id="status11"<%=cked(clng(tstatus and 1)<>0)%> /><label for="status11">����</label>����<input type="radio" name="status1" value="0" id="status12"<%=cked(clng(tstatus and 1)=0)%> /><label for="status12">�ر�</label> (�ر�ʱ"�ÿ�����Ȩ��"������Ч)</span>
				</div>
				<div class="field">
					<span class="label">�ÿ�����Ȩ�ޣ�</span>
					<span class="value"><input type="radio" name="status2" value="2" id="status21"<%=cked(clng(tstatus and 2)<>0)%> /><label for="status21">����</label>����<input type="radio" name="status2" value="0" id="status22"<%=cked(clng(tstatus and 2)=0)%> /><label for="status22">�ر�</label></span>
				</div>
				<div class="field">
					<span class="label">�ÿ���������Ȩ�ޣ�</span>
					<span class="value"><input type="radio" name="status3" value="4" id="status31"<%=cked(clng(tstatus and 4)<>0)%> /><label for="status31">����</label>����<input type="radio" name="status3" value="0" id="status32"<%=cked(clng(tstatus and 4)=0)%> /><label for="status32">�ر�</label></span>
				</div>
				<div class="field">
					<span class="label">�ÿ�ͷ���ܣ�</span>
					<span class="value"><input type="radio" name="status7" value="8" id="status71"<%=cked(clng(tstatus and 8)<>0)%> /><label for="status71">����</label>����<input type="radio" name="status7" value="0" id="status72"<%=cked(clng(tstatus and 8)=0)%> /><label for="status72">�ر�</label></span>
				</div>
				<div class="field">
					<span class="label">������ʾǰ����ˣ�</span>
					<span class="value"><input type="radio" name="status4" value="16" id="status41"<%=cked(clng(tstatus and 16)<>0)%> /><label for="status41">���</label>����<input type="radio" name="status4" value="0" id="status42"<%=cked(clng(tstatus and 16)=0)%> /><label for="status42">�����</label></span>
				</div>
				<div class="field">
					<span class="label">���Ļ����ܣ�</span>
					<span class="value"><input type="radio" name="status5" value="32" id="status51"<%=cked(clng(tstatus and 32)<>0)%> /><label for="status51">����</label>����<input type="radio" name="status5" value="0" id="status52"<%=cked(clng(tstatus and 32)=0)%> /><label for="status52">�ر�</label></span>
				</div>
				<div class="field">
					<span class="label">�ÿ�Ϊ���Ļ����ܣ�</span>
					<span class="value"><input type="radio" name="status6" value="64" id="status61"<%=cked(clng(tstatus and 64)<>0)%> /><label for="status61">����</label>����<input type="radio" name="status6" value="0" id="status62"<%=cked(clng(tstatus and 64)=0)%> /><label for="status62">��ֹ</label></span>
				</div>
				<div class="field">
					<span class="label">����ÿͻظ���</span>
					<span class="value"><input type="radio" name="status8" value="128" id="status81"<%=cked(clng(tstatus and 128)<>0)%> /><label for="status81">����</label>����<input type="radio" name="status8" value="0" id="status82"<%=cked(clng(tstatus and 128)=0)%> /><label for="status82">��ֹ</label></span>
				</div>
				<div class="field">
					<span class="label">����ͳ�ƣ�</span>
					<span class="value"><input type="radio" name="status9" value="256" id="status91"<%=cked(clng(tstatus and 256)<>0)%> /><label for="status91">����</label>����<input type="radio" name="status9" value="0" id="status92"<%=cked(clng(tstatus and 256)=0)%> /><label for="status92">�ر�</label></span>
				</div>
				<div class="field">
					<span class="label">������վLogo��ַ��</span>
					<span class="value"><input type="text" size="<%=ConfigTextWidth%>" maxlength="127" name="homelogo" value="<%=rs("homelogo")%>" /></span>
				</div>
				<div class="field">
					<span class="label">������վ���ƣ�</span>
					<span class="value"><input type="text" size="<%=ConfigTextWidth%>" maxlength="30" name="homename" value="<%=rs("homename")%>" /></span>
				</div>
				<div class="field">
					<span class="label">������վ��ַ��</span>
					<span class="value"><input type="text" size="<%=ConfigTextWidth%>" maxlength="127" name="homeaddr" value="<%=rs("homeaddr")%>" /></span>
				</div>
				<%adminlimit=rs("adminhtml")%>
				<div class="field">
					<span class="label">����Ա��ȫ�����ã�</span>
					<span class="value"><input type="checkbox" value="1" name="adminviewcode" id="adminviewcode"<%=cked(clng(adminlimit and 8)<>0)%> /><label for="adminviewcode">�ÿ�������ʾʵ��HTML��UBB����</label></span>
				</div>
				<div class="field">
					<span class="label">����ԱĬ��HTMLȨ�ޣ�</span>
					<span class="value">
						<span class="row"><input type="checkbox" value="1" name="adminhtml" id="adminhtml"<%=cked(clng(adminlimit and 1)<>0)%> /><label for="adminhtml">����Ա�ظ�������Ĭ��֧��HTML</label></span>
						<span class="row"><input type="checkbox" value="1" name="adminubb" id="adminubb"<%=cked(clng(adminlimit and 2)<>0)%> /><label for="adminubb">����Ա�ظ�������Ĭ��֧��UBB</label></span>
						<span class="row"><input type="checkbox" value="1" name="adminertn" id="adminertn"<%=cked(clng(adminlimit and 4)<>0)%> /><label for="adminertn">����Ա��֧��HTML��UBBʱ��Ĭ������س�����</label></span>
					</span>
				</div>
				<%guestlimit=rs("guesthtml")%>
				<div class="field">
					<span class="label">�ÿ�HTMLȨ�ޣ�</span>
					<span class="value">
						<span class="row"><input type="checkbox" value="1" name="guesthtml" id="guesthtml"<%=cked(clng(guestlimit and 1)<>0)%> /><label for="guesthtml">�ÿ�����֧��HTML</label></span>
						<span class="row"><input type="checkbox" value="1" name="guestubb" id="guestubb"<%=cked(clng(guestlimit and 2)<>0)%> /><label for="guestubb">�ÿ�����֧��UBB</label></span>
						<span class="row"><input type="checkbox" value="1" name="guestertn" id="guestertn"<%=cked(clng(guestlimit and 4)<>0)%> /><label for="guestertn">�ÿͲ�֧��HTML��UBBʱ������س�����</label></span>
					</span>
				</div>
				<div class="field">
					<span class="label">����Ա��¼��ʱ��</span>
					<span class="value"><input type="text" size="4" maxlength="4" name="admintimeout" value="<%=rs("admintimeout")%>" />�� (Ĭ��=20)</span>
				</div>
				<div class="field">
					<span class="label">Ϊ�ÿ���ʾIPv4ǰ��</span>
					<span class="value"><input type="text" size="4" maxlength="1" name="showipv4" value="<%=ShowIPv4%>" />�ֽ� (��ѡֵ��0��4)</span>
				</div>
				<div class="field">
					<span class="label">Ϊ�ÿ���ʾIPv6ǰ��</span>
					<span class="value"><input type="text" size="4" maxlength="1" name="showipv6" value="<%=ShowIPv6%>" />�� (��ѡֵ��0��8)</span>
				</div>
				<div class="field">
					<span class="label">Ϊ����Ա��ʾIPv4ǰ��</span>
					<span class="value"><input type="text" size="4" maxlength="1" name="adminshowipv4" value="<%=AdminShowIPv4%>" />�ֽ� (��ѡֵ��0��4)</span>
				</div>
				<div class="field">
					<span class="label">Ϊ����Ա��ʾIPv6ǰ��</span>
					<span class="value"><input type="text" size="4" maxlength="1" name="adminshowipv6" value="<%=AdminShowIPv6%>" />�� (��ѡֵ��0��8)</span>
				</div>
				<div class="field">
					<span class="label">Ϊ����Ա��ʾԭIPv4ǰ��</span>
					<span class="value"><input type="text" size="4" maxlength="1" name="adminshoworiginalipv4" value="<%=AdminShowOriginalIPv4%>" />�ֽ� (��ѡֵ��0��4,ʹ�ô��������ʱ������ʾԭʼIP)</span>
				</div>
				<div class="field">
					<span class="label">Ϊ����Ա��ʾԭIPv6ǰ��</span>
					<span class="value"><input type="text" size="4" maxlength="1" name="adminshoworiginalipv6" value="<%=AdminShowOriginalIPv6%>" />�� (��ѡֵ��0��8,ʹ�ô��������ʱ������ʾԭʼIP)</span>
				</div>
				<div class="field">
					<span class="label">��¼��֤�볤�ȣ�</span>
					<span class="value"><input type="text" size="4" maxlength="2" name="vcodecount" value="<%=clng(rs("vcodecount") and &H0F)%>" />λ (��ѡֵ��0��10)</span>
				</div>
				<div class="field">
					<span class="label">������֤�볤�ȣ�</span>
					<span class="value"><input type="text" size="4" maxlength="2" name="writevcodecount" value="<%=clng(rs("vcodecount") and &HF0) \ &H10%>" />λ (��ѡֵ��0��10)</span>
				</div>
			<%
			end if

			if clng(showpage and 2)<> 0 then
			%>
				<%tmailflag=rs("mailflag")%>
				<h4>�ʼ�֪ͨ</h4>
				<div class="field">
					<span class="label">�����Ե���֪ͨ��</span>
					<span class="value"><input type="checkbox" value="1" name="mailnewinform" id="mailnewinform"<%=cked(clng(tmailflag and 1)<>0)%> /><label for="mailnewinform">����</label></span>
				</div>
				<div class="field">
					<span class="label">�����ظ�֪ͨѡ�</span>
					<span class="value"><input type="checkbox" value="1" name="mailreplyinform" id="mailreplyinform"<%=cked(clng(tmailflag and 2)<>0)%> /><label for="mailreplyinform">����</label></span>
				</div>
				<div class="field">
					<span class="label">�ʼ����������</span>
					<span class="value"><input type="radio" value="0" name="mailcomponent" id="mailcomponent0"<%=cked(clng(tmailflag and 4)=0)%> /><label for="mailcomponent0">JMail</label>��<input type="radio" value="1" name="mailcomponent" id="mailcomponent1"<%=cked(clng(tmailflag and 4)<>0)%> /><label for="mailcomponent1">CDO</label></span>
				</div>
				<div class="field">
					<span class="label">������֪ͨ���յ�ַ��</span>
					<span class="value"><input type="text" size="<%=ConfigTextWidth%>" maxlength="48" name="mailreceive" value="<%=rs("mailreceive")%>" /></span>
				</div>
				<div class="field">
					<span class="label">�����˵�ַ��</span>
					<span class="value"><input type="text" size="<%=ConfigTextWidth%>" maxlength="48" name="mailfrom" value="<%=rs("mailfrom")%>" /></span>
				</div>
				<div class="field">
					<span class="label">������SMTP��������ַ��</span>
					<span class="value"><input type="text" size="<%=ConfigTextWidth%>" maxlength="48" name="mailsmtpserver" value="<%=rs("mailsmtpserver")%>" /></span>
				</div>
				<div class="field">
					<span class="label">��¼�û���(����Ҫ)��</span>
					<span class="value"><input type="text" size="<%=ConfigTextWidth%>" maxlength="48" name="mailuserid" value="<%=rs("mailuserid")%>" /></span>
				</div>
				<div class="field">
					<span class="label">��¼����(����Ҫ)��</span>
					<span class="value"><input type="password" size="<%=ConfigTextWidth%>" maxlength="48" name="mailuserpass" value="<%=rs("mailuserpass")%>" /></span>
				</div>
				<div class="field">
					<span class="label">�ʼ������̶ȣ�</span>
					<span class="value"><input type="text" size="4" maxlength="1" name="maillevel" value="<%=rs("maillevel")%>" /> (1��졫5����)</span>
				</div>
			<%end if

			if clng(showpage and 4)<> 0 then
			%>
				<h4>����ߴ�</h4>
				<div class="field">
					<span class="label">��������</span>
					<span class="value"><input type="text" size="10" maxlength="30" name="cssfontfamily" value="<%=rs("cssfontfamily")%>" /></span>
				</div>
				<div class="field">
					<span class="label">�����С��</span>
					<span class="value"><input type="text" size="10" maxlength="30" name="cssfontsize" value="<%=rs("cssfontsize")%>" /></span>
				</div>
				<div class="field">
					<span class="label">�����м�ࣺ</span>
					<span class="value"><input type="text" size="10" maxlength="30" name="csslineheight" value="<%=rs("csslineheight")%>" /></span>
				</div>
				<div class="field">
					<span class="label">���Ա�����ȣ�</span>
					<span class="value"><input type="text" size="10" maxlength="5" name="tablewidth" value="<%=rs("tablewidth")%>" /> (Ĭ��=630,���ðٷֱ�)</span>
				</div>
				<div class="field">
					<span class="label">���������ࣺ</span>
					<span class="value"><input type="text" size="10" maxlength="3" name="windowspace" value="<%=rs("windowspace")%>" /> (Ĭ��=20,��λ:����)</span>
				</div>
				<div class="field">
					<span class="label">���Ա��󴰸��ȣ�</span>
					<span class="value"><input type="text" size="10" maxlength="5" name="tableleftwidth" value="<%=rs("tableleftwidth")%>" /> (Ĭ��=150,���ðٷֱ�)</span>
				</div>
				<div class="field">
					<span class="label">��д���ԡ��ı����ȣ�</span>
					<span class="value"><input type="text" size="10" maxlength="3" name="leavetextwidth" value="<%=rs("leavetextwidth")%>" /> (Ĭ��=20,��λ:��ĸ���)</span>
				</div>
				<div class="field">
					<span class="label">����֤�롱�ı����ȣ�</span>
					<span class="value"><input type="text" size="10" maxlength="3" name="leavevcodewidth" value="<%=rs("leavevcodewidth")%>" /> (Ĭ��=10,��λ:��ĸ���)</span>
				</div>
				<div class="field">
					<span class="label">���������ݡ��ı��߶ȣ�</span>
					<span class="value"><input type="text" size="10" maxlength="3" name="leavecontentheight" value="<%=rs("leavecontentheight")%>" /> (Ĭ��=7,��λ:��ĸ�߶�)</span>
				</div>
				<div class="field">
					<span class="label">�������ȣ�</span>
					<span class="value"><input type="text" size="10" maxlength="3" name="searchtextwidth" value="<%=rs("searchtextwidth")%>" /> (Ĭ��=20,��λ:��ĸ���)</span>
				</div>
				<div class="field">
					<span class="label">���߼�ɾ�������ı���</span>
					<span class="value"><input type="text" size="10" maxlength="3" name="advdeltextwidth" value="<%=rs("advdeltextwidth")%>" /> (Ĭ��=20,��λ:��ĸ���)</span>
				</div>
				<div class="field">
					<span class="label">���޸����ϡ����ı���</span>
					<span class="value"><input type="text" size="10" maxlength="3" name="setinfotextwidth" value="<%=rs("setinfotextwidth")%>" /> (Ĭ��=40,��λ:��ĸ���,���������)</span>
				</div>
				<div class="field">
					<span class="label">�����������ı���ȣ�</span>
					<span class="value"><input type="text" size="10" maxlength="3" name="configtextwidth" value="<%=rs("configtextwidth")%>" /> (Ĭ��=75,��λ:��ĸ���,��"��������"ҳ���ı�)</span>
				</div>
				<div class="field">
					<span class="label">�����ݹ��ˡ����ı���</span>
					<span class="value"><input type="text" size="10" maxlength="3" name="filtertextwidth" value="<%=rs("filtertextwidth")%>" /> (Ĭ��=62,��λ:��ĸ���)</span>
				</div>
				<div class="field">
					<span class="label">�ظ�������༭��߶ȣ�</span>
					<span class="value"><input type="text" size="10" maxlength="3" name="replytextheight" value="<%=rs("replytextheight")%>" /> (Ĭ��=10,��λ:��ĸ�߶�)</span>
				</div>
				<div class="field">
					<span class="label">ÿҳ��ʾ����������</span>
					<span class="value"><input type="text" size="10" maxlength="5" name="itemsperpage" value="<%=rs("itemsperpage")%>" /> (Ĭ��=5)</span>
				</div>
				<div class="field">
					<span class="label">ÿҳ��ʾ�ı�������</span>
					<span class="value"><input type="text" size="10" maxlength="5" name="titlesperpage" value="<%=rs("titlesperpage")%>" /> (Ĭ��=20)</span>
				</div>
				<div class="field">
					<span class="label">ͷ��ÿ����ʾ����Ŀ��</span>
					<span class="value"><input type="text" size="10" maxlength="3" name="picturesperrow" value="<%=rs("picturesperrow")%>" /> (Ĭ��=5)</span>
				</div>
				<div class="field">
					<span class="label">���������ͷ������</span>
					<span class="value"><input type="text" size="10" maxlength="3" name="frequentfacecount" value="<%=rs("frequentfacecount")%>" /> (Ĭ��=15,����"����ͷ��"��ʾȫ��)</span>
				</div>
			<%
			end if

			if clng(showpage and 8)<> 0 then
			%>
				<h4>��������</h4>
				<%tvisualflag=rs("visualflag")%>
				<div class="field">
					<span class="label">Ĭ�ϰ���ģʽ��</span>
					<span class="value"><input type="radio" name="displaymode" value="1" id="displaymode1"<%=cked(clng(tvisualflag and 1024)<>0)%> /><label for="displaymode1">��̳�б���ʽ</label>����<input type="radio" name="displaymode" value="0" id="displaymode2"<%=cked(clng(tvisualflag and 1024)=0)%> /><label for="displaymode2">���Ա���ʽ</label></span>
				</div>
				<div class="field">
					<span class="label">�ظ�������ʾλ�ã�</span>
					<span class="value"><input type="radio" name="replyinword" value="1" id="replyinword1"<%=cked(clng(tvisualflag and 1)<>0)%> /><label for="replyinword1">��Ƕ�ڷÿ�����</label>����<input type="radio" name="replyinword" value="0" id="replyinword2"<%=cked(clng(tvisualflag and 1)=0)%> /><label for="replyinword2">��ʾ�ڷÿ������·�</label></span>
				</div>
				<div class="field">
					<span class="label">�������ݱ����ص����ԣ�</span>
					<span class="value"><input type="radio" name="hidehidden" value="1" id="hidehidden1"<%=cked(clng(tvisualflag and 128)<>0)%> /><label for="hidehidden1" >����</label>����<input type="radio" name="hidehidden" value="0" id="hidehidden2"<%=cked(clng(tvisualflag and 128)=0)%> /><label for="hidehidden2" >��ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">���ش�������ԣ�</span>
					<span class="value"><input type="radio" name="hideaudit" value="1" id="hideaudit1"<%=cked(clng(tvisualflag and 256)<>0)%> /><label for="hideaudit1" >����</label>����<input type="radio" name="hideaudit" value="0" id="hideaudit2"<%=cked(clng(tvisualflag and 256)=0)%> /><label for="hideaudit2" >��ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">���ذ���δ�ظ����Ļ���</span>
					<span class="value"><input type="radio" name="hidewhisper" value="1" id="hidewhisper1"<%=cked(clng(tvisualflag and 512)<>0)%> /><label for="hidewhisper1" >����</label>����<input type="radio" name="hidewhisper" value="0" id="hidewhisper2"<%=cked(clng(tvisualflag and 512)=0)%> /><label for="hidewhisper2" >��ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">Logo��ʾģʽ��</span>
					<span class="value"><input type="radio" name="logobannermode" value="1" id="logobannermode1"<%=cked(clng(tvisualflag and 2048)<>0)%> /><label for="logobannermode1" >���ģʽ</label>����<input type="radio" name="logobannermode" value="0" id="logobannermode2"<%=cked(clng(tvisualflag and 2048)=0)%> /><label for="logobannermode2" >ͼƬģʽ</label></span>
				</div>
				<div class="field">
					<span class="label">��ҳ������ʾλ�ã�</span>
					<span class="value"><input type="radio" name="showpagelist" value="3" id="showpagelist3"<%=cked(clng(tvisualflag and 12)=12)%> /><label for="showpagelist3">���·�</label>��<input type="radio" name="showpagelist" value="1" id="showpagelist1"<%=cked(clng(tvisualflag and 12)=4)%> /><label for="showpagelist1">�Ϸ�</label>����<input type="radio" name="showpagelist" value="2" id="showpagelist2"<%=cked(clng(tvisualflag and 12)=8)%> /><label for="showpagelist2">�·�</label></span>
				</div>
				<div class="field">
					<span class="label">�ÿ�����������ʾλ�ã�</span>
					<span class="value"><input type="radio" name="showsearchbox" value="3" id="showsearchbox3"<%=cked(clng(tvisualflag and 48)=48)%> /><label for="showsearchbox3">���·�</label>��<input type="radio" name="showsearchbox" value="1" id="showsearchbox1"<%=cked(clng(tvisualflag and 48)=16)%> /><label for="showsearchbox1">�Ϸ�</label>����<input type="radio" name="showsearchbox" value="2" id="showsearchbox2"<%=cked(clng(tvisualflag and 48)=32)%> /><label for="showsearchbox2">�·�</label></span>
				</div>
				<div class="field">
					<span class="label">��ҳ�б�ģʽ��</span>
					<span class="value"><input type="radio" name="advpagelist" value="1" id="advpagelist1"<%=cked(clng(tvisualflag and 64)<>0)%> /><label for="advpagelist1">����ʽ</label>��<input type="radio" name="advpagelist" value="0" id="advpagelist2"<%=cked(clng(tvisualflag and 64)=0)%> /><label for="advpagelist2">ƽ��ʽ</label></span>
				</div>
				<div class="field">
					<span class="label">����ʽ��ҳ������</span>
					<span class="value"><input type="text" size="5" maxlength="3" name="advpagelistcount" value="<%=rs("advpagelistcount")%>" /> (Ĭ��=10)</span>
				</div>
				<div class="field">
					<span class="label">UBB��������</span>
					<span class="value"><input type="radio" name="showubbtool" value="1" id="showubbtool1"<%=cked(clng(tvisualflag and 2)<>0)%> /><label for="showubbtool1">��ʾ</label>����<input type="radio" name="showubbtool" value="0" id="showubbtool2"<%=cked(clng(tvisualflag and 2)=0)%> /><label for="showubbtool2">����</label></span>
				</div>
				<div class="field">
					<span class="label">UBB����(������UBB)��</span>
					<span class="value">
						<span class="row">
							<input type="checkbox" name="ubbflag_image" id="ubbflag_image" value="1"<%=cked(UbbFlag_image)%> /><label for="ubbflag_image">ͼƬ</label>
							<input type="checkbox" name="ubbflag_url" id="ubbflag_url" value="1"<%=cked(UbbFlag_url)%> /><label for="ubbflag_url">URL��Email</label>
							<input type="checkbox" name="ubbflag_autourl" id="ubbflag_autourl" value="1"<%=cked(UbbFlag_autourl)%> /><label for="ubbflag_autourl">�Զ�ʶ����ַ</label>
						</span>
						<span class="row">
							<input type="checkbox" name="ubbflag_player" id="ubbflag_player" value="1"<%=cked(UbbFlag_player)%> /><label for="ubbflag_player">���ſؼ�</label>
							<input type="checkbox" name="ubbflag_paragraph" id="ubbflag_paragraph" value="1"<%=cked(UbbFlag_paragraph)%> /><label for="ubbflag_paragraph">������ʽ</label>
							<input type="checkbox" name="ubbflag_fontstyle" id="ubbflag_fontstyle" value="1"<%=cked(UbbFlag_fontstyle)%> /><label for="ubbflag_fontstyle">������ʽ</label>
						</span>
						<span class="row">
							<input type="checkbox" name="ubbflag_fontcolor" id="ubbflag_fontcolor" value="1"<%=cked(UbbFlag_fontcolor)%> /><label for="ubbflag_fontcolor">������ɫ</label>
							<input type="checkbox" name="ubbflag_alignment" id="ubbflag_alignment" value="1"<%=cked(UbbFlag_alignment)%> /><label for="ubbflag_alignment">���뷽ʽ</label>
							<input type="checkbox" name="ubbflag_face" id="ubbflag_face" value="1"<%=cked(UbbFlag_face)%> /><label for="ubbflag_face">����ͼ��</label>
						</span>
					</span>
				</div>
				<%talign=rs("tablealign")%>
				<div class="field">
					<span class="label">���Ա����뷽ʽ��</span>
					<span class="value"><input type="radio" name="tablealign" value="left" id="align1"<%=cked(talign<>"center" and talign<>"right")%> /><label for="align1">�����</label>��<input type="radio" name="tablealign" value="center" id="align2"<%=cked(talign="center")%> /><label for="align2">����</label>����<input type="radio" name="tablealign" value="right" id="align3"<%=cked(talign="right")%> /><label for="align3">�Ҷ���</label></span>
				</div>
				<%tpagecontrol=rs("pagecontrol")%>
				<div class="field">
					<span class="label">���Ա��߿��ߣ�</span>
					<span class="value"><input type="radio" name="showborder" value="1" id="showborder1"<%=cked(clng(tpagecontrol and 1)<>0)%> /><label for="showborder1">��ʾ</label>����<input type="radio" name="showborder" value="0" id="showborder2"<%=cked(clng(tpagecontrol and 1)=0)%> /><label for="showborder2">����</label></span>
				</div>
				<div class="field">
					<span class="label">���Ա��ܱ��⣺</span>
					<span class="value"><input type="radio" name="showtitle" value="1" id="showtitle1"<%=cked(clng(tpagecontrol and 2)<>0)%> /><label for="showtitle1">��ʾ</label>����<input type="radio" name="showtitle" value="0" id="showtitle2"<%=cked(clng(tpagecontrol and 2)=0)%> /><label for="showtitle2">����</label></span>
				</div>
				<div class="field">
					<span class="label">��ҳ�Ҽ��˵���</span>
					<span class="value"><input type="radio" name="showcontext" value="1" id="showcontext1"<%=cked(clng(tpagecontrol and 4)<>0)%> /><label for="showcontext1">����</label>����<input type="radio" name="showcontext" value="0" id="showcontext2"<%=cked(clng(tpagecontrol and 4)=0)%> /><label for="showcontext2">����</label></span>
				</div>
				<div class="field">
					<span class="label">ѡ����ҳ���ݣ�</span>
					<span class="value"><input type="radio" name="selectcontent" value="1" id="selectcontent1"<%=cked(clng(tpagecontrol and 8)<>0)%> /><label for="selectcontent1">����</label>����<input type="radio" name="selectcontent" value="0" id="selectcontent2"<%=cked(clng(tpagecontrol and 8)=0)%> /><label for="selectcontent2">��ֹ</label></span>
				</div>
				<div class="field">
					<span class="label">������ҳ(Ctrl-C)��</span>
					<span class="value"><input type="radio" name="copycontent" value="1" id="copycontent1"<%=cked(clng(tpagecontrol and 16)<>0)%> /><label for="copycontent1">����</label>����<input type="radio" name="copycontent" value="0" id="copycontent2"<%=cked(clng(tpagecontrol and 16)=0)%> /><label for="copycontent2">��ֹ</label></span>
				</div>
				<div class="field">
					<span class="label">��IFrame��Ƕ��</span>
					<span class="value"><input type="radio" name="beframed" value="1" id="beframed1"<%=cked(clng(tpagecontrol and 32)<>0)%> /><label for="beframed1">����</label>����<input type="radio" name="beframed" value="0" id="beframed2"<%=cked(clng(tpagecontrol and 32)=0)%> /><label for="beframed2">��ֹ</label></span>
				</div>
				<div class="field">
					<span class="label">�����������ƣ�</span>
					<span class="value"><input type="text" size="10" maxlength="10" name="wordslimit" value="<%=rs("wordslimit")%>" /> (0=����,�س�����ռ2��,��UBB����ǰͳ��)</span>
				</div>
				<%tdelconfirm=rs("delconfirm")%>
				<div class="field">
					<span class="label">����ͨ�����ǰ��ʾ��</span>
					<span class="value"><input type="radio" name="passaudittip" value="1" id="passaudittip1"<%=cked(clng(tdelconfirm and 16)<>0)%> /><label for="passaudittip1">��ʾ</label>����<input type="radio" name="passaudittip" value="0" id="passaudittip2"<%=cked(clng(tdelconfirm and 16)=0)%> /><label for="passaudittip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">ѡ������ͨ�������ʾ��</span>
					<span class="value"><input type="radio" name="passseltip" value="1" id="passseltip1"<%=cked(clng(tdelconfirm and 32)<>0)%> /><label for="passseltip1">��ʾ</label>����<input type="radio" name="passseltip" value="0" id="passseltip2"<%=cked(clng(tdelconfirm and 32)=0)%> /><label for="passseltip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">ɾ������ʱ��ʾ��</span>
					<span class="value"><input type="radio" name="deltip" value="1" id="deltip1"<%=cked(clng(tdelconfirm and 1)<>0)%> /><label for="deltip1">��ʾ</label>����<input type="radio" name="deltip" value="0" id="deltip2"<%=cked(clng(tdelconfirm and 1)=0)%> /><label for="deltip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">ɾ���ظ�ʱ��ʾ��</span>
					<span class="value"><input type="radio" name="delretip" value="1" id="delretip1"<%=cked(clng(tdelconfirm and 2)<>0)%> /><label for="delretip1">��ʾ</label>����<input type="radio" name="delretip" value="0" id="delretip2"<%=cked(clng(tdelconfirm and 2)=0)%> /><label for="delretip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">ɾ��ѡ������ʱ��ʾ��</span>
					<span class="value"><input type="radio" name="delseltip" value="1" id="delseltip1"<%=cked(clng(tdelconfirm and 4)<>0)%> /><label for="delseltip1">��ʾ</label>����<input type="radio" name="delseltip" value="0" id="delseltip2"<%=cked(clng(tdelconfirm and 4)=0)%> /><label for="delseltip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">ִ�и߼�ɾ��ʱ��ʾ��</span>
					<span class="value"><input type="radio" name="deladvtip" value="1" id="deladvtip1"<%=cked(clng(tdelconfirm and 8)<>0)%> /><label for="deladvtip1">��ʾ</label>����<input type="radio" name="deladvtip" value="0" id="deladvtip2"<%=cked(clng(tdelconfirm and 8)=0)%> /><label for="deladvtip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">�������Ļ�ʱ��ʾ��</span>
					<span class="value"><input type="radio" name="pubwhispertip" value="1" id="pubwhispertip1"<%=cked(clng(tdelconfirm and 64)<>0)%> /><label for="pubwhispertip1">��ʾ</label>����<input type="radio" name="pubwhispertip" value="0" id="pubwhispertip2"<%=cked(clng(tdelconfirm and 64)=0)%> /><label for="pubwhispertip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">�ö�����ʱ��ʾ��</span>
					<span class="value"><input type="radio" name="lock2toptip" value="1" id="lock2toptip1"<%=cked(clng(tdelconfirm and 256)<>0)%> /><label for="lock2toptip1">��ʾ</label>����<input type="radio" name="lock2toptip" value="0" id="lock2toptip2"<%=cked(clng(tdelconfirm and 256)=0)%> /><label for="lock2toptip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">��ǰ����ʱ��ʾ��</span>
					<span class="value"><input type="radio" name="bring2toptip" value="1" id="bring2toptip1"<%=cked(clng(tdelconfirm and 128)<>0)%> /><label for="bring2toptip1">��ʾ</label>����<input type="radio" name="bring2toptip" value="0" id="bring2toptip2"<%=cked(clng(tdelconfirm and 128)=0)%> /><label for="bring2toptip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">��������˳��ʱ��ʾ��</span>
					<span class="value"><input type="radio" name="reordertip" value="1" id="reordertip1"<%=cked(clng(tdelconfirm and 512)<>0)%> /><label for="reordertip1">��ʾ</label>����<input type="radio" name="reordertip" value="0" id="reordertip2"<%=cked(clng(tdelconfirm and 512)=0)%> /><label for="reordertip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">���Ա���ɫ������</span>
					<span class="value">
						<select name="style">
						<%
						styleid=rs("styleid")
						rs.Close
						rs.Open sql_adminconfig_style,cn,,,1

						dim onestyleid,onestylename
						while rs.EOF=false
							onestyleid=rs("styleid")
							onestylename=rs("stylename")
							Response.Write "<option value=" &chr(34)& onestyleid &chr(34)
							if onestyleid=styleid or onestyleid="" then Response.Write " selected=""selected"""
							Response.Write ">" &onestylename& "</option>"
							rs.MoveNext
						wend
						%>
						</select>
					</span>
				</div>
			<%end if%>

			<input type="hidden" name="page" value="<%=showpage%>" />
			<div class="command"><input value="��������" type="submit" name="submit1" /></div>
			</form>
		</div>
	</div>

<%rs.Close : cn.Close : set rs=nothing : set cn=nothing%>
</div>

<script type="text/javascript">
function check()
{
	var tv,showpage=<%=showpage%>;
	if ((showpage & 1) != 0)
	{
		if (isNaN(tv=Number(document.configform.admintimeout.value)))
			{alert('������Ա��¼��ʱ������Ϊ���֡�');document.configform.admintimeout.select();return false;}
		else if (tv<1 || tv>1440)
			{alert('������Ա��¼��ʱ��������1��1440�ķ�Χ�ڡ�');document.configform.admintimeout.select();return false;}

		if (isNaN(tv=Number(document.configform.showipv4.value)))
			{alert('��Ϊ�ÿ���ʾIPv4������Ϊ���֡�');document.configform.showipv4.select();return false;}
		else if (tv<0 || tv>4 || document.configform.showipv4.value==='')
			{alert('��Ϊ�ÿ���ʾIPv4��������0��4�ķ�Χ�ڡ�');document.configform.showipv4.select();return false;}

		if (isNaN(tv=Number(document.configform.showipv6.value)))
			{alert('��Ϊ�ÿ���ʾIPv6������Ϊ���֡�');document.configform.showipv6.select();return false;}
		else if (tv<0 || tv>8 || document.configform.showipv6.value==='')
			{alert('��Ϊ�ÿ���ʾIPv6��������0��8�ķ�Χ�ڡ�');document.configform.showipv6.select();return false;}

		if (isNaN(tv=Number(document.configform.adminshowipv4.value)))
			{alert('��Ϊ����Ա��ʾIPv4������Ϊ���֡�');document.configform.adminshowipv4.select();return false;}
		else if (tv<0 || tv>4 || document.configform.adminshowipv4.value==='')
			{alert('��Ϊ����Ա��ʾIPv4��������0��4�ķ�Χ�ڡ�');document.configform.adminshowipv4.select();return false;}

		if (isNaN(tv=Number(document.configform.adminshowipv6.value)))
			{alert('��Ϊ����Ա��ʾIPv6������Ϊ���֡�');document.configform.adminshowipv6.select();return false;}
		else if (tv<0 || tv>8 || document.configform.adminshowipv6.value==='')
			{alert('��Ϊ����Ա��ʾIPv6��������0��8�ķ�Χ�ڡ�');document.configform.adminshowipv6.select();return false;}

		if (isNaN(tv=Number(document.configform.adminshoworiginalipv4.value)))
			{alert('��Ϊ����Ա��ʾԭIPv4������Ϊ����');document.configform.adminshoworiginalipv4.select();return false;}
		else if (tv<0 || tv>4 || document.configform.adminshoworiginalipv4.value==='')
			{alert('��Ϊ����Ա��ʾԭIPv4��������0��4�ķ�Χ�ڡ�');document.configform.adminshoworiginalipv4.select();return false;}

		if (isNaN(tv=Number(document.configform.adminshoworiginalipv6.value)))
			{alert('��Ϊ����Ա��ʾԭIPv6������Ϊ����');document.configform.adminshoworiginalipv6.select();return false;}
		else if (tv<0 || tv>8 || document.configform.adminshoworiginalipv6.value==='')
			{alert('��Ϊ����Ա��ʾԭIPv6��������0��8�ķ�Χ�ڡ�');document.configform.adminshoworiginalipv6.select();return false;}

		if (isNaN(tv=Number(document.configform.vcodecount.value)))
			{alert('����¼��֤�볤�ȡ�����Ϊ���֡�');document.configform.vcodecount.select();return false;}
		else if (tv<0 || tv>10)
			{alert('����¼��֤�볤�ȡ�������0��10�ķ�Χ�ڡ�');document.configform.vcodecount.select();return false;}

		if (isNaN(tv=Number(document.configform.writevcodecount.value)))
			{alert('��������֤�볤�ȡ�����Ϊ���֡�');document.configform.writevcodecount.select();return false;}
		else if (tv<0 || tv>10)
			{alert('��������֤�볤�ȡ�������0��10�ķ�Χ�ڡ�');document.configform.writevcodecount.select();return false;}
	}
	if ((showpage & 2) != 0)
	{
		if (isNaN(tv=Number(document.configform.maillevel.value)))
			{alert('���ʼ������̶ȡ�����Ϊ���֡�');document.configform.maillevel.select();return false;}
		else if (tv<1 || tv>5)
			{alert('���ʼ������̶ȡ�������1��5�ķ�Χ�ڡ�');document.configform.maillevel.select();return false;}	
	}
	if ((showpage & 4) != 0)
	{
		if (isNaN(tv=Number(document.configform.tablewidth.value)))
			{if (/^\d+%$/.test(document.configform.tablewidth.value)==false) {alert('�����Ա�����ȡ�����Ϊ���������ٷֱȡ�');document.configform.tablewidth.select();return false;}}
		else if (tv<1)
			{alert('�����Ա�����ȡ���������㡣');document.configform.tablewidth.select();return false;}

		if (isNaN(tv=Number(document.configform.tableleftwidth.value)))
			{if (/^\d+%$/.test(document.configform.tableleftwidth.value)==false) {alert('�����Ա��󴰸��ȡ�����Ϊ���������ٷֱȡ�');document.configform.tableleftwidth.select();return false;}}
		else if (tv<1)
			{alert('�����Ա��󴰸��ȡ���������㡣');document.configform.tableleftwidth.select();return false;}

		if (isNaN(tv=Number(document.configform.windowspace.value)))
			{alert('�����������ࡱ����Ϊ���֡�');document.configform.windowspace.select();return false;}
		else if (tv<1 || tv>255)
			{alert('�����������ࡱ������1��255�ķ�Χ�ڡ�');document.configform.windowspace.select();return false;}

		if (isNaN(tv=Number(document.configform.leavetextwidth.value)))
			{alert('����д���ԡ��ı����ȡ�����Ϊ���֡�');document.configform.leavetextwidth.select();return false;}
		else if (tv<1 || tv>255)
			{alert('����д���ԡ��ı����ȡ�������1��255�ķ�Χ�ڡ�');document.configform.leavetextwidth.select();return false;}

		if (isNaN(tv=Number(document.configform.leavevcodewidth.value)))
			{alert('������֤�롯�ı����ȡ�����Ϊ���֡�');document.configform.leavevcodewidth.select();return false;}
		else if (tv<1 || tv>255)
			{alert('������֤�롯�ı����ȡ�������1��255�ķ�Χ�ڡ�');document.configform.leavevcodewidth.select();return false;}

		if (isNaN(tv=Number(document.configform.leavecontentheight.value)))
			{alert('�����������ݡ��ı��߶ȡ�����Ϊ���֡�');document.configform.leavecontentheight.select();return false;}
		else if (tv<1 || tv>255)
			{alert('�����������ݡ��ı��߶ȡ�������1��255�ķ�Χ�ڡ�');document.configform.leavecontentheight.select();return false;}

		if (isNaN(tv=Number(document.configform.searchtextwidth.value)))
			{alert('���������ȡ�����Ϊ���֡�');document.configform.searchtextwidth.select();return false;}
		else if (tv<1 || tv>255)
			{alert('���������ȡ�������1��255�ķ�Χ�ڡ�');document.configform.searchtextwidth.select();return false;}

		if (isNaN(tv=Number(document.configform.advdeltextwidth.value)))
			{alert('�����߼�ɾ�������ı�������Ϊ���֡�');document.configform.advdeltextwidth.select();return false;}
		else if (tv<1 || tv>255)
			{alert('�����߼�ɾ�������ı���������1��255�ķ�Χ�ڡ�');document.configform.advdeltextwidth.select();return false;}

		if (isNaN(tv=Number(document.configform.setinfotextwidth.value)))
			{alert('�����޸����ϡ����ı�������Ϊ���֡�');document.configform.setinfotextwidth.select();return false;}
		else if (tv<1 || tv>255)
			{alert('�����޸����ϡ����ı���������1��255�ķ�Χ�ڡ�');document.configform.setinfotextwidth.select();return false;}

		if (isNaN(tv=Number(document.configform.configtextwidth.value)))
			{alert('�������������ı���ȡ�����Ϊ���֡�');document.configform.configtextwidth.select();return false;}
		else if (tv<1 || tv>255)
			{alert('�������������ı���ȡ�������1��255�ķ�Χ�ڡ�');document.configform.configtextwidth.select();return false;}

		if (isNaN(tv=Number(document.configform.filtertextwidth.value)))
			{alert('�������ݹ��ˡ����ı�������Ϊ���֡�');document.configform.filtertextwidth.select();return false;}
		else if (tv<1 || tv>255)
			{alert('�������ݹ��ˡ����ı���������1��255�ķ�Χ�ڡ�');document.configform.filtertextwidth.select();return false;}

		if (isNaN(tv=Number(document.configform.replytextheight.value)))
			{alert('���ظ�������༭��߶ȡ�����Ϊ���֡�');document.configform.replytextheight.select();return false;}
		else if (tv<1 || tv>255)
			{alert('���ظ�������༭��߶ȡ�������1��255�ķ�Χ�ڡ�');document.configform.replytextheight.select();return false;}

		if (isNaN(tv=Number(document.configform.itemsperpage.value)))
			{alert('��ÿҳ��ʾ��������������Ϊ���֡�');document.configform.itemsperpage.select();return false;}
		else if (tv<1 || tv>32767)
			{alert('��ÿҳ��ʾ����������������1��32767�ķ�Χ�ڡ�');document.configform.itemsperpage.select();return false;}

		if (isNaN(tv=Number(document.configform.titlesperpage.value)))
			{alert('��ÿҳ��ʾ�ı�����������Ϊ���֡�');document.configform.titlesperpage.select();return false;}
		else if (tv<1 || tv>32767)
			{alert('��ÿҳ��ʾ�ı�������������1��32767�ķ�Χ�ڡ�');document.configform.titlesperpage.select();return false;}

		if (isNaN(tv=Number(document.configform.picturesperrow.value)))
			{alert('��ͷ��ÿ����ʾ����Ŀ������Ϊ���֡�');document.configform.picturesperrow.select();return false;}
		else if (tv<1 || tv>255)
			{alert('��ͷ��ÿ����ʾ����Ŀ��������1��255�ķ�Χ�ڡ�');document.configform.picturesperrow.select();return false;}

		if (isNaN(tv=Number(document.configform.frequentfacecount.value)))
			{alert('�����������ͷ����������Ϊ���֡�');document.configform.frequentfacecount.select();return false;}
		else if (tv<0 || tv>255)
			{alert('�����������ͷ������������0��255�ķ�Χ�ڡ�');document.configform.frequentfacecount.select();return false;}
	}
	if ((showpage & 8) != 0)
	{
		if (isNaN(tv=Number(document.configform.advpagelistcount.value)))
			{alert('������ʽ��ҳ����������Ϊ���֡�');document.configform.advpagelistcount.select();return false;}
		else if (tv<1 || tv>255 || document.configform.wordslimit.value=='')
			{alert('������ʽ��ҳ������������1��255�ķ�Χ�ڡ�');document.configform.advpagelistcount.select();return false;}

		if (isNaN(tv=Number(document.configform.wordslimit.value)))
			{alert('�������������ơ�����Ϊ���֡�');document.configform.wordslimit.select();return false;}
		else if (tv<0 || tv>2147483647 || document.configform.wordslimit.value=='')
			{alert('�������������ơ�������0��2147483647�ķ�Χ�ڡ�');document.configform.wordslimit.select();return false;}
	}
	document.configform.submit1.disabled=true;
	return true;
}
</script>

<!-- #include file="bottom.asp" -->
</body>
</html>