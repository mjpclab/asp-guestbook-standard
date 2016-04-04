<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/sql/admin_config.asp" -->
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
	<title><%=HomeName%> ���Ա� ���Ա�����</title>
	<!-- #include file="inc_admin_stylesheet.asp" -->
</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<div id="outerborder" class="outerborder">

	<%if ShowTitle then%><%Call InitHeaderData("����")%><!-- #include file="include/template/header.inc" --><%end if%>
	<div id="mainborder" class="mainborder">
	<!-- #include file="include/template/admin_mainmenu.inc" -->
	<%
	set cn=server.CreateObject("ADODB.Connection")
	set rs=server.CreateObject("ADODB.Recordset")

	Call CreateConn(cn)
	rs.Open sql_adminconfig_config,cn,,,1
	tstatus=rs("status")
	adminlimit=rs("adminhtml")
	guestlimit=rs("guesthtml")
	tmailflag=rs("mailflag")
	tvisualflag=rs("visualflag")
	tpagecontrol=rs("pagecontrol")
	tdelconfirm=rs("delconfirm")
	%>

	<div class="region region-config admin-tools">
		<h3 class="title">���Ա���������</h3>
		<div class="content">
			<form method="post" action="admin_saveconfig.asp" name="configform" onsubmit="return check();">

			<div id="tabContainer"></div>

			<div id="tab-switch">
				<h4>����</h4>
				<div class="field">
					<span class="label">���Ա�״̬</span>
					<span class="value"><input type="radio" name="status1" value="1" id="status11"<%=cked(CBool(tstatus AND 1))%> /><label for="status11">����</label>����<input type="radio" name="status1" value="0" id="status12"<%=cked(Not CBool(tstatus AND 1))%> /><label for="status12">�ر�</label> (�ر�ʱ"�ÿ�����Ȩ��"������Ч)</span>
				</div>
				<div class="field">
					<span class="label">�ÿ�����Ȩ��</span>
					<span class="value"><input type="radio" name="status2" value="1" id="status21"<%=cked(CBool(tstatus AND 2))%> /><label for="status21">����</label>����<input type="radio" name="status2" value="0" id="status22"<%=cked(Not CBool(tstatus AND 2))%> /><label for="status22">�ر�</label></span>
				</div>
				<div class="field">
					<span class="label">�ÿ���������Ȩ��</span>
					<span class="value"><input type="radio" name="status3" value="1" id="status31"<%=cked(CBool(tstatus AND 4))%> /><label for="status31">����</label>����<input type="radio" name="status3" value="0" id="status32"<%=cked(Not CBool(tstatus AND 4))%> /><label for="status32">�ر�</label></span>
				</div>
				<div class="field">
					<span class="label">�ÿ�ͷ����</span>
					<span class="value"><input type="radio" name="status4" value="1" id="status41"<%=cked(CBool(tstatus AND 8))%> /><label for="status41">����</label>����<input type="radio" name="status4" value="0" id="status42"<%=cked(Not CBool(tstatus AND 8))%> /><label for="status42">�ر�</label></span>
				</div>
				<div class="field">
					<span class="label">������ʾǰ�����</span>
					<span class="value"><input type="radio" name="status5" value="1" id="status51"<%=cked(CBool(tstatus AND 16))%> /><label for="status51">���</label>����<input type="radio" name="status5" value="0" id="status52"<%=cked(Not CBool(tstatus AND 16))%> /><label for="status52">�����</label></span>
				</div>
				<div class="field">
					<span class="label">���Ļ�����</span>
					<span class="value"><input type="radio" name="status6" value="1" id="status61"<%=cked(CBool(tstatus AND 32))%> /><label for="status61">����</label>����<input type="radio" name="status6" value="0" id="status62"<%=cked(Not CBool(tstatus AND 32))%> /><label for="status62">�ر�</label></span>
				</div>
				<div class="field">
					<span class="label">�ÿ�Ϊ���Ļ�����</span>
					<span class="value"><input type="radio" name="status7" value="1" id="status71"<%=cked(CBool(tstatus AND 64))%> /><label for="status71">����</label>����<input type="radio" name="status7" value="0" id="status72"<%=cked(Not CBool(tstatus AND 64))%> /><label for="status72">��ֹ</label></span>
				</div>
				<div class="field">
					<span class="label">����ÿͻظ�</span>
					<span class="value"><input type="radio" name="status8" value="1" id="status81"<%=cked(CBool(tstatus AND 128))%> /><label for="status81">����</label>����<input type="radio" name="status8" value="0" id="status82"<%=cked(Not CBool(tstatus AND 128))%> /><label for="status82">��ֹ</label></span>
				</div>
				<div class="field">
					<span class="label">����ͳ��</span>
					<span class="value"><input type="radio" name="status9" value="1" id="status91"<%=cked(CBool(tstatus AND 256))%> /><label for="status91">����</label>����<input type="radio" name="status9" value="0" id="status92"<%=cked(Not CBool(tstatus AND 256))%> /><label for="status92">�ر�</label></span>
				</div>
			</div>

			<div id="tab-navigation">
				<h4>����</h4>
				<div class="field">
					<span class="label">������վLogo��ַ</span>
					<span class="value"><input type="text" class="longtext" maxlength="127" name="homelogo" value="<%=rs("homelogo")%>" /></span>
				</div>
				<div class="field">
					<span class="label">Logo��ʾģʽ</span>
					<span class="value"><input type="radio" name="logobannermode" value="1" id="logobannermode1"<%=cked(CBool(tvisualflag AND 2048))%> /><label for="logobannermode1" >���ģʽ</label>����<input type="radio" name="logobannermode" value="0" id="logobannermode2"<%=cked(Not CBool(tvisualflag AND 2048))%> /><label for="logobannermode2" >ͼƬģʽ</label></span>
				</div>
				<div class="field">
					<span class="label">������վ����</span>
					<span class="value"><input type="text" class="longtext" maxlength="30" name="homename" value="<%=rs("homename")%>" /></span>
				</div>
				<div class="field">
					<span class="label">������վ��ַ</span>
					<span class="value"><input type="text" class="longtext" maxlength="127" name="homeaddr" value="<%=rs("homeaddr")%>" /></span>
				</div>
			</div>

			<div id="tab-code">
				<h4>����</h4>
				<div class="field">
					<span class="label">����Ա��ȫ������</span>
					<span class="value"><input type="checkbox" value="1" name="adminviewcode" id="adminviewcode"<%=cked(CBool(adminlimit AND 8))%> /><label for="adminviewcode">�ÿ�������ʾʵ��HTML��UBB����</label></span>
				</div>
				<div class="field">
					<span class="label">����ԱĬ��HTMLȨ��</span>
					<span class="value">
						<span class="row"><input type="checkbox" value="1" name="adminhtml" id="adminhtml"<%=cked(CBool(adminlimit AND 1))%> /><label for="adminhtml">����Ա�ظ�������Ĭ��֧��HTML</label></span>
						<span class="row"><input type="checkbox" value="1" name="adminubb" id="adminubb"<%=cked(CBool(adminlimit AND 2))%> /><label for="adminubb">����Ա�ظ�������Ĭ��֧��UBB</label></span>
						<span class="row"><input type="checkbox" value="1" name="adminertn" id="adminertn"<%=cked(CBool(adminlimit AND 4))%> /><label for="adminertn">����Ա��֧��HTML��UBBʱ��Ĭ������س�����</label></span>
					</span>
				</div>
				<div class="field">
					<span class="label">�ÿ�HTMLȨ��</span>
					<span class="value">
						<span class="row"><input type="checkbox" value="1" name="guesthtml" id="guesthtml"<%=cked(CBool(guestlimit AND 1))%> /><label for="guesthtml">�ÿ�����֧��HTML</label></span>
						<span class="row"><input type="checkbox" value="1" name="guestubb" id="guestubb"<%=cked(CBool(guestlimit AND 2))%> /><label for="guestubb">�ÿ�����֧��UBB</label></span>
						<span class="row"><input type="checkbox" value="1" name="guestertn" id="guestertn"<%=cked(CBool(guestlimit AND 4))%> /><label for="guestertn">�ÿͲ�֧��HTML��UBBʱ������س�����</label></span>
					</span>
				</div>
				<div class="field">
					<span class="label">UBB������</span>
					<span class="value"><input type="radio" name="showubbtool" value="1" id="showubbtool1"<%=cked(CBool(tvisualflag AND 2))%> /><label for="showubbtool1">��ʾ</label>����<input type="radio" name="showubbtool" value="0" id="showubbtool2"<%=cked(Not CBool(tvisualflag AND 2))%> /><label for="showubbtool2">����</label></span>
				</div>
				<div class="field">
					<span class="label">UBB����(������UBB)</span>
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
			</div>

			<div id="tab-security">
				<h4>��ȫ</h4>
				<div class="field">
					<span class="label">����Ա��¼��ʱ</span>
					<span class="value"><input type="text" size="4" maxlength="4" name="admintimeout" value="<%=rs("admintimeout")%>" />�� (Ĭ��=20)</span>
				</div>
				<div class="field">
					<span class="label">Ϊ�ÿ���ʾIPv4ǰ</span>
					<span class="value"><input type="text" size="4" maxlength="1" name="showipv4" value="<%=ShowIPv4%>" />�ֽ� (��ѡֵ��0��4)</span>
				</div>
				<div class="field">
					<span class="label">Ϊ�ÿ���ʾIPv6ǰ</span>
					<span class="value"><input type="text" size="4" maxlength="1" name="showipv6" value="<%=ShowIPv6%>" />�� (��ѡֵ��0��8)</span>
				</div>
				<div class="field">
					<span class="label">Ϊ����Ա��ʾIPv4ǰ</span>
					<span class="value"><input type="text" size="4" maxlength="1" name="adminshowipv4" value="<%=AdminShowIPv4%>" />�ֽ� (��ѡֵ��0��4)</span>
				</div>
				<div class="field">
					<span class="label">Ϊ����Ա��ʾIPv6ǰ</span>
					<span class="value"><input type="text" size="4" maxlength="1" name="adminshowipv6" value="<%=AdminShowIPv6%>" />�� (��ѡֵ��0��8)</span>
				</div>
				<div class="field">
					<span class="label">Ϊ����Ա��ʾԭIPv4ǰ</span>
					<span class="value"><input type="text" size="4" maxlength="1" name="adminshoworiginalipv4" value="<%=AdminShowOriginalIPv4%>" />�ֽ� (��ѡֵ��0��4,ʹ�ô��������ʱ������ʾԭʼIP)</span>
				</div>
				<div class="field">
					<span class="label">Ϊ����Ա��ʾԭIPv6ǰ</span>
					<span class="value"><input type="text" size="4" maxlength="1" name="adminshoworiginalipv6" value="<%=AdminShowOriginalIPv6%>" />�� (��ѡֵ��0��8,ʹ�ô��������ʱ������ʾԭʼIP)</span>
				</div>
				<div class="field">
					<span class="label">��¼��֤�볤��</span>
					<span class="value"><input type="text" size="4" maxlength="2" name="vcodecount" value="<%=rs("vcodecount") and &H0F%>" />λ (��ѡֵ��0��10)</span>
				</div>
				<div class="field">
					<span class="label">������֤�볤��</span>
					<span class="value"><input type="text" size="4" maxlength="2" name="writevcodecount" value="<%=(rs("vcodecount") and &HF0) \ &H10%>" />λ (��ѡֵ��0��10)</span>
				</div>
			</div>

			<div id="tab-ui">
				<h4>����</h4>
				<div class="field">
					<span class="label">�����б�(","�ָ�)</span>
					<span class="value"><input type="text" class="longtext" maxlength="48" name="cssfontfamily" value="<%=rs("cssfontfamily")%>" /></span>
				</div>
				<div class="field">
					<span class="label">�����С</span>
					<span class="value"><input type="text" size="10" maxlength="8" name="cssfontsize" value="<%=rs("cssfontsize")%>" /></span>
				</div>
				<div class="field">
					<span class="label">�����м��</span>
					<span class="value"><input type="text" size="10" maxlength="8" name="csslineheight" value="<%=rs("csslineheight")%>" /></span>
				</div>
				<div class="field">
					<span class="label">���Ա������</span>
					<span class="value"><input type="text" size="10" maxlength="5" name="tablewidth" value="<%=rs("tablewidth")%>" /> (Ĭ��=630,���ðٷֱ�)</span>
				</div>
				<div class="field">
					<span class="label">����������</span>
					<span class="value"><input type="text" size="10" maxlength="3" name="windowspace" value="<%=rs("windowspace")%>" /> (Ĭ��=20,��λ:����)</span>
				</div>
				<div class="field">
					<span class="label">���Ա��󴰸���</span>
					<span class="value"><input type="text" size="10" maxlength="5" name="tableleftwidth" value="<%=rs("tableleftwidth")%>" /> (Ĭ��=150,���ðٷֱ�)</span>
				</div>
				<div class="field">
					<span class="label">���������ݡ��ı��߶�</span>
					<span class="value"><input type="text" size="10" maxlength="3" name="leavecontentheight" value="<%=rs("leavecontentheight")%>" /> (Ĭ��=7,��λ:��ĸ�߶�)</span>
				</div>
				<div class="field">
					<span class="label">��������</span>
					<span class="value"><input type="text" size="10" maxlength="3" name="searchtextwidth" value="<%=rs("searchtextwidth")%>" /> (Ĭ��=20,��λ:��ĸ���)</span>
				</div>
				<div class="field">
					<span class="label">�ظ�������༭��߶�</span>
					<span class="value"><input type="text" size="10" maxlength="3" name="replytextheight" value="<%=rs("replytextheight")%>" /> (Ĭ��=10,��λ:��ĸ�߶�)</span>
				</div>
				<div class="field">
					<span class="label">ÿҳ��ʾ��������</span>
					<span class="value"><input type="text" size="10" maxlength="5" name="itemsperpage" value="<%=rs("itemsperpage")%>" /> (Ĭ��=5)</span>
				</div>
				<div class="field">
					<span class="label">ÿҳ��ʾ�ı�����</span>
					<span class="value"><input type="text" size="10" maxlength="5" name="titlesperpage" value="<%=rs("titlesperpage")%>" /> (Ĭ��=20)</span>
				</div>
				<div class="field">
					<span class="label">ͷ��ÿ����ʾ����Ŀ</span>
					<span class="value"><input type="text" size="10" maxlength="3" name="picturesperrow" value="<%=rs("picturesperrow")%>" /> (Ĭ��=5)</span>
				</div>
				<div class="field">
					<span class="label">���������ͷ����</span>
					<span class="value"><input type="text" size="10" maxlength="3" name="frequentfacecount" value="<%=rs("frequentfacecount")%>" /> (Ĭ��=15,����"����ͷ��"��ʾȫ��)</span>
				</div>
				<div class="field">
					<span class="label">���Ա���ɫ����</span>
					<span class="value">
						<select name="style">
						<%
						styleid=rs("styleid")

						set rs1=server.CreateObject("ADODB.Recordset")
						rs1.Open sql_adminconfig_style,cn,,,1

						dim onestyleid,onestylename
						while rs1.EOF=false
							onestyleid=rs1("styleid")
							onestylename=rs1("stylename")
							Response.Write "<option value=""" & onestyleid & """"
							Response.Write seled(onestyleid=styleid or onestyleid="")
							Response.Write ">" &onestylename& "</option>"
							rs1.MoveNext
						wend
						rs1.Close : set rs1=nothing
						%>
						</select>
					</span>
				</div>
			</div>

			<div id="tab-behavior">
				<h4>��Ϊ</h4>
				<div class="field">
					<span class="label">������ʱ��ƫ��</span>
					<span class="value"><input type="text" size="10" maxlength="5" name="servertimezoneoffset" value="<%=rs("servertimezoneoffset")%>" /> (Ĭ��=0,��λ:����)</span>
				</div>
				<div class="field">
					<span class="label">��ʾʱ��ƫ��</span>
					<span class="value"><input type="text" size="10" maxlength="5" name="displaytimezoneoffset" value="<%=rs("displaytimezoneoffset")%>" /> (Ĭ��=0,��λ:����)</span>
				</div>
				<div class="field">
					<span class="label">Ĭ�ϰ���ģʽ</span>
					<span class="value"><input type="radio" name="displaymode" value="1" id="displaymode1"<%=cked(CBool(tvisualflag AND 1024))%> /><label for="displaymode1">����ģʽ</label>����<input type="radio" name="displaymode" value="0" id="displaymode2"<%=cked(Not CBool(tvisualflag AND 1024))%> /><label for="displaymode2">����ģʽ</label></span>
				</div>
				<div class="field">
					<span class="label">�ظ�������ʾλ��</span>
					<span class="value"><input type="radio" name="replyinword" value="1" id="replyinword1"<%=cked(CBool(tvisualflag AND 1))%> /><label for="replyinword1">��Ƕ�ڷÿ�����</label>����<input type="radio" name="replyinword" value="0" id="replyinword2"<%=cked(Not CBool(tvisualflag AND 1))%> /><label for="replyinword2">��ʾ�ڷÿ������·�</label></span>
				</div>
				<div class="field">
					<span class="label">�������ݱ����ص�����</span>
					<span class="value"><input type="radio" name="hidehidden" value="1" id="hidehidden1"<%=cked(CBool(tvisualflag AND 128))%> /><label for="hidehidden1" >����</label>����<input type="radio" name="hidehidden" value="0" id="hidehidden2"<%=cked(Not CBool(tvisualflag AND 128))%> /><label for="hidehidden2" >��ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">���ش��������</span>
					<span class="value"><input type="radio" name="hideaudit" value="1" id="hideaudit1"<%=cked(CBool(tvisualflag AND 256))%> /><label for="hideaudit1" >����</label>����<input type="radio" name="hideaudit" value="0" id="hideaudit2"<%=cked(Not CBool(tvisualflag AND 256))%> /><label for="hideaudit2" >��ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">���ذ���δ�ظ����Ļ�</span>
					<span class="value"><input type="radio" name="hidewhisper" value="1" id="hidewhisper1"<%=cked(CBool(tvisualflag AND 512))%> /><label for="hidewhisper1" >����</label>����<input type="radio" name="hidewhisper" value="0" id="hidewhisper2"<%=cked(Not CBool(tvisualflag AND 512))%> /><label for="hidewhisper2" >��ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">��ҳ������ʾλ��</span>
					<span class="value"><input type="radio" name="showpagelist" value="3" id="showpagelist3"<%=cked((tvisualflag and 12)=12)%> /><label for="showpagelist3">���·�</label>��<input type="radio" name="showpagelist" value="1" id="showpagelist1"<%=cked((tvisualflag and 12)=4)%> /><label for="showpagelist1">�Ϸ�</label>����<input type="radio" name="showpagelist" value="2" id="showpagelist2"<%=cked((tvisualflag and 12)=8)%> /><label for="showpagelist2">�·�</label></span>
				</div>
				<div class="field">
					<span class="label">�ÿ�����������ʾλ��</span>
					<span class="value"><input type="radio" name="showsearchbox" value="3" id="showsearchbox3"<%=cked((tvisualflag and 48)=48)%> /><label for="showsearchbox3">���·�</label>��<input type="radio" name="showsearchbox" value="1" id="showsearchbox1"<%=cked((tvisualflag and 48)=16)%> /><label for="showsearchbox1">�Ϸ�</label>����<input type="radio" name="showsearchbox" value="2" id="showsearchbox2"<%=cked((tvisualflag and 48)=32)%> /><label for="showsearchbox2">�·�</label></span>
				</div>
				<div class="field">
					<span class="label">��ҳ�б�ģʽ</span>
					<span class="value"><input type="radio" name="advpagelist" value="1" id="advpagelist1"<%=cked(CBool(tvisualflag AND 64))%> /><label for="advpagelist1">����ʽ</label>��<input type="radio" name="advpagelist" value="0" id="advpagelist2"<%=cked(Not CBool(tvisualflag AND 64))%> /><label for="advpagelist2">ƽ��ʽ</label></span>
				</div>
				<div class="field">
					<span class="label">����ʽ��ҳ����</span>
					<span class="value"><input type="text" size="5" maxlength="3" name="advpagelistcount" value="<%=rs("advpagelistcount")%>" /> (Ĭ��=10)</span>
				</div>
				<div class="field">
					<span class="label">���Ա�ҳͷ</span>
					<span class="value"><input type="radio" name="showtitle" value="1" id="showtitle1"<%=cked(CBool(tpagecontrol AND 2))%> /><label for="showtitle1">��ʾ</label>����<input type="radio" name="showtitle" value="0" id="showtitle2"<%=cked(Not CBool(tpagecontrol AND 2))%> /><label for="showtitle2">����</label></span>
				</div>
				<div class="field">
					<span class="label">��ҳ�Ҽ��˵�</span>
					<span class="value"><input type="radio" name="showcontext" value="1" id="showcontext1"<%=cked(CBool(tpagecontrol AND 4))%> /><label for="showcontext1">����</label>����<input type="radio" name="showcontext" value="0" id="showcontext2"<%=cked(Not CBool(tpagecontrol AND 4))%> /><label for="showcontext2">����</label></span>
				</div>
				<div class="field">
					<span class="label">ѡ����ҳ����</span>
					<span class="value"><input type="radio" name="selectcontent" value="1" id="selectcontent1"<%=cked(CBool(tpagecontrol AND 8))%> /><label for="selectcontent1">����</label>����<input type="radio" name="selectcontent" value="0" id="selectcontent2"<%=cked(Not CBool(tpagecontrol AND 8))%> /><label for="selectcontent2">��ֹ</label></span>
				</div>
				<div class="field">
					<span class="label">������ҳ(Ctrl-C)</span>
					<span class="value"><input type="radio" name="copycontent" value="1" id="copycontent1"<%=cked(CBool(tpagecontrol AND 16))%> /><label for="copycontent1">����</label>����<input type="radio" name="copycontent" value="0" id="copycontent2"<%=cked(Not CBool(tpagecontrol AND 16))%> /><label for="copycontent2">��ֹ</label></span>
				</div>
				<div class="field">
					<span class="label">��IFrame��Ƕ</span>
					<span class="value"><input type="radio" name="beframed" value="1" id="beframed1"<%=cked(CBool(tpagecontrol AND 32))%> /><label for="beframed1">����</label>����<input type="radio" name="beframed" value="0" id="beframed2"<%=cked(Not CBool(tpagecontrol AND 32))%> /><label for="beframed2">��ֹ</label></span>
				</div>
				<div class="field">
					<span class="label">������������</span>
					<span class="value"><input type="text" size="10" maxlength="10" name="wordslimit" value="<%=rs("wordslimit")%>" /> (0=����,�س�����ռ2��,��UBB����ǰͳ��)</span>
				</div>
			</div>

			<div id="tab-mail">
				<h4>�ʼ�</h4>
				<div class="field">
					<span class="label">�����Ե���֪ͨ����</span>
					<span class="value"><input type="checkbox" value="1" name="mailnewinform" id="mailnewinform"<%=cked(CBool(tmailflag AND 1))%> /><label for="mailnewinform">����</label></span>
				</div>
				<div class="field">
					<span class="label">�����ظ�֪ͨ������</span>
					<span class="value"><input type="checkbox" value="1" name="mailreplyinform" id="mailreplyinform"<%=cked(CBool(tmailflag AND 2))%> /><label for="mailreplyinform">����</label></span>
				</div>
				<div class="field">
					<span class="label">�ʼ��������</span>
					<span class="value"><input type="radio" value="0" name="mailcomponent" id="mailcomponent0"<%=cked(Not CBool(tmailflag AND 4))%> /><label for="mailcomponent0">JMail</label>��<input type="radio" value="1" name="mailcomponent" id="mailcomponent1"<%=cked(CBool(tmailflag AND 4))%> /><label for="mailcomponent1">CDO</label></span>
				</div>
				<div class="field">
					<span class="label">������֪ͨ���յ�ַ</span>
					<span class="value"><input type="text" class="longtext" maxlength="48" name="mailreceive" value="<%=rs("mailreceive")%>" /></span>
				</div>
				<div class="field">
					<span class="label">�����˵�ַ</span>
					<span class="value"><input type="text" class="longtext" maxlength="48" name="mailfrom" value="<%=rs("mailfrom")%>" /></span>
				</div>
				<div class="field">
					<span class="label">������SMTP��������ַ</span>
					<span class="value"><input type="text" class="longtext" maxlength="48" name="mailsmtpserver" value="<%=rs("mailsmtpserver")%>" /></span>
				</div>
				<div class="field">
					<span class="label">��¼�û���(����Ҫ)</span>
					<span class="value"><input type="text" class="longtext" maxlength="48" name="mailuserid" value="<%=rs("mailuserid")%>" /></span>
				</div>
				<div class="field">
					<span class="label">��¼����(����Ҫ)</span>
					<span class="value"><input type="password" class="longtext" maxlength="48" name="mailuserpass" value="<%=rs("mailuserpass")%>" /></span>
				</div>
				<div class="field">
					<span class="label">�ʼ������̶�</span>
					<span class="value"><input type="text" size="4" maxlength="1" name="maillevel" value="<%=rs("maillevel")%>" /> (1��졫5����)</span>
				</div>
			</div>

			<div id="tab-confirm">
				<h4>ȷ��</h4>
				<div class="field">
					<span class="label">ͨ���������</span>
					<span class="value"><input type="radio" name="passaudittip" value="1" id="passaudittip1"<%=cked(CBool(tdelconfirm AND 16))%> /><label for="passaudittip1">��ʾȷ��</label>����<input type="radio" name="passaudittip" value="0" id="passaudittip2"<%=cked(Not CBool(tdelconfirm AND 16))%> /><label for="passaudittip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">ͨ�����ѡ������</span>
					<span class="value"><input type="radio" name="passseltip" value="1" id="passseltip1"<%=cked(CBool(tdelconfirm AND 32))%> /><label for="passseltip1">��ʾȷ��</label>����<input type="radio" name="passseltip" value="0" id="passseltip2"<%=cked(Not CBool(tdelconfirm AND 32))%> /><label for="passseltip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">ɾ������</span>
					<span class="value"><input type="radio" name="deltip" value="1" id="deltip1"<%=cked(CBool(tdelconfirm AND 1))%> /><label for="deltip1">��ʾȷ��</label>����<input type="radio" name="deltip" value="0" id="deltip2"<%=cked(Not CBool(tdelconfirm AND 1))%> /><label for="deltip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">ɾ���ظ�</span>
					<span class="value"><input type="radio" name="delretip" value="1" id="delretip1"<%=cked(CBool(tdelconfirm AND 2))%> /><label for="delretip1">��ʾȷ��</label>����<input type="radio" name="delretip" value="0" id="delretip2"<%=cked(Not CBool(tdelconfirm AND 2))%> /><label for="delretip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">ɾ��ѡ������</span>
					<span class="value"><input type="radio" name="delseltip" value="1" id="delseltip1"<%=cked(CBool(tdelconfirm AND 4))%> /><label for="delseltip1">��ʾȷ��</label>����<input type="radio" name="delseltip" value="0" id="delseltip2"<%=cked(Not CBool(tdelconfirm AND 4))%> /><label for="delseltip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">ִ�и߼�ɾ��</span>
					<span class="value"><input type="radio" name="deladvtip" value="1" id="deladvtip1"<%=cked(CBool(tdelconfirm AND 8))%> /><label for="deladvtip1">��ʾȷ��</label>����<input type="radio" name="deladvtip" value="0" id="deladvtip2"<%=cked(Not CBool(tdelconfirm AND 8))%> /><label for="deladvtip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">�������Ļ�</span>
					<span class="value"><input type="radio" name="pubwhispertip" value="1" id="pubwhispertip1"<%=cked(CBool(tdelconfirm AND 64))%> /><label for="pubwhispertip1">��ʾȷ��</label>����<input type="radio" name="pubwhispertip" value="0" id="pubwhispertip2"<%=cked(Not CBool(tdelconfirm AND 64))%> /><label for="pubwhispertip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">�ö�����</span>
					<span class="value"><input type="radio" name="lock2toptip" value="1" id="lock2toptip1"<%=cked(CBool(tdelconfirm AND 256))%> /><label for="lock2toptip1">��ʾȷ��</label>����<input type="radio" name="lock2toptip" value="0" id="lock2toptip2"<%=cked(Not CBool(tdelconfirm AND 256))%> /><label for="lock2toptip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">��ǰ����</span>
					<span class="value"><input type="radio" name="bring2toptip" value="1" id="bring2toptip1"<%=cked(CBool(tdelconfirm AND 128))%> /><label for="bring2toptip1">��ʾȷ��</label>����<input type="radio" name="bring2toptip" value="0" id="bring2toptip2"<%=cked(Not CBool(tdelconfirm AND 128))%> /><label for="bring2toptip2">����ʾ</label></span>
				</div>
				<div class="field">
					<span class="label">��������˳��</span>
					<span class="value"><input type="radio" name="reordertip" value="1" id="reordertip1"<%=cked(CBool(tdelconfirm AND 512))%> /><label for="reordertip1">��ʾȷ��</label>����<input type="radio" name="reordertip" value="0" id="reordertip2"<%=cked(Not CBool(tdelconfirm AND 512))%> /><label for="reordertip2">����ʾ</label></span>
				</div>
			</div>

			<input type="hidden" name="tabIndex" id="tabIndex" value="<%=Request.QueryString("tabIndex")%>" />
			<div class="command"><input value="��������" type="submit" name="submit1" /></div>
			</form>
		</div>
	</div>
	</div>

	<!-- #include file="include/template/footer.inc" -->
</div>
<%rs.Close : cn.Close : set rs=nothing : set cn=nothing%>

<script type="text/javascript" src="asset/js/tabcontrol.js"></script>
<script type="text/javascript">
var tab=new TabControl('tabContainer');

tab.addPage('tab-switch','����');
tab.addPage('tab-navigation','����');
tab.addPage('tab-code','����');
tab.addPage('tab-security','��ȫ');
tab.addPage('tab-ui','����');
tab.addPage('tab-behavior','��Ϊ');
tab.addPage('tab-mail','�ʼ�');
tab.addPage('tab-confirm','ȷ��');

tab.restoreFromField('tabIndex');

function check()
{
	function checkRange(tabIndex, name, textbox, min, max) {
		var value;
		if(isNaN(value = parseInt(textbox.value, 10))) {
			tab.selectPage(tabIndex);
			textbox.select();
			alert(name + ' ����Ϊ���֡�');
			return false;
		} else if(value<min || value>max) {
			tab.selectPage(tabIndex);
			textbox.select();
			alert(name + ' ������' + min + '��' + max + '�ķ�Χ�ڡ�');
			return false;
		}

		return true;
	}
	function checkCssSize(tabIndex, name, textbox) {
		var value;
		if(isNaN(value = parseInt(textbox.value, 10))) {
			if (!/^\s*\d+%\s*$/.test(value)) {
				tab.selectPage(tabIndex);
				textbox.select();
				alert(name + ' ����Ϊ���������ٷֱȡ�');
				return false;
			}
		} else if (value <= 0) {
			tab.selectPage(tabIndex);
			textbox.select();
			alert(name + ' ��������㡣');
			return false;
		}

		return true;
	}

	var frm=document.configform;
	return checkRange(3, "����Ա��¼��ʱ", frm.admintimeout, 1, 1440) &&
		checkRange(3, "Ϊ�ÿ���ʾIPv4", frm.showipv4, 0, 4) &&
		checkRange(3, "Ϊ�ÿ���ʾIPv6", frm.showipv6, 0, 8) &&
		checkRange(3, "Ϊ����Ա��ʾIPv4", frm.adminshowipv4, 0, 4) &&
		checkRange(3, "Ϊ����Ա��ʾIPv6", frm.adminshowipv6, 0, 8) &&
		checkRange(3, "Ϊ����Ա��ʾԭIPv4", frm.adminshoworiginalipv4, 0, 4) &&
		checkRange(3, "Ϊ����Ա��ʾԭIPv6", frm.adminshoworiginalipv6, 0, 8) &&
		checkRange(3, "��¼��֤�볤��", frm.vcodecount, 0, 10) &&
		checkRange(3, "������֤�볤��", frm.writevcodecount, 0, 10) &&
		checkCssSize(4, "���Ա������", frm.tablewidth) &&
		checkRange(4, "����������", frm.windowspace, 1, 255) &&
		checkCssSize(4, "���Ա��󴰸���", frm.tableleftwidth) &&
		checkRange(4, "���������ݡ��ı��߶�", frm.leavecontentheight, 1, 255) &&
		checkRange(4, "��������", frm.searchtextwidth, 1, 255) &&
		checkRange(4, "�ظ�������༭��߶�", frm.replytextheight, 1, 255) &&
		checkRange(4, "ÿҳ��ʾ��������", frm.itemsperpage, 1, 32767) &&
		checkRange(4, "ÿҳ��ʾ�ı�����", frm.titlesperpage, 1, 32767) &&
		checkRange(4, "ͷ��ÿ����ʾ����Ŀ", frm.picturesperrow, 1, 255) &&
		checkRange(4, "���������ͷ����", frm.frequentfacecount, 0, 255) &&
		checkRange(5, "������ʱ��ƫ��", frm.servertimezoneoffset, -1440, 1440) &&
		checkRange(5, "��ʾʱ��ƫ��", frm.displaytimezoneoffset, -1440, 1440) &&
		checkRange(5, "����ʽ��ҳ����", frm.advpagelistcount, 1, 255) &&
		checkRange(5, "������������", frm.wordslimit, 0, 2147483647) &&
		checkRange(6, "�ʼ������̶�", frm.maillevel, 1, 5) &&
		(document.configform.submit1.disabled=true, true);
}
</script>
</body>
</html>