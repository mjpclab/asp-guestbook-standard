<!-- #include file="include/template/page_instruction.inc" -->
<!-- #include file="config/system.asp" -->
<!-- #include file="config/database.asp" -->
<!-- #include file="include/sql/init.asp" -->
<!-- #include file="include/sql/admin_verify.asp" -->
<!-- #include file="include/utility/database.asp" -->
<!-- #include file="include/utility/backend.asp" -->
<!-- #include file="include/utility/frontend.asp" -->
<!-- #include file="include/utility/book.asp" -->
<!-- #include file="loadconfig.asp" -->
<!-- #include file="admin_verify.asp" -->
<%
Response.Expires=-1
%>

<!-- #include file="include/template/dtd.inc" -->
<html>
<head>
	<!-- #include file="include/template/metatag.inc" -->
	<title><%=HomeName%> ���Ա� ���ɵ��ô���</title>
	<!-- #include file="inc_admin_stylesheet.asp" -->

	<script type="text/javascript">
	function generateCallCode()
	{
		var frm_n=document.getElementById('frm_n');
		var frm_len=document.getElementById('frm_len');
		var frm_code=document.getElementById('frm_code');

		var mode=document.getElementById('frm_mode').value;
		var baseurl=document.getElementById('frm_baseurl').value;
		var n=frm_n.value;
		var len=frm_len.value;
		var prefix=document.getElementById('frm_prefix').value;
		var target=document.getElementById('frm_target').value;

		if(n.length===0)
		{
			alert('��������ʾ������');
			frm_n.focus();
			return false;
		}
		else if(isNaN(n))
		{
			alert('��ʾ�����������顣');
			frm_n.select();
			return false;
		}
		else if(isNaN(len))
		{
			alert('���������������顣');
			frm_len.select();
			return false;
		}

		if(baseurl.substr(-1)!=='/') {
			baseurl+='/';
		}

		var url=baseurl + 'tlist.asp?mode=' + mode + '&n=' + n;
		if(len) {
			url+='&len='+len;
		}
		if(prefix) {
			url+='&prefix='+encodeURIComponent(prefix)
		}
		if(target) {
			url+='&target='+target;
		}

		if(mode==='iframe') {
			frm_code.value='<iframe src="' +url+ '" frameborder="0"><\/iframe>';
		}
		else if(mode==='js') {
			frm_code.value='<' + 'script type="text/javascript" src="' +url+ '" charset="gbk"><\/script>';
		}
		else if(mode==='json') {
			frm_code.value='����GET��ʽ��������URL��\n' + url;
		}
	}
	</script>
</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<div id="outerborder" class="outerborder">

<%if ShowTitle then%><%Call InitHeaderData("����")%><!-- #include file="include/template/header.inc" --><%end if%>
<div id="mainborder" class="mainborder">
<!-- #include file="include/template/admin_mainmenu.inc" -->
<div class="region region-longtext region-callgen">
	<h3 class="title">���ɵ��ô���</h3>
	<div class="content">
		<p>��ҳ�����������Ա����Ա�����ô��롣��������ô�������Ĳ��������С���ʾ������Ϊ�����</p>

		<form>
		<div class="field field-mode">
			<span class="label">ʹ��ģʽ��</span>
			<span class="value">
				<select name="frm_mode" id="frm_mode">
					<option value="iframe">iframe</option>
					<option value="js">JS���</option>
					<option value="json">JSON����</option>
				</select>
			</span>
		</div>
		<div class="field field-base-url">
			<span class="label">���Ա���URL��</span>
			<span class="value"><input type="text" name="frm_baseurl" id="frm_baseurl" value="<%=geturlpath%>" /></span>
		</div>
		<div class="field field-n">
			<span class="label">��ʾ������<span class="required">*</span></span>
			<span class="value"><input type="text" name="frm_n" id="frm_n" maxlength="10" value="10"/></span>
		</div>
		<div class="field field-len">
			<span class="label">�����������ƣ�</span>
			<span class="value"><input type="text" name="frm_len" id="frm_len"/></span>
		</div>
		<div class="field field-prefix">
			<span class="label">����ǰ׺��</span>
			<span class="value"><input type="text" name="frm_prefix" id="frm_prefix"/></span>
		</div>
		<div class="field field-target">
			<span class="label">�򿪴��ڣ�</span>
			<span class="value">
				<select name="frm_target" id="frm_target">
					<option value="">(Ĭ��)</option>
					<option value="_blank">����ҳ��(_blank)</option>
					<option value="_self">��ͬ���ڻ��ܴ���(_self)</option>
					<option value="_top">�������������(_top)</option>
					<option value="_parent">����ܴ���(_parent)</option>
				</select>
			</span>
		</div>
		<div class="field-command">
			<input type="button" name="btn_generate" id="btn_generate" value="���ɵ��ô���" onclick="generateCallCode();" />
		</div>
		<div class="field">
			<span class="row">�����ɵĵ��ô��룺</span>
			<span class="value"><textarea readonly="readonly" name="frm_code" id="frm_code" rows="<%=ReplyTextHeight%>"></textarea></span>
		</div>
		</form>
	</div>
</div>
</div>

<!-- #include file="include/template/footer.inc" -->
</div>
</body>
</html>