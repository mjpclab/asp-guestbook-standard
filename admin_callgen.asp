<!-- #include file="config.asp" -->
<!-- #include file="admin_verify.asp" -->
<!-- #include file="common2.asp" -->

<%
Response.Expires=-1
%>

<!-- #include file="inc_dtd.asp" -->
<html>
<head>
	<!-- #include file="inc_metatag.asp" -->
	<title><%=HomeName%> ���Ա� ���ɵ��ô���</title>
	<!-- #include file="inc_admin_stylesheet.asp" -->

	<script type="text/javascript">
	function swIframe(jschk)
	{
		var span=document.getElementById('iframeSettings');
		if(jschk.checked)
			span.style.visibility='hidden';
		else
			span.style.visibility='visible';
	}
	
	function urlEncode(url)
	{
		return url.replace('%','%25').replace('?','%3F').replace(' ','%20').replace('#','%23').replace('&','%26').replace('\'','%27').replace('/','%2F').replace('<','%3C').replace('>','%3E');
	}
	
	function generateCallCode()
	{
		var frm_n=document.getElementById('frm_n');
		var frm_len=document.getElementById('frm_len');
		var frm_code=document.getElementById('frm_code');
	
		var url=document.getElementById('url').value;
		var n=document.getElementById('frm_n').value;
		var target=document.getElementById('frm_target').value;
		var prefix=document.getElementById('frm_prefix').value;
		var len=document.getElementById('frm_len').value;
		var nobr=document.getElementById('frm_nobr').checked;
		var js=document.getElementById('frm_js').checked;
		var width=document.getElementById('frm_width').value;
		var height=document.getElementById('frm_height').value;
		
		if(n==='')
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
		else
		{
			var temp;
			if(js) temp='<' + 'script type="text/javascript" src="{0}" charset="gbk"><\/script>';
			else temp='<iframe width="{w}" height="{h}" src="{0}" frameborder="0"><\/iframe>';
			
			var src=
				url +
				'?n=' + n +
				(js ? '&js=yes' : '') +
				(target!='' ? '&target='+target : '') +
				(prefix!='' ? '&prefix='+urlEncode(prefix) : '') +
				(len!='' ? '&len='+len : '') +
				(nobr ? '&nobr=yes' : '');
			
			frm_code.value=temp.replace('{w}',width).replace('{h}',height).replace('{0}',src);
			return false;
		}
	}
	</script>

</head>

<body<%=bodylimit%> onload="<%=framecheck%>">

<div id="outerborder" class="outerborder">

<%if ShowTitle=true then show_book_title 3,"����"%>
<!-- #include file="admincontrols.inc" -->

<div class="region region-callgen">
	<h3 class="title">���ɵ��ô���</h3>
	<div class="content">
		<p>��ҳ�����������Ա����Ա�����ô��롣��������ô�������Ĳ��������С���ʾ������Ϊ�����</p>

		<form method="post" action="" onsubmit="return generateCallCode();">
		<input type="hidden" name="IsPostBack" id="IsPostBack" value="1" />
		<input type="hidden" name="url" id="url" value="<%=geturlpath%>tlist.asp" />
		<div class="field">
			<span class="label">��ʾ������<span class="required">*</span></span>
			<span class="value"><input type="text" name="frm_n" id="frm_n" size="10" maxlength="10" value="10" /></span>
		</div>
		<div class="field">
			<span class="label">�򿪴��ڣ�</span>
			<span class="value">
				<select name="frm_target" id="frm_target">
					<option value="" <%=seled(Request.Form("frm_target")="")%>>(Ĭ��)</option>
					<option value="_blank" <%=seled(Request.Form("frm_target")="_blank")%>>����ҳ��</option>
					<option value="_self" <%=seled(Request.Form("frm_target")="_self")%>>��ͬ���ڻ��ܴ���</option>
					<option value="_top" <%=seled(Request.Form("frm_target")="_top")%>>�������������</option>
					<option value="_parent" <%=seled(Request.Form("frm_target")="_parent")%>>����ܴ���</option>
				</select>
			</span>
		</div>
		<div class="field">
			<span class="label">����ǰ׺��</span>
			<span class="value"><input type="text" name="frm_prefix" id="frm_prefix" size="10" value="<%=Request.Form("frm_prefix")%>" /></span>
		</div>
		<div class="field">
			<span class="label">�������ƣ�</span>
			<span class="value"><input type="text" name="frm_len" id="frm_len" size="10" value="<%=Request.Form("frm_len")%>" /></span>
		</div>
		<div class="field">
			<span class="label">��������У�</span>
			<span class="value"><input type="checkbox" name="frm_nobr" id="frm_nobr" value="1" <%=cked(Request.Form("frm_nobr")="1")%> /></span>
		</div>
		<div class="field">
			<span class="label">ʹ��JSģʽ��</span>
			<span class="value"><input type="checkbox" name="frm_js" id="frm_js" value="1" <%=cked(Request.Form("frm_js")="1")%> onclick="swIframe(this);" /></span>
		</div>
		<div id="iframeSettings">
			<div class="field">
				<span class="label">���ڿ�ȣ�</span>
				<span class="value"><input type="text" name="frm_width" id="frm_width" size="10" value="<%=Request.Form("frm_width")%>" /></span>
			</div>
			<div class="field">
				<span class="label">���ڸ߶ȣ�</span>
				<span class="value"><input type="text" name="frm_height" id="frm_height" size="10" value="<%=Request.Form("frm_height")%>" /></span>
			</div>
		</div>
		<div class="field-command">
			<input type="submit" name="submit1" id="submit1" value="���ɵ��ô���" />
		</div>
		<div class="field">
			<span class="row">�����ɵĵ��ô��룺</span>
			<span class="value"><textarea readonly="readonly" name="frm_code" id="frm_code" rows="<%=ReplyTextHeight%>"></textarea></span>
		</div>
		</form>
	</div>
</div>

</div>

<!-- #include file="bottom.asp" -->
</body>
</html>